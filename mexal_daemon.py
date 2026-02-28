import tkinter as tk
from tkinter import ttk, messagebox
import os
import time
import json
import getpass
import shutil
import re
import subprocess
import ctypes
import urllib.parse
import smtplib
import ssl
import sys
from email.message import EmailMessage
from dataclasses import dataclass
from typing import Optional
from datetime import datetime

from PyPDF2 import PdfReader, PdfWriter

try:
    import win32com.client  # type: ignore
except Exception:
    win32com = None


def _smtp_config() -> dict[str, object]:
    host = os.environ.get("SMTP_HOST", "").strip()
    port_raw = os.environ.get("SMTP_PORT", "").strip()
    user = os.environ.get("SMTP_USER", "").strip()
    password = os.environ.get("SMTP_PASS", "").strip()
    from_addr = os.environ.get("SMTP_FROM", "").strip() or user

    port = 0
    if port_raw:
        try:
            port = int(port_raw)
        except Exception:
            port = 0

    return {
        "host": host,
        "port": port,
        "user": user,
        "password": password,
        "from_addr": from_addr,
    }


def send_email_smtp(
    *,
    host: str,
    port: int,
    user: str,
    password: str,
    from_addr: str,
    to_addrs: list[str],
    subject: str,
    body: str,
    attachment_path: str,
) -> None:
    msg = EmailMessage()
    msg["From"] = from_addr
    msg["To"] = ", ".join([a for a in to_addrs if a])
    msg["Subject"] = subject
    msg.set_content(body or "")

    with open(attachment_path, "rb") as f:
        data = f.read()

    filename = os.path.basename(attachment_path)
    msg.add_attachment(data, maintype="application", subtype="pdf", filename=filename)

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(host, port, context=context) as server:
        server.login(user, password)
        server.send_message(msg)


def _powershell_escape(s: str) -> str:
    return s.replace("'", "''")


def _create_shortcut_ps(
    *,
    link_path: str,
    target_path: str,
    arguments: str,
    working_dir: str,
    description: str,
) -> None:
    link_path = os.path.abspath(link_path)
    target_path = os.path.abspath(target_path)
    working_dir = os.path.abspath(working_dir)
    ps = (
        "$WshShell = New-Object -ComObject WScript.Shell;"
        f"$Shortcut = $WshShell.CreateShortcut('{_powershell_escape(link_path)}');"
        f"$Shortcut.TargetPath = '{_powershell_escape(target_path)}';"
        f"$Shortcut.Arguments = '{_powershell_escape(arguments)}';"
        f"$Shortcut.WorkingDirectory = '{_powershell_escape(working_dir)}';"
        f"$Shortcut.Description = '{_powershell_escape(description)}';"
        "$Shortcut.Save();"
    )
    subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps],
        check=True,
        creationflags=0x08000000,
    )


def _remove_file_silent(path: str) -> None:
    try:
        if os.path.isfile(path):
            os.remove(path)
    except Exception:
        pass


def install_windows_shortcuts() -> None:
    script_path = os.path.abspath(__file__)
    workdir = os.path.dirname(script_path)

    py_exe = sys.executable
    pyw_exe = py_exe
    if py_exe.lower().endswith("python.exe"):
        cand = py_exe[:-10] + "pythonw.exe"
        if os.path.isfile(cand):
            pyw_exe = cand

    startup_dir = os.path.join(os.environ.get("APPDATA", ""), "Microsoft", "Windows", "Start Menu", "Programs", "Startup")
    desktop_dir = os.path.join(os.path.expanduser("~"), "Desktop")
    os.makedirs(startup_dir, exist_ok=True)
    os.makedirs(desktop_dir, exist_ok=True)

    args = f'"{script_path}"'
    _create_shortcut_ps(
        link_path=os.path.join(startup_dir, "Mexal Automation Daemon.lnk"),
        target_path=pyw_exe,
        arguments=args,
        working_dir=workdir,
        description="Mexal Automation Daemon",
    )
    _create_shortcut_ps(
        link_path=os.path.join(desktop_dir, "Mexal Automation Daemon.lnk"),
        target_path=py_exe,
        arguments=args,
        working_dir=workdir,
        description="Mexal Automation Daemon",
    )


def uninstall_windows_shortcuts() -> None:
    startup_dir = os.path.join(os.environ.get("APPDATA", ""), "Microsoft", "Windows", "Start Menu", "Programs", "Startup")
    desktop_dir = os.path.join(os.path.expanduser("~"), "Desktop")
    _remove_file_silent(os.path.join(startup_dir, "Mexal Automation Daemon.lnk"))
    _remove_file_silent(os.path.join(desktop_dir, "Mexal Automation Daemon.lnk"))


USER = getpass.getuser()
BASE_PATH = os.path.join("C:/Users", USER, "Desktop", "AMMINISTRAZIONE_2025")
BOLLE_DIR = os.path.join(BASE_PATH, "BOLLE_2025")
FATTURE_DIR = os.path.join(BASE_PATH, "FATTURE_2025")
PREVENTIVI_DIR = os.path.join(BASE_PATH, "PREVENTIVI_2025")
ORDINI_DIR = os.path.join(BASE_PATH, "ORDINI_2025")
ORDINI_FORNITORI_DIR = os.path.join(BASE_PATH, "ORDINI FORNITORI_2025")

PATHS = {
    # descrizioni
    "Bolla": BOLLE_DIR,
    "DDT": BOLLE_DIR,
    "Fattura": FATTURE_DIR,
    "Preventivo": PREVENTIVI_DIR,
    "Ordine cliente": ORDINI_DIR,
    "Ordine fornitore": ORDINI_FORNITORI_DIR,
    # codici
    "DDT": BOLLE_DIR,
    "FC": FATTURE_DIR,
    "PC": PREVENTIVI_DIR,
    "OC": ORDINI_DIR,
    "OF": ORDINI_FORNITORI_DIR,
}

LOG_FILE = "mexal_daemon.log"


def _log(msg: str) -> None:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def _safe_get_mtime(path: str) -> Optional[float]:
    try:
        return os.path.getmtime(path)
    except Exception:
        return None


def _safe_get_size(path: str) -> Optional[int]:
    try:
        return os.path.getsize(path)
    except Exception:
        return None

def _detect_mexal_temp_dir() -> str:
    override = os.environ.get("MEXAL_TEMP")
    if override and os.path.isdir(override):
        _log(f"MEXAL_TEMP override: {override}")
        return override

    candidates = [
        r"C:\Passepartout\PassClient\mxdesk1205143000\temp",
        r"C:\Passepartout\PassClient1\mxdesk1205143000\temp",
    ]

    def newest_pdf_mtime(base_dir: str) -> float:
        newest = 0.0
        try:
            for root_dir, _, files in os.walk(base_dir):
                for fn in files:
                    if fn.lower().endswith(".pdf"):
                        full_path = os.path.join(root_dir, fn)
                        # NOTE: this runs at import time; keep it independent of later definitions.
                        try:
                            mt = os.path.getmtime(full_path)
                        except Exception:
                            mt = None
                        if mt and mt > newest:
                            newest = mt
        except Exception:
            return 0.0
        return newest

    existing = [c for c in candidates if os.path.isdir(c)]
    if not existing:
        return candidates[0]

    existing.sort(key=newest_pdf_mtime, reverse=True)
    chosen = existing[0]
    _log(f"MEXAL_TEMP autodetect candidates={existing} chosen={chosen}")
    return chosen


MEXAL_TEMP = _detect_mexal_temp_dir()
STATE_FILE = "documenti_state.json"
SEEN_FILE = "watcher_seen.json"

_log(f"Startup. MEXAL_TEMP={MEXAL_TEMP}")


@dataclass(frozen=True)
class ParsedDoc:
    source_path: str
    created_at: float
    doc_code: str
    doc_type: str
    doc_number: str
    doc_date: str
    recipient: str


def _load_json(path: str, default):
    if not os.path.exists(path):
        return default
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def _save_json(path: str, data) -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def _safe_get_mtime(path: str) -> Optional[float]:
    try:
        return os.path.getmtime(path)
    except Exception:
        return None


def _safe_get_size(path: str) -> Optional[int]:
    try:
        return os.path.getsize(path)
    except Exception:
        return None


def extract_first_page_text(pdf_path: str) -> str:
    reader = PdfReader(pdf_path)
    if not reader.pages:
        return ""
    text = reader.pages[0].extract_text() or ""
    return text


_HEADER_RE = re.compile(
    r"^(?P<tipo>[A-Za-zÀ-ÿ]+)\b.*?\b(n\.?|nr\.?|n°)\s*(?P<num>[0-9/]+)\b.*?\bdel\s+(?P<data>\d{2}/\d{2}/\d{4})\b",
    re.IGNORECASE,
)

_DOCNUM_RE = re.compile(
    r"\b(?:n\s*[\.:°º]?|nr\s*[\.:°º]?)\s*(?P<num>[0-9/]+)\b",
    re.IGNORECASE,
)
_DOCDATE_RE = re.compile(
    r"\b(?:del|data)\b\.?\s*(?P<data>\d{2}[\./-]\d{2}[\./-]\d{4})\b",
    re.IGNORECASE,
)
_ANYDATE_RE = re.compile(r"\b(?P<data>\d{2}[\./-]\d{2}[\./-]\d{4})\b")
_NUM_AFTER_N_RE = re.compile(r"\bn\b[^0-9]*(?P<num>[0-9/]+)", re.IGNORECASE)
_SERIES_PROG_RE = re.compile(r"\b(?P<serie>\d+)\s*/\s*(?P<prog>\d+)\b")


def _doc_code_from_lines(lines: list[str]) -> tuple[str, str]:
    first_lines_raw = [ln.lower() for ln in lines[:10]]
    first_compact = [re.sub(r"[^a-z0-9]+", "", ln) for ln in first_lines_raw]
    first_norm = [
        re.sub(r"\s+", " ", re.sub(r"[^a-z0-9]+", " ", ln)).strip()
        for ln in first_lines_raw
    ]
    first5_joined_norm = " ".join(first_norm[:5]).strip()
    first5_joined_compact = re.sub(r"[^a-z0-9]+", "", first5_joined_norm)

    header_raw = " ".join(lines[:20]).lower()
    header_norm = re.sub(r"[^a-z0-9]+", " ", header_raw)
    header_norm = re.sub(r"\s+", " ", header_norm).strip()
    header_compact = re.sub(r"[^a-z0-9]+", "", header_raw)

    # Priorità assoluta: se nelle prime righe c'è "fattura" allora è una fattura.
    # Serve a evitare falsi positivi DDT quando l'estrazione testo contiene "ddt" in altre zone.
    if any("fattura" in c for c in first_compact) or any("fattura" in n for n in first_norm):
        return "FC", "Fattura"

    # Caso noto: "D.D.T. consegna" (deve comparire come etichetta in alto, non altrove)
    # Gestisce anche il caso in cui "D.D.T." e "consegna" siano su righe diverse.
    if (
        "ddtconsegna" in first5_joined_compact
        or "ddt consegna" in first5_joined_norm
        or "d d t consegna" in first5_joined_norm
        or any("ddt consegna" in n or "d d t consegna" in n for n in first_norm[:5])
    ):
        return "DDT", "DDT"

    if "preventivo" in header_compact or "preventivo" in header_norm:
        return "PC", "Preventivo"

    # Ordini: evitiamo falsi positivi (es. DDT che contiene la parola "cliente" o riferimenti ad "ordine").
    # Richiediamo la dicitura in alto (prime righe) e in forma coerente.
    if (
        any("ordine cliente" in n for n in first_norm)
        or any("ordinecliente" in c for c in first_compact)
        or "ordinecliente" in first5_joined_compact
    ):
        return "OC", "Ordine cliente"
    if (
        any("ordine fornitore" in n for n in first_norm)
        or any("ordinefornitore" in c for c in first_compact)
        or any("ordine forn" in n for n in first_norm)
        or any("ordineforn" in c for c in first_compact)
        or "ordinefornitore" in first5_joined_compact
    ):
        return "OF", "Ordine fornitore"

    # Fattura: fallback (in teoria già coperta sopra dalle prime righe)
    if ("fattura" in header_compact or "fattura" in header_norm) and "ordine" not in header_norm:
        return "FC", "Fattura"

    # DDT: spesso appare come D.D.T. nel PDF.
    # Evitiamo falsi positivi: se troviamo "fattura" nel testo compatto non deve diventare DDT.
    if "ddt" in header_compact and "fattura" not in header_compact:
        return "DDT", "DDT"
    if "documento di trasporto" in header_raw:
        return "DDT", "DDT"
    if "bolla" in header_norm:
        return "DDT", "DDT"
    return "?", "Sconosciuto"


def parse_mexal_pdf(pdf_path: str) -> Optional[ParsedDoc]:
    mtime = _safe_get_mtime(pdf_path)
    if mtime is None:
        return None

    try:
        text = extract_first_page_text(pdf_path)
    except Exception:
        return None

    raw_lines = [ln.strip() for ln in text.splitlines()]
    lines = [re.sub(r"\s+", " ", ln) for ln in raw_lines if ln.strip()]

    doc_code, doc_type = _doc_code_from_lines(lines)
    doc_number = ""
    doc_date = ""

    if doc_code == "?":
        preview = " | ".join(lines[:20])
        _log(f"Doc type UNKNOWN. file={pdf_path} preview={preview}")

    # Numero/data possono essere sulla stessa riga (es. Preventivo) oppure su righe separate (es. DDT).
    for ln in lines[:20]:
        m = _HEADER_RE.match(ln)
        if m:
            doc_number = doc_number or m.group("num")
            doc_date = doc_date or m.group("data")
        else:
            if not doc_number:
                mnum = _DOCNUM_RE.search(ln)
                if mnum:
                    doc_number = mnum.group("num")
                else:
                    # fallback per casi dove l'estrazione separa "n" dalla punteggiatura
                    mnum2 = _NUM_AFTER_N_RE.search(ln)
                    if mnum2:
                        doc_number = mnum2.group("num")
            if not doc_date:
                mdat = _DOCDATE_RE.search(ln)
                if mdat:
                    doc_date = mdat.group("data").replace(".", "/").replace("-", "/")

        if doc_number and doc_date:
            break

    # Fallback: alcuni DDT potrebbero non contenere chiaramente "del" nell'estrazione; prendiamo la prima data trovata.
    if doc_code == "DDT" and not doc_date:
        for ln in lines[:25]:
            m_any = _ANYDATE_RE.search(ln)
            if m_any:
                doc_date = m_any.group("data").replace(".", "/").replace("-", "/")
                break

    # Serie/progressivo: spesso il numero è nel formato "3/ 1234" dove 3 è la serie e 1234 è il progressivo.
    # Per OC/OF/FC lo normalizziamo come "3-1234".
    if doc_code in {"OC", "OF", "FC"}:
        for ln in lines[:20]:
            m_sp = _SERIES_PROG_RE.search(ln)
            if m_sp:
                doc_number = f"{m_sp.group('serie')}-{m_sp.group('prog')}"
                break

    recipient = ""
    dest_idx = None
    for i, ln in enumerate(lines):
        if "destinatario" in ln.lower():
            dest_idx = i
            break

    if dest_idx is not None:
        for ln in lines[dest_idx + 1 : dest_idx + 8]:
            if ln.strip():
                recipient = ln.strip()
                break

    if not recipient:
        recipient = "(Destinatario non trovato)"
    else:
        # Mexal spesso rende "Destinatario" e "Destinazione" come due colonne.
        # Nell'estrazione testo possono finire sulla stessa riga in duplicato.
        # In tal caso prendiamo la prima colonna.
        parts = [p.strip() for p in re.split(r"\t+|\s{2,}", recipient) if p.strip()]
        if len(parts) >= 2:
            recipient = parts[0]
        else:
            # Caso: la stessa stringa ripetuta due volte nella stessa riga
            m_dup = re.match(r"^(?P<a>.+?)\s+(?P=a)\s*$", recipient)
            if m_dup:
                recipient = m_dup.group("a").strip()

    return ParsedDoc(
        source_path=pdf_path,
        created_at=mtime,
        doc_code=doc_code,
        doc_type=doc_type,
        doc_number=doc_number,
        doc_date=doc_date,
        recipient=recipient,
    )


def save_first_page_only(input_path: str, output_path: str) -> None:
    reader = PdfReader(input_path)
    writer = PdfWriter()
    writer.add_page(reader.pages[0])
    with open(output_path, "wb") as f_out:
        writer.write(f_out)


def print_pdf(path: str, copies: int = 1) -> None:
    copies = max(1, int(copies or 1))

    # Prefer SumatraPDF if installed: supports silent printing and copies.
    sumatra_candidates = [
        r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
        r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
    ]
    sumatra = next((p for p in sumatra_candidates if os.path.isfile(p)), None)
    if sumatra:
        # -print-settings supports copies; use default printer.
        args = [
            sumatra,
            "-silent",
            "-print-to-default",
            "-print-settings",
            f"{copies}x",
            path,
        ]
        subprocess.run(args, check=True)
        return

    # Fallback: uses the default PDF handler. Copies may be ignored by some viewers.
    for _ in range(copies):
        try:
            os.startfile(path, "print")
        except OSError as e:
            # WinError 1155: no application associated with the specified file for this operation
            if getattr(e, "winerror", None) == 1155:
                raise RuntimeError(
                    "Nessuna applicazione associata alla stampa PDF (WinError 1155). "
                    "Installa un lettore PDF che supporti la stampa da shell (consigliato: SumatraPDF) "
                    "oppure imposta un'app predefinita per i PDF con la funzione di stampa."
                )
            raise
        time.sleep(1.0)


class MexalDaemonApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.withdraw()

        self.state = _load_json(STATE_FILE, {"docs": {}})
        # seen[path] = last_mtime_processed
        self.seen = _load_json(SEEN_FILE, {"seen": {}})
        self._size_history: dict[str, list[int]] = {}

        self.overlay: Optional[tk.Toplevel] = None
        self.list_window: Optional[tk.Toplevel] = None
        self.list_tree: Optional[ttk.Treeview] = None

        self._last_detected: list[ParsedDoc] = []
        self._current_doc: Optional[ParsedDoc] = None
        self._tick_count = 0

        self.root.after(1000, self._tick)

    def _refresh_list_tree(self) -> None:
        tree = self.list_tree
        if not tree or not tree.winfo_exists():
            return

        try:
            for iid in list(tree.get_children()):
                tree.delete(iid)
        except Exception:
            return

        docs = self._collect_last_docs(limit=5)
        for doc in docs:
            created_str = time.strftime("%d/%m/%Y %H:%M:%S", time.localtime(doc.created_at))
            tree.insert(
                "",
                "end",
                iid=self._doc_id(doc),
                values=(doc.doc_code, doc.doc_type, doc.recipient, doc.doc_date, created_str),
            )

        if tree.get_children():
            tree.selection_set(tree.get_children()[0])
            tree.event_generate("<<TreeviewSelect>>")

    def _tick(self):
        try:
            self._tick_count += 1
            new_docs = self._scan_for_new_docs()
            if new_docs:
                self._last_detected = new_docs
                self._current_doc = new_docs[0]
                self._show_overlay(new_docs[0])
        finally:
            self.root.after(1000, self._tick)

    def _scan_for_new_docs(self) -> list[ParsedDoc]:
        if not os.path.isdir(MEXAL_TEMP):
            if self._tick_count % 10 == 0:
                _log(f"Scan: MEXAL_TEMP non esiste: {MEXAL_TEMP}")
            return []

        pdfs = []
        for root_dir, _, files in os.walk(MEXAL_TEMP):
            for fn in files:
                if fn.lower().endswith(".pdf"):
                    full_path = os.path.join(root_dir, fn)
                    mtime = _safe_get_mtime(full_path)
                    if mtime is None:
                        continue
                    pdfs.append((full_path, mtime))

        pdfs.sort(key=lambda x: x[1], reverse=True)

        if self._tick_count % 10 == 0:
            _log(f"Scan: found_pdfs={len(pdfs)} (showing up to 20)")

        parsed: list[ParsedDoc] = []
        for path, mtime in pdfs[:20]:
            seen_map = self.seen.setdefault("seen", {})
            last_mtime = seen_map.get(path)
            if last_mtime is not None and mtime <= float(last_mtime):
                continue

            size = _safe_get_size(path)
            if size is None:
                continue

            hist = self._size_history.setdefault(path, [])
            hist.append(size)
            if len(hist) > 3:
                del hist[0]

            if len(hist) < 2 or hist[-1] != hist[-2]:
                if self._tick_count % 10 == 0:
                    _log(f"Scan: not_stable_yet: {os.path.basename(path)} size_hist={hist}")
                continue

            doc = parse_mexal_pdf(path)
            if not doc:
                _log(f"Scan: parse_failed: {path}")
                continue

            parsed.append(doc)
            seen_map[path] = mtime

        if parsed:
            _log(f"Scan: new_docs={len(parsed)} first={os.path.basename(parsed[0].source_path)} type={parsed[0].doc_type}")
            _save_json(SEEN_FILE, self.seen)

        return parsed

    def _show_overlay(self, doc: ParsedDoc) -> None:
        if self.overlay and self.overlay.winfo_exists():
            self.overlay.lift()
            return

        win = tk.Toplevel(self.root)
        self.overlay = win
        win.title("Nuovo documento Mexal")
        win.attributes("-topmost", True)
        win.resizable(False, False)

        frm = ttk.Frame(win, padding=12)
        frm.grid(row=0, column=0, sticky="nsew")

        msg = (
            "È stato emesso un nuovo documento.\n"
            "Vuoi processarlo adesso?\n\n"
            f"Tipo: {doc.doc_code} ({doc.doc_type})\n"
            f"Destinatario: {doc.recipient}\n"
        )
        ttk.Label(frm, text=msg, justify="left").grid(row=0, column=0, columnspan=2, sticky="w")

        ttk.Button(frm, text="No", command=self._overlay_no).grid(row=1, column=0, sticky="ew", pady=(12, 0), padx=(0, 8))
        ttk.Button(frm, text="Sì", command=self._overlay_yes).grid(row=1, column=1, sticky="ew", pady=(12, 0))

        win.update_idletasks()
        w = win.winfo_width()
        h = win.winfo_height()
        x = int((win.winfo_screenwidth() - w) / 2)
        y = int((win.winfo_screenheight() - h) / 2)
        win.geometry(f"{w}x{h}+{x}+{y}")

        win.protocol("WM_DELETE_WINDOW", self._overlay_no)

    def _overlay_no(self):
        if self.overlay and self.overlay.winfo_exists():
            self.overlay.destroy()
        self.overlay = None

    def _overlay_yes(self):
        self._overlay_no()
        self._show_list_window()

    def _doc_id(self, doc: ParsedDoc) -> str:
        base = os.path.basename(doc.source_path)
        return f"{base}|{int(doc.created_at)}"

    def _preferred_save_dir(self, doc: ParsedDoc) -> str:
        if doc.doc_code in PATHS:
            chosen = PATHS[doc.doc_code]
            _log(f"SaveDir: doc_code={doc.doc_code} doc_type={doc.doc_type} chosen={chosen}")
            return chosen
        if doc.doc_type in PATHS:
            chosen = PATHS[doc.doc_type]
            _log(f"SaveDir: doc_code={doc.doc_code} doc_type={doc.doc_type} chosen={chosen}")
            return chosen
        chosen = os.path.join(BASE_PATH, "DOCUMENTI_2025")
        _log(f"SaveDir: doc_code={doc.doc_code} doc_type={doc.doc_type} chosen={chosen}")
        return chosen

    def _get_doc_state(self, doc_id: str) -> dict:
        return self.state.setdefault("docs", {}).setdefault(
            doc_id,
            {
                "saved": False,
                "printed": False,
                "emailed": False,
                "dest_path": "",
                "meta": {},
            },
        )

    def _show_list_window(self):
        if self.list_window and self.list_window.winfo_exists():
            self._refresh_list_tree()
            try:
                self.list_window.deiconify()
            except Exception:
                pass
            self.list_window.lift()
            try:
                self.list_window.focus_force()
            except Exception:
                pass
            return

        win = tk.Toplevel(self.root)
        self.list_window = win
        win.title("Documenti Mexal - ultimi")
        win.attributes("-topmost", True)
        try:
            win.deiconify()
            win.focus_force()
        except Exception:
            pass

        main = ttk.Frame(win, padding=10)
        main.grid(row=0, column=0, sticky="nsew")

        cols = ("tipo", "descr", "destinatario", "data_doc", "creato")
        tree = ttk.Treeview(main, columns=cols, show="headings", height=6)
        self.list_tree = tree
        tree.heading("tipo", text="Tipo")
        tree.heading("descr", text="Descrizione")
        tree.heading("destinatario", text="Destinatario")
        tree.heading("data_doc", text="Data doc")
        tree.heading("creato", text="Creato")

        tree.column("tipo", width=60, stretch=False)
        tree.column("descr", width=140, stretch=False)
        tree.column("destinatario", width=300, stretch=True)
        tree.column("data_doc", width=100, stretch=False)
        tree.column("creato", width=140, stretch=False)

        tree.grid(row=0, column=0, columnspan=4, sticky="nsew")

        docs = self._collect_last_docs(limit=5)
        for doc in docs:
            created_str = time.strftime("%d/%m/%Y %H:%M:%S", time.localtime(doc.created_at))
            tree.insert(
                "",
                "end",
                iid=self._doc_id(doc),
                values=(doc.doc_code, doc.doc_type, doc.recipient, doc.doc_date, created_str),
            )

        if tree.get_children():
            tree.selection_set(tree.get_children()[0])

        btn_save = ttk.Button(main, text="1) Salva", command=lambda: self._action_save(tree))
        btn_print = ttk.Button(main, text="2) Stampa", command=lambda: self._action_print(tree))
        btn_email = ttk.Button(main, text="3) Email", command=lambda: self._action_email(tree))
        btn_view = ttk.Button(main, text="4) Vedi", command=lambda: self._action_view(tree))

        btn_save.grid(row=1, column=0, sticky="ew", pady=(10, 0), padx=(0, 6))
        btn_print.grid(row=1, column=1, sticky="ew", pady=(10, 0), padx=(0, 6))
        btn_email.grid(row=1, column=2, sticky="ew", pady=(10, 0), padx=(0, 6))
        btn_view.grid(row=1, column=3, sticky="ew", pady=(10, 0))

        def refresh_buttons(*_):
            sel = tree.selection()
            if not sel:
                btn_save.state(["disabled"])
                btn_print.state(["disabled"])
                btn_email.state(["disabled"])
                btn_view.state(["disabled"])
                return

            doc_id = sel[0]
            st = self._get_doc_state(doc_id)
            if st.get("saved"):
                btn_save.state(["disabled"])
                btn_print.state(["!disabled"])
                btn_email.state(["!disabled"])
                btn_view.state(["!disabled"])
            else:
                btn_save.state(["!disabled"])
                btn_print.state(["disabled"])
                btn_email.state(["disabled"])
                btn_view.state(["disabled"])

        tree.bind("<<TreeviewSelect>>", refresh_buttons)
        refresh_buttons()

        win.protocol("WM_DELETE_WINDOW", win.withdraw)

    def _collect_last_docs(self, limit: int = 5) -> list[ParsedDoc]:
        pdfs = []
        if not os.path.isdir(MEXAL_TEMP):
            return []

        for root_dir, _, files in os.walk(MEXAL_TEMP):
            for fn in files:
                if fn.lower().endswith(".pdf"):
                    full_path = os.path.join(root_dir, fn)
                    mtime = _safe_get_mtime(full_path)
                    if mtime is None:
                        continue
                    pdfs.append((full_path, mtime))

        pdfs.sort(key=lambda x: x[1], reverse=True)

        docs: list[ParsedDoc] = []
        for path, _ in pdfs:
            doc = parse_mexal_pdf(path)
            if doc:
                docs.append(doc)
            if len(docs) >= limit:
                break

        # Forza l'ultimo documento rilevato in cima (deve sempre apparire nel modale)
        current = self._current_doc or (self._last_detected[0] if self._last_detected else None)
        if current:
            current_id = self._doc_id(current)
            docs = [d for d in docs if self._doc_id(d) != current_id]
            docs.insert(0, current)
            docs = docs[:limit]

        return docs

    def _find_doc_by_id(self, doc_id: str) -> Optional[ParsedDoc]:
        docs = self._collect_last_docs(limit=10)
        for d in docs:
            if self._doc_id(d) == doc_id:
                return d
        return None

    def _action_save(self, tree: ttk.Treeview):
        sel = tree.selection()
        if not sel:
            return
        doc_id = sel[0]
        doc = self._find_doc_by_id(doc_id)
        if not doc:
            messagebox.showerror("Errore", "Documento non trovato.")
            return

        st = self._get_doc_state(doc_id)
        if st.get("saved"):
            return

        save_dir = self._preferred_save_dir(doc)
        _log(f"ActionSave: doc_code={doc.doc_code} doc_type={doc.doc_type} number={doc.doc_number} dest={save_dir}")
        os.makedirs(save_dir, exist_ok=True)

        numero = (doc.doc_number or "").strip()
        intestatario = (doc.recipient or "").strip()
        data_doc = (doc.doc_date or "").strip()

        if doc.doc_code == "DDT":
            parts = [numero, intestatario, data_doc]
        elif doc.doc_code in {"PC", "FC"}:
            parts = [numero, intestatario]
        else:
            # Default: tipo documento; numero; intestatario; data documento
            tipo = (doc.doc_code or "?").strip()
            parts = [tipo, numero, intestatario, data_doc]

        parts = [p for p in parts if p]
        if not parts:
            parts = ["Documento"]

        filename = " ".join(parts) + ".pdf"
        filename = re.sub(r"[\\/:*?\"<>|]", "-", filename)
        filename = re.sub(r"\s+", " ", filename)
        dest_path = os.path.join(save_dir, filename)

        try:
            if doc.doc_type.lower() == "fattura":
                save_first_page_only(doc.source_path, dest_path)
            else:
                shutil.copy2(doc.source_path, dest_path)
        except Exception as e:
            messagebox.showerror("Errore", f"Errore durante il salvataggio:\n{e}")
            return

        st["saved"] = True
        st["dest_path"] = dest_path
        st["meta"] = {
            "doc_code": doc.doc_code,
            "doc_type": doc.doc_type,
            "doc_number": doc.doc_number,
            "recipient": doc.recipient,
            "doc_date": doc.doc_date,
            "source_path": doc.source_path,
            "created_at": doc.created_at,
        }
        _save_json(STATE_FILE, self.state)

        messagebox.showinfo("Completato", f"Documento salvato in:\n{dest_path}")
        tree.event_generate("<<TreeviewSelect>>")

    def _action_print(self, tree: ttk.Treeview):
        sel = tree.selection()
        if not sel:
            return
        doc_id = sel[0]
        st = self._get_doc_state(doc_id)
        if not st.get("saved") or not st.get("dest_path"):
            return

        copies = self._ask_copies()
        if copies is None:
            return

        try:
            print_pdf(st["dest_path"], copies=copies)
        except Exception as e:
            messagebox.showerror("Errore", f"Errore stampa:\n{e}")
            return

        st["printed"] = True
        _save_json(STATE_FILE, self.state)

    def _ask_copies(self) -> Optional[int]:
        dlg = tk.Toplevel(self.root)
        dlg.title("Stampa")
        dlg.attributes("-topmost", True)
        dlg.resizable(False, False)

        frm = ttk.Frame(dlg, padding=10)
        frm.grid(row=0, column=0, sticky="nsew")

        ttk.Label(frm, text="Numero copie:").grid(row=0, column=0, sticky="w")
        var = tk.StringVar(value="1")
        entry = ttk.Entry(frm, textvariable=var, width=6)
        entry.grid(row=0, column=1, sticky="w", padx=(8, 0))
        entry.focus_set()

        result: dict[str, Optional[int]] = {"value": None}

        def ok():
            try:
                n = int(var.get().strip())
                if n < 1:
                    raise ValueError
            except Exception:
                messagebox.showwarning("Attenzione", "Inserisci un numero valido (>= 1).")
                return
            result["value"] = n
            dlg.destroy()

        def cancel():
            dlg.destroy()

        ttk.Button(frm, text="Annulla", command=cancel).grid(row=1, column=0, pady=(10, 0), sticky="ew", padx=(0, 8))
        ttk.Button(frm, text="OK", command=ok).grid(row=1, column=1, pady=(10, 0), sticky="ew")

        dlg.grab_set()
        self.root.wait_window(dlg)
        return result["value"]

    def _ask_email(self, doc: ParsedDoc) -> Optional[dict[str, str]]:
        dlg = tk.Toplevel(self.root)
        dlg.title("Email")
        dlg.attributes("-topmost", True)
        dlg.resizable(False, False)

        frm = ttk.Frame(dlg, padding=10)
        frm.grid(row=0, column=0, sticky="nsew")

        ttk.Label(frm, text="A:").grid(row=0, column=0, sticky="w")
        to_var = tk.StringVar(value="")
        to_entry = ttk.Entry(frm, textvariable=to_var, width=42)
        to_entry.grid(row=0, column=1, sticky="ew", padx=(8, 0))

        subj_default = f"{(doc.doc_code or '').strip()} {(doc.doc_number or '').strip()} {(doc.recipient or '').strip()}".strip()
        ttk.Label(frm, text="Oggetto:").grid(row=1, column=0, sticky="w", pady=(8, 0))
        subj_var = tk.StringVar(value=subj_default)
        subj_entry = ttk.Entry(frm, textvariable=subj_var, width=42)
        subj_entry.grid(row=1, column=1, sticky="ew", padx=(8, 0), pady=(8, 0))

        ttk.Label(frm, text="Testo:").grid(row=2, column=0, sticky="nw", pady=(8, 0))
        body = tk.Text(frm, width=42, height=6)
        body.grid(row=2, column=1, sticky="ew", padx=(8, 0), pady=(8, 0))

        recip = (doc.recipient or "").strip()
        default_body = (
            f"Buongiorno {recip},\n"
            "in allegato troverà il documento in oggetto.\n\n"
            "Cordiali saluti\n"
            "Fior d'Acqua Team"
        ).strip()
        body.insert("1.0", default_body)

        result: dict[str, Optional[dict[str, str]]] = {"value": None}

        def ok():
            result["value"] = {
                "to": to_var.get().strip(),
                "subject": subj_var.get().strip(),
                "body": body.get("1.0", "end").strip(),
            }
            dlg.destroy()

        def cancel():
            dlg.destroy()

        ttk.Button(frm, text="Annulla", command=cancel).grid(row=3, column=0, pady=(10, 0), sticky="ew", padx=(0, 8))
        ttk.Button(frm, text="OK", command=ok).grid(row=3, column=1, pady=(10, 0), sticky="ew")

        dlg.grab_set()
        to_entry.focus_set()
        self.root.wait_window(dlg)
        return result["value"]

    def _action_email(self, tree: ttk.Treeview):
        sel = tree.selection()
        if not sel:
            return
        doc_id = sel[0]
        st = self._get_doc_state(doc_id)
        if not st.get("saved") or not st.get("dest_path"):
            return

        doc = self._find_doc_by_id(doc_id)
        if not doc:
            messagebox.showerror("Errore", "Documento non trovato.")
            return

        fields = self._ask_email(doc)
        if not fields:
            return

        to_addr = fields.get("to", "").strip()
        subject = fields.get("subject", "").strip()
        body = fields.get("body", "").strip()
        attachment_path = st["dest_path"]

        to_addrs = [a.strip() for a in re.split(r"[;,\s]+", to_addr) if a.strip()]
        if not to_addrs:
            messagebox.showwarning("Attenzione", "Inserisci almeno un destinatario valido (campo A:).")
            return

        cfg = _smtp_config()
        host = str(cfg.get("host") or "").strip()
        port = int(cfg.get("port") or 0)
        user = str(cfg.get("user") or "").strip()
        password = str(cfg.get("password") or "").strip()
        from_addr = str(cfg.get("from_addr") or "").strip()

        if not host or not port or not user or not password or not from_addr:
            messagebox.showerror(
                "Errore",
                "Configurazione SMTP mancante. Imposta le variabili d'ambiente:\n"
                "SMTP_HOST=smtp.gmail.com\n"
                "SMTP_PORT=465\n"
                "SMTP_USER=<la mail gmail>\n"
                "SMTP_PASS=<app password gmail>\n"
                "(opzionale) SMTP_FROM=<mittente visualizzato>\n",
            )
            return

        try:
            send_email_smtp(
                host=host,
                port=port,
                user=user,
                password=password,
                from_addr=from_addr,
                to_addrs=to_addrs,
                subject=subject,
                body=body,
                attachment_path=os.path.abspath(attachment_path),
            )
        except Exception as e:
            _log(f"SMTP send failed: {e}")
            messagebox.showerror("Errore", f"Invio email fallito:\n{e}")
            return

        st["emailed"] = True
        _save_json(STATE_FILE, self.state)
        tree.event_generate("<<TreeviewSelect>>")
        messagebox.showinfo("Email", "Email inviata.")

    def _action_view(self, tree: ttk.Treeview):
        sel = tree.selection()
        if not sel:
            return
        doc_id = sel[0]
        st = self._get_doc_state(doc_id)
        if not st.get("saved") or not st.get("dest_path"):
            return
        try:
            os.startfile(st["dest_path"])
        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile aprire il file:\n{e}")


if __name__ == "__main__":
    if "--install-startup" in sys.argv:
        install_windows_shortcuts()
        print("OK: collegamenti creati (Startup + Desktop).")
        raise SystemExit(0)

    if "--uninstall-startup" in sys.argv:
        uninstall_windows_shortcuts()
        print("OK: collegamenti rimossi (Startup + Desktop).")
        raise SystemExit(0)

    root = tk.Tk()
    ttk.Style().theme_use("clam")
    MexalDaemonApp(root)
    root.mainloop()
