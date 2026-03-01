# Mexal Automation Daemon

Daemon Windows (Tkinter) che monitora la cartella temporanea di Mexal per nuovi PDF e mostra un toast/modale per:
- Salvataggio in cartelle dedicate con naming automatico
- Stampa
- Invio Email via SMTP Gmail con allegato

## Requisiti
- Windows
- Python 3.10+ consigliato
- (Consigliato) SumatraPDF per stampa affidabile

## Installazione
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Configurazione SMTP (Gmail)
Puoi configurare SMTP in 2 modi:
- Variabili d’ambiente Windows
- File `.env` nella cartella dell’app (utile per versione portable)

Se premi **Email** e la configurazione manca, l’app propone un wizard per inserirla e salvarla in `.env`.

### Variabili d’ambiente (App Password Gmail)

```powershell
$env:SMTP_HOST="smtp.gmail.com"
$env:SMTP_PORT="465"
$env:SMTP_USER="<gmail>"
$env:SMTP_PASS="<app_password>"
# opzionale
$env:SMTP_FROM="Fior d'Acqua Team <gmail>"
```

### File `.env`
```ini
SMTP_HOST=smtp.gmail.com
SMTP_PORT=465
SMTP_USER=<gmail>
SMTP_PASS=<app_password>
SMTP_FROM=Fior d'Acqua Team <gmail>
```

Nota: **non** salvare password nel repository.

## Avvio
```powershell
python .\mexal_daemon.py
```

## Installazione collegamenti (auto-avvio + Desktop)
Crea:
- Startup: avvio invisibile (pythonw)
- Desktop: avvio manuale

```powershell
python .\mexal_daemon.py --install-startup
```

Per rimuovere:
```powershell
python .\mexal_daemon.py --uninstall-startup
```

## Migrazione su un altro PC
Vedi: `INSTALL_PC_LAVORO.md`
