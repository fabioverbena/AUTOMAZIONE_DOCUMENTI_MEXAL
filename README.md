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
Impostare le variabili d’ambiente (App Password Gmail):

```powershell
$env:SMTP_HOST="smtp.gmail.com"
$env:SMTP_PORT="465"
$env:SMTP_USER="<gmail>"
$env:SMTP_PASS="<app_password>"
# opzionale
$env:SMTP_FROM="Fior d'Acqua Team <gmail>"
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
1. Copia la cartella del progetto (o clona da GitHub)
2. Crea venv e installa requirements
3. Imposta variabili SMTP (a livello Utente o Sistema)
4. (Consigliato) installa SumatraPDF
5. Esegui `--install-startup`
