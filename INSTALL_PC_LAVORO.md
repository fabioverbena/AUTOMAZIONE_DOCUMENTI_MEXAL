# Installazione su PC lavoro (Windows)

Questa guida serve per installare e configurare **Mexal Automation Daemon** su un nuovo PC Windows.

## 1) Prerequisiti

- Python 3.10+ (consigliato)
- Accesso a Internet (solo per installare le dipendenze)
- (Consigliato) **SumatraPDF** per una stampa affidabile da riga di comando

## 2) Ottenere il progetto

### Opzione A: clonare da GitHub
1. Installa Git (se non presente)
2. Clona il repository in una cartella a tua scelta, ad esempio:

```powershell
git clone <URL_REPO_GITHUB>
cd <CARTELLA_REPO>
```

### Opzione B: copiare la cartella
Copia l’intera cartella del progetto in una posizione stabile (es. `C:\Tools\MEXAL_AUTOMATION`).

## 3) Creare l’ambiente virtuale e installare le dipendenze

Dalla cartella del progetto:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## 4) Configurare SMTP Gmail (invio email)

Il daemon legge la configurazione SMTP da **variabili d’ambiente**.

Valori richiesti:
- `SMTP_HOST` (es. `smtp.gmail.com`)
- `SMTP_PORT` (es. `465`)
- `SMTP_USER` (email Gmail completa)
- `SMTP_PASS` (App Password Gmail)
- `SMTP_FROM` (opzionale)

### Impostazione (consigliata) permanente a livello utente
Esegui una volta (PowerShell):

```powershell
[Environment]::SetEnvironmentVariable('SMTP_HOST','smtp.gmail.com','User')
[Environment]::SetEnvironmentVariable('SMTP_PORT','465','User')
[Environment]::SetEnvironmentVariable('SMTP_USER','<GMAIL>','User')
[Environment]::SetEnvironmentVariable('SMTP_PASS','<APP_PASSWORD>','User')
# opzionale
[Environment]::SetEnvironmentVariable('SMTP_FROM','<NOME> <GMAIL>','User')
```

Dopo averle impostate, **chiudi e riapri** eventuali terminali/app per farle leggere.

## 5) (Consigliato) Installare SumatraPDF per la stampa

Installa SumatraPDF in uno dei percorsi standard:
- `C:\Program Files\SumatraPDF\SumatraPDF.exe`
- `C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe`

Il daemon lo usa automaticamente per stampare in modo silenzioso e gestire le copie.

## 6) Avvio manuale (test)

Dalla cartella del progetto:

```powershell
python .\mexal_daemon.py
```

Genera un documento da Mexal e verifica:
- toast (overlay)
- apertura modale con “Sì”
- Salva / Stampa / Email

## 7) Installare avvio automatico + collegamento Desktop

Quando tutto funziona, crea i collegamenti:

```powershell
python .\mexal_daemon.py --install-startup
```

Risultato:
- Avvio automatico: esecuzione invisibile (pythonw) all’accesso utente
- Collegamento Desktop: avvio manuale

Per rimuovere:

```powershell
python .\mexal_daemon.py --uninstall-startup
```

## 8) Note operative

- Se non compare nulla all’avvio è normale: il daemon è “silenzioso”.
- Vedrai finestre solo quando Mexal genera un nuovo PDF nella cartella temp.
- Log: `mexal_daemon.log` nella cartella del progetto.
