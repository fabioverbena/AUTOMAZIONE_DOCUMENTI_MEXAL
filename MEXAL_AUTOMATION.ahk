#NoEnv
#SingleInstance Force
#Persistent
SendMode Input
SetWorkingDir %A_ScriptDir%

; Cartelle di destinazione
BolleFolder := "C:\Users\Fabio\Desktop\AMMINISTRAZIONE_2025\BOLLE_2025"
FattureFolder := "C:\Users\Fabio\Desktop\AMMINISTRAZIONE_2025\FATTURE_2025"

; Monitora l'apertura di Adobe Reader
SetTimer, CheckForPDF, 2000
return

CheckForPDF:
    ; Controlla se Adobe Reader è aperto con un PDF
    WinGet, readerWindows, List, ahk_exe AcroRd32.exe
    if (readerWindows > 0) {
        ; Ferma il timer per evitare duplicati
        SetTimer, CheckForPDF, Off
        
        ; Aspetta un momento che il PDF si carichi completamente
        Sleep, 1000
        
        ; Mostra la finestra di input
        Gosub, ShowInputDialog
    }
return

ShowInputDialog:
    ; Distruggi GUI esistente se presente
    Gui, Destroy
    
    ; Crea la GUI per inserire i dati
    Gui, Add, Text, x10 y10 w200 Center, === GESTIONE DOCUMENTO PDF ===
    
    Gui, Add, Text, x10 y40, Tipo Documento:
    Gui, Add, Radio, x10 y60 vRadioBolla Checked, Bolla
    Gui, Add, Radio, x10 y80 vRadioFattura, Fattura
    
    Gui, Add, Text, x10 y110, Numero Documento:
    Gui, Add, Edit, x10 y130 w200 vEditNumero
    
    Gui, Add, Text, x10 y160, Nome Cliente:
    Gui, Add, Edit, x10 y180 w200 vEditCliente
    
    Gui, Add, Button, x10 y220 w90 h30 gProcessaDocumento, Salva e Stampa
    Gui, Add, Button, x110 y220 w90 h30 gChiudiFinestra, Annulla
    
    Gui, Show, w230 h270, Gestione Documento PDF
    return

ProcessaDocumento:
    ; Ottieni i valori inseriti
    Gui, Submit
    
    ; Validazione input
    if (EditNumero = "" || EditCliente = "") {
        MsgBox, 16, Errore, Inserire numero documento e nome cliente!
        Gosub, ShowInputDialog
        return
    }
    
    ; Determina tipo documento e cartella
    if (RadioBolla = 1) {
        TipoDoc := "Bolla"
        DestFolder := BolleFolder
    } else {
        TipoDoc := "Fattura" 
        DestFolder := FattureFolder
    }
    
    ; Crea il nome del file
    NomeFile := TipoDoc . "-" . EditNumero . " " . EditCliente . ".pdf"
    PercorsoCompleto := DestFolder . "\" . NomeFile
    
    ; Attiva la finestra di Adobe Reader
    WinActivate, ahk_exe AcroRd32.exe
    WinWaitActive, ahk_exe AcroRd32.exe, , 3
    
    if ErrorLevel {
        MsgBox, 16, Errore, Impossibile attivare Adobe Reader!
        Gosub, ResetTimer
        return
    }
    
    ; Salva il documento - comando corretto per Adobe Reader
    Send, ^+s
    Sleep, 1500
    
    ; Controlla se si è aperta una finestra di dialogo
    ; Proviamo con titoli diversi che Adobe Reader potrebbe usare
    IfWinExist, Salva con nome
    {
        WinActivate, Salva con nome
        goto SalvataggioTrovato
    }
    
    IfWinExist, Save As
    {
        WinActivate, Save As  
        goto SalvataggioTrovato
    }
    
    IfWinExist, Salva
    {
        WinActivate, Salva
        goto SalvataggioTrovato
    }
    
    ; Se non funziona Shift+Ctrl+S, proviamo con menu File
    Send, !f
    Sleep, 500
    Send, a
    Sleep, 1500
    
    IfWinExist, Salva con nome
    {
        WinActivate, Salva con nome
        goto SalvataggioTrovato
    }
    
    IfWinExist, Save As
    {
        WinActivate, Save As
        goto SalvataggioTrovato  
    }
    
    ; Se ancora non funziona
    MsgBox, 16, Errore, Impossibile aprire la finestra di salvataggio!`n`nProva manualmente:`n1. Premi Shift+Ctrl+S in Adobe Reader`n2. Salva come: %NomeFile%`n3. Nella cartella: %DestFolder%
    Gosub, ResetTimer
    return
    
    SalvataggioTrovato:
    
    ; Inserisce il percorso completo nel campo nome file
    Send, ^a
    Sleep, 200
    Send, %PercorsoCompleto%
    Sleep, 500
    Send, {Enter}
    
    ; Aspetta che il salvataggio sia completato
    Sleep, 2000
    
    ; Stampa il documento
    WinActivate, ahk_exe AcroRd32.exe
    Sleep, 500
    Send, ^p
    
    ; Aspetta la finestra di stampa
    Sleep, 2000
    WinWaitActive, Stampa, , 5
    if !ErrorLevel {
        Sleep, 1000
        Send, {Enter}  ; Conferma stampa con stampante predefinita
    }
    
    ; Aspetta un po' e chiude Adobe Reader
    Sleep, 3000
    WinClose, ahk_exe AcroRd32.exe
    
    ; Mostra messaggio di conferma
    MsgBox, 64, Completato, Documento salvato e stampato!`n`nFile: %NomeFile%`nCartella: %DestFolder%
    
    ; Riattiva il timer per il prossimo documento
    Gosub, ResetTimer
    return

ChiudiFinestra:
    Gui, Destroy
    Gosub, ResetTimer
    return

ResetTimer:
    SetTimer, CheckForPDF, 2000
    return

; Hotkey per chiudere lo script (Ctrl+Alt+X)
^!x::ExitApp

; Hotkey per mostrare manualmente la finestra (Ctrl+Alt+M)  
^!m::Gosub, ShowInputDialog

GuiClose:
    Gui, Destroy
    Gosub, ResetTimer
    return