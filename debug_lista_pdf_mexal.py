
import os
from tkinter import messagebox, Tk

MEXAL_TEMP = r"C:\Passepartout\PassClient\mxdesk1205143000\temp"

def get_all_pdfs():
    found = []
    for root, dirs, files in os.walk(MEXAL_TEMP):
        for file in files:
            if file.lower().endswith(".pdf"):
                full_path = os.path.join(root, file)
                found.append(full_path)
    return found

pdfs = get_all_pdfs()

root = Tk()
root.withdraw()

if pdfs:
    msg = "PDF trovati:\n\n" + "\n".join(pdfs)
else:
    msg = "⚠️ Nessun file PDF trovato in:\n" + MEXAL_TEMP

messagebox.showinfo("Risultato scansione", msg)
