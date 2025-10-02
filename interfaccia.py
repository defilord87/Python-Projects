# importo i vari pezzi di tkinter per la finestra di dialogo:
from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog
import sys # per chiudere lo script Python

def finestra():

    # definisco la funzione per aprire il programma
    def apri_file():
        global percorso_file
        file = filedialog.askopenfilename(
            title="Seleziona il programma da importare",
            defaultextension="xlsx",
            filetypes=[("File Excel", "*.xlsx"), ("Tutti i file", "*.*")],
            initialdir=r"\\iglomfs\Produzione\FILTRAZIONE\COMPUTER LAB\CREAZIONE CARICHI SCARICHI"
        )
        if file:
            lbl.config(text=f"Programma caricato:\n{file}")
    
    # definisco la funzione che sarà chiamata dal pulsante Annulla per chiudere tutto lo script
    def annulla():
        root.destroy()
        sys.exit()
    
    root = Tk()
    frame = Frame(root, width=300, height=200, padding=10)
    frame.grid()
    lbl = Label(frame, text="File non caricato")
    lbl.grid(column=0, row=0)
    btn_apri = Button(frame, text="Carica un altro programma", command=apri_file)
    btn_apri.grid(column=0, row=2)
    btn_ok = Button(frame, text="Ok", command=root.destroy)
    btn_ok.grid(column=1, row=2)
    btn_annulla = Button(frame, text="Annulla", command=annulla)
    btn_annulla.grid(column=2, row=2)
    file = filedialog.askopenfilename(
        title="Seleziona il programma da importare",
        defaultextension="xlsx",
        filetypes=[("File Excel", "*.xlsx"), ("Tutti i file", "*.*")],
        initialdir=r"\\iglomfs\Produzione\FILTRAZIONE\COMPUTER LAB\CREAZIONE CARICHI SCARICHI"
    )
    lbl.config(text=f"File caricato: {file}")

    root.mainloop()

    return file