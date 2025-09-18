"""
SCRIPT PER AGGIORNARE I CoA DEI SERBATOI 410 E 411 QUANDO ARRIVANO LE ANALISI DA VADO E NEL FRATTEMPO HO GIÀ FATTO DEI TRASFERIMENTI
PUÒ LAVORARE IN TANDEM CON coa_serbatoi.py (PRIMA LANCIO QUELLO E POI QUESTO)

by Simone Barondi
s.barondi@iglom.it
"""

# IMPORTAZIONE MODULI

import pandas as pd # Per gestire i file Excel
import glob # per listare le cartelle
from pathlib import Path # per prendere solo il nome del file senza estensione
import shutil # per copiare e incollare i file
import sys # per fermare l'esecuzione se c'è qualcosa che non va
import os, fnmatch # per cercare il FdM corrispondente al batch
import math # per il logaritmo

# DEFINIZIONE FUNZIONI

# Definisco la funzione per dividere il lotto in due (es. 910714-1 -> 910714 e -1)
def spezza(lotto:str) -> list:
    """ Prende il batch con il trattino e lo divide nelle due parti (batch e numero del campione).
        Se il lotto non ha un trattino restituisce solo il batch.
        Argomenti:
        lotto (str): è il batch intero da dividere
        Restituisce:
        lotto_diviso (list): una lista i cui elementi sono le due porzioni divise del batch
    """
    lotto_diviso = lotto.split('-')
    try:
        lotto_diviso[1] = '-' + lotto_diviso[1]
    except IndexError: # se non trova il trattino evita l'errore mettendo una stringa vuota alla fine
        lotto_diviso.append('')
    return lotto_diviso

# Definisco la funzione per trovare il foglio di marcia
def trova(pattern:str, path:str) -> str:
    """ Trova un file fornendo una porzione del nome e il percorso in cui va cercato
        Argomenti:
        pattern (str): porzione del nome file da cercare, in questo caso le due parti di batch separate dalla wildcard *
        path (str): percorso nel quale cercare il file, in questo caso la cartella dei fogli di marcia
        Restituisce:
        result[0] (str): la funzione fnmatch restituisce una lista con i vari risultati, ma dato che il batch è univoco
        sono sicuro di trovare un solo risultato, quindi prendo sempre il primo elemento della lista sapendo che è l'unico
    """
    result = []
    for root, dirs, files in os.walk(path):
        for name in sorted(files,reverse=True):
            if fnmatch.fnmatch(name, pattern):
                result.append(os.path.join(root, name))
    return result[0]

# Definisco la funzione per calcolare i valori delle analisi (tranne Kv100)
def calcolo(per_iniz:float,per_trasf:float,val_iniz:int,val_trasf:int) -> float:
    """ Funzione che calcola i valori finali del serbatoio considerando le analisi del campione, le analisi precedenti
        del serbatoio e la percentuale in peso del prodotto trasferito rispetto alla quantità totale nel serbatoio
        Argomenti:
        per_iniz (float): percentuale del prodotto presente nel serbatoio prima di trasferire rispetto a quello
                          totale dopo il trasferimento
        per_trasf (float): percentuale del prodotto trasferito rispetto al totale dopo il trasferimento
        val_iniz (int): chili nel serbatoio prima di trasferire
        val_trasf (int): chili di prodotto trasferiti
        Restituisce:
        il valore dell'analisi finale utilizzando la formula di diluizione (float)
        """
    return val_iniz * per_iniz / 100 + val_trasf * per_trasf / 100

# Definisco la funzione per calcolare la Kv100
def kv100(per_iniz,per_trasf,val_iniz,val_trasf):
    """ Funzione analoga a calcolo() con gli stessi argomenti, ma la viscosità ha una proporzionalità di tipo logaritmico
        e quindi la formula è diversa """
    return math.pow(10, math.log10(val_iniz)*per_iniz/100+math.log10(val_trasf)*per_trasf/100)

# ESECUZIONE PROGRAMMA

# Chiedo se è modalità test
while True:
    test = input("Vuoi lavorare in modalità test? S/N: ").lower()
    # Salvo i percorsi principali come variabili, a seconda se sono in modalità test o no
    if test == "s":
        path_main = r"C:\Users\s.barondi\Documents\TEST"
        break
    elif test == "n":
        path_main = r"\\iglomfs\Produzione\FILTRAZIONE\COMPUTER LAB"
        break
    else:
        print("Devi scegliere un'opzione tra S e N!")
path_fdm = fr"{path_main}\Fogli di marcia"

# Faccio il giro completo prima per il 410 e poi per il 411
tanks = ["410", "411"]
for tank in tanks:
    if tank == "410":
        prodotto = "D3336F"
    else:
        prodotto = "P6072F"
    path_coa = fr"{path_main}\{tank}_TRASFERIMENTI {prodotto}"
    elenco = glob.glob(fr"{path_coa}\*.xlsx")
    elenco = sorted(elenco)

    # Controllo il numero dei certificati VADO (deve essere 1 per forza)
    conta_vado = sum(1 for a in elenco if "VADO" in a)
    if conta_vado == 0:
        sys.exit("Nella cartella non c'è un certificato VADO!\nPer favore inseriscilo e riprova.")
    elif conta_vado > 1:
        sys.exit("Nella cartella c'è più di un certificato VADO!\nPer favore lascia solo l'ultimo e riprova.")

    # Inizio il ciclo di copia - apertura - modifica - chiusura dei file Excel
    analisi = pd.DataFrame(columns=['vecchio','nuovo']) # Creo il DataFrame per le analisi
    provv = fr"{path_coa}\provv.xlsx" # nome del file provvisorio
    for file in elenco:
        # Ogni volta creo un file provvisorio che poi vado a chiamare con il nome del CoA successivo
        # Con il CoA di Vado faccio solo questo passaggio
        if "VADO" not in file: # con i CoA non VADO rinomino il vecchio file come old e provv.xlsx con il nome del CoA
            nomefile = Path(file).stem
            file_old = fr"{path_coa}\{nomefile} old.xlsx"
            shutil.move(file, file_old)
            shutil.move(provv, file)
            # Inizio a estrapolare i dati dai file Excel e metterli nel DataFrame delle analisi
            # succhia è la variabile in cui ogni volta estraggo il file Excel interessato
            # taglia è la variabile in cui butto la parte interessata del file (prima i dati del lotto e poi le analisi)
            succhia = pd.read_excel(file_old, engine='openpyxl')
            taglia_prod = succhia.iloc[7:11, 3]
            peso_iniz = succhia.iloc[7, 3]
            peso_trasf = succhia.iloc[8, 3]
            peso_tot = peso_iniz + peso_trasf
            perc_iniz = peso_iniz / peso_tot * 100
            perc_trasf = peso_trasf / peso_tot * 100
            taglia_prod[9] = '=D9+D10'
            batch = str(succhia.iloc[10, 3])
            succhia = pd.read_excel(file,engine='openpyxl')
            taglia_analisi = succhia.iloc[17:28, 3]
            analisi['vecchio'] = taglia_analisi.reset_index(drop=True)
            batch_diviso = spezza(batch)
            fdm = trova(f'*{batch_diviso[0]}*{batch_diviso[1]}*',path_fdm) # cerco il FdM corrispondente al batch
            succhia = pd.read_excel(fdm,engine='openpyxl',sheet_name='trasferimento')
            taglia_analisi = succhia.iloc[0:11,2]
            analisi['nuovo'] = taglia_analisi
            analisi['risultato'] = analisi['vecchio']
            mask = analisi['risultato'] != 'Pass'
            analisi.loc[mask, 'risultato'] = calcolo(perc_iniz, perc_trasf, analisi.loc[mask, 'vecchio'], analisi.loc[mask, 'nuovo'])
            if tank == "410": # la viscosità è in posizione diversa a seconda che sia D3336F o P6072F
                analisi.iloc[3,2] = kv100(perc_iniz,perc_trasf,analisi.iloc[3,0],analisi.iloc[3,1])
            else:
                analisi.iloc[5, 2] = kv100(perc_iniz, perc_trasf, analisi.iloc[5, 0], analisi.iloc[5, 1])
            with pd.ExcelWriter(file,mode='a',if_sheet_exists='overlay',engine='openpyxl') as writer:
                taglia_prod.to_excel(writer,sheet_name=prodotto,startrow=8,startcol=3,header=False,index=False)
                analisi['risultato'].to_excel(writer,sheet_name=prodotto,startrow=18,startcol=3,header=False,index=False)
        shutil.copy(file, provv)
    
    # PULISCO I FILE DA CANCELLARE
    elenco = glob.glob(fr"{path_coa}\*.xlsx")
    for file in elenco:
        if "old" in file or "provv" in file:
            os.remove(file)