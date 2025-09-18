"""
SCRIPT PER LEGGERE I CoA DEI SERBATOI 410 E 411 E CREARE I FILE EXCEL CORRISPONDENTI
DOPO QUESTO SCRIPT POSSO LANCIARE trasferimenti.py PER AGGIORNARE EVENTUALI CoA EXCEL CREATI A PARTIRE DA CoA _VADO PRECEDENTI

by Simone Barondi
s.barondi@iglom.it
"""

# Importo i moduli necessari
import pymupdf # per leggere i dati contenuti nei file .pdf
import glob # per cercare i file nella cartella
import shutil # per copiare i file Excel creati nelle cartelle di destinazione
import pandas as pd # per salvare le analisi come Series ed esportarle nel file Excel del CoA
from datetime import datetime # per trasformare la data da formato ggmmaa (letta nel nome del file) in gg/mm/aaaa
import sys # per uscire dal programma se non c'è esattamente un certificato nella cartella

# Creo una classe per l'exception che viene generata se nella cartella non ci sono esattamente due CoA
class CoaNumber(Exception):
    pass

""" Creo due dizionari per le analisi da cercare, rispettivamente per il 410 e per il 411
I valori saranno assegnati alle rispettive chiavi quando le analisi saranno trovate nel CoA pdf """
analisi_410 = {
    'Appearance': 'Pass',
    'Total Base Number': None,
    'Calcium': None,
    'Kv 100': None,
    'Magnesium': None,
    'Molybdenum': None, # da dividere per 10000 perché nell'Excel è in %, ma il primo risultato che trova nel pdf è in ppm
    'Nitrogen / 1': None,
    'Phosphorus': None,
    'Zinc': None
}
analisi_411 = {
    'Appearance': 'Pass',
    'Total Base Number': None,
    'Boron': None,
    'Calcium': None,
    'IR': 'Pass',
    'Kv 100': None,
    'Nitrogen / 1': None,
    'Phosphorus': None,
    'S_Ash': None,
    'Zinc': None,
    'Magnesium': None, # da moltiplicare per 10000 perché nell'Excel è in ppm, ma il primo risultato che trova nel pdf è in %
    'Water': None
}

# Chiedo se si vuole lavorare in modalità offline oppure no
# In base alla scelta imposto il percorso di lavoro
while True:
    offline = input("Vuoi lavorare in modalità offline? (S/N) ").lower()
    if offline == 's':
        path = r'D:\Documenti\Python'
        break
    elif offline == 'n':
        path = r'\\iglomfs\Produzione\FILTRAZIONE\COMPUTER LAB'
        break
    else:
        print("Devi scegliere un'opzione tra S e N!")

# Cerco i due file dei CoA nella cartella assegnata
try:
    coa = glob.glob(fr'{path}\CoA serbatoi\*.pdf')
    if len(coa) != 2:
        raise CoaNumber()
except CoaNumber:
    sys.exit("Nella cartella devono esserci esattamente due CoA! Controlla e poi rilancia lo script.")
for c in coa:
    # In base al nome del file imposto serbatoio e prodotto, che mi servanno per le cartelle e i nomi dei file
    if '410' in c:
        serb = '410'
        tank = 'TK410'
        product = 'D3336F'
        analisi = analisi_410
    else:
        serb = '411'
        tank = 'TK411'
        product = 'P6072F'
        analisi = analisi_411

    # Apro il pdf con pymupdf, cerco le analisi in base alle keys del dizionario e popolo i valori con i risultati trovati nel CoA
    pdf = pymupdf.open(c)
    for pagina in pdf: # Le analisi vanno cercate nelle varie pagine
        for chiave in analisi:
            """ search_for restituisce sempre una lista, controllo se la lista non è vuota
                In caso affermativo prendo il primo risultato trovato """
            result = pagina.search_for(chiave)
            if len(result) > 0:
                rect_result = result[0]
                rect_valore = pymupdf.Rect(
                    # Traslo le coordinate del rettangolo per prendere quello con il valore dell'analisi
                    x0=rect_result.x0+169.65,
                    y0=rect_result.y0+0.48,
                    x1=rect_result.x1+180,
                    y1=rect_result.y1+0.26
                )
                valore = pagina.get_textbox(rect_valore).strip() # get_textbox prende il testo contenuto nel rettangolo indicato
                try:
                    valore = float(valore)
                except ValueError:
                    pass
                if product == 'D3336F' and chiave == 'Molybdenum':
                    # da dividere per 10000 perché nell'Excel è in %, ma il primo risultato che trova nel pdf è in ppm
                    valore = valore / 10000
                if product == 'P6072F' and chiave == 'Magnesium':
                    # da moltiplicare per 10000 perché nell'Excel è in ppm, ma il primo risultato che trova nel pdf è in %
                    valore = valore * 10000
                analisi[chiave] = valore
    
    # Parte relativa alla creazione del CoA Excel e all'inserimento dei valori
    file = fr'{path}\CoA serbatoi\{tank}_.xlsx' # File da copiare
    # Prendo la data dal nome di uno dei due CoA e la trasformo da ggmmaa a gg/mm/aaaa
    data = datetime.strptime(c[-10:-4], '%d%m%y')
    data = data.strftime('%d/%m/%Y')
    nomefile = c[-16:-4] # prendo solo il nome del file senza estensione, che è anche il batch
    analisi_series = pd.Series(analisi.values()) # Creo una Series di pandas dal dizionario contenente i valori delle analisi
    shutil.copy(fr'{path}\CoA serbatoi\{tank}_.xlsx', fr'{path}\{serb}_TRASFERIMENTI {product}\{nomefile} VADO.xlsx')
    with pd.ExcelWriter(file,mode='a',if_sheet_exists='overlay',engine='openpyxl') as writer:
        analisi_series.to_excel(writer,sheet_name=product,startrow=18,startcol=3,header=False,index=False)
        # Scrivo data e batch come valori singoli direttamente dalle variabili utilizzando il writer
        workbook = writer.book
        worksheet = writer.sheets[product]
        worksheet["D7"] = data
        worksheet["D13"] = nomefile