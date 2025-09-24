"""
SCRIPT PER CREARE FOGLI DI MARCIA, ETICHETTE E POPOLARE IL BOLLETTONE ANALISI M30B A PARTIRE DAL CERTIFICATO IN pdf
IMPORTA IL FILE anagrafica.py CONTENENTE L'ANAGRAFICA DELLE ANALISI E DEI PRODOTTI

by Simone Barondi
s.barondi@iglom.it
"""

# IMPORTAZIONE MODULI
import anagrafica as an         # importo l'anagrafica (vedi docstring in alto)
import os                       # per cercare i file
import shutil                   # per copiare-incollare-rinominare i file
from pathlib import Path        # per prendere il nome del file trovato dal percorso completo
import pymupdf                  # per leggere il certificato pdf ed estrarre i dati
import pandas as pd             # per gestire i dati estratti dal pdf ed esportarli nel file Excel del foglio di marcia
from datetime import datetime   # per gestire le date

# DEFINIZIONE VARIABILI GLOBALI
PERCORSO_COA = r"\\vm-cegeka\COA"
PERCORSO_MAIN = r"C:\Users\s.barondi\Documents\Python" # --- STO TESTANDO, PER ORA LAVORO IN LOCALE
# PERCORSO_MAIN = r"\\iglomfs\Produzione\FILTRAZIONE\COMPUTER LAB"
DICT_ANALISI = {'ANALISI': [], 'VALORE': []}    # inizializzo il dizionario che poi andrò a riempire con analisi e relativi valori

#CREAZIONE CLASSI
class CoaNotOk(Exception):

    """ Creo un'eccezione custom nel caso la delivery inserita non esista oppure non sia univoca """
    pass


class FiltroNotOk(Exception):

    """ Creo un'eccezione custom nel caso il filtro inserito non sia 1, 2 o 3 """
    pass


class Coa:

    """ Inizializzo una lista e un dizionario per il recappone per collezionare le istanze create.
        Mi serve per poter capire quali delivery fanno parte di un blenderone e di quante ATB è composto. """
    lista_istanze = []
    dict_recap = {'Delivery': [], 'Batch': [], 'Filtro': []}

    @classmethod
    def recappone(cls):
        cls.df_recap = pd.DataFrame.from_dict(cls.dict_recap)
        cls.df_recap['num_batch'] = cls.df_recap.groupby('Batch').cumcount()
        cls.df_recap['Batch'] = cls.df_recap.apply(
            lambda row: f"{row['Batch']}-{row['num_batch']}" if row['num_batch'] > 0 else row['Batch'],
            axis=1
        )
        cls.df_recap = cls.df_recap.drop(columns='num_batch')
        for i in cls.lista_istanze:
            i.batch = cls.df_recap['Batch'][cls.lista_istanze.index(i)]
        return cls.df_recap

    """ Nel metodo costruttore sono inserite anche le istruzioni per cercare il file pdf corrispondente e prelevare il prodotto e il nome del file, assegnandoli all'istanza """
    def __init__(self, delivery:str):
        Coa.lista_istanze.append(self) # aggiungo l'istanza alla lista per il recappone
        self.delivery = delivery
        result = [] # Inizializza la lista vuota dei risultati della ricerca
        for trova in os.scandir(PERCORSO_COA):
            if trova.is_file() and self.delivery in trova.name and trova.name.endswith('.pdf'):
                result.append(trova.path)
        if not result:
            raise CoaNotOk(f"La delivery {self.delivery} non esiste!")
        elif len(result) != 1:
            raise CoaNotOk(f"La delivery {self.delivery} non è univoca!\nHo trovato {len(result)} risultati, ma dovrei trovare un solo CoA. Verifica di avere inserito il numero di delivery per intero.")
        self.file = result[0]
        self.nomefile = Path(self.file).name
        self.prodotto = self.nomefile[16:22]
        self.filtrato = self.prodotto.replace('C', 'F')
        # inizializzo alcune variabili che verranno poi definite nella funzione di creazione creazione()
        self.data = ''
        self.filtro = ''

    """ La seguente funzione cerca il prodotto nel modulo anagrafica.py; se non lo trova (ovvero se self.istanza_prodotto rimane una stringa vuota) genera un errore.
        Questo check è fondamentale perché nella cartella dei CoA in vm-cegeka ci sono anche i prodotti per infustamento e si potrebbe selezionare uno di quelli per errore. """
    def check_prodotto(self):
        self.istanza_prodotto = '' # inizializzo la variabile
        for p in an.prodotti:
            if self.prodotto == p.nome:
                self.istanza_prodotto = p
                self.blendable = p.blendable # per capire se può essere un blenderone oppure no in base al prodotto
                self.riga = p.riga
                self.blend = False # lo imposto False per default, poi da input chiederò se è effettivamente un blenderone oppure no
                self.contabatch = 1 # valore di default, serve per contare quanti scarichi fanno parte di un blenderone
        if self.istanza_prodotto == '':
            raise CoaNotOk(f"Il prodotto {self.prodotto} non è presente in anagrafica!")

    """ La seguente funzione processa il certificato pdf con i seguenti passaggi:
        1) apre il certificato pdf trovato e trova le analisi previste per quel prodotto con i rispettivi valori
        2) estrapola nomi analisi e rispettivi valori in un dizionario di liste, una per le analisi e una per i valori
            ** NB: IL DIZIONARIO è FORMATTATO IN QUESTO MODO PER ESSERE GIà PRONTO DA ESPORTARE COME DATAFRAME: LE CHIAVI SARANNO GLI INDICI DELLE COLONNE E GLI ELEMENTI DELLE LISTE SARANNO LE RIGHE **
        3) crea e restituisce un DataFrame di pandas a partire dal dizionario ottenuto
        """
    def processa(self):
        coa_pdf = pymupdf.open(self.file)
        risultati = coa_pdf[0].search_for('Batch No.')
        cerca_batch = risultati[0]
        rettangolo_batch = pymupdf.Rect(x0=cerca_batch.x0+11.65, y0=cerca_batch.y0+10.45, x1=cerca_batch.x1+90, y1=cerca_batch.y1+12.45)
        self.batch = coa_pdf[0].get_textbox(rettangolo_batch).strip()
        # cerco i valori delle analisi
        for analisi in self.istanza_prodotto.lista_analisi:
            pagina = 0
            risultati = coa_pdf.search_page_for(pagina, analisi)
            if risultati == []:
                pagina = 1
                risultati = coa_pdf.search_page_for(pagina, analisi)
            try:
                testo = risultati[0]
                rettangolo_valore = pymupdf.Rect(x0=testo.x0+252, y0=testo.y0+12, x1=testo.x1+250, y1=testo.y1+12)
                valore = coa_pdf[pagina].get_textbox(rettangolo_valore).strip()
                # Faccio un po' di pulizia togliendo i caratteri a capo:
                valore = valore.split('\n')
                valore = valore[0]
            except IndexError:
                valore = ''
            # Metto il punto ai valori numerici e li converto da stringhe a float:
            try:
                valore = valore.replace(',', '.')
                valore = float(valore)
            except ValueError:
                pass
            except AttributeError:
                pass
            # Popolo il dizionario delle analisi che sarà poi esportato come DataFrame:
            DICT_ANALISI['ANALISI'].append(analisi)
            DICT_ANALISI['VALORE'].append(valore)
        self.df_analisi = pd.DataFrame(DICT_ANALISI)
        # popolo il dizionario recappone con i valori dell'istanza
        Coa.dict_recap['Delivery'].append(self.delivery)
        Coa.dict_recap['Batch'].append(self.batch)
        Coa.dict_recap['Filtro'].append(self.filtro)
        return self.df_analisi
    
    """ La seguente funzione lavora sul foglio di marcia con i seguenti passaggi:
        1) crea il FdM a partire dal foglio vergine
        2) lo copia-incolla-rinomina nella cartella principale dei FdM
        3) lo popola con i dati del batch e i valori delle analisi
        4) lo replica più volte nel caso si tratti di un blenderone """
    def crea_fdm(self):
        percorso_fdm = PERCORSO_MAIN + r"\Fogli di marcia"
        for trova in os.scandir(fr"{percorso_fdm}\Vergini 2023"):
            if trova.is_file() and self.filtrato in trova.name and trova.name.endswith('.xlsx'):
                result = trova.path
        # Creo le altre versioni del batch (batchcorto: senza tank, batch_compresso: intero senza interpunzioni)
        self.batchcorto = self.batch.split('  / ')[0]
        self.batch_compresso = self.batch.replace('  / ', '')[6:]
        fdm_finale = fr"{percorso_fdm}\{self.filtrato} 1-{self.batch_compresso}.xlsx"
        shutil.copy(result, fdm_finale) # devo ancora definire il batch
        self.df_analisi.loc[len(self.df_analisi)] = ['AUTO', 'AUTO'] # per far capire che il CoA è stato generato in automatico
        with pd.ExcelWriter(fdm_finale,mode='a',if_sheet_exists='overlay',engine='openpyxl') as writer:
            worksheet = writer.sheets[self.filtrato]
            worksheet["C5"] = self.batch.replace('  / ', ' / ') # elimino uno spazio di troppo
            worksheet["C7"] = self.data
            worksheet["C8"] = self.filtro
            self.df_analisi['VALORE'].to_excel(writer,sheet_name=self.filtrato,startrow=self.riga,startcol=2,header=False,index=False)

    def __str__(self):
        return (
            f"\n[Certificato di analisi pdf]\n"
            f"Delivery: {self.delivery}\n"
            f"Nome prodotto: {self.prodotto}\n"
            f"Nome del file: {self.nomefile}\n"
            f"Blenderone: {self.blend}\n"
        )
    
    def __repr__(self):
        return f"Coa({self.delivery})"

# DEFINISCO LA FUNZIONE DI CREAZIONE DELL'ISTANZA DI CLASSE, DA CHIAMARE IMPORTANDO QUESTO SCRIPT COME MODULO
def creazione(delivery:str, data:datetime, filtro:int):
    # Creo l'istanza di classe e faccio i controlli sul prodotto (che sia nell'anagrafica e che possa avere un blenderone)
    certificato = Coa(delivery)
    certificato.check_prodotto()
    certificato.data = data
    certificato.filtro = filtro
    return certificato