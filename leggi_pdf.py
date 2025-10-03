"""
SCRIPT PER CREARE FOGLI DI MARCIA, ETICHETTE E POPOLARE IL BOLLETTONE ANALISI M30B A PARTIRE DAL CERTIFICATO IN pdf
IMPORTA IL FILE anagrafica.py CONTENENTE L'ANAGRAFICA DELLE ANALISI E DEI PRODOTTI

by Simone Barondi
s.barondi@iglom.it
"""

# IMPORTAZIONE MODULI
import anagrafica as an         # importo l'anagrafica (vedi docstring in alto)
import os                       # per cercare i file e per prendere il nome utente di Windows
import shutil                   # per copiare-incollare-rinominare i file
from pathlib import Path        # per prendere il nome del file trovato dal percorso completo
import pymupdf                  # per leggere il certificato pdf ed estrarre i dati
import pandas as pd             # per gestire i dati estratti dal pdf ed esportarli nel file Excel del foglio di marcia
from datetime import datetime   # per gestire le date
import hashlib                  # per vedere se nel bollettone sono già stati inseriti i prodotti (programma C/S già esportato)

# DEFINIZIONE VARIABILI GLOBALI E IMPOSTAZIONE TIMESET LOCALE
# PERCORSO_COA = r"C:\Users\s.barondi\Documents\Python\COA" # --- PER QUANDO TESTO DA CASA
PERCORSO_COA = r"\\vm-cegeka\COA"
# PERCORSO_MAIN = fr"C:\Users\{os.getlogin()}\Documents\Python" # --- PER TESTARE IN LOCALE, ANCHE A LAVORO
PERCORSO_MAIN = r"\\iglomfs\Produzione\FILTRAZIONE\COMPUTER LAB"

#CREAZIONE CLASSI
class CoaNotOk(Exception):

    """ Creo un'eccezione custom nel caso la delivery non esista oppure non sia univoca """
    pass


class FiltroNotOk(Exception):

    """ Creo un'eccezione custom nel caso il filtro non sia 1, 2 o 3 """
    pass


""" Creo la classe Coa sulle cui istanze faccio tutto il processo (cerca pdf, importa analisi, esporta su Excel ecc.)"""
class Coa:

    """ Inizializzo una lista e un dizionario per il recappone per collezionare le istanze create.
        Mi serve per poter capire quali delivery fanno parte di un blenderone e di quante ATB è composto. """
    lista_istanze = []
    dict_recap = {'Delivery': [], 'Batch': [], 'Batchcorto': [], 'Filtro': []}

    """ Creo la funzione recappone con i seguenti scopi:
        1) Replicare il batch per le varie ATB dello stesso blenderone
        2) Popolare il dizionario recappone con i dati delle istanze create
        3) Popolare il bollettone delle analisi M30B con i dati dell'istanza """
    @classmethod
    def recappone(cls):
        # Metto i trattini ai batch del blenderone
        # --- N.B è fatto IN MODO SEQUENZIALE a seconda dell'ordine in cui le delivery compaiono nel programma ---
        cls.df_recap = pd.DataFrame.from_dict(cls.dict_recap)
        cls.df_recap['num_batch'] = cls.df_recap.groupby('Batch').cumcount()
        cls.df_recap['Batch'] = cls.df_recap.apply(
            lambda row: f"{row['Batch']}-{row['num_batch']}" if row['num_batch'] > 0 else row['Batch'],
            axis=1
        )
        cls.df_recap['Batchcorto'] = cls.df_recap.apply(
            lambda row: f"{row['Batchcorto']}-{row['num_batch']}" if row['num_batch'] > 0 else row['Batchcorto'],
            axis=1
        )
        cls.df_recap = cls.df_recap.drop(columns='num_batch')
        aggiunta_bollettone = {'Prodotto ': [], 'Data': [], 'Batch': [], 'Filtro': []}
        bollettone = fr"{PERCORSO_MAIN}\M30B Bollettino d'analisi interno {datetime.now().year}.xlsx"
        cls.df_bollettone = pd.read_excel(bollettone, usecols=(1,2,3,4,5,6), skiprows=(range(0,22)))
        indice_finedati = cls.df_bollettone['Prodotto '].isna().idxmax()
        for i in cls.lista_istanze:
            i.batch = cls.df_recap['Batch'][cls.lista_istanze.index(i)]
            i.batchcorto = cls.df_recap['Batchcorto'][cls.lista_istanze.index(i)]
            i.batch_compresso = i.batch.replace('  / ', '')[6:]
            for j in range(0, 9):
                aggiunta_bollettone['Prodotto '].append(i.filtrato)
                breakpoint()
                aggiunta_bollettone['Data'].append(datetime.strftime(i.data, '%d-%b'))
                aggiunta_bollettone['Batch'].append(i.batchcorto)
                aggiunta_bollettone['Filtro'].append(f"TK{i.filtro}2")
        cls.df_aggiunta = pd.DataFrame(aggiunta_bollettone)
        percorso_log = fr"{PERCORSO_MAIN}\Automatizzazione fogli di marcia\log_bollettone.csv"
        hash_aggiunta = hashlib.md5(cls.df_aggiunta.to_csv(index=False).encode("utf-8")).hexdigest()
        df_log = pd.read_csv(percorso_log, dtype=str, sep=";")
        if hash_aggiunta not in df_log['batch_id'].values:
            with pd.ExcelWriter(bollettone,mode='a',if_sheet_exists='overlay',engine='openpyxl') as writer:
                cls.df_aggiunta.iloc[:, 0:3].to_excel(writer,sheet_name="ANALISI",startrow=indice_finedati+23,startcol=1,header=False,index=False)
                cls.df_aggiunta.loc[:, 'Filtro'].to_excel(writer,sheet_name="ANALISI",startrow=indice_finedati+23,startcol=6,header=False,index=False)
            nuova_riga_log = pd.DataFrame([{'batch_id': hash_aggiunta, 'descrizione': f'Automatizzato in data {datetime.now()}'}])
            df_log = pd.concat([df_log, nuova_riga_log], ignore_index=True)
            df_log.to_csv(percorso_log, index=False, sep=";")
        return cls.df_recap

    """ Nel metodo costruttore sono inserite anche le istruzioni per cercare il file pdf corrispondente e prelevare il prodotto e il nome del file, assegnandoli all'istanza """
    def __init__(self, delivery:str, data:datetime, filtro:int):
        Coa.lista_istanze.append(self) # aggiungo l'istanza alla lista per il recappone
        self.delivery = delivery
        self.data = data
        self.filtro = filtro
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

    """ La seguente funzione processa il certificato pdf con i seguenti passaggi:
        1) apre il certificato pdf trovato e trova le analisi previste per quel prodotto con i rispettivi valori
        2) estrapola nomi analisi e rispettivi valori in un dizionario di liste, una per le analisi e una per i valori
            ** NB: IL DIZIONARIO è FORMATTATO IN QUESTO MODO PER ESSERE GIà PRONTO DA ESPORTARE COME DATAFRAME: LE CHIAVI SARANNO GLI INDICI DELLE COLONNE E GLI ELEMENTI DELLE LISTE SARANNO LE RIGHE **
        3) crea e restituisce un DataFrame di pandas a partire dal dizionario ottenuto
        """
    def processa(self):
        for p in an.prodotti:
            if self.prodotto == p.nome:
                self.istanza_prodotto = p
        coa_pdf = pymupdf.open(self.file)
        risultati = coa_pdf[0].search_for('Batch No.')
        cerca_batch = risultati[0]
        rettangolo_batch = pymupdf.Rect(x0=cerca_batch.x0+11.65, y0=cerca_batch.y0+10.45, x1=cerca_batch.x1+90, y1=cerca_batch.y1+12.45)
        self.batch = coa_pdf[0].get_textbox(rettangolo_batch).strip()
        # Creo le altre versioni del batch (batchcorto: senza tank, batch_compresso: intero senza interpunzioni)
        self.batchcorto = self.batch.split('  / ')[0]
        self.batch_compresso = self.batch.replace('  / ', '')[6:]
        # inizializzo il dizionario che poi andrò a riempire con analisi e relativi valori
        self.dict_analisi = {'ANALISI': [], 'VALORE': []}
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
            self.dict_analisi['ANALISI'].append(analisi)
            self.dict_analisi['VALORE'].append(valore)
        self.df_analisi = pd.DataFrame(self.dict_analisi)
        # popolo il dizionario recappone con i valori dell'istanza
        Coa.dict_recap['Delivery'].append(self.delivery)
        Coa.dict_recap['Batch'].append(self.batch)
        Coa.dict_recap['Batchcorto'].append(self.batchcorto)
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
        fdm_finale = fr"{percorso_fdm}\{self.filtrato} 1-{self.batch_compresso}.xlsx"
        shutil.copy(result, fdm_finale) # devo ancora definire il batch
        self.df_analisi.loc[len(self.df_analisi)] = ['AUTO', 'AUTO'] # per far capire che il CoA è stato generato in automatico
        with pd.ExcelWriter(fdm_finale,mode='a',if_sheet_exists='overlay',engine='openpyxl') as writer:
            worksheet = writer.sheets[self.filtrato]
            worksheet["C5"] = self.batch.replace('  / ', ' / ') # elimino uno spazio di troppo
            worksheet["C7"] = self.data
            worksheet["C8"] = self.filtro
            self.df_analisi['VALORE'].to_excel(writer,sheet_name=self.filtrato,startrow=self.istanza_prodotto.riga,startcol=2,header=False,index=False)

    def __str__(self):
        return (
            f"\n[Certificato di analisi pdf]\n"
            f"Delivery: {self.delivery}\n"
            f"Nome prodotto: {self.prodotto}\n"
            f"Nome del file: {self.nomefile}\n"
        )
    
    def __repr__(self):
        return f"Coa({self.delivery})"

# DEFINISCO LA FUNZIONE DI INPUT, DA CHIAMARE QUANDO QUESTO SCRIPT VIENE ESEGUITO DIRETTAMENTE
def inserisci():
    tasks = []
    while True:
        chiedi_filtro = True
        chiedi_data = True
        delivery = input("Inserisci una delivery o scrivi OK per confermare: ").lower()
        if delivery == 'ok':
            return tasks
        else:
            while chiedi_filtro:
                filtro = int(input(f"Inserisci il filtro per la delivery {delivery} (1, 2, 3): "))
                if filtro == 1 or filtro == 2 or filtro == 3:
                    while chiedi_data:
                        data_in = input(f"Inserisci la data per la delivery {delivery} nel formato gg/mm/aa: ")
                        try:
                            data = datetime.strptime(data_in, '%d/%m/%y')
                            tasks.append({'delivery': delivery, 'filtro': filtro, 'data': data})
                            chiedi_filtro = False
                            chiedi_data = False
                        except ValueError:
                            print("Devi inserire una data nel formato gg/mm/aa")
                else:
                    print("Devi inserire 1, 2 o 3 come valore per il filtro")


""" Se eseguo questo script direttamente chiedo manualmente le delivery.
    Se viene importato come modulo le prendo in automatico dal programma carichi scarichi."""
if __name__ == "__main__":
    tasks = inserisci()
    print(tasks)
    for t in range(0, len(tasks)):
        istanza = Coa(tasks[t]['delivery'], tasks[t]['data'], tasks[t]['filtro'])
        istanza.processa()
    recappone = Coa.recappone()
    lista_istanze = Coa.lista_istanze
    for i in lista_istanze:
        i.crea_fdm()