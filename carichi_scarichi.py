import pandas as pd
import interfaccia # importo le funzioni per l'interfaccia grafica (da togliere un giorno?)
import leggi_pdf as lp # importo la funzione per processare i CoA e fare tutto il giro
import re # regular expression per prendere la data dalle note, es. "1° filtrazione del 17/10"

""" Avvio l'interfaccia per aprire il programma carichi/scarichi e ritorno il path del file """
file = interfaccia.finestra()

""" Importo il contenuto del file Excel in un DataFrame, saltando le prime due righe di intestazione e
    prendendo solo le colonne di prodotto, delivery e serbatoio (-> filtro) """
leggi = pd.read_excel(file, usecols=(0, 1, 4, 5), skiprows=(0,1), sheet_name='PROGRAMMA UNICO')
# Prendo l'indice della riga dove iniziano i carichi per tagliarli via e tenere solo gli scarichi:
indice = leggi.loc[leggi['Delivery'] == 'Cliente'].index[0]
# Prendo la data:
data = leggi.iloc[indice-2, 1]
# Taglio le ultime tre righe per pulire il DataFrame
scarichi = leggi.iloc[:indice-3, :].reset_index(drop=True)
# Riformatto la colonna del prodotto 'Infineum XXXXX' -> 'XXXXX' in modo da cercarlo nella lista prodotti in anagrafica:
scarichi["Material Description"] = scarichi["Material Description"].str.replace("Infineum ", "", regex=False)
# Filtro solo i prodotti in anagrafica in modo da scartare l'olio SN150, i prodotti dalla Francia ecc.:
scarichi = scarichi[scarichi["Material Description"].isin(lp.an.lista_prodotti)]
# Metto il df in ordine di delivery crescente per gestire bene i blenderoni
scarichi = scarichi.sort_values(by="Delivery").reset_index(drop=True)
# Esporto il DataFrame in un dizionario per prelevare delivery e filtro da mandare al costruttore dell'istanza CoA del prodotto:
dict_scarichi = scarichi.to_dict()
print("Programma letto correttamente, inizio a prelevare i dati.")

""" Prendo il numero degli scarichi e itero su tutti questi prendendo ogni volta delivery e filtro e inizializzando l'istanza.
    Dopo averla inizializzata chiamo il metodo processa() dell'istanza per importare le analisi dal CoA pdf corrispondente """
numero_scarichi = len(dict_scarichi['Delivery'])
for i in range(numero_scarichi):
    delivery = str(dict_scarichi['Delivery'][i])
    tank = dict_scarichi['Serbatoio'][i]
    nota = dict_scarichi['Note'][i]
    if nota:
        pattern = r'\b(\d{1,2})/(\d{1,2})\b'
        match = re.search(pattern, nota)
        if match:    
            giorno, mese = map(int, match.groups())
            data = data.replace(day=giorno, month=mese)
    filtro = int(tank[2])
    istanza = lp.Coa(delivery, data, filtro) # creo l'istanza della classe Coa
    istanza.processa()
print("Certificati letti correttamente.")

""" Chiamo il metodo di classe recappone() per individuare eventuali blenderoni e popolare il bollettone M30B,
    quindi creo il foglio di marcia per ogni scarico (iterando nella lista delle istanze) """
recappone = lp.Coa.recappone()
for i in lp.Coa.lista_istanze:
    i.crea_fdm()
print("Fogli di marcia creati, puoi chiudere la finestra.")