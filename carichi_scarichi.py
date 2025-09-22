import pandas as pd
import leggi_pdf as lp # importo la funzione per processare i CoA e fare tutto il giro

file = r"C:\Users\s.barondi\Documents\Python\CREAZIONE CARICHI SCARICHI - 22.09.25_MOD.xlsx"

nomi_colonne = ['DELIVERY', 'TANK']
leggi = pd.read_excel(file, usecols=(1, 5), header=None, sheet_name='PROGRAMMA UNICO', names=nomi_colonne)
data = leggi['DELIVERY'][0]
indice = leggi.loc[leggi['DELIVERY'] == 'Cliente'].index[0]
scarichi = leggi.iloc[3:indice-2, :].reset_index(drop=True)
dict_scarichi = scarichi.to_dict()
numero_scarichi = len(dict_scarichi['DELIVERY'])
lista = list(dict_scarichi.values())
istanze = []
conta_batch = {}
for i in range(numero_scarichi):
    delivery = str(dict_scarichi['DELIVERY'][i])
    tank = dict_scarichi['TANK'][i]
    filtro = tank[2]
    istanza = lp.creazione(delivery, data, filtro)
    istanza.processa()
    if istanza.batch in conta_batch.keys():
        conta_batch[istanza.batch] += 1
    else:
        conta_batch[istanza.batch] = 1
        istanze.append(istanza)
breakpoint()