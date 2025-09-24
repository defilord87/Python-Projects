import pandas as pd
import leggi_pdf as lp # importo la funzione per processare i CoA e fare tutto il giro

file = r"C:\Users\s.barondi\Documents\Python\CREAZIONE CARICHI SCARICHI - 22.09.25_MOD.xlsx"

leggi = pd.read_excel(file, usecols=(1, 5), skiprows=(0,1), sheet_name='PROGRAMMA UNICO')
data = leggi['Delivery'][0]
indice = leggi.loc[leggi['Delivery'] == 'Cliente'].index[0]
scarichi = leggi.iloc[:indice-2, :].reset_index(drop=True)
dict_scarichi = scarichi.to_dict()
numero_scarichi = len(dict_scarichi['Delivery'])
lista = list(dict_scarichi.values())
for i in range(numero_scarichi):
    delivery = str(dict_scarichi['Delivery'][i])
    tank = dict_scarichi['Serbatoio'][i]
    filtro = tank[2]
    istanza = lp.creazione(delivery, data, filtro)
    istanza.processa()
recappone = lp.Coa.recappone()
lista_istanze = lp.Coa.lista_istanze
for i in lista_istanze:
    i.crea_fdm()