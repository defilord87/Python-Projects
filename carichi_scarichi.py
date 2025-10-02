import pandas as pd
import interfaccia # importo le funzioni per l'interfaccia grafica (da togliere un giorno?)
import leggi_pdf as lp # importo la funzione per processare i CoA e fare tutto il giro

file = interfaccia.finestra()

leggi = pd.read_excel(file, usecols=(0, 1, 5), skiprows=(0,1), sheet_name='PROGRAMMA UNICO')
indice = leggi.loc[leggi['Delivery'] == 'Cliente'].index[0]
data = leggi.iloc[indice-2, 0]
scarichi = leggi.iloc[:indice-3, :].reset_index(drop=True)
scarichi["Material Description"] = scarichi["Material Description"].str.replace("Infineum ", "", regex=False)
scarichi = scarichi[scarichi["Material Description"].isin(lp.an.lista_prodotti)]
dict_scarichi = scarichi.to_dict()
numero_scarichi = len(dict_scarichi['Delivery'])
lista = list(dict_scarichi.values())
for i in range(numero_scarichi):
    delivery = str(dict_scarichi['Delivery'][i])
    tank = dict_scarichi['Serbatoio'][i]
    filtro = int(tank[2])
    istanza = lp.Coa(delivery, data, filtro)
    istanza.processa()
recappone = lp.Coa.recappone()
lista_istanze = lp.Coa.lista_istanze
for i in lista_istanze:
    i.crea_fdm()