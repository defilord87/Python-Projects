import pandas as pd

file = r"D:\Documenti\Python\CREAZIONE CARICHI SCARICHI - 22.09.25.xlsx"

nomi_colonne = ['DELIVERY', 'TANK']
leggi = pd.read_excel(file, usecols=(1, 5), skiprows=0, sheet_name='PROGRAMMA UNICO', names=nomi_colonne)
indice = leggi.loc[leggi['DELIVERY'] == 'Cliente'].index[0]
scarichi = leggi.iloc[2:indice-3, :].reset_index(drop=True)
dict_scarichi = scarichi.to_dict()
numero_scarichi = len(dict_scarichi['DELIVERY'])
lista = list(dict_scarichi.values())
for i in range(numero_scarichi):
    delivery = dict_scarichi['DELIVERY'][i]
    tank = dict_scarichi['TANK'][i]
    filtro = tank[2]