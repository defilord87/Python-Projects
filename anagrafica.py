"""
MODULO CONTENENTE L'ANAGRAFICA DELLE ANALISI E DEI PRODOTTI
DA IMPORTARE COME MODULO NEGLI ALTRI SCRIPT

by Simone Barondi
s.barondi@iglom.it
"""

# Creo la lista con tutte le analisi esistenti tra i vari prodotti, ordinata alfabeticamente. Accanto ci sono i vari indici riportati come commenti per comodità.
ANALISI = [
    'Appearance', # 0
    'Base Number', # 1
    'Boron', # 2
    'Calcium', # 3
    'Chlorine', # 4
    'Color, diluted', # 5
    'Density', # 6
    'Flash Point', # 7
    'Infrared', # 8
    'Kinematic Viscosity @ 100 C', # 9
    'Magnesium', # 10
    'Molybdenum', # 11
    'Nitrogen', # 12
    'Phosphorus', # 13
    'Silicon', # 14
    'Sulfated Ash', # 15
    'Sulfur', # 16
    'Zinc', # 17
    'Water', # 18
]

class Prodotto:

    def __init__(self, nome:str, classe:str, *, blendable:bool=False, riga:int, analisi:tuple):
        """ Se la classe del prodotto indicata non è 'salicilato' o 'solfonato' il costruttore restituisce un ValueError """
        self.nome = nome
        if classe == 'solfonato' or classe == 'salicilato':
            self.classe = classe
        else:
            raise ValueError('La classe del prodotto deve essere solfonato o salicilato!')
        self.blendable = blendable
        self.riga = riga
        self.analisi = analisi
        self.lista_analisi = []
        for a in analisi:
            self.lista_analisi.append(ANALISI[a])

    def __repr__(self):
        return f"Prodotto({self.nome}, {self.classe}, blendable={self.blendable}, analisi={self.analisi})"
    
    def __str__(self):
        return (
            f"Nome prodotto: {self.nome}\n"
            f"Classe prodotto: {self.classe}\n"
            f"Può essere un blenderone: {self.blendable}\n"
            f"Analisi: {self.lista_analisi}\n"
        )

prodotti = [
    Prodotto('D3336C', 'salicilato', blendable=True, riga=71, analisi=(0, 1, 3, 6, 9, 10, 11, 12, 13, 17)),
    Prodotto('P6072C', 'salicilato', blendable=False, riga=72, analisi=(0, 6, 7, 3, 2, 13, 16, 17, 12, 10, 1, 15, 9, 18)),
    Prodotto('P6571C', 'salicilato', blendable=False, riga=72, analisi=(0, 6, 7, 3, 2, 13, 16, 17, 12, 4, 1, 5, 9)),
    Prodotto('P6052C', 'salicilato', blendable=False, riga=71, analisi=(0, 6, 7, 3, 10, 13, 17, 12, 1, 16, 2, 14, 9))
]