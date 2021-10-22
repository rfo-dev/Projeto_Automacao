import pandas as pd
import tabula
from tabula.io import read_pdf

df = tabula.read_pdf(r"C:\Users\rafael.oliveira\Documents\Projeto_Automacao\FOLHA_REF_03.2020_CAMPEIRO_MATRIZ.pdf")
print(df)