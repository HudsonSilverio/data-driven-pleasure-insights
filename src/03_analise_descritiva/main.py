# poetry run main.py

import pandas as pd

# Use r"" para evitar problema com as barras invertidas
caminho = "C:/Users/Administrador/Desktop/Sourcer of Plasure.xlsx"

# Ler o arquivo Excel
df = pd.read_excel(caminho)

print(df.dtypes)