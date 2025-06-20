import os 
import requests
import pandas as pd
import openpyxl
import time 

script_folder = os.path.dirname(os.path.abspath(__file__))

archivo_excel = os.path.join(script_folder, "SRLAB.xlsx")
df = pd.read_excel(archivo_excel)

filtered_df = df[df["IMPUESTOS"] == 0].reset_index(drop=True)
print(filtered_df)

for index, row in filtered_df.iterrows():
    number = row["NUMERO_ENVIO"]
    