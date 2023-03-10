import pandas as pd
import decimal
import requests
import math
from geraToken import myNewToken

df = pd.read_excel("C:\GitHub\myxlsm\inventario_externo.xlsx")
url = "https://jpautomacao-getcard02.getcard.uniplusweb.com/public-api/v1/produtos"


headers = {
    "Content-Type": "application/json",
    "Authorization": "Bearer" + myNewToken() 
}

for index, row in df.iterrows():
    try:
        decimal_value = decimal.Decimal(str(row['Venda Unitário']))
        rounded_value = round(decimal_value, 3)
        precoDeVenda = float(rounded_value)

        if precoDeVenda == 0 or precoDeVenda == "" or math.isnan(precoDeVenda):
            precoDeVenda = 0.01
        
        payload = {"produto": {
            "codigo": str(row['id']),
            "nome": row['Código'],
            "preco": precoDeVenda,
            "nomeEcf": str(row['Especificação']),
            "observacao": str(row['Especificação']),
            "unidadeMedida": "UN"
        }}

        response = requests.request("POST", url, json=payload, headers=headers)

        if response.status_code != 200:
            print("erro neste lancamento")
            print(response.text)
            print(row['ID'])
            print(row['CÓDIGO'])
            print(precoDeVenda)
            
            break

    except Exception as e:
        print(e)
        print("erro neste lancamento")
        print(row['ID'])
        print(row['CÓDIGO'])
        print(precoDeVenda)
        
        
        break    


def is_number(s):
    try:
        float(s)
        return float(s)
    except ValueError:
        return 0.01
    
def is_integer(s):
    try:
        int(s)
        return int(s)
    except ValueError:
        return 0.01

