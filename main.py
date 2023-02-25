import pandas as pd
import decimal
import requests
import math


df = pd.read_excel("C:\GitHub\myxlsm\inventario_externo.xlsx")
url = "https://jpautomacao-getcard02.getcard.uniplusweb.com/public-api/v1/produtos"


headers = {
    "Content-Type": "application/json",
    "Authorization": "Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJhdWQiOlsidW5pcGx1c3dlYiJdLCJzY29wZSI6WyJwZHYiLCJqb2JzLXBkdiIsIm1vYmlsZSIsInNob3AiLCJwdWJsaWMtYXBpIl0sImV4cCI6MTY3NzM2MjY1MSwianRpIjoiOGZlZjA5ZDktYzRiYi00MjVjLTg3ODktNWU3MjRhZmEzYTI1IiwidGVuYW50IjpudWxsLCJjbGllbnRfaWQiOiIwMjc0MjM2MTAwMDEzOSJ9.JveL_e3mI93B7G_4yNXFe78DYdU2MUlrRkLuNn1TTvA"
}

for index, row in df.iterrows():
    try:
        decimal_value = decimal.Decimal(str(row['VENDA']))
        rounded_value = round(decimal_value, 3)
        precoDeVenda = float(rounded_value)

        decimal_value = decimal.Decimal(str(row['LOCAÇÃO']))
        rounded_value = round(decimal_value, 3)
        precoDeLocacao = float(rounded_value)

        if precoDeVenda == 0 or precoDeVenda == "" or math.isnan(precoDeVenda):
            precoDeVenda = 0.01

        if precoDeLocacao == 0 or precoDeLocacao == "" or math.isnan(precoDeLocacao):
            precoDeLocacao = 0.01
        
        payload = {"produto": {
            "codigo": str(row['ID']),
            "nome": row['CÓDIGO'],
            "preco": precoDeVenda,
            "precoPauta1": precoDeLocacao,
            "unidadeMedida": "UN"
        }}

        response = requests.request("POST", url, json=payload, headers=headers)

        if response.status_code != 200:
            print("erro neste lancamento")
            print(response.text)
            print(row['ID'])
            print(row['CÓDIGO'])
            print(precoDeVenda)
            print(precoDeLocacao)
            break

    except Exception as e:
        print(e)
        print("erro neste lancamento")
        print(row['ID'])
        print(row['CÓDIGO'])
        print(precoDeVenda)
        print(precoDeLocacao)
        
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

