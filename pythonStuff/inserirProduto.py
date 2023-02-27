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

payload = {"produto": {"codigo": "22","nome": "hhh","unidadeMedida": "UN", "preco": 5}}

response = requests.request("POST", url, json=payload, headers=headers)
print(response.text)

    