import pandas as pd
import decimal
import requests
import math
from geraToken import myNewToken

df = pd.read_excel("C:\GitHub\myxlsm\inventario_externo.xlsx")
url = "https://jpautomacao-getcard02.getcard.uniplusweb.com/public-api/v1/produtos"


headers = {
    "Content-Type": "application/json",
    "Authorization": "Bearer " + myNewToken() 
}

for index, row in df.iterrows():
    try:
        
        url = "https://jpautomacao-getcard02.getcard.uniplusweb.com/public-api/v1/produtos/" + str(row['id'])
        
        response = requests.request("DELETE", url, headers=headers)

        if response.status_code != 200:
            print("erro neste lancamento")
            print(response.text)
            print((row['id']))
            break

    except Exception as e:
        print(e)
        print((row['id']))
        print("erro neste lancamento")

        
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

