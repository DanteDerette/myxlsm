import pandas as pd
import decimal
import requests
import math


# df = pd.read_excel("C:\GitHub\myxlsm\inventario_externo.xlsx")
url = "https://jpautomacao-getcard02.getcard.uniplusweb.com/public-api/v1/davs"


headers = {
    "Content-Type": "application/json",
    "Authorization": "Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJhdWQiOlsidW5pcGx1c3dlYiJdLCJzY29wZSI6WyJwZHYiLCJqb2JzLXBkdiIsIm1vYmlsZSIsInNob3AiLCJwdWJsaWMtYXBpIl0sImV4cCI6MTY3NzQxNzQ0MCwianRpIjoiYTFjYzA0ZmUtY2Y2Mi00ZmRiLWJkY2YtODBhMzdlNzE3NDExIiwidGVuYW50IjpudWxsLCJjbGllbnRfaWQiOiIwMjc0MjM2MTAwMDEzOSJ9.mHIqJMcdgZzo_91chjpkbs69opegoEbFT94jIp8_azc"
}

        
payload = {
        "dav": 
            {"codigo": "1",
            "tipoDocumento": 6,
            "data":"2023-01-01",
            "itens":
                [
                    {
                        "produto":"1008",
                        "quantidade":1,
                        "precoUnitario":12.34
                    }
                ]
            }
        }



response = requests.request("POST", url, json=payload, headers=headers)

if response.status_code != 200:
    print("erro neste lancamento")
    print(response.text)
    




