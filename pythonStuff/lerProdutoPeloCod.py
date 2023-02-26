import pandas as pd
import decimal
import requests
import math
import json

url = "https://jpautomacao-getcard02.getcard.uniplusweb.com/public-api/v1/produtos"


headers = {
    "Content-Type": "application/json",
    "Authorization": "Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJhdWQiOlsidW5pcGx1c3dlYiJdLCJzY29wZSI6WyJwZHYiLCJqb2JzLXBkdiIsIm1vYmlsZSIsInNob3AiLCJwdWJsaWMtYXBpIl0sImV4cCI6MTY3NzQyNDU5NywianRpIjoiNTk4YTlkYWMtMDFmNS00ZmIyLTkzYTktYTc2YTFmM2Y3NmI4IiwidGVuYW50IjpudWxsLCJjbGllbnRfaWQiOiIwMjc0MjM2MTAwMDEzOSJ9.NijXZ5BVJHVPun-mIHmXy3da7pDG88MFOB9VnJOGG-Y"
}

payload = {"produto": {"nome": "PFG132"}}
response = requests.request("GET", url, json=payload, headers=headers)
print(json.loads(response.text)[0]['codigo'])


# {“{“produto”: {“codigo”: “999000”, “nome”:”PRODUTO TESTE”, “preco”: 10.23}}

 



