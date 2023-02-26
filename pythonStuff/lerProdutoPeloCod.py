import pandas as pd
import decimal
import requests
import math
import json
from geraToken import myNewToken

url = "https://jpautomacao-getcard02.getcard.uniplusweb.com/public-api/v1/produtos"


headers = {
    "Content-Type": "application/json",
    "Authorization": "Bearer " + myNewToken() 
}

payload = {"produto": {"nome": "PFG132"}}
response = requests.request("GET", url, json=payload, headers=headers)

print(json.loads(response.text)[0]['codigo'])
