
import requests
import json

# curl -k -X POST -H "Authorization: Basic MDI3NDIzNjEwMDAxMzk6YjcxN2E4NzYtNWYwZC00NzMxLTlmNjgtOGMzZWE0ZjU2Yzkw" -H "Content-Type: application/x-www-form-urlencoded" -d "grant_type=client_credentials" "https://jpautomacao-getcard02.getcard.uniplusweb.com/oauth/token"

def myNewToken():
    headers = "Basic MDI3NDIzNjEwMDAxMzk6YjcxN2E4NzYtNWYwZC00NzMxLTlmNjgtOGMzZWE0ZjU2Yzkw" 
    ContentType = "application/x-www-form-urlencoded"
    url = "https://jpautomacao-getcard02.getcard.uniplusweb.com/oauth/token"
    data = "grant_type=client_credentials"

    response = requests.post(url, headers={"Authorization":headers, "Content-Type": "application/x-www-form-urlencoded"}, data=data)

    access_token = str(json.loads(response.text)["access_token"])
    return access_token