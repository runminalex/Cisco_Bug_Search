import requests
import json

def get_token():
    url = 'https://cloudsso.cisco.com/as/token.oauth2'
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    payload = {
        'grant_type': 'client_credentials',
        'client_id': 'bcqs6ejwhj9tufavrknxnm4u',
        'client_secret': 'Zs3nyzjg7MP2kQgszcWcycBk'
    }
    r = requests.post(url, headers=headers, data=payload)
    result = json.loads(r.content)
    # result = r.content['access_token']
    # print(result)
    # print(result['access_token'])
    return result["token_type"] +' ' + result["access_token"]