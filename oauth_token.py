import requests
import json
from oauth_cred import get_cred

def get_token():
    cred=get_cred()
    url = 'https://cloudsso.cisco.com/as/token.oauth2'
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    payload = {
        'grant_type': 'client_credentials',
        'client_id': cred['client_id'],
        'client_secret': cred['client_secret']
    }
    r = requests.post(url, headers=headers, data=payload)
    result = json.loads(r.content)
    # result = r.content['access_token']
    # print(result)
    # print(result['access_token'])
    return result["token_type"] +' ' + result["access_token"]