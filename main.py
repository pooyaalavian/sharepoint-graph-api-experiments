import logging
logging.basicConfig(level=logging.WARN)
logging.info("App Started")

import dotenv
dotenv.load_dotenv()

import os 
from azure.identity import ClientSecretCredential
from azure.identity._credentials.certificate import load_pkcs12_certificate
from azure.core.exceptions import AzureError
import requests
import msal
import json

SHAREPOINT_BASE = os.getenv('SHAREPOINT_BASE')
assert SHAREPOINT_BASE, "SHAREPOINT_BASE is not set"
logging.info(f"SHAREPOINT_BASE: {SHAREPOINT_BASE}")

SHAREPOINT_ADMIN_BASE = os.getenv('SHAREPOINT_ADMIN_BASE')
assert SHAREPOINT_ADMIN_BASE, "SHAREPOINT_ADMIN_BASE is not set"
logging.info(f"SHAREPOINT_ADMIN_BASE: {SHAREPOINT_ADMIN_BASE}")

SHAREPOINT_SITE = os.getenv('SHAREPOINT_SITE')
logging.info(f"SHAREPOINT_SITE: {SHAREPOINT_SITE}")

CLIENT_ID = os.getenv('CLIENT_ID')
assert CLIENT_ID, "CLIENT_ID is not set"
logging.info(f"CLIENT_ID: {CLIENT_ID}")

TENANT_ID = os.getenv('TENANT_ID')
assert TENANT_ID, "TENANT_ID is not set"
logging.info(f"TENANT_ID: {TENANT_ID}")

CLIENT_SECRET = os.getenv('CLIENT_SECRET')
assert CLIENT_SECRET, "CLIENT_SECRET is not set"
logging.info(f"CLIENT_SECRET: {CLIENT_SECRET[0:5]}***")

CLIENT_CERT_PATH = os.getenv('CLIENT_CERT_PATH')
assert CLIENT_CERT_PATH, "CLIENT_CERT_PATH is not set"
logging.info(f"CLIENT_CERT_PATH: {CLIENT_CERT_PATH}")

CLIENT_CERT_THUMBPRINT = os.getenv('CLIENT_CERT_THUMBPRINT')
assert CLIENT_CERT_THUMBPRINT, "CLIENT_CERT_THUMBPRINT is not set"
logging.info(f"CLIENT_CERT_THUMBPRINT: {CLIENT_CERT_THUMBPRINT[0:5]}***")

def get_access_token_secret(scopes):
    try:
        credential = ClientSecretCredential(
            tenant_id=os.getenv('TENANT_ID'), 
            client_id=os.getenv('CLIENT_ID'), 
            client_secret=os.getenv('CLIENT_SECRET'))
        token = credential.get_token(scopes)
        return token.token
    except AzureError as ex:
        print(f"An error occurred: {ex}")

def get_access_token_cert(scopes):
    cert = load_pkcs12_certificate(open(os.getenv('CLIENT_CERT_PATH'),'rb').read(), )
    key = str(cert.pem_bytes, 'utf-8')
    key = key.split('-----BEGIN CERTIFICATE-----')[0]
    app = msal.ConfidentialClientApplication(
        os.getenv('CLIENT_ID'), 
        authority=f'https://login.microsoftonline.com/{os.getenv("TENANT_ID")}',
        client_credential={
            "thumbprint": os.getenv('CLIENT_CERT_THUMBPRINT'), 
            "private_key": key
        },
    )

    result = app.acquire_token_for_client(scopes.split(' '))
    
    if "access_token" in result:
        return result['access_token']
    else:
        logging.error(result.get("error"))
        logging.error(result.get("error_description"))
        logging.error(result.get("correlation_id"))  # You may need this when reporting a bug
        return None        

def get_access_token(scopes, use_cert = True):
    if use_cert:
        token = get_access_token_cert(scopes)
    else:
        token = get_access_token_secret(scopes)
    return token

def make_api_call(token, api, method, payload=None):
    headers = {
        'Authorization': 'Bearer ' + token,
        'Accept': 'application/json',
    }
    logging.info(f"Calling {method} on '{api}'")
    if method == 'GET':
        res = requests.get(api, headers=headers)
    elif method == 'POST':
        res = requests.post(api, headers=headers, json=payload)
    else:
        raise Exception('Invalid method')
    if res.status_code < 400:
        logging.debug(res.text)
        ans= res.json()
        return ans
    else:
        logging.error(res.status_code, res.text)
        raise Exception(f'{res.status_code}: {res.text}')


def call_sharepoint(token, api, method='GET',*, site=None):
    sp_site = site or os.getenv('SHAREPOINT_SITE')
    base = f'https://{SHAREPOINT_BASE}/sites/{sp_site}/_api'

    res = make_api_call(token, base + api, method)
    return res

def call_sharepoint_admin(token, api, method='GET'):
    base = f'https://{SHAREPOINT_ADMIN_BASE}/_api'

    res = make_api_call(token, base + api, method)
    return res

def get_permissions(token_sp, token_spa, site=None):
    site_info = call_sharepoint(token_sp, f'/site/id', site=site)
    site_id = site_info['value']
    
    url = f"/SPO.Tenant/sites/GetSiteUserGroups?siteId=%27{site_id}%27&userGroupIds=[0,1,2]"
    data = call_sharepoint_admin(token_spa, url)
    owners = data['value'][0]['userGroup']
    members = data['value'][1]['userGroup']
    readers = data['value'][2]['userGroup']
    def get_list(group):
        ans = []
        for l in group:
            ans.append({
                "name":l.get('name',''),
                "upn":l.get('userPrincipalName',''),
                'login':l.get('loginName',''),
            })
        return ans

    return {
        "owners": get_list(owners),
        "members": get_list(members),
        "readers": get_list(readers),
    }
        
def main():
    # graph_scopes = "https://graph.microsoft.com/.default"
    # token_gr = get_access_token(graph_scopes, use_cert=True)

    sp_scopes = f"https://{SHAREPOINT_BASE}/.default"
    token_sp = get_access_token(sp_scopes, use_cert=True)

    spadmin_scopes = f"https://{SHAREPOINT_ADMIN_BASE}/.default"
    token_spa = get_access_token(spadmin_scopes, use_cert=True)


    res = get_permissions(token_sp, token_spa)
    print(json.dumps(res, indent=2))


if __name__ == "__main__":
    main()