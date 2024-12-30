import msal
import json
import os

def get_msal_app(client_id, client_secret, authority):
    return msal.ConfidentialClientApplication(
        client_id=client_id,
        client_credential=client_secret,
        authority=authority
    )

def get_access_token():
    # Load credentials from a JSON file
    credentials_path = os.path.join(os.path.dirname(__file__), '..', '..', '..', 'credentials.json')
    with open(credentials_path) as f:
        credentials = json.load(f)

    client_id = credentials['client_id']
    client_secret = credentials['client_secret']
    authority = 'https://login.microsoftonline.com/common'
    scopes = ['Files.ReadWrite.All']

    app = get_msal_app(client_id, client_secret, authority)

    # Check for cached token
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes, account=accounts[0])
    else:    
        # Get the authorization URL
        auth_url = app.get_authorization_request_url(scopes, redirect_uri='http://localhost')
        print(f"Please go to this URL and authorize the app: {auth_url}")
        
        # Capture the authorization code from the user
        code = input("Enter the authorization code: ")
        
        # Get the access token
        result = app.acquire_token_by_authorization_code(code, scopes=scopes, redirect_uri='http://localhost')
        
    if 'access_token' in result:
        return result['access_token']
    else:
        raise Exception("Failed to acquire access token")