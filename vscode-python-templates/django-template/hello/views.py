import requests
import json
import os
from django.shortcuts import render, redirect
from .auth import get_access_token

def get_file_path():
    credentials_path = os.path.join(os.path.dirname(__file__), '..', '..', '..', 'credentials.json')
    with open(credentials_path) as f:
        credentials = json.load(f)
    return credentials['file_path']

def read_excel(request):
    access_token = get_access_token()
    print(f"Access Token: {access_token}")  # Debugging access token
    file_path = get_file_path()
    workbook_url = f'https://graph.microsoft.com/v1.0/me/drive/root:{file_path}:/workbook/worksheets/Sheet1/range(address=\'A1\')'

    headers = {
        'Authorization': f'Bearer {access_token}'
    }

    response = requests.get(workbook_url, headers=headers)
    
    # Log the response for debugging
    print(response.json())

    if response.status_code == 200:
        cell_value = response.json().get('values', [['']])[0][0]
        return render(request, 'hello/read_excel.html', {'cell_value': cell_value})
    else:
        error_message = response.json().get('error', {}).get('message', 'Unknown error')
        return render(request, 'hello/read_excel.html', {'error': error_message})

def update_excel(request):
    if request.method == 'POST':
        new_value = request.POST['new_value']
        access_token = get_access_token()
        file_path = get_file_path()
        update_url = f'https://graph.microsoft.com/v1.0/me/drive/root:{file_path}:/workbook/worksheets/Sheet1/range(address=\'A1\')'

        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }

        data = {
            'values': [[new_value]]
        }

        response = requests.patch(update_url, json=data, headers=headers)
        if response.status_code == 200:
            return redirect('read_excel')
        else:
            error_message = response.json().get('error', {}).get('message', 'Unknown error')
            return render(request, 'hello/update_excel.html', {'error': error_message})
    else:
        return render(request, 'hello/update_excel.html')