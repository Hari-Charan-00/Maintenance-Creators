import requests
import json
import pandas as pd
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BaseUrl = "https://netenrich.opsramp.com/"
OpsRampSecret = ''
OpsRampKey = ''

def create_maintenance_schedule(data):
    token_url = BaseUrl + "auth/oauth/token"
    auth_data = {
        'client_secret': OpsRampSecret,
        'grant_type': 'client_credentials',
        'client_id': OpsRampKey
    }
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    token_response = requests.post(token_url, data=auth_data, headers=headers, verify=False)

    if token_response.status_code == 200:
        access_token = token_response.json().get('access_token')
        if access_token:
            auth_header = {'Authorization': 'Bearer ' + access_token, 'Content-Type': 'application/json'}
            api_endpoint = f'https://netenrich.api.opsramp.com/api/v2/tenants/{data["Client_ID"]}/scheduleMaintenances'

            payload = {
                "name": data['Name'],
                "description": "Script",
                "runRBA": False,
                "installPatch": False,
                "devices": [
                    {
                        "hostName": data['Device_Name'],
                        "uniqueId": data['unique_Id']
                    }
                ],
                "schedule": {
                    "type": "one-time",
                    "startTime": "2024-08-06T22:40:00+0530",
                    "endTime": "2024-08-06T22:45:00+0530",
                    "timezone": "Asia/Calcutta",
                }
            }

            response = requests.post(api_endpoint, headers=auth_header, json=payload, verify=False)

            if response.status_code == 200:
                print("Maintenance schedule created successfully!")
            else:
                print(f"Failed to create maintenance schedule. Status code: {response.status_code}")
                print("Error message:", response.text)
        else:
            print("Access token not found in the response.")
    else:
        print(f"Failed to obtain access token. Status code: {token_response.status_code}")
        print("Error message:", token_response.text)

# Read input data from Excel file
input_file = 'Maintenance.xlsx'  # Update with your file path
df = pd.read_excel(input_file)

# Loop through the rows and call create_maintenance_schedule for each row
for index, row in df.iterrows():
    data = {
        'Name': row['Name'],
        'Client_ID': row['client_id'],  # Ensure this matches the column name in your Excel file
        'Device_Name': row['Device_Name'],
        'unique_Id': row['unique_Id']  # Ensure this matches the column name in your Excel file
    }

    # Call the function with the provided payload
    create_maintenance_schedule(data)
