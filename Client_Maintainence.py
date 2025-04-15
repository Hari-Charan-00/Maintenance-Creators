import requests
import json
import pandas as pd
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BaseUrl = "https://netenrich.opsramp.com/"
OpsRampSecret = ''
OpsRampKey = ''

def create_one_time_maintenance_schedule(data):
    token_url = BaseUrl + "auth/oauth/token"
    auth_data = {
        'client_secret': OpsRampSecret,
        'grant_type': 'client_credentials',
        'client_id': OpsRampKey
    }
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    token_response = requests.post(token_url, data=auth_data, headers=headers, verify=True)

    success = True  # To track if all maintenance schedules are created successfully

    if token_response.status_code == 200:
        access_token = token_response.json().get('access_token')
        if access_token:
            auth_header = {'Authorization': 'Bearer ' + access_token, 'Content-Type': 'application/json'}

            for index, row in data.iterrows():
                payload = {
                    "description": row['Description'],
                    "name": row['Name'],
                    "schedule": {
                        "startTime": row['Start_Time'],
                        "endTime": row['End_Time'],
                        "timezone": row['Timezone'],
                        "type": "One-Time"
                    }
                }

                api_endpoint = f'https://netenrich.api.opsramp.com/api/v2/tenants/{client_id}/scheduleMaintenances'

                # Make the API call to create the maintenance schedule
                response = requests.post(api_endpoint, data=json.dumps(payload), headers=auth_header, verify=True)

                if response.status_code != 200:
                    print(f"Failed to create maintenance schedule for client {client_id}.")
                    print(f"Response: {response.text}")
                    success = False

            if success:
                print(f"Successfully created the maintenance on all resources for the client {client_id}.")
            else:
                print(f"Failed to create maintenance on all resources for the client {client_id}.")
        else:
            print("Failed to obtain access token.")
    else:
        print(f"Failed to authenticate. Status Code: {token_response.status_code}")
        print(f"Response: {token_response.text}")

# Read data from the Excel sheet
excel_file_path = 'client_mai.xlsx'
data = pd.read_excel(excel_file_path)

# Get the 'Client_ID' from the user as input
client_id = input("Enter the Client ID: ")

create_one_time_maintenance_schedule(data)
