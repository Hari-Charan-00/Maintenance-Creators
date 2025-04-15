import requests
import json
import pandas as pd
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BaseUrl = "https://netenrich.opsramp.com/"
OpsRampSecret = 'RETp4NzhXvrpUBJzhJXsfTawypEt2ZQYWNBRf9qktcDeGwPmEmFtw9KqAvChyYxW'
OpsRampKey = 'wyJ9bKRZGzby55K5xxuxa9Dkh53RJv46'

def create_maintenance_schedule(data):
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
        print (access_token)
        if access_token:
            auth_header = {'Authorization': 'Bearer ' + access_token, 'Content-Type': 'application/json'}

            for index, row in data.iterrows():
                client_id = input(f"Enter the Client ID for '{row['aliasName']}': ")
                num_conditions = int(input(f"Enter number of conditions for '{row['aliasName']}': "))
                matching_type = input(f"Enter matching type (ANY/ALL) for '{row['aliasName']}': ").upper()
                
                if 1 <= num_conditions <= 4:
                    conditions = []
                    for i in range(num_conditions):
                        key = row[f"condition_{i + 1}_key"]
                        operator = row[f"condition_{i + 1}_operator"]
                        value = row[f"condition_{i + 1}_value"]
                        conditions.append({
                            "key": key,
                            "operator": operator,
                            "value": value
                        })

                    maintenance_schedule_data = {
                        "name": row['name'],
                        "description": row['description'],
                        "devices": [
                            {
                                "aliasName": row['aliasName']
                            }
                        ],
                        "schedule": {
                            "type": "One-Time",
                            "startTime": row['startTime'],
                            "endTime": row['endTime'],
                            "timezone": row['timezone']
                        },
                        "alertConditions": {
                            "matchingType": matching_type,  # Use matching type from user input
                            "rules": conditions
                        }
                    }

                    create_schedule_url = f'https://netenrich.api.opsramp.com/api/v2/tenants/{client_id}/scheduleMaintenances'
                    create_schedule_response = requests.post(create_schedule_url, json=maintenance_schedule_data, headers=auth_header, verify=True)

                    if create_schedule_response.status_code != 200:
                        success = False
                        print(f"Failed to create schedule for alias: {row['aliasName']}")
                        print (create_schedule_response)

                else:
                    print("Number of conditions must be between 1 and 4.")

            if success:
                print("All maintenance schedules created successfully.")
        else:
            print("Access token not received.")
    else:
        print("Token request failed.")

# Read Excel file into a DataFrame
excel_file = "C:\\Users\\hari.boddu\\Downloads\\Alert_Maintenance.xlsx"
data = pd.read_excel(excel_file)

# Call the function with the DataFrame
create_maintenance_schedule(data)
