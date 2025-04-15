import requests
import json
import pandas as pd
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BaseUrl = "https://netenrich.opsramp.com/"
OpsRampSecret = 'RETp4NzhXvrpUBJzhJXsfTawypEt2ZQYWNBRf9qktcDeGwPmEmFtw9KqAvChyYxW'
OpsRampKey = 'wyJ9bKRZGzby55K5xxuxa9Dkh53RJv46'

def create_maintenance_schedule(data, client_id):
    token_url = BaseUrl + "auth/oauth/token"
    auth_data = {
        'client_secret': OpsRampSecret,
        'grant_type': 'client_credentials',
        'client_id': OpsRampKey
    }
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    token_response = requests.post(token_url, data=auth_data, headers=headers, verify=True)

    if token_response.status_code == 200:
        access_token = token_response.json().get('access_token')
        if access_token:
            auth_header = {'Authorization': 'Bearer ' + access_token, 'Content-Type': 'application/json'}

            for index, row in data.iterrows():
                num_conditions = int(row['Num_Conditions'])
                if num_conditions < 1 or num_conditions > 4:
                    print(f"Invalid number of conditions ({num_conditions}) for '{row['deviceName']}' in client '{client_id}'.")
                    continue  # Skip this entry and move to the next
                matching_type = row['Matching_Type']
                name = row['name']
                description = row['description']
                device_name = row['deviceName']
                pattern_type = row['Pattern_Type']  # Column for pattern type
                pattern_weekdays = row['Pattern_WeekDays']  # Column for weekdays

                # Constructing the maintenance schedule data
                conditions = []
                for i in range(1, num_conditions + 1):
                    condition_key = row[f"condition_{i}_key"]
                    condition_operator = row[f"condition_{i}_operator"]
                    condition_value = row[f"condition_{i}_value"]
                    conditions.append({
                        "key": condition_key,
                        "operator": condition_operator,
                        "value": condition_value
                    })

                maintenance_schedule_data = {
                    "name": name,
                    "description": description,
                    "startTime": row['Start_Time'],
                    "endTime": row['End_Time'],
                    "timezone": row['Schedule_Timezone'],
                    "pattern": {
                        "type": pattern_type,
                        "weekDays": pattern_weekdays
                    },
                    "devices": [
                        {
                            "aliasName": device_name
                        }
                    ],
                    "alertConditions": {
                        "matchingType": matching_type,
                        "rules": conditions
                    }
                }

                # Make the API call to create the maintenance schedule
                api_endpoint = f'https://netenrich.api.opsramp.com/api/v2/tenants/{client_id}/scheduleMaintenances'
                response = requests.post(api_endpoint, data=json.dumps(maintenance_schedule_data), headers=auth_header, verify=True)

                if response.status_code == 200:
                    print(f"Successfully created the maintenance schedule for '{device_name}' in client '{client_id}'.")
                else:
                    # Handle the case where the schedule creation fails
                    try:
                        error_data = response.json()
                        error_message = error_data.get('message', 'Unknown error')
                    except json.JSONDecodeError:
                        error_message = 'Failed to decode the response content'
                    print(f"Failed to create maintenance schedule for '{device_name}' in client '{client_id}'. Reason: {error_message}")
        else:
            print("Failed to obtain access token.")
    else:
        print(f"Failed to authenticate. Status Code: {token_response.status_code}")
        print(f"Response: {token_response.text}")

# Read data from the Excel sheet
excel_file = "C:\\Users\\hari.boddu\\Downloads\\Alert_Maintenance_Recuuring.xlsx"

try:
    data = pd.read_excel(excel_file)
    print("Read data from Excel successfully.")

    # Call the function to create the maintenance schedules for each client_id in the Excel sheet
    for index, row in data.iterrows():
        client_id = row['Client_ID']
        create_maintenance_schedule(data, client_id)

except Exception as e:
    print(f"An error occurred: {str(e)}")
