import requests
import json
import pandas as pd
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BaseUrl = "https://netenrich.opsramp.com/"
OpsRampSecret = 'RETp4NzhXvrpUBJzhJXsfTawypEt2ZQYWNBRf9qktcDeGwPmEmFtw9KqAvChyYxW'
OpsRampKey = 'wyJ9bKRZGzby55K5xxuxa9Dkh53RJv46'

def create_maintenance_schedule(data):
    success = True  # Initialize success as True
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
                try:
                    schedule_type = input("Enter Schedule Type (one-time/recurring): ").lower()  # Ask the user to specify the schedule type

                    if schedule_type == 'one-time':
                        payload = {
                            "description": row['Description'],
                            "name": row['Name'],
                            "deviceGroups": [
                                {
                                    "name": row['group_name']
                                },
                                {
                                    "id": row['group_id']
                                }
                            ],
                            "schedule": {
                                "type": "one-time",
                                "startTime": row['Start_Time'],
                                "endTime": row['End_Time'],
                                "timezone": row['Timezone']
                            }
                        }
                    elif schedule_type == 'recurring':
                        # Ask the user to specify the pattern type
                        pattern_type = input("Enter Pattern Type (daily/weekly/monthly): ").lower()
                        
                        if pattern_type not in ['daily', 'weekly', 'monthly']:
                            print("Invalid Pattern Type specified.")
                            continue
                        
                        # Manually specify the schedule type and additional fields based on the schedule type
                        payload = {
                            "name": row['Name'],
                            "description": row['Description'],
                            "deviceGroups": [
                                {
                                    "name": row['group_name']
                                },
                                {
                                    "id": row['group_id']
                                }
                            ],
                            "schedule": {
                                "type": "recurring",
                                "startTime": row['Start_Time'],
                                "endTime": row['End_Time'],
                                "timezone": row['Timezone'],
                                "pattern": {
                                    "type": pattern_type
                                }
                            }
                        }
                        
                        # Manually specify the additional fields based on the pattern type
                        if pattern_type == 'daily':
                            payload['schedule']['pattern']['dayFrequency'] = input("Enter Day Frequency: ")
                        elif pattern_type == 'weekly':
                            payload['schedule']['pattern']['weekDays'] = input("Enter Week Days: ")
                        elif pattern_type == 'monthly':
                            day_type = input("Enter Day Type (Day_Of_Week or Day_Of_Month): ")
                            if day_type == 'Day_Of_Week':
                                payload['schedule']['pattern']['dayOfWeek'] = input("Enter Day Of Week: ")
                            elif day_type == 'Day_Of_Month':
                                payload['schedule']['pattern']['dayOfMonth'] = input("Enter Day Of Month: ")
                            else:
                                print("Invalid Day Type specified.")
                                continue

                    else:
                        print(f"Invalid schedule type for client {row['Client_ID']}: {schedule_type}")
                        continue  # Skip processing for invalid schedule types

                    client_id = row['Client_ID']  # Get the client_id from the Excel file

                    api_endpoint = f'https://netenrich.api.opsramp.com/api/v2/tenants/{client_id}/scheduleMaintenances'

                    # Make the API call to create the maintenance schedule
                    response = requests.post(api_endpoint, data=json.dumps(payload), headers=auth_header, verify=True)

                    if response.status_code != 200:
                        print(f"Failed to create maintenance schedule for group {row['group_name']} in client {client_id}.")
                        print(f"Response: {response.text}")
                        success = False

                except Exception as e:
                    print(f"Error processing row {index}: {str(e)}")
                    print(f"Row Data: {row}")

            if success:
                print(f"Successfully created the maintenance schedules for all groups.")
            else:
                print(f"Failed to create maintenance schedules for some groups.")
        else:
            print("Failed to obtain access token.")
    else:
        print(f"Failed to authenticate. Status Code: {token_response.status_code}")
        print(f"Response: {token_response.text}")


# Read data from the Excel sheet
excel_file_path = 'C:\\Users\\hari.boddu\\Downloads\\Group_Maintenance.xlsx'
data = pd.read_excel(excel_file_path)

# Call the function to create maintenance schedules
create_maintenance_schedule(data)


