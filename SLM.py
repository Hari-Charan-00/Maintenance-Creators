import requests
import pandas as pd
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BaseUrl = "https://netenrich.opsramp.com/"
OpsRampSecret = ''
OpsRampKey = ''

# Function to obtain the OpsRamp API token
def get_opsramp_token():
    token_url = BaseUrl + "auth/oauth/token"
    auth_data = {
        'client_secret': OpsRampSecret,
        'grant_type': 'client_credentials',
        'client_id': OpsRampKey
    }
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    token_response = requests.post(token_url, data=auth_data, headers=headers, verify=True)

    if token_response.status_code == 200:
        return token_response.json().get('access_token')
    else:
        print("Failed to obtain access token. Status code:", token_response.status_code)
        print("Error message:", token_response.text)
        return None

# Function to read the payload and client ID from the Excel file
def read_excel_data():
    # Replace with your Excel file path
    excel_data = pd.read_excel('Site-Level-Maintenance.xlsx')
    
    # Extract client ID and payload data
    client_id = excel_data.iloc[0]['client_id']
    
    # Extract the payload data (excluding client ID column)
    payload_data = excel_data.drop(columns=['client_id'])
    
    return client_id, payload_data

# Function to create the payload from the payload data
def create_payload(payload_data):
    payloads = []
    
    # Iterate over each row (site) in the payload data
    for index, row in payload_data.iterrows():
        name = row['name']
        description = row['description']
        location_name = row['location_name']
        start_time = row['start_time']
        end_time = row['end_time']
        timezone = row['timezone']
        
        # Construct the payload for this site
        payload = {
            "name": name,
            "description": description,
            "locations": [{"name": location_name}],
            "schedule": {
                "type": "one-time",
                "startTime": start_time,
                "endTime": end_time,
                "timezone": timezone
            }
        }
        
        # Add this payload to the list of payloads
        payloads.append(payload)
    
    return payloads

# Function to make the API call for each payload (site)
def send_api_requests(client_id, payloads):

    access_token = get_opsramp_token()
    
    if access_token:

        for payload in payloads:
            api_url = f'https://netenrich.api.opsramp.com/api/v2/tenants/{client_id}/scheduleMaintenances'
            auth_header = {'Authorization': 'Bearer ' + access_token, 'Content-Type': 'application/json'}

            response = requests.post(api_url, json=payload, headers=auth_header, verify=True)
            
            # Print the response for this payload (site)
            print(f"API Response for {payload['name']} ({payload['locations'][0]['name']}):")
            print(response.status_code)
            if response.status_code==200:
                print("Successfully created the Maintenance on the site")
    else:
        print("Failed to obtain OpsRamp API token.")

if __name__ == "__main__":
    client_id, payload_data = read_excel_data()
    payloads = create_payload(payload_data)
    send_api_requests(client_id, payloads)
