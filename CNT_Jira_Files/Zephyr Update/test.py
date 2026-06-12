import requests

def fetch_with_bearer_token():
    url = "https://jira.honda-ivi.com/rest/atm/1.0/testrun/CNTKPIT-C1213/testcase/CNTKPIT-T68/testresult"
    
    # Replace 'YOUR_TOKEN_HERE' with your actual Personal Access Token
    token = "MDE3NTE2MTg1MjQ5Oqc4yGWxGSApiFdsaSE/F4GPXrwK"
    
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json"
    }

    try:
        response = requests.get(url, headers=headers)
        
        if response.status_code == 200:
            print("Success!")
            print(response.json())
        else:
            print(f"Failed. Status Code: {response.status_code}")
            print(f"Response Body: {response.text}")
            
    except Exception as e:
        print(f"Error: {e}")

fetch_with_bearer_token()
# This is AI generated code, please refer KPIT AI Policy before using this in your projects