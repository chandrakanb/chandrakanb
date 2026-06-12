import requests
import json

# 1. Use 'username' instead of 'query' in the URL
target_user = "kpt_rathod_na" 
url = f"https://jira.honda-ivi.com/rest/api/2/user/search?username={target_user}"

# 2. Use your Personal Access Token (PAT)
token = "MDE3NTE2MTg1MjQ5Oqc4yGWxGSApiFdsaSE/F4GPXrwK"

# 3. Use Bearer authentication
headers = {
    "Authorization": f"Bearer {token}",
    "Accept": "application/json"
}

try:
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        users = response.json()
        if users:
            user_data = users[0]
            print("Success!")
            print(f"Display Name: {user_data.get('displayName')}")
            # In Data Center, 'name' or 'key' is usually the unique ID used for automation
            print(f"User ID (name/key): {user_data.get('name') or user_data.get('key')}")
            print(f"Full User Data: {json.dumps(user_data, indent=2)}")
        else:
            print(f"No user found matching: {target_user}")
    elif response.status_code == 400:
        print("Error 400: Bad Request. The server didn't like the parameters. Check the URL format.")
    elif response.status_code == 401:
        print("Error 401: Unauthorized. Check your PAT token.")
    else:
        print(f"Error {response.status_code}: {response.text}")

except Exception as e:
    print(f"An error occurred: {e}")

# This is AI generated code, please refer KPIT AI Policy before using this in your projects