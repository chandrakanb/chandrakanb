import requests
from requests.auth import HTTPBasicAuth

# Jira Config
JIRA_DOMAIN = "https://kpithondajapan.atlassian.net"
API_USER_EMAIL = "vitthalh@kpit.com"
API_TOKEN = "ATATT3xFfGF0Tc7G0IKJ284eB1FcHIPyAaSuR9bRDQhUzEjgJq24L08qXWR75TBIhpXyn5xeGRaISqd20oZ3wxYbWeILDc2SRtiEC4VXnG8aSYewOTbTNRFttXV1JfltJuqWj8fpKnAxWU7Rzvo1S7tOz5KxkjWV1zma7BOjWAAwD9F-gg3NjhY=82D11A9C"

# Authentication and headers for API
auth = HTTPBasicAuth(API_USER_EMAIL, API_TOKEN)
headers = {
    "Accept": "application/json",
    "Content-Type": "application/json"
}

# Function to get accountId based on email
def get_assignee_account_id(email):
    user_url = f"{JIRA_DOMAIN}/rest/api/3/user?accountId={email}"  # Use accountId query parameter
    
    # Make the GET request to fetch user details
    user_response = requests.get(user_url, headers=headers, auth=auth)

    if user_response.status_code == 200:
        user_data = user_response.json()
        account_id = user_data['accountId']
        print(f"User {email} has account ID: {account_id}")
        return account_id
    else:
        print(f"❌ Failed to get user info for {email}: {user_response.status_code} - {user_response.text}")
        return None

# Example: Use a valid email to get the account ID
email = "user@example.com"  # Replace with the actual email of the user
account_id = get_assignee_account_id(email)

if account_id:
    print(f"Account ID for {email}: {account_id}")
else:
    print("No account ID found.")
