import requests
from requests.auth import HTTPBasicAuth

# Jira credentials
JIRA_URL = 'https://kpithondajapan.atlassian.net'
USERNAME = 'chandrakanb@kpit.com'
API_TOKEN = 'ATATT3xFfGF00OVb1BHWgz1klgLPAPkq8W8nFtC2j6k_ZYhvBxAWBFwEnCRk18fPD1ctA4TWpWJ79K4s4OPXJOiskD0GInEa14A0ex8puHBIQRLaGOsmNkVRY4ij9rKJr1q2JgK9D4zR-nK0rpl4QevUi5_g2PWiSM5HB7qEZAOMUaWUirk1vD4=54745E1F'

# New reporter's accountId
NEW_REPORTER_ID = '712020:9000a5c2-2cc0-40f8-a94d-4c00bc685071'

# List of issue keys
issue_keys = [
    "DRT-4110", "DRT-4111", "DRT-4112",
    "DRT-4113", "DRT-4114", "DRT-4115", "DRT-4116"
]

# Headers
headers = {
    "Accept": "application/json",
    "Content-Type": "application/json"
}

# Loop through issues and update reporter
for key in issue_keys:
    url = f"{JIRA_URL}/rest/api/3/issue/{key}"
    payload = {
        "fields": {
            "reporter": {
                "id": NEW_REPORTER_ID
            }
        }
    }

    response = requests.put(
        url,
        json=payload,
        headers=headers,
        auth=HTTPBasicAuth(USERNAME, API_TOKEN)
    )

    if response.status_code == 204:
        print(f"✅ {key} → Reporter updated successfully.")
    else:
        print(f"❌ {key} → Failed: {response.status_code} | {response.text}")
