import os
import requests
from requests.auth import HTTPBasicAuth

JIRA_EMAIL = "chandrakanb@kpit.com"
JIRA_URL = "https://kpithondajapan.atlassian.net/rest/api/2/field"
API_TOKEN_FILE = os.getenv("JIRA_API_TOKEN")
if API_TOKEN_FILE:
    print("Token loaded from Environment Variable")
else:
    if os.path.isfile("API_TOKEN_FILE.txt"):
        with open("API_TOKEN_FILE.txt", "r") as f:
            API_TOKEN_FILE = f.read().strip()
        print("Token loaded from API_TOKEN_FILE.txt")
    else:
        raise SystemExit("❌ API token not found in environment variable or API_TOKEN_FILE.txt")

auth = HTTPBasicAuth(JIRA_EMAIL, API_TOKEN_FILE)
response = requests.get(JIRA_URL, auth=auth)

response.raise_for_status()
fields = response.json()

for field in fields:
    print(f"{field['name']} → {field['id']}")
