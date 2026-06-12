import os
import requests
from requests.auth import HTTPBasicAuth

JIRA_EMAIL = "chandrakanb@kpit.com"
JIRA_URL = "https://kpithondajapan.atlassian.net/rest/api/2/issue/"  # Base URL for issues
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

def get_field_values(issue_key, field_ids):
    """
    Retrieves the current values of specified fields for a given JIRA issue.

    Args:
        issue_key (str): The key of the JIRA issue (e.g., "PROJECT-123").
        field_ids (list): A list of field IDs to retrieve values for.

    Returns:
        dict: A dictionary where keys are field IDs and values are the
              corresponding field values.  Returns an empty dictionary if
              the issue or fields are not found.
    """
    url = f"{JIRA_URL}{issue_key}?fields={','.join(field_ids)}"  # Construct the API URL
    try:
        response = requests.get(url, auth=auth)
        response.raise_for_status()  # Raise an exception for bad status codes
        issue = response.json()
        field_values = {}
        for field_id in field_ids:
            if field_id in issue['fields']:
                field_values[field_id] = issue['fields'][field_id]
            else:
                print(f"Field with ID {field_id} not found in issue {issue_key}")
        return field_values
    except requests.exceptions.RequestException as e:
        print(f"Error fetching issue or fields: {e}")
        return {}

# Example Usage:
issue_key = "DRT-7527"  # Replace with your JIRA issue key
field_ids_to_retrieve = ["summary", "status", "priority", "customfield_14461", "customfield_10851"]  # Replace with the desired field IDs

field_values = get_field_values(issue_key, field_ids_to_retrieve)

if field_values:
    print(f"Field Values for Issue {issue_key}:")
    for field_id, value in field_values.items():
        print(f"{field_id} → {value}")
else:
    print(f"Could not retrieve field values for issue {issue_key}.")