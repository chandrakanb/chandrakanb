import requests
from requests.auth import HTTPBasicAuth
import pandas as pd
import sys

# JIRA details
jira_url = "https://kpithondajapan.atlassian.net"  # Change this to your JIRA instance URL

# Authentication details
username = "chandrakanb@kpit.com"  # Your JIRA email
api_token = "ATATT3xFfGF0OB0rbo6JnQkMqLWvZVPVWXzrfm4irzK_FIFWR3w4kfQ3uug6MgffT25YppnLmp0ZjVJoZO2cpbpBmn4P--s-zdpbkUOJEoUfWzRLRySLFoHQRR3s3UnMyApSMbmild4QhPkhG-NePjSGD0UqwZEQnDlZuxlx-dz5LVuhKbGlMEI=63668BB4"  # Generate it from https://id.atlassian.com/manage-profile/security/api-tokens

# Get Excel file name from argument
if len(sys.argv) < 2:
    print("Usage: python script_name.py <excel_filename>")
    sys.exit(1)

excel_filename = sys.argv[1]

# Read issue keys from Excel
issue_keys = pd.read_excel(excel_filename, header=None).iloc[:, 0].tolist()

# Loop through each issue key and post comment
for ticket_id in issue_keys:
    print(f"Posting comment to ticket: {ticket_id}")

    api_endpoint = f"{jira_url}/rest/api/3/issue/{ticket_id}/comment"

    payload = {
        "body": {
            "type": "doc",
            "version": 1,
            "content": [
                {
                    "type": "table",
                    "content": [
                        {
                            "type": "tableRow",
                            "content": [
                                {
                                    "type": "tableCell",
                                    "content": [
                                        {
                                            "type": "paragraph",
                                            "content": [
                                                {
                                                    "type": "text",
                                                    "text": "Date: 19th February 2025(1405 Dev)"
                                                }
                                            ]
                                        }
                                    ]
                                }
                            ]
                        },
                        {
                            "type": "tableRow",
                            "content": [
                                {
                                    "type": "tableCell",
                                    "content": [
                                        {
                                            "type": "paragraph",
                                            "content": [
                                                {
                                                    "type": "text",
                                                    "text": "Status: Passed"
                                                }
                                            ]
                                        }
                                    ]
                                }
                            ]
                        }
                    ]
                }
            ]
        }
    }

    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json"
    }

    response = requests.post(
        api_endpoint,
        headers=headers,
        json=payload,
        auth=HTTPBasicAuth(username, api_token)
    )

    if response.status_code == 201:
        print(f"Comment added successfully to {ticket_id}!")
    else:
        print(f"Failed to add comment to {ticket_id}: {response.status_code} {response.text}")
