import requests
from requests.auth import HTTPBasicAuth

# JIRA details
jira_url = "https://kpithondajapan.atlassian.net"  # Change this to your JIRA instance URL
ticket_url = "https://kpithondajapan.atlassian.net/browse/DRT-1745"  # Replace with your ticket URL
comment = "This is a test comment from Python script."

# Extracting ticket ID from the URL
ticket_id = ticket_url.split('/')[-1]

# Authentication details
username = "chandrakanb@kpit.com"  # Your JIRA email
api_token = "ATATT3xFfGF0OB0rbo6JnQkMqLWvZVPVWXzrfm4irzK_FIFWR3w4kfQ3uug6MgffT25YppnLmp0ZjVJoZO2cpbpBmn4P--s-zdpbkUOJEoUfWzRLRySLFoHQRR3s3UnMyApSMbmild4QhPkhG-NePjSGD0UqwZEQnDlZuxlx-dz5LVuhKbGlMEI=63668BB4"  # Generate it from https://id.atlassian.com/manage-profile/security/api-tokens

# API endpoint to add a comment
api_endpoint = f"{jira_url}/rest/api/3/issue/{ticket_id}/comment"

# Request headers and payload
headers = {
    "Accept": "application/json",
    "Content-Type": "application/json"
}

payload = {
    "body": {
        "type": "doc",
        "version": 1,
        "content": [
            {
                "type": "table",
                "content": [
                    # First row for Date
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
                                                "text": "Date: 12th February 2025"
                                            }
                                        ]
                                    }
                                ]
                            }
                        ]
                    },
                    # Second row for Status
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


# Sending POST request to add comment
response = requests.post(
    api_endpoint,
    headers=headers,
    json=payload,
    auth=HTTPBasicAuth(username, api_token)
)

# Checking response
if response.status_code == 201:
    print("Comment added successfully!")
else:
    print("Failed to add comment:", response.status_code, response.text)
