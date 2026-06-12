import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
import json
from datetime import date
import urllib3
import sys
import os  # For environment variables

def update_defect(issue_key, assignee, username=None, api_token=None):
    """
    Updates the assignee of a Jira issue.

    Args:
        issue_key (str): The key of the Jira issue to update (e.g., "PROJECT-123").
        assignee (str): The account ID of the user to assign the issue to.
        username (str, optional): The email address for Jira authentication.
                                        Defaults to environment variable JIRA_USER_MAIL.
        api_token (str, optional): The API token for Jira authentication.
                                   Defaults to environment variable JIRA_API_TOKEN.

    Returns:
        dict: A dictionary containing the status of the update.
              {"success": True, "message": "Issue updated successfully"}
              or
              {"success": False, "message": "Error message"}
    """

    jira_server = f'https://kpithondajapan.atlassian.net/rest/api/3/issue/{issue_key}'

    # Get credentials from environment variables
    username = "lumesh.jorwar@kpit.com"
    api_token = "ATATT3xFfGF0vfNlOj36xqi_knX5L3YBkNQ0jadQHnhoe74f-OjkvG7Ap6R8muUGmAjBBnSDVV68SZTSZY4AnBjfGtXNdhMKZCnQO6yjEh4KyyOuRvCmmr90sgKeDA9dyAahv8ksxC5EVtFcEQFL_2c-98rF81GcFFwsheHjWQWZhV46ih6XSGg=2440DFB7"

    auth = HTTPBasicAuth(username, api_token)
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json"
    }

    try:
        # Fetch issue data
        response = requests.get(jira_server, headers=headers, auth=auth, verify=True)
        response.raise_for_status()  # Raise HTTPError for bad responses (4xx or 5xx)

        # Construct the payload
        payload = {
            "fields": {
                "assignee": {"accountId": assignee}
            }
        }

        # Update the issue
        response = requests.put(jira_server, json=payload, headers=headers, auth=auth, verify=True)
        response.raise_for_status() # Raise HTTPError for bad responses (4xx or 5xx)

        if response.status_code == 204:
            return {"success": True, "message": f"Issue {issue_key} updated successfully."}
        else:
            return {"success": False, "message": f"Failed to update issue. Response: {response.text}"}

    except requests.exceptions.RequestException as e:
        return {"success": False, "message": f"Request failed: {e}"}
    except Exception as e:
        return {"success": False, "message": f"An unexpected error occurred: {e}"}

# Example Usage (replace with your actual values)
issue_key = "DRT-3914"
assignee = "712020:7ebd6444-898e-4d1e-ba7a-2710a262ffa7"
result = update_defect(issue_key, assignee)
if result["success"]:
    print(result["message"])
else:
    print(f"Error: {result['message']}")
