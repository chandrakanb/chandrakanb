from jira import JIRA
import sys
import os
import requests
from requests.auth import HTTPBasicAuth

jira_server = "https://kpithondajapan.atlassian.net"
username = "lumesh.jorwar@kpit.com"
api_token = "ATATT3xFfGF0vfNlOj36xqi_knX5L3YBkNQ0jadQHnhoe74f-OjkvG7Ap6R8muUGmAjBBnSDVV68SZTSZY4AnBjfGtXNdhMKZCnQO6yjEh4KyyOuRvCmmr90sgKeDA9dyAahv8ksxC5EVtFcEQFL_2c-98rF81GcFFwsheHjWQWZhV46ih6XSGg=2440DFB7"

try:
    jira = JIRA(server=jira_server, basic_auth=(username, api_token))
    print(f"Successfully connected to Jira server: {jira_server}")
except Exception as e:
    print(f"Failed to connect to Jira: {e}")
    sys.exit(1)  # Exit the script if connection fails

issue_key = "A15-3565"  # Replace with the issue key you want to delete
delete_url = f"{jira_server}/rest/api/3/issue/{issue_key}"  # API v3 endpoint

try:
    response = requests.delete(delete_url, auth=HTTPBasicAuth(username, api_token))
    response.raise_for_status()  # Raise an exception for bad status codes (4xx or 5xx)
    print(f"Successfully deleted issue: {issue_key}")
except requests.exceptions.RequestException as e:
    print(f"Failed to delete issue {issue_key}: {e}")