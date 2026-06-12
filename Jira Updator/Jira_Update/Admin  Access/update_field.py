import os
import sys
import requests
from requests.auth import HTTPBasicAuth
from jira import JIRA
from datetime import datetime
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

jira_server = "https://kpithondajapan.atlassian.net"
username = "lumesh.jorwar@kpit.com"
api_token = "ATATT3xFfGF0vfNlOj36xqi_knX5L3YBkNQ0jadQHnhoe74f-OjkvG7Ap6R8muUGmAjBBnSDVV68SZTSZY4AnBjfGtXNdhMKZCnQO6yjEh4KyyOuRvCmmr90sgKeDA9dyAahv8ksxC5EVtFcEQFL_2c-98rF81GcFFwsheHjWQWZhV46ih6XSGg=2440DFB7"

try:
    jira = JIRA(server=jira_server, basic_auth=(username, api_token))
    logging.info(f"Successfully connected to Jira server: {jira_server}.")
except Exception as e:
    logging.error(f"Failed to connect to Jira: {e}.")
    sys.exit(1)  # Exit the script if connection fails

def update_custom_field(issue_id, custom_field_id, new_value):
    """
    Updates a custom field in a Jira issue.
    """
    try:
        issue = jira.issue(issue_id)
        issue.update(fields={custom_field_id: new_value})
        logging.info(f"Successfully updated custom field '{custom_field_id}' for issue {issue_id} to '{new_value}'.")
    except JIRAError as e:
        logging.error(f"Jira error updating custom field for issue {issue_id}: {e}")
    except Exception as e:
        logging.error(f"Error updating custom field for issue {issue_id}: {e}")

if __name__ == "__main__":
    if len(sys.argv) != 4:
        logging.warning("Usage: python update_field.py <issue_id> <custom_field_id> <new_value>")
        sys.exit(1)

    issue_id = sys.argv[1]
    custom_field_id = sys.argv[2]
    new_value = int(sys.argv[3])

    update_custom_field(issue_id, custom_field_id, new_value)