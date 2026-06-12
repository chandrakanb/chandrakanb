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

def add_label_to_jira_ids(issue_id, label_to_add):
    """
    Adds a label to a Jira issue.  If the label already exists, it's added again (no change).
    """
    try:
        issue = jira.issue(issue_id)
        existing_labels = issue.fields.labels
        logging.info(f"Existing Labels for {issue_id}: {existing_labels}")

        updated_labels = existing_labels + [label_to_add]  # Always add the label
        logging.info(f"Adding label '{label_to_add}' to issue {issue_id}.")


        if updated_labels != existing_labels:  # Only update if something changed
            issue.update(fields={"labels": updated_labels})
            logging.info(f"Successfully updated Labels for issue {issue_id}.")
        else:
            logging.info(f"Label '{label_to_add}' already present, no changes made for issue {issue_id}.")

    except JIRAError as e:
        logging.error(f"Jira error fetching/updating/adding labels for issue {issue_id}: {e}")
    except Exception as e:
        logging.error(f"Error fetching/updating/adding labels for issue {issue_id}: {e}")


issue_id = sys.argv[1] if len(sys.argv) > 1 else None
label_to_add = sys.argv[2] if len(sys.argv) > 2 else None

if not issue_id:
    logging.warning("Please provide Issue ID to add/remove the Label.")
elif not label_to_add:
    logging.warning("Please provide Label to add/remove.")
else:
    add_label_to_jira_ids(issue_id, label_to_add) # Pass both arguments