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
    log.error(f"Failed to connect to Jira: {e}")
    sys.exit(1)  # Exit the script if connection fails

issues = [
    "DRT-1816",
    "DRT-1840",
    "DRT-2538",
    "DRT-3309",
    "DRT-3838",
    "DRT-3948",
    "DRT-4090",
    "DRT-4092",
    "DRT-4426",
    "DRT-5308",
    "DRT-7601",
    "DRT-7602",
    "DRT-7603"
]
# issues = ["DRT-3932", "DRT-3931", "DRT-2306", "DRT-2139", "DRT-1820", "DRT-1818", "DRT-1811", "DRT-1800" ]

#target_transitions = ["Analysis In Progress", "Analysis Completed", "Issue Fixing In Progress", "Sending for Review", "Under Observation for 3 Days"]
target_transitions = ["Moving_In_Progress"]
#target_transitions = ["Reopend"]

for issue in issues:
    for transition in target_transitions:
        try:
            jira.transition_issue(issue, transition)
        except Exception as e:
            print(f"Failed to connect to Jira: {e}")