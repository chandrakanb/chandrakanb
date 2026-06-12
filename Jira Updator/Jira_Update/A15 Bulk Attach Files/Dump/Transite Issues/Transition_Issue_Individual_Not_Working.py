from jira import JIRA
import sys
import os
import requests
from requests.auth import HTTPBasicAuth

jira_server = "https://kpithondajapan.atlassian.net"
username = "chandrakanb@kpit.com"
api_token = "ATATT3xFfGF0zhq0vBH0KDXc4vYh0oPaxKRg7S07GsQt7G86B8f9X1Pd-UrFL0R6VorR1XQmbRYCMnG0LVmOeUCEpg1WhE2XWjpMIIMD8d5r3pKM5_NNo4ITxT5TCN9p9CLYP2SgShP7RaJGKEaZOVb6cVFgbOFtKM6GIQ4X56S4SpYmLkgzvy4=778CDED8"

try:
    jira = JIRA(server=jira_server, basic_auth=(username, api_token))
    print(f"Successfully connected to Jira server: {jira_server}")
except Exception as e:
    log.error(f"Failed to connect to Jira: {e}")
    sys.exit(1)  # Exit the script if connection fails

issues = ["A15-2483"] #, "A15-1956", "A15-1886", "A15-1873"]

target_transitions = "TS_CREATION/MODIFICATION_IN_PROGRESS" #["TS Development In Is Progress", "TS Completed", "Execution is in progress", "Sent for L1 Review with 1 Iterations"]

for issue in issues:
    for transition in target_transitions:
        try:
            jira.transition_issue(issue, transition)
        except Exception as e:
            print(f"Failed to connect to Jira: {e}")