from jira import JIRA
import sys
import os
import requests
from requests.auth import HTTPBasicAuth

jira_server = "https://kpithondajapan.atlassian.net"
username = "chandrakanb@kpit.com"
api_token = os.getenv("JIRA_API_TOKEN")

try:
    jira = JIRA(server=jira_server, basic_auth=(username, api_token))
    print(f"Successfully connected to Jira server: {jira_server}")
except Exception as e:
    log.error(f"Failed to connect to Jira: {e}")
    sys.exit(1)  # Exit the script if connection fails

# issues = ["DRT-1800"]
issues = ["A15-2311", "A15-1908", "A15-1778", "A15-1793", "A15-1909", "A15-1905", "A15-1776", "A15-1784", "A15-1904", "A15-1903", "A15-1902", "A15-1777", "A15-2383", "A15-2376", "A15-1349", "A15-2145", "A15-1427", "A15-1951"]

# target_transitions = ["Analysis In Progress", "Analysis Completed", "Issue Fixing In Progress", "Sending for Review", "Stabilization not required"]
# target_transitions = ["Moving_In_Progress"]
# target_transitions = ["Reopend"]
target_transitions = ["TS Development Is In Progress", "TS Completed", "Execution is in progress", "Sent for L1 Review with 1 Iterations"]
for issue in issues:
    for transition in target_transitions:
        try:
            jira.transition_issue(issue, transition)
        except Exception as e:
            print(f"Failed to connect to Jira: {e}")