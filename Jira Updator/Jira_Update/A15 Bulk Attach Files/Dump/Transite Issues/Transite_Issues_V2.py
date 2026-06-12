from jira import JIRA
import sys
import pandas as pd
import os
import logging
from datetime import datetime

# --- Configure logging ---
os.makedirs("logs", exist_ok=True)
timestamp = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
log_file_path = os.path.join("logs", f"{timestamp}.log")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[
        logging.FileHandler(log_file_path, encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
log = logging.getLogger()

# --- Configuration (ideally loaded from a config file) ---
excel_path = "Transite_Issues.xlsx"
sheet_name = "Sheet1"
jira_server = "https://kpithondajapan.atlassian.net"
username = "chandrakanb@kpit.com"
status_mapping = {
    "TS_ASSIGNED" : "TS_READY_FOR_L1_REVIEW",
    "TS_CREATION/MODIFICATION_IN_PROGRESS" : "TS Completed",
    "TS_COMPLETED" : "Execution is in progress",
    "EXECUTION_IS_IN_PROGRESS" : "Sent for L1 Review with 1 Iterations"
}

def get_api_token():
    """Retrieves the JIRA API token from an environment variable."""
    api_token = os.getenv("JIRA_API_TOKEN")
    if not api_token:
        log.error("API token not found in environment variable JIRA_API_TOKEN")
        raise SystemExit("❌ API token not found in environment variable JIRA_API_TOKEN")
    log.info("Token loaded from Environment Variable")
    return api_token

def read_excel_data(excel_path, sheet_name):
    """Reads issue IDs from an Excel file."""
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name, dtype=str, keep_default_na=False).fillna("NA")
        log.info(f"Successfully read Excel file: {excel_path}")
        return df
    except FileNotFoundError:
        log.error(f"Excel file not found: {excel_path}")
        sys.exit(1)
    except Exception as e:
        log.error(f"Error reading Excel file: {e}")
        sys.exit(1)

def get_jira_issue(jira, ticket_id):
    """Retrieves a JIRA issue by ID."""
    try:
        issue = jira.issue(ticket_id)
        return issue
    except Exception as e:
        log.error(f"Failed to retrieve issue {ticket_id}: {e}")
        return None

def transition_issue(jira, issue, next_status_name):
    """Transitions a JIRA issue to the next status."""
    transitions = jira.transitions(issue)
    target_transition = None
    for transition in transitions:
        if transition['name'] == next_status_name:
            target_transition = transition
            break

    if target_transition:
        jira.transition_issue(issue, target_transition['id'])
        log.info(f"Transitioned {issue.key} to '{next_status_name}'")
        return True
    else:
        log.warning(f"Transition '{next_status_name}' not found for {issue.key}")
        return False

# --- Main execution ---
if __name__ == "__main__":
    api_token = get_api_token()

    try:
        jira = JIRA(server=jira_server, basic_auth=(username, api_token))
        log.info(f"Successfully connected to Jira server: {jira_server}")
    except Exception as e:
        log.error(f"Failed to connect to Jira: {e}")
        sys.exit(1)

    df = read_excel_data(excel_path, sheet_name)

    for index, row in df.iterrows():
        ticket_id = row['Jira ID']
        issue = get_jira_issue(jira, ticket_id)

        if issue:
            current_status = issue.fields.status.name
            log.info(f"Processing ticket {ticket_id} with current status: {current_status}")

            while current_status != "TS_READY_FOR_L1_REVIEW":
                next_status_name = status_mapping.get(current_status)

                if next_status_name:
                    if transition_issue(jira, issue, next_status_name):
                        issue = get_jira_issue(jira, ticket_id)  # Refresh issue after transition
                        current_status = issue.fields.status.name
                    else:
                        log.warning(f"Failed to transition {ticket_id} from {current_status}")
                        break  # Exit the inner loop if transition fails
                else:
                    log.warning(f"No mapping found for status: {current_status}")
                    break  # Exit the inner loop if no mapping is found

            log.info(f"Ticket {ticket_id} is transited till status: {current_status}")

    log.info(f"✅ JIRA ISSUE TRANSITED...\n")