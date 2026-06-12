from jira import JIRA
import sys
import pandas as pd
import os
import logging
import requests
from requests.auth import HTTPBasicAuth
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

# --- JIRA CREDENTIALS ---
excel_path = "Transite_Issues.xlsx"  # Updated file name
sheet_name = "Sheet1"
jira_server = "https://kpithondajapan.atlassian.net"
username = "chandrakanb@kpit.com"
api_token = os.getenv("JIRA_API_TOKEN")

if api_token:
    log.info("Token loaded from Environment Variable")
else:
    if os.path.isfile("api_token.txt"):
        with open("api_token.txt", "r") as f:
            api_token = f.read().strip()
        log.info("Token loaded from api_token.txt")
    else:
        log.error("API token not found in environment variable or api_token.txt")
        raise SystemExit("❌ API token not found in environment variable or api_token.txt")

# --- CONNECT TO JIRA ---
try:
    jira = JIRA(server=jira_server, basic_auth=(username, api_token))
    log.info(f"Successfully connected to Jira server: {jira_server}")
except Exception as e:
    log.error(f"Failed to connect to Jira: {e}")
    sys.exit(1)  # Exit the script if connection fails

# --- STATUS MAPPING ---
status_mapping = {
    "TS_ASSIGNED" : "TS_READY_FOR_L1_REVIEW",
    "TS_CREATION/MODIFICATION_IN_PROGRESS" : "TS Completed",
    "TS_COMPLETED" : "Execution is in progress",
    "EXECUTION_IS_IN_PROGRESS" : "Sent for L1 Review with 1 Iterations"
}

# --- READ EXCEL FILE ---
try:
    df = pd.read_excel(excel_path, sheet_name=sheet_name, dtype=str, keep_default_na=False).fillna("NA")
    log.info(f"Successfully read Excel file: {excel_path}")
except FileNotFoundError:
    log.error(f"Excel file not found: {excel_path}")
    sys.exit(1)  # Exit the script if the file isn't found
except Exception as e:
    log.error(f"Error reading Excel file: {e}")
    sys.exit(1)

# --- PROCESS EACH TICKET ---
for index, row in df.iterrows():
    try:
        ticket_id = row['Jira ID']
        issue = jira.issue(ticket_id)
        current_status = issue.fields.status.name
        if current_status != "TS_READY_FOR_L1_REVIEW":
            log.info(f"Processing ticket {ticket_id} with current status: {current_status}")
        
        while current_status != "TS_READY_FOR_L1_REVIEW":
            next_status_name = status_mapping.get(current_status)
                
            if next_status_name:
                # Find the target status ID
                transitions = jira.transitions(issue)
                target_transition = None
                for transition in transitions:
                    if transition['name'] == next_status_name:
                        target_transition = transition
                        break

                if target_transition:
                    jira.transition_issue(issue, target_transition['id'])
                    issue = jira.issue(ticket_id)
                    new_status = issue.fields.status.name
                    log.info(f"Transitioned {ticket_id} from '{current_status}' to '{new_status}'")
                else:
                    log.warning(f"Transition '{next_status_name}' not found for {ticket_id} from '{current_status}'")
            issue = jira.issue(ticket_id)
            current_status = issue.fields.status.name
        
        issue = jira.issue(ticket_id)
        current_status = issue.fields.status.name
        log.info(f"Ticket {ticket_id} is transited till status: {current_status}")

    except Exception as e:
        log.error(f"Failed on {ticket_id}: {e}", exc_info=True)  # Log the full exception traceback

log.info(f"✅ JIRA ISSUE TRANSITED...\n")