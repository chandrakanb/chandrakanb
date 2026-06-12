import os
import sys
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
from jira import JIRA
import logging
from datetime import datetime

# --- CONFIG ---
excel_path = "failure_tickets.xlsx"
sheet_name = sys.argv[1]  
jira_server = "https://kpithondajapan.atlassian.net"
username = "chandrakanb@kpit.com"

api_token = os.getenv("JIRA_API_TOKEN")
print(api_token)
if api_token:
    print("Token loaded from Environment Variable")
else:
    if os.path.isfile("api_token.txt"):
        with open("api_token.txt", "r") as f:
            api_token = f.read().strip()
        print("Token loaded from api_token.txt")
    else:
        raise SystemExit("❌ API token not found in environment variable or api_token.txt")

# --- CONNECT TO JIRA ---
jira = JIRA(server=jira_server, basic_auth=(username, api_token))

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

# --- READ EXCEL FILE ---
df = pd.read_excel(excel_path, sheet_name=sheet_name, dtype=str, keep_default_na=False).fillna("NA")

# --- PROCESS EACH TICKET ---
for index, row in df.iterrows():
    try:
        ticket_id = row['Jira ID']
        test_script_id = row['Test Script ID']
        #test_pack_name = row['Test Pack Name']
        date = row['Date']
        result = row['Result']
        failure_category = row['Failure Category']
        failure_sub_category = row['Failure Sub Category']
        root_cause_details = row['Root Cause Details']
        fixation_performed = row['Fixation Performed']
        remarks = row['Remarks']
        reviewer = row['Reviewer']
        build_number = row['Build Number']
        
        issue = jira.issue(ticket_id)
        
        fields = {}
        fields["customfield_11462"] = str(root_cause_details)
        fields["customfield_11461"] = str(fixation_performed)
        
        issue.update(fields=fields)
        log.info(f"  ✅ Updated fields for {ticket_id}: '{test_script_id}.")
        
    except Exception as e:
        log.error(f"[ERROR] Could not process {ticket_id}: '{test_script_id} : {e}.\n")

log.info(f"  ✅ JIRA UPDATE COMPLETED...\n")