import os
import sys
import pandas as pd
from jira import JIRA
import logging
from datetime import datetime

# --- CONFIG ---
excel_path = "OVN_Analysis.xlsx"
sheet_name = sys.argv[1]
jira_server = "https://kpithondajapan.atlassian.net"
username = "chandrakanb@kpit.com"

# Read API token
with open("api_token.txt", "r") as f:
    JIRA_API_TOKEN = f.read().strip()

# Connect to JIRA
jira = JIRA(server=jira_server, basic_auth=(username, JIRA_API_TOKEN))

# Logging setup
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

# Load Excel
df = pd.read_excel(excel_path, sheet_name=sheet_name, dtype=str, keep_default_na=False).fillna("NA")

# Process rows
for index, row in df.iterrows():
    try:
        ticket_id = row['Jira ID']
        test_script_id = row['Test Script ID']
        result = row['Result']
        failure_category = row['Failure Category']

        if ticket_id.strip() and result == "Fail":
            issue = jira.issue(ticket_id)
            fields = {}
            if failure_category:
                fields["customfield_12208"] = [{"value": str(failure_category)}]
                log.info(f"⏳ Updating Failure Category for {ticket_id} - {test_script_id}: {failure_category}")
                issue.update(fields=fields)
                log.info(f"✅ Updated Failure Category for {ticket_id} - {test_script_id}\n")
        else:
            log.info(f"⚠️ Skipped (No ID or not a 'Fail'): {test_script_id}")

    except Exception as e:
        log.error(f"[ERROR] Could not process {ticket_id} - {test_script_id}: {e}\n")
