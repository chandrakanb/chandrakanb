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

# Capitalize helper
def cap(text):
    return ' '.join(word.capitalize() for word in text.strip().split()) if text and text.strip().lower() != "na" else None

# Load Excel
df = pd.read_excel(excel_path, sheet_name=sheet_name, dtype=str, keep_default_na=False).fillna("NA")

# Process rows
for index, row in df.iterrows():
    try:
        ticket_id = row['Jira ID'].strip()
        test_script_id = row['Test Script ID']
        result = row['Result'].strip()

        if not ticket_id or result.lower() != "fail":
            log.info(f"⚠️ Skipping {test_script_id} (no ID or not Fail)")
            continue

        failure_cat1 = row.get('Failure Category1', '').strip()
        failure_cat2 = row.get('Failure Category2', '')
        failure_cat3 = row.get('Failure Category3', '')

        issue = jira.issue(ticket_id)
        fields = {}

        # --- Set Current Failure Category (as-is) ---
        if failure_cat1 and failure_cat1.lower() != "na":
            fields["customfield_11454"] = {"value": failure_cat1}

        # --- Get existing Previous Failure Category values ---
        existing_prev = getattr(issue.fields, "customfield_12208", [])
        existing_prev_vals = set((v.value) for v in existing_prev if v and v.value)

        # --- Merge with Category2 & 3 ---
        to_add = {v for v in [failure_cat2, failure_cat3] if v}
        final_prev = sorted(existing_prev_vals.union(to_add))

        fields["customfield_12208"] = [{"value": v} for v in final_prev]

        # --- Update JIRA ---
        log.info(f"⏳ Updating {ticket_id} - {test_script_id}")
        log.info(f"    ➤ Current Failure Category : {failure_cat1}")
        log.info(f"    ➤ Previous Failure Category: {final_prev}")
        issue.update(fields=fields)
        log.info(f"✅ Updated {ticket_id} - {test_script_id}\n")

    except Exception as e:
        log.error(f"[ERROR] Could not process {ticket_id} - {test_script_id}: {e}\n")
