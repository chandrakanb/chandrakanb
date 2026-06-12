import os
import sys
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
from jira import JIRA
import logging
from datetime import datetime

# --- CONFIG ---
excel_path = "OVN_Analysis.xlsx"
sheet_name = sys.argv[1]
jira_server = "https://kpithondajapan.atlassian.net"
username = "chandrakanb@kpit.com"

# --- AUTHENTICATION ---
with open("api_token.txt", "r") as f:
    JIRA_API_TOKEN = f.read().strip()

jira = JIRA(server=jira_server, basic_auth=(username, JIRA_API_TOKEN))

# --- LOGGING SETUP ---
os.makedirs("logs", exist_ok=True)
timestamp = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
log_file = os.path.join("logs", f"{timestamp}.log")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[logging.FileHandler(log_file, encoding='utf-8'), logging.StreamHandler(sys.stdout)]
)
log = logging.getLogger()

# --- LOAD EXCEL ---
df = pd.read_excel(excel_path, sheet_name=sheet_name, dtype=str, keep_default_na=False).fillna("NA")

# --- PROCESS EACH ROW ---
for idx, row in df.iterrows():
    try:
        ticket_id = row['Jira ID']
        test_script_id = row['Test Script ID']
        date = row['Date']
        analyst = row['Analysis Done By']
        result = row['Result'].strip()
        failure_category = row['Failure Category']
        failure_sub_category = row['Failure Sub Category']
        root_cause = row['Root Cause Details']
        fixation = row['Fixation Performed']

        if not ticket_id.strip():
            log.warning(f"⚠️ Missing Jira ID for '{test_script_id}' on {date}")
            continue

        issue = jira.issue(ticket_id)
        log.info(f"🔧 Processing: {ticket_id} - {test_script_id} | Result: {result}")

        # Clean values
        def clean(val): return val.strip() if val and val.strip().lower() != "na" else None
        failure_category = clean(failure_category)
        failure_sub_category = clean(failure_sub_category)
        root_cause = clean(root_cause)
        fixation = clean(fixation)

        # --- FIELD UPDATES ---
        if result.lower() == "fail":
            fields = {}

            # Handle Previous Failure Category (multi-select, capitalized)
            if failure_category:
                new_vals = [v.strip() for v in failure_category.split(",") if v.strip()]
                capitalized_vals = [' '.join(w.capitalize() for w in val.split()) for val in new_vals]
                existing_vals = [v.value for v in (issue.fields.customfield_12208 or [])]
                merged_vals = sorted(set(existing_vals + capitalized_vals))
                fields["customfield_12208"] = [{"value": v} for v in merged_vals]

            if failure_sub_category:
                fields["customfield_11460"] = failure_sub_category
            if root_cause:
                fields["customfield_11462"] = root_cause
            if fixation:
                fields["customfield_11461"] = fixation

            if fields:
                log.info(f"📥 Field Updates for {ticket_id}:")
                for key, val in fields.items():
                    log.info(f"    {key} : {val}")
                issue.update(fields=fields)
                log.info(f"✅ Fields updated for {ticket_id}\n")

        # --- COMMENTING ---
        if result.lower() == "fail":
            comment_text = (
                f"| *Date* | *Analysis Done By* | *Result* | *Failure Category* | *Failure Sub Category* | *Root Cause Details* | *Fixation Performed* |\n"
                f"| {date} | {analyst} | {result} | {failure_category} | {failure_sub_category} | {root_cause} | {fixation} |"
            )
        else:
            comment_text = (
                f"| *Date* | *Analysis Done By* | *Result* |\n"
                f"| {date} | {analyst} | {result} |"
            )

        comment_url = f"{jira_server}/rest/api/2/issue/{ticket_id}/comment"
        comment_response = requests.post(
            comment_url,
            json={"body": comment_text},
            auth=HTTPBasicAuth(username, JIRA_API_TOKEN),
            headers={"Content-Type": "application/json"}
        )

        if comment_response.status_code in [200, 201]:
            log.info(f"💬 Comment added for {ticket_id}: {test_script_id}")
            log.info(f"    ↪ {comment_text}")
        else:
            log.error(f"❌ Comment failed for {ticket_id} | Status: {comment_response.status_code} | Response: {comment_response.text}")

    except Exception as e:
        log.error(f"[ERROR] Could not process {ticket_id} - {test_script_id}: {e}")
