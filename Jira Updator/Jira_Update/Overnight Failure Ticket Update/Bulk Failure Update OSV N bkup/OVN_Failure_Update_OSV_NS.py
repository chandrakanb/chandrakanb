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
username = "Neha.Suryawanshi@kpit.com"

# Read the actual API token from the file
with open("api_token_ns.txt", "r") as f:
    JIRA_API_TOKEN = f.read().strip()
print(JIRA_API_TOKEN)

# --- CONNECT TO JIRA ---
jira = JIRA(server=jira_server, basic_auth=(username, JIRA_API_TOKEN))

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

        issue = jira.issue(ticket_id)
        log.info(f"⏳ Updating '{ticket_id}: {test_script_id}'...")
        
        # Clean all fields to avoid NaN and ensure proper formatting
        def clean_value(val):
            return None if pd.isna(val) else val

        #test_pack_name = clean_value(test_pack_name)
        failure_category = clean_value(failure_category)
        failure_sub_category = clean_value(failure_sub_category)
        root_cause_details = clean_value(root_cause_details)
        fixation_performed = clean_value(fixation_performed)

        # Update fields
        fields = {}
        #if test_pack_name:
        #    fields["customfield_11042"] = str(test_pack_name)
        """if failure_category:
            fields["customfield_11454"] = {"value": str(failure_category)}
        if failure_sub_category:
            fields["customfield_11460"] = str(failure_sub_category)
        if root_cause_details:
            fields["customfield_11462"] = str(root_cause_details)
        if fixation_performed:
            fields["customfield_11461"] = str(fixation_performed)"""
        
        log.info(f"  🔄 Updating fields for {ticket_id}: '{test_script_id}...")
        #log.info(f"    Test Pack Name : {test_pack_name}")
        log.info(f"    Failure Category : {failure_category}")
        log.info(f"    Failure Sub Category : {failure_sub_category}")
        log.info(f"    Root Cause Details : {root_cause_details}")
        log.info(f"    Fixation Performed : {fixation_performed}")
        issue.update(fields=fields)
        log.info(f"  ✅ Updated fields for {ticket_id}: '{test_script_id}.")

        # Comment message
        comment_text = (
            f" |  *Date*  |  *Result*  |  *Failure Category*  |  *Failure Sub Category*  |  *Root Cause Details*  |  *Fixation Performed*  | \n"
            f" |  {date}  |  {result}  |  {failure_category}  |  {failure_sub_category}  |  {root_cause_details}  |  {fixation_performed}  | "
        )
        
        # Add comment
        log.info(f"  🔄 Commenting on {ticket_id}: '{test_script_id}...")
        log.info( f"    Date : {date}" )
        log.info( f"    Result : {result}" )
        log.info( f"    Failure Category : {failure_category}" )
        log.info( f"    Failure Sub Category : {failure_sub_category}" )
        log.info( f"    Root Cause Details : {root_cause_details}" )
        log.info( f"    Fixation Performed : {fixation_performed}" )
        comment_url = f"{jira_server}/rest/api/2/issue/{ticket_id}/comment"
        comment_response = requests.post(
            comment_url,
            json={"body": comment_text},
            auth=HTTPBasicAuth(username, JIRA_API_TOKEN),
            headers={"Content-Type": "application/json"}
        )

        if comment_response.status_code in [200, 201]:
            log.info(f"  ✅ Commented on {ticket_id}: '{test_script_id}.\n")
        else:
            log.info(f"❌ Comment failed for {ticket_id}: '{test_script_id} | Status: {comment_response.status_code} | Response: {comment_response.text}.\n")

    except Exception as e:
        log.error(f"[ERROR] Could not process {ticket_id}: '{test_script_id} : {e}.\n")
