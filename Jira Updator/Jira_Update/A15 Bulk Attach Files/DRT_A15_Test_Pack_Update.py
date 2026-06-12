import os
import sys
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
from jira import JIRA
import logging
from datetime import datetime

# --- CONFIG ---
excel_path = "DRT_A15_Test_Pack_Update.xlsx"
sheet_name = "Sheet1"  
jira_server = "https://kpithondajapan.atlassian.net"
username = "chandrakanb@kpit.com"

# Read the actual API token from ENV variable or file
api_token = os.getenv("JIRA_API_TOKEN")

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
        test_pack_name = row['Test Pack Name']
        ts_modification_by= row['TS Modification By']
        ts_modifications_on= row['TS Modifications On']
        tc_created_by= row['TC Created By']
        tc_created_on= row['TC Created On']
        ts_created_by= row['TS Created by']
        ts_created_on= row['TS Created On']
        executed_by= row['Executed by']
        executed_on= row['Executed on']
        automation_tl_name= row['Automation TL Name']
        assignee= row['Assignee']
        test_pack_name= row['Test Pack Name']

        issue = jira.issue(ticket_id)
        log.info(f"⏳ Updating '{ticket_id}: {test_script_id}'...")
        
        # Clean all fields to avoid NaN and ensure proper formatting
        def clean_value(val):
            return None if pd.isna(val) else val

        ts_modification_by =clean_value(ts_modification_by)
        ts_modifications_on =clean_value(ts_modifications_on)
        tc_created_by =clean_value(tc_created_by)
        tc_created_on =clean_value(tc_created_on)
        ts_created_by =clean_value(ts_created_by)
        ts_created_on =clean_value(ts_created_on)
        executed_by =clean_value(executed_by)
        executed_on =clean_value(executed_on)
        automation_tl_name =clean_value(automation_tl_name)
        assignee =clean_value(assignee)
        test_pack_name =clean_value(test_pack_name)

        # Update fields
        fields = {}            
        if ts_modification_by:
            fields[" customfield_14170"] = str(ts_modification_by)
            log.info(f"  🔄 Updating TS Modification By for {ticket_id}: '{test_script_id} with {ts_modification_by}")
        if ts_modifications_on:
            fields["customfield_14171"] = str(ts_modifications_on)
            log.info(f"  🔄 Updating TS Modifications On for {ticket_id}: '{test_script_id} with {ts_modifications_on}")
        if tc_created_by:
            fields["customfield_11003"] = str(tc_created_by)
            log.info(f"  🔄 Updating TC Created By for {ticket_id}: '{test_script_id} with {tc_created_by}")
        if tc_created_on:
            fields["customfield_11008"] = str(tc_created_on)
            log.info(f"  🔄 Updating TC Created On for {ticket_id}: '{test_script_id} with {tc_created_on}")
        if ts_created_by:
            fields["customfield_11088"] = str(ts_created_by)
            log.info(f"  🔄 Updating TS Created by for {ticket_id}: '{test_script_id} with {ts_created_by}")
        if ts_created_on:
            fields["customfield_11464"] = str(ts_created_on)
            log.info(f"  🔄 Updating TS Created On for {ticket_id}: '{test_script_id} with {ts_created_on}")
        if executed_by:
            fields["customfield_12084"] = str(executed_by)
            log.info(f"  🔄 Updating Executed by for {ticket_id}: '{test_script_id} with {executed_by}")
        if executed_on:
            fields["customfield_12085"] = str(executed_on)
            log.info(f"  🔄 Updating Executed on for {ticket_id}: '{test_script_id} with {executed_on}")
        if automation_tl_name:
            fields["customfield_11087"] = str(automation_tl_name)
            log.info(f"  🔄 Updating Automation TL Name for {ticket_id}: '{test_script_id} with {automation_tl_name}")
        if test_pack_name:
            fields["customfield_11042"] = str(test_pack_name)
            log.info(f"  🔄 Updating Test Pack Name for {ticket_id}: '{test_script_id} with {test_pack_name}")
        
        issue.update(fields=fields)
        log.info(f"  🔄 Updated TS Modification By for {ticket_id}: '{test_script_id} with {ts_modification_by}")
        log.info(f"  🔄 Updated TS Modifications On for {ticket_id}: '{test_script_id} with {ts_modifications_on}")
        log.info(f"  🔄 Updated TC Created By for {ticket_id}: '{test_script_id} with {tc_created_by}")
        log.info(f"  🔄 Updated TC Created On for {ticket_id}: '{test_script_id} with {tc_created_on}")
        log.info(f"  🔄 Updated TS Created by for {ticket_id}: '{test_script_id} with {ts_created_by}")
        log.info(f"  🔄 Updated TS Created On for {ticket_id}: '{test_script_id} with {ts_created_on}")
        log.info(f"  🔄 Updated Executed by for {ticket_id}: '{test_script_id} with {executed_by}")
        log.info(f"  🔄 Updated Executed on for {ticket_id}: '{test_script_id} with {executed_on}")
        log.info(f"  🔄 Updated Automation TL Name for {ticket_id}: '{test_script_id} with {automation_tl_name}")
        log.info(f"  🔄 Updated Test Pack Name for {ticket_id}: '{test_script_id} with {test_pack_name}")
        
    except Exception as e:
        log.error(f"[ERROR] Could not process {ticket_id}: '{test_script_id} : {e}.\n")
    