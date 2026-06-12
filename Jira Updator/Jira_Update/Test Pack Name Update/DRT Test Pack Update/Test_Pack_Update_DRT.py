import os
import sys
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
from jira import JIRA
import logging
from datetime import datetime

# --- CONFIG ---
excel_path = "Test_Pack_Update_DRT.xlsx"
sheet_name = "Sheet1"  
jira_server = "https://kpithondajapan.atlassian.net"
username = "chandrakanb@kpit.com"

# Read the actual API token from the file
JIRA_API_TOKEN = os.getenv("JIRA_API_TOKEN")
 
if JIRA_API_TOKEN:
    print("Token loaded from Environment Variable")
else:
    if os.path.isfile("JIRA_API_TOKEN.txt"):
        with open("JIRA_API_TOKEN.txt", "r") as f:
            JIRA_API_TOKEN = f.read().strip()
        print("Token loaded from JIRA_API_TOKEN.txt")
    else:
        raise SystemExit("❌ API token not found in environment variable or JIRA_API_TOKEN.txt")

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
        test_pack_name = row['Test Pack Name']

        issue = jira.issue(ticket_id)
        log.info(f"⏳ Updating '{ticket_id}: {test_script_id}'...")
        
        # Clean all fields to avoid NaN and ensure proper formatting
        def clean_value(val):
            return None if pd.isna(val) else val

        test_pack_name = clean_value(test_pack_name)

        # Update fields
        fields = {}
        if test_pack_name:
            fields["customfield_10502"] = str(test_pack_name)
        
        log.info(f"  🔄 Updating Test Pack for {ticket_id}: '{test_script_id}...")
        log.info(f"    Test Pack Name : {test_pack_name}")
        issue.update(fields=fields)
        log.info(f"  ✅ Updated Test Pack for {ticket_id}: '{test_script_id}.\n")

    except Exception as e:
        log.error(f"[ERROR] Could not process {ticket_id}: '{test_script_id} : {e}.\n")
