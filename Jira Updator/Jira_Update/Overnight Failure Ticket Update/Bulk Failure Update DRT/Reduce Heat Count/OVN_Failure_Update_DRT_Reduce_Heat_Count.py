import os
import sys
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
from jira import JIRA
from datetime import datetime

# --- CONFIG ---
excel_path = "input.xlsx"
sheet_name = "Sheet1"
jira_server = "https://kpithondajapan.atlassian.net"
username = "chandrakanb@kpit.com"

# Read the actual API token from the file
with open("api_token.txt", "r") as f:
    JIRA_API_TOKEN = f.read().strip()

# --- CONNECT TO JIRA ---
jira = JIRA(server=jira_server, basic_auth=(username, JIRA_API_TOKEN))

# --- READ EXCEL FILE ---
df = pd.read_excel(excel_path, sheet_name=sheet_name, dtype=str, keep_default_na=False).fillna("NA")

# --- PROCESS EACH TICKET ---
for index, row in df.iterrows():
    try:
        ticket_id = row['Jira ID']
        test_script_id = row['Test Script ID']
        heat_count = float(row['Heat Count'])

        try:
            issue = jira.issue(ticket_id)
        except Exception as e:
            print(f"Error getting issue {ticket_id}: {e}")
            exit()
        
        # Extract the 'Heat Count' custom field
        try:
            heat_count = issue.fields.customfield_10612  # Use the custom field ID directly
            print(f"Current Heat Count: {heat_count}")
        except Exception as e:
            print(f"Error extracting Heat Count: {e}")
            
        # reduce heat count by one
        heat_count = heat_count - 1

        print(f"⏳ Updating '{ticket_id}: {test_script_id}' with New Heat Count: {heat_count}")
            
        # Update Heat Count
        if heat_count:
            issue.update(fields={
                "customfield_10612": heat_count
            })
            
    except Exception as e:
        print(f"[ERROR] Could not process {ticket_id}: '{test_script_id} : {e}.\n")