import pandas as pd
import requests
from requests.auth import HTTPBasicAuth

# Jira API credentials
base_url = "https://jira.honda-ivi.com/rest/api/2/issue/"
user = "kpt_jorawar_lu"
pw = "Iamlj!2024"
auth = HTTPBasicAuth(user, pw)

# Input Excel path
excel_file = "Jira_Update_Input.xlsx"  # Change to your file path

# Read the Excel file
df = pd.read_excel(excel_file)

# Check required columns
required_columns = ['ZephyrKey', 'BenchNumber', 'TestPackName']
if not all(col in df.columns for col in required_columns):
    print(f"Missing required columns. Required: {required_columns}")
    exit(1)

# Iterate and update each Jira ticket
for _, row in df.iterrows():
    issue_key = row['ZephyrKey']
    bench_value = row['BenchNumber']
    pack_value = row['TestPackName']

    update_url = base_url + issue_key
    payload = {
        "fields": {
            "customfield_10505": [{"value": bench_value}],
            "customfield_10502": pack_value
        }
    }

    headers = {"Content-Type": "application/json"}
    response = requests.put(update_url, json=payload, auth=auth, headers=headers)

    if response.status_code == 204:
        print(f"[SUCCESS] {issue_key} updated.")
    else:
        print(f"[FAIL] {issue_key} update failed. Status: {response.status_code}, Response: {response.text}")
