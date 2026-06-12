import os
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth

# === CONFIGURATION ===
JIRA_URL = "https://kpithondajapan.atlassian.net"
JIRA_USER_EMAIL = "sahilb@kpit.com"
api_token = "api_token.txt"

# Read the actual API token from the file
with open(api_token, "r") as f:
    JIRA_API_TOKEN = f.read().strip()

EXCEL_FILE = os.path.join("input", "input.xlsx")
FILES_DIR = os.path.join("files")

# === READ EXCEL ===
df = pd.read_excel(EXCEL_FILE)

for _, row in df.iterrows():
    jira_id = row["JIRA_ID"]
    file_name = row["File_Name"]
    comment_text = row.get("Comment", "Please see the attached file.")
    file_path = os.path.join(FILES_DIR, file_name)

    if not os.path.isfile(file_path):
        print(f"❌ File not found: {file_name}")
        continue

    # Step 1: Attach file
    attach_url = f"{JIRA_URL}/rest/api/2/issue/{jira_id}/attachments"
    headers = {"X-Atlassian-Token": "no-check"}
    with open(file_path, "rb") as f:
        files = {"file": (file_name, f, "application/pdf")}
        attach_response = requests.post(
            attach_url,
            headers=headers,
            auth=HTTPBasicAuth(JIRA_USER_EMAIL, JIRA_API_TOKEN),
            files=files
        )

    if attach_response.status_code not in [200, 201]:
        print(f"❌ Failed to attach {file_name} to {jira_id}")
        continue

    # Step 2: Add comment (standard text comment)
    comment_url = f"{JIRA_URL}/rest/api/2/issue/{jira_id}/comment"
    comment_payload = {
        "body": comment_text  # Just plain text
    }

    comment_response = requests.post(
        comment_url,
        json=comment_payload,
        auth=HTTPBasicAuth(JIRA_USER_EMAIL, JIRA_API_TOKEN),
        headers={"Content-Type": "application/json"}
    )

    if comment_response.status_code in [200, 201]:
        print(f"✅ Successfully attached and commented on {jira_id}")
    else:
        print(f"❌ Comment failed for {jira_id} | Status: {comment_response.status_code} | Response: {comment_response.text}")
