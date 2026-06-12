import os
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth

JIRA_URL = "https://kpithondajapan.atlassian.net"
JIRA_USER_EMAIL = "sahilb@kpit.com"
api_token = "api_token.txt"

if not os.path.isfile(API_TOKEN_FILE):
    raise SystemExit(f"API token file not found: {API_TOKEN_FILE}")
with open(API_TOKEN_FILE, "r") as f:
    JIRA_API_TOKEN = f.read().strip()
if not JIRA_API_TOKEN:
    raise SystemExit("API token is empty")

EXCEL_FILE = os.path.join("input", "input.xlsx")
FILES_DIR = os.path.join("files")

if not os.path.isfile(EXCEL_FILE):
    raise SystemExit(f"Input Excel not found: {EXCEL_FILE}")
df = pd.read_excel(EXCEL_FILE)

session = requests.Session()
session.auth = HTTPBasicAuth(JIRA_USER_EMAIL, JIRA_API_TOKEN)
session.headers.update({"X-Atlassian-Token": "no-check"})

for _, row in df.iterrows():
    jira_id = row.get("JIRA_ID")
    if pd.isna(jira_id) or str(jira_id).strip() == "":
        print("❌ Missing JIRA_ID in row, skipping")
        continue
    jira_id = str(jira_id).strip()

    pdf_file_name = str(row.get("File_Name", "")).strip()
    excel_file_name = str(row.get("Excel_Name", "")).strip()
    comment_text = row.get("Comment", "Please see the attached file.")

    attach_url = f"{JIRA_URL}/rest/api/2/issue/{jira_id}/attachments"

    files_to_send = []
    pdf_file_path = os.path.join(FILES_DIR, pdf_file_name) if pdf_file_name else ""
    excel_file_path = os.path.join(FILES_DIR, excel_file_name) if excel_file_name else ""

    if pdf_file_name:
        if os.path.isfile(pdf_file_path):
            print(f"ℹ️ Found PDF: {pdf_file_name} (will attach)")
            files_to_send.append(("file", (pdf_file_name, open(pdf_file_path, "rb"), "application/pdf")))
        else:
            print(f"❌ PDF not found: {pdf_file_name} (JIRA: {jira_id})")
    if excel_file_name:
        if os.path.isfile(excel_file_path):
            print(f"ℹ️ Found Excel: {excel_file_name} (will attach)")
            files_to_send.append(("file", (excel_file_name, open(excel_file_path, "rb"), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")))
        else:
            print(f"❌ Excel not found: {excel_file_name} (JIRA: {jira_id})")

    if files_to_send:
        try:
            attach_resp = session.post(attach_url, files=files_to_send, timeout=60)
        finally:
            for _, filetuple in files_to_send:
                fileobj = filetuple[1]
                try:
                    fileobj.close()
                except Exception:
                    pass
        if attach_resp.status_code in (200, 201):
            try:
                resp_json = attach_resp.json()
                attached_names = [item.get("filename") for item in resp_json] if isinstance(resp_json, list) else []
                if attached_names:
                    for name in attached_names:
                        print(f"✅ Attached: {name} -> {jira_id}")
                else:
                    print(f"✅ Attach succeeded for {jira_id} (no filenames returned)")
            except Exception:
                print(f"✅ Attach succeeded for {jira_id} (response not JSON)")
        else:
            print(f"❌ Failed to attach files to {jira_id} | Status: {attach_resp.status_code} | {attach_resp.text}")
            continue
    else:
        print(f"ℹ️ No files to attach for {jira_id}")

    comment_url = f"{JIRA_URL}/rest/api/2/issue/{jira_id}/comment"
    comment_payload = {"body": comment_text}
    comment_resp = session.post(comment_url, json=comment_payload, headers={"Content-Type": "application/json"}, timeout=30)
    if comment_resp.status_code in (200, 201):
        print(f"✅ Comment posted for {jira_id}\n")
    else:
        print(f"❌ Comment failed for {jira_id} | Status: {comment_resp.status_code} | {comment_resp.text}\n")
