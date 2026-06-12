    import os
    import pandas as pd
    import requests
    from requests.auth import HTTPBasicAuth

    # ---------------- CONFIG ----------------
    JIRA_URL = "https://kpithondajapan.atlassian.net"
    JIRA_USER_EMAIL = "chandrakanb@kpit.com"

    JIRA_API_TOKEN = os.getenv("JIRA_API_TOKEN")
    if not JIRA_API_TOKEN:
        raise SystemExit("❌ JIRA_API_TOKEN environment variable not set")

EXCEL_FILE = "input.xlsx"
# ----------------------------------------

if not os.path.isfile(EXCEL_FILE):
    raise SystemExit(f"❌ Input Excel not found: {EXCEL_FILE}")

df = pd.read_excel(EXCEL_FILE)

session = requests.Session()
session.auth = HTTPBasicAuth(JIRA_USER_EMAIL, JIRA_API_TOKEN)
session.headers.update({"X-Atlassian-Token": "no-check"})

for _, row in df.iterrows():
    jira_id = row.get("JIRA_ID")

    if pd.isna(jira_id) or str(jira_id).strip() == "":
        print("❌ Missing JIRA_ID, skipping row")
        continue

    jira_id = str(jira_id).strip()

    pdf_file_name = str(row.get("File Name", "")).strip()
    comment_text = row.get("Comment", "Please see the attached file.")
    build_number = str(row.get("Build Number", "")).strip()
    date = str(row.get("Date", "")).strip()

    attach_url = f"{JIRA_URL}/rest/api/2/issue/{jira_id}/attachments"

    files_to_send = []

    pdf_file_path = os.path.join("files", pdf_file_name) if pdf_file_name else ""

    if pdf_file_name:
        if os.path.isfile(pdf_file_path):
            print(f"ℹ️ Found PDF: {pdf_file_name}")
            files_to_send.append(
                ("file", (pdf_file_name, open(pdf_file_path, "rb"), "application/pdf"))
            )
        else:
            print(f"❌ PDF not found: {pdf_file_name} | {jira_id}")

    attached_names = []

    if files_to_send:
        try:
            attach_resp = session.post(attach_url, files=files_to_send, timeout=60)
        finally:
            for _, filetuple in files_to_send:
                try:
                    filetuple[1].close()
                except Exception:
                    pass

        if attach_resp.status_code in (200, 201):
            try:
                resp_json = attach_resp.json()
                attached_names = [item.get("filename") for item in resp_json]
                for name in attached_names:
                    print(f"✅ Attached: {name} -> {jira_id}")
            except Exception:
                print(f"⚠️ Attach success but no JSON for {jira_id}")
        else:
            print(f"❌ Attachment failed for {jira_id} | {attach_resp.status_code}")
            continue
    else:
        print(f"ℹ️ No files found to attach for {jira_id}")

    # ---------------- COMMENT ONLY IF FILES ATTACHED ----------------
    if attached_names:
        file_wiki = "\n".join([f"!{name}!" for name in attached_names])
        final_comment = f"{comment_text}\nDate: {date}\nBuild: {build_number}\n{file_wiki}"

        comment_url = f"{JIRA_URL}/rest/api/2/issue/{jira_id}/comment"
        payload = {"body": final_comment}

        comment_resp = session.post(
            comment_url,
            json=payload,
            headers={"Content-Type": "application/json"},
            timeout=30
        )

        if comment_resp.status_code in (200, 201):
            print(f"💬 Comment posted for {jira_id}\n")
            print(f{final_comment})
        else:
            print(f"❌ Comment failed for {jira_id} | {comment_resp.status_code}\n")
    else:
        print(f"⏭️ Comment skipped for {jira_id} (no attachments)\n")
