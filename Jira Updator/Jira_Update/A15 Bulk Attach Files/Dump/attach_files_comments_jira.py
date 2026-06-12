import os
import sys
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
from jira import JIRA
import logging
from datetime import datetime

# --- Configuration (Environment Variables) ---
JIRA_URL = os.getenv("JIRA_URL")
JIRA_USER_EMAIL = os.getenv("JIRA_USER_EMAIL")
JIRA_API_TOKEN = os.getenv("JIRA_API_TOKEN")
EXCEL_FILE = os.path.join("input", "input.xlsx")
FILES_DIR = os.path.join("files")

# --- Logging Setup (as before) ---
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

# --- Assignee Mapping (Move to config file in production) ---
ASSIGNEE_MAPPING = {
    "Amit Yadav": "712020:b6811490-5959-4378-99ab-e920231f0c7c",
    "Lumesh Jorwar": "712020:9000a5c2-2cc0-40f8-a94d-4c00bc685071",
    "Yukta Hindurao Rane": "712020:5a091caa-0568-47a1-8e2e-9d6959c43015",
    "Rajashri Bhamare": "712020:b70ff7b8-8f22-4c1e-a03c-791ef73ac3ef",
    "Sahil Bhalekar": "712020:5e45192c-a1d0-4e92-b528-97cb5d418265",
    "CHANDRAKANT PRADIP BHALEKAR": "712020:7ebd6444-898e-4d1e-ba7a-2710a262ffa7",
    "Akash Bobade": "712020:79dea246-e8bf-460f-908c-d7b1d8b836ac",
    "Aishwarya Dhole": "712020:12ecdb66-fe0f-45de-bbfc-99054c287639",
    "Srushti Potadar": "712020:6083a942-b00a-4e8d-9f97-2949883f07b6",
    "Amol Chandrakant Kore": "712020:351350cc-71d7-4581-be87-923dd393501d",
    "Gayathri G": "712020:6121a91a-06b2-4ab2-ba3f-6d56c491eadc"
}

# --- Functions ---

def read_excel_data(excel_file):
    """Reads data from the Excel file."""
    try:
        df = pd.read_excel(excel_file)
        return df
    except FileNotFoundError:
        log.error(f"Error: Excel file not found: {excel_file}")
        raise

def authenticate_jira(jira_url, user_email, api_token):
    """Authenticates with the Jira API."""
    try:
        jira = JIRA(
            server=jira_url,
            basic_auth=(user_email, api_token)
        )
        session = requests.Session()
        session.auth = HTTPBasicAuth(user_email, api_token)
        session.headers.update({"X-Atlassian-Token": "no-check"})
        return jira, session
    except Exception as e:
        log.error(f"Error authenticating with Jira: {e}")
        raise

def update_jira_issue(jira_issue, row_data):
    """Updates the Jira issue with data from the Excel row."""

    # Clean all fields to avoid NaN and ensure proper formatting
    def clean_value(val):
        return None if pd.isna(val) else val

    # Format date fields
    def format_date(date_string):
        try:
            date_object = datetime.strptime(date_string, '%Y-%m-%d %H:%M:%S')
            formatted_date = date_object.strftime('%a, %d %b %Y %H:%M:%S +0000')
            return formatted_date
        except ValueError:
            return "Invalid date format. Please use YYYY-MM-DD HH:MM:SS"

    ts_modification_by = clean_value(row_data.get('TS Modification By', ""))
    ts_modifications_on = clean_value(row_data.get('TS Modifications On', ""))
    tc_created_by = clean_value(row_data.get('TC Created By', ""))
    tc_created_on = clean_value(row_data.get('TC Created On', ""))
    ts_created_by = clean_value(row_data.get('TS Created by', ""))
    ts_created_on = clean_value(row_data.get('TS Created On', ""))
    executed_by = clean_value(row_data.get('Executed by', ""))
    executed_on = clean_value(row_data.get('Executed on', ""))
    automation_tl_name = clean_value(row_data.get('Automation TL Name', ""))
    assignee = clean_value(row_data.get('Assignee', ""))
    test_pack_name = clean_value(row_data.get('Test Pack Name', ""))

    fields = {}
    if ts_modification_by:
        fields["customfield_14170"] = str(ts_modification_by)
        log.info(f"  🔄 Updating TS Modification By: {ts_modification_by}")
    if ts_modifications_on:
        ts_modifications_on = format_date(ts_modifications_on)
        fields["customfield_14171"] = str(ts_modifications_on)
        log.info(f"  🔄 Updating TS Modifications On: {ts_modifications_on}")
    if tc_created_by:
        fields["customfield_11003"] = str(tc_created_by)
        log.info(f"  🔄 Updating TC Created By: {tc_created_by}")
    if tc_created_on:
        tc_created_on = format_date(tc_created_on)
        fields["customfield_11008"] = str(tc_created_on)
        log.info(f"  🔄 Updating TC Created On: {tc_created_on}")
    if ts_created_by:
        fields["customfield_11088"] = str(ts_created_by)
        log.info(f"  🔄 Updating TS Created by: {ts_created_by}")
    if ts_created_on:
        ts_created_on = format_date(ts_created_on)
        fields["customfield_11464"] = str(ts_created_on)
        log.info(f"  🔄 Updating TS Created On: {ts_created_on}")
    if executed_by:
        fields["customfield_12084"] = str(executed_by)
        log.info(f"  🔄 Updating Executed by: {executed_by}")
    if executed_on:
        executed_on = format_date(executed_on)
        fields["customfield_12085"] = str(executed_on)
        log.info(f"  🔄 Updating Executed on: {executed_on}")
    if automation_tl_name:
        fields["customfield_11087"] = str(automation_tl_name)
        log.info(f"  🔄 Updating Automation TL Name: {automation_tl_name}")
    if assignee:
        fields["assignee"] = ASSIGNEE_MAPPING.get(assignee, "-1")  # Use mapping
        log.info(f"  🔄 Updating Assignee: {assignee}")
    if test_pack_name:
        fields["customfield_11042"] = str(test_pack_name)
        log.info(f"  🔄 Updating Test Pack Name: {test_pack_name}")

    try:
        jira_issue.update(fields=fields)
        log.info(f"  ✅ Updated fields for {jira_issue.key}")
    except Exception as e:
        log.error(f"[ERROR] Could not update issue {jira_issue.key}: {e}")

def attach_files_to_jira(jira_issue, pdf_file_name, excel_file_name, files_dir, session):
    """Attaches PDF and Excel files to the Jira issue."""
    attach_url = f"{JIRA_URL}/rest/api/2/issue/{jira_issue.key}/attachments"
    files_to_send = []

    pdf_file_path = os.path.join(files_dir, "PDF", pdf_file_name) if pdf_file_name else ""
    excel_file_path = os.path.join(files_dir, "EXCEL", excel_file_name) if excel_file_name else ""

    if pdf_file_name:
        if os.path.isfile(pdf_file_path):
            log.info(f"ℹ️ Found PDF: {pdf_file_name} (will attach)")
            files_to_send.append(("file", (pdf_file_name, open(pdf_file_path, "rb"), "application/pdf")))
        else:
            log.info(f"❌ PDF not found: {pdf_file_name} (JIRA: {jira_issue.key})")
    if excel_file_name:
        if os.path.isfile(excel_file_path):
            log.info(f"ℹ️ Found Excel: {excel_file_name} (will attach)")
            files_to_send.append(("file", (excel_file_name, open(excel_file_path, "rb"),
                                             "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")))
        else:
            log.info(f"❌ Excel not found: {excel_file_name} (JIRA: {jira_issue.key})")

    attached_names = []
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
                for name in attached_names:
                    log.info(f"✅ Attached: {name} -> {jira_issue.key}")
            except Exception:
                log.warning(f"⚠️ Attach succeeded for {jira_issue.key} (no JSON returned)")
        else:
            log.info(f"❌ Failed to attach files to {jira_issue.key} | Status: {attach_resp.status_code} | {attach_resp.text}")
            return []  # Return empty list if attachment fails
    return attached_names

def add_comment_to_jira(jira_issue, comment_text, attached_names, session):
    """Adds a comment to the Jira issue with file previews."""
    comment_url = f"{JIRA_URL}/rest/api/2/issue/{jira_issue.key}/comment"
    payload = {
        "body": comment_text
    }

    if attached_names:
        file_wiki_section = "\n".join([f"!{name}!" for name in attached_names])
        comment_text = f"{comment_text}\n\nAttached Files:\n{file_wiki_section}"

    comment_resp = session.post(comment_url, json=payload, headers={"Content-Type": "application/json"}, timeout=30)

    if comment_resp.status_code in (200, 201):
        log.info(f"💬 Comment posted with file preview for {jira_issue.key}\n")
    else:
        log.info(f"❌ Comment failed for {jira_issue.key} | Status: {comment_resp.status_code} | {comment_resp.text}\n")

# --- Main Function ---

def main():
    """Main function to orchestrate the process."""
    try:
        df = read_excel_data(EXCEL_FILE)
        jira, session = authenticate_jira(JIRA_URL, JIRA_USER_EMAIL, JIRA_API_TOKEN)

        for _, row in df.iterrows():
            ticket_id = row.get("Jira ID")
            if pd.isna(ticket_id) or str(ticket_id).strip() == "":
                log.info("❌ Missing JIRA_ID in row, skipping")
                continue
            ticket_id = str(ticket_id).strip()

            try:
                issue = jira.issue(ticket_id)
                log.info(f"⏳ Updating '{ticket_id}'...")
            except Exception as e:
                log.error(f"[ERROR] Could not find issue {ticket_id}: {e}")
                continue

            update_jira_issue(issue, row)
            attached_names = attach_files_to_jira(issue, str(row.get("File Name", "")).strip(),
                                                  str(row.get("Excel Name", "")).strip(), FILES_DIR, session)
            comment_text = row.get("Comment", "Please see the attached file.")
            add_comment_to_jira(issue, comment_text, attached_names, session)

    except Exception as e:
        log.error(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    main()