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
username = "chandrakanb@kpit.com"

api_token = os.getenv("JIRA_API_TOKEN")
print(api_token)
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
        #test_pack_name = row['Test Pack Name']
        date = row['Date']
        result = row['Result']
        failure_category = row['Failure Category']
        failure_sub_category = row['Failure Sub Category']
        root_cause_details = row['Root Cause Details']
        fixation_performed = row['Fixation Performed']
        remarks = row['Remarks']
        reviewer = row['Reviewer']
        build_number = row['Build Number']
        
        issue = jira.issue(ticket_id)
        
        if result == "Fail":
            # Extract, combine, and deduplicate failure category values using sets for uniqueness. 
            combined_values = []
            current_failure_category = issue.fields.customfield_10851
            if current_failure_category != None:
                current_failure_values = [option.value for option in current_failure_category]
                previous_failure_category = issue.fields.customfield_12208
                if previous_failure_category != None:
                    existing_previous_values = [option.value for option in previous_failure_category]
                    combined_values = list(set(existing_previous_values + current_failure_values))
                else:
                    combined_values = current_failure_values
                normalized_values = []
                for value in combined_values:
                    if value == "KITE Issue":
                        # normalized_values.append("KITE Issues")
                        normalized_values.append("KITE Issue")
                    else:
                        normalized_values.append(value)
                combined_values = list(set(normalized_values))

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
            #    fields["customfield_10502"] = str(test_pack_name)
            if failure_category:
                fields["customfield_10851"] = [{"value": str(failure_category)}]
            if failure_sub_category:
                fields["customfield_10617"] = str(failure_sub_category)
            if root_cause_details:
                fields["customfield_10614"] = str(root_cause_details)
            if fixation_performed:
                fields["customfield_10618"] = str(fixation_performed)
            if combined_values:
                fields["customfield_12208"] = [{"value": str(v)} for v in combined_values]
            print(fields)
            log.info(f"  🔄 Updating fields for {ticket_id}: '{test_script_id}...")
            #log.info(f"    Test Pack Name : {test_pack_name}")
            log.info(f"    Failure Category : {failure_category}")
            log.info(f"    Failure Sub Category : {failure_sub_category}")
            log.info(f"    Previous Failure Category : {combined_values}")
            log.info(f"    Root Cause Details : {root_cause_details}")
            log.info(f"    Fixation Performed : {fixation_performed}")
            issue.update(fields=fields)
            log.info(f"  ✅ Updated fields for {ticket_id}: '{test_script_id}.")

            # Comment message
            comment_text = (
                f" |  *Date*  |  *Result*  |  *Failure Category*  |  *Failure Sub Category*  |  *Root Cause Details*  |  *Fixation Performed*  | *Remarks* | *Reviewer* | \n"
                f" |  {date}  *[{build_number}]*  |  {result}  |  {failure_category}  |  {failure_sub_category}  |  {root_cause_details}  |  {fixation_performed}  | {remarks} | {reviewer} | "
            )
            
            # Add comment
            log.info(f"  🔄 Commenting on {ticket_id}: '{test_script_id}...")
            log.info( f"    Date {build_number} : {date}" )
            log.info( f"    Result : {result}" )
            log.info( f"    Failure Category : {failure_category}" )
            log.info( f"    Failure Sub Category : {failure_sub_category}" )
            log.info( f"    Root Cause Details : {root_cause_details}" )
            log.info( f"    Fixation Performed : {fixation_performed}" )
            log.info(f"    Remarks : {remarks}")
            log.info(f"    Reviewer : {reviewer}")
        
        elif result == "Pass":
            # Comment message
            # Comment message
            comment_text = (
                f" |  *Date*  |  *Result*  |  *Build Number*  |\n"
                f" |  {date}  |  {result}  |  {build_number}  |"
            )
            
            # Add comment
            log.info(f"  🔄 Commenting on {ticket_id}: '{test_script_id}...")
            log.info( f"    Date : {date}" )
            log.info( f"    Result : {result}" )
            log.info( f"    Boild Number : {build_number}" )
        
        """comment_url = f"{jira_server}/rest/api/2/issue/{ticket_id}/comment"
        comment_response = requests.post(
            comment_url,
            json={"body": comment_text},
            auth=HTTPBasicAuth(username, api_token),
            headers={"Content-Type": "application/json"}
        )"""

        if comment_response.status_code in [200, 201]:
            log.info(f"  ✅ Commented on {ticket_id}: '{test_script_id}.\n")
        else:
            log.info(f"❌ Comment failed for {ticket_id}: '{test_script_id} | Status: {comment_response.status_code} | Response: {comment_response.text}.\n")

    except Exception as e:
        log.error(f"[ERROR] Could not process {ticket_id}: '{test_script_id} : {e}.\n")

log.info(f"  ✅ JIRA UPDATE COMPLETED...\n")