import os
import sys
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
from jira import JIRA
import logging
from datetime import datetime

# --- CONFIG ---
excel_path = "A14_Overnight_Analysis.xlsx"
sheet_name = "Sheet1"
# sheet_name = sys.argv[1]  
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

failure_category_dict = {}
failure_category_dict = {
    "KITE issue": 16506,
    "TS issue": 16507,
    "ICB issue": 16512,
    "Test Env issue": 16508,
    "Human error": 16509
}

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

def find_failure_sub_category_id(failure_sub_category):
    failure_sub_category_dict = {}

    if failure_category == "KITE issue":
        failure_sub_category_dict = {
            "CAN signal issue": 16522,
            "Image comparison": 16523,
            "Audio comparison": 16524,
            "Video comparison": 16525,
            "Robot operation": 16526,
            "Device Mapping issue": 16527,
            "KITE Application issue": 16528
        }
    elif failure_category == "TS issue":
        failure_sub_category_dict = {
            "TS logic incorrect" : 16529
        }
    elif failure_category == "ICB issue":
        failure_sub_category_dict = {
            "Once seen" : 16518,
            "Always" : 16520,
            "GAS implementation" : 16519,
            "Timing issue" : 16521,
            "SRL/CR changes" : 16554
        }
    elif failure_category == "Test Env issue":
        failure_sub_category_dict = {
            "Automation Setup malfunction" : 16514,
            "External HW" : 16515,
            "External Application issue" : 16516,
            "Network Issue" : 16517,
            "Cascading Issue" : 19025
        }
    
    return failure_sub_category_dict

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
            """# Extract, combine, and deduplicate failure category values using sets for uniqueness. 
            combined_values = []
            current_failure_category = []
            current_failure_category.append(issue.fields.customfield_14461)
            current_failure_category_values = []
            if current_failure_category != [None]:
                current_failure_category_values.append(issue.fields.customfield_14461.value)
                for value in current_failure_category_values:                
                    normalized_values = []
                    ICB_Issue_normalization_map = {
                        "Once seen" : "ICB Issue (Once)",
                        "Always" : "ICB Issue (Always)",
                        "GAS implementation" : "ICB Issue (Once)",
                        "Timing issue" : "ICB Issue (Once)",
                        "SRL/CR changes" : "SRL/CR Changes",
                    }
                    value_lower = value.lower()
                    if value_lower == "icb issue" :
                        current_failure_sub_category = issue.fields.customfield_14461.child.value
                        normalized_value = ICB_Issue_normalization_map[current_failure_sub_category]
                        if normalized_value not in normalized_values:
                            normalized_values.append(normalized_value)                
                    else:
                        normalized_value = value
                        if normalized_value not in normalized_values:
                            normalized_values.append(normalized_value)                
                current_failure_category_values = normalized_value
                previous_failure_category_list = []
                previous_failure_category_list.append(issue.fields.customfield_12208)
                previous_failure_category_values = []
                if previous_failure_category_list != [None]:
                    previous_failure_category_values = [item.value for sublist in previous_failure_category_list for item in sublist]
                    for previous_failure_category_value in previous_failure_category_values:
                        combined_values.append(previous_failure_category_value)
                    combined_values.append(current_failure_category_values)
                else:
                    combined_values.append(current_failure_category_values)

                current_failure_category_values_list = []
                for value in combined_values:
                    normalization_map = {
                        "ts issue": "TS Issue",
                        "kite issue": "KITE Issue",
                        "icb issue (once)": "ICB Issue (Once)",
                        "icb issue (always)": "ICB Issue (Always)",
                        "human error": "Human Error",
                        "test env issue": "Test Env Issue",
                        "srl/cr changes": "SRL/CR Changes"
                    }
                    value_lower = value.lower()
                    if value_lower in normalization_map:
                        normalized_value = normalization_map[value_lower]
                        if normalized_value not in current_failure_category_values_list:
                            current_failure_category_values_list.append(normalized_value)
            else:
                previous_failure_category_list = []
                previous_failure_category_list.append(issue.fields.customfield_12208)
                previous_failure_category_values = []
                if previous_failure_category_list != [None]:
                    previous_failure_category_values = [item.value for sublist in previous_failure_category_list for item in sublist]
                    for previous_failure_category_value in previous_failure_category_values:
                        combined_values.append(previous_failure_category_value)

                current_failure_category_values_list = []
                if combined_values != [None]:
                    for value in combined_values:
                        normalization_map = {
                            "ts issue": "TS Issue",
                            "kite issue": "KITE Issue",
                            "icb issue (once)": "ICB Issue (Once)",
                            "icb issue (always)": "ICB Issue (Always)",
                            "human error": "Human Error",
                            "test env issue": "Test Env Issue",
                            "srl/cr changes": "SRL/CR Changes"
                        }
                        value_lower = value.lower()
                        if value_lower in normalization_map:
                            normalized_value = normalization_map[value_lower]
                            if normalized_value not in current_failure_category_values_list:
                                current_failure_category_values_list.append(normalized_value)
                
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
            if failure_category:
                failure_category_id = failure_category_dict[failure_category]
                if failure_category == "Human error":
                    fields["customfield_14461"] = {"id": str(failure_category_id)}
                elif failure_category == "TS issue":
                    failure_sub_category_temp = failure_sub_category
                    failure_sub_category = "TS logic incorrect"
                    failure_sub_category_dict = find_failure_sub_category_id(failure_sub_category)
                    failure_sub_category_id = failure_sub_category_dict[failure_sub_category]
                    fields["customfield_14461"] = {"id": str(failure_category_id), "child": {"id": str(failure_sub_category_id)}}
                    failure_sub_category = failure_sub_category_temp
                else :
                    failure_sub_category_dict = find_failure_sub_category_id(failure_sub_category)
                    failure_sub_category_id = failure_sub_category_dict[failure_sub_category]
                    fields["customfield_14461"] = {"id": str(failure_category_id), "child": {"id": str(failure_sub_category_id)}}   
            if root_cause_details:
                fields["customfield_10614"] = str(root_cause_details)
            if fixation_performed:
                fields["customfield_10618"] = str(fixation_performed)
            if current_failure_category_values_list:
                fields["customfield_12208"] = [{"value": str(v)} for v in current_failure_category_values_list]
            log.info(f"  🔄 Updating fields for {ticket_id}: '{test_script_id}...")
            #log.info(f"    Test Pack Name : {test_pack_name}")
            log.info(f"    Failure Category : {failure_category}")
            log.info(f"    Failure Sub Category : {failure_sub_category}")
            log.info(f"    Previous Failure Category : {current_failure_category_values_list}")
            log.info(f"    Root Cause Details : {root_cause_details}")
            log.info(f"    Fixation Performed : {fixation_performed}")
            # issue.update(fields=fields)
            log.info(f"  ✅ Updated fields for {ticket_id}: '{test_script_id}.")
            """
            # Comment message
            comment_text = (
                f" |  *Date*  |  *Result*  |  *Failure Category*  |  *Failure Sub Category*  |  *Root Cause Details*  |  *Fixation Performed*  | *Remarks* | *Reviewer* | \n"
                f" |  {date}  *[{build_number}]*  |  {result}  |  {failure_category}  |  ~{failure_sub_category}~  |  {root_cause_details}  |  {fixation_performed}  | {remarks} | {reviewer} | "
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
        
        comment_url = f"{jira_server}/rest/api/2/issue/{ticket_id}/comment"
        comment_response = requests.post(
            comment_url,
            json={"body": comment_text},
            auth=HTTPBasicAuth(username, api_token),
            headers={"Content-Type": "application/json"}
        )

        if comment_response.status_code in [200, 201]:
            log.info(f"  ✅ Commented on {ticket_id}: '{test_script_id}.\n")
        else:
            log.info(f"❌ Comment failed for {ticket_id}: '{test_script_id} | Status: {comment_response.status_code} | Response: {comment_response.text}.\n")

    except Exception as e:
        log.error(f"[ERROR] Could not process {ticket_id}: '{test_script_id} : {e}.\n")

log.info(f"  ✅ JIRA UPDATE COMPLETED...\n")