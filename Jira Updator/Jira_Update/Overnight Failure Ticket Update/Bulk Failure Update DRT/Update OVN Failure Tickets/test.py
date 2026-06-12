import os
from jira import JIRA
import logging

logging.basicConfig(level=logging.INFO)

jira_server = "https://kpithondajapan.atlassian.net"
username = "chandrakanb@kpit.com"
api_token = os.getenv("JIRA_API_TOKEN")

issue_key = "DRT-7527"

failure_category_dict = {}

failure_category_dict = {
    "KITE issue": 16506,
    "TS issue": 16507,
    "ICB issue": 16512,
    "Test Env issue": 16508,
    "Human error": 16509
}

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

def update_current_failure_category_sub_category(failure_category_id,failure_sub_category_id):
    try:

        jira = JIRA(
            server=jira_server,
            basic_auth=(username, api_token)
        )

        issue = jira.issue(issue_key)

        fields = {}
        # fields["customfield_14461"] = {"id": str(failure_category_id), "child": {"id": str(failure_sub_category_id)}}
        fields["customfield_14461"] = {"id": str(failure_category_id)}
        print(fields)
        issue.update(fields=fields)

        logging.info(f"{issue_key} updated successfully")

    except Exception as e:
        logging.error(f"Error updating issue: {e}")

failure_category = "KITE issue"
failure_sub_category = "CAN signal issue"

failure_category_id = failure_category_dict[failure_category]

print(find_failure_sub_category_id(failure_sub_category))
failure_sub_category_dict = find_failure_sub_category_id(failure_sub_category)


failure_sub_category_id = failure_sub_category_dict[failure_sub_category]

update_current_failure_category_sub_category(failure_category_id,failure_sub_category_id)        