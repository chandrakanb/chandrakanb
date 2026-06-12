import requests
import pandas as pd

# ===== CONFIG =====
JIRA_URL = "https://kpithondajapan.atlassian.net"
EMAIL = "chandrakanb@kpit.com"
API_TOKEN = "ATATT3xFfGF0Ig79bW5jI19wSjZgmxXCqKyysJhUqiAUOrCHnYWeja9LW5xni_jNJeSKuhtsujOJ3A1a5Ogu_RCLNLxlONZrvy1p5d-A9vc_HW0DJmoIG3walwwNUAMyq9vC_71dC5UUAGg8zdIALIxeNDXCzLs-NC5beisZm-3m1CNLmohgeXc=47FD5B12"

JQL = 'issuetype = "Test Cases" AND project = OSV AND assignee = 712020:5e45192c-a1d0-4e92-b528-97cb5d418265 AND status = TS_Rework_Required_L1 ORDER BY created DESC'

KPIT_TC_ID_FIELD = "customfield_10697"
BENCH_NUMBER_FIELD = "customfield_11463"
TEST_PACK_NAME_FIELD = "customfield_11042"

auth = (EMAIL, API_TOKEN)
headers = {"Accept": "application/json"}

# ====== FETCH ISSUES ======
start_at = 0
max_results = 100
issues = []

while True:
    params = {
        "jql": JQL,
        "startAt": start_at,
        "maxResults": max_results,
        "fields": ["summary", KPIT_TC_ID_FIELD, BENCH_NUMBER_FIELD, TEST_PACK_NAME_FIELD, "status"]
    }
    response = requests.get(f"{JIRA_URL}/rest/api/2/search", headers=headers, params=params, auth=auth)
    data = response.json()

    issues.extend(data['issues'])
    if start_at + max_results >= data['total']:
        break
    start_at += max_results

# ====== PROCESS DATA ======
output = []
for issue in issues:
    key = issue["key"]
    fields = issue["fields"]
    summary = fields["summary"]
    kpit_tc_id = fields.get(KPIT_TC_ID_FIELD, "N/A")
    test_pack_name = fields.get(TEST_PACK_NAME_FIELD, "N/A")
    parent_status = fields.get("status", {}).get("name", "N/A")

    # Fetch child tickets (sub-tasks)
    issue_details = requests.get(f"{JIRA_URL}/rest/api/2/issue/{key}", headers=headers, auth=auth).json()
    subtasks = issue_details.get("fields", {}).get("subtasks", [])
    num_children = len(subtasks)

    if num_children > 0:
        for child in subtasks:
            child_key = child["key"]
            # Fetch child ticket details
            child_details = requests.get(f"{JIRA_URL}/rest/api/2/issue/{child_key}", headers=headers, auth=auth).json()
            child_fields = child_details.get("fields", {})
            child_status = child_fields.get("status", {}).get("name", "N/A")
            child_summary = child_fields.get("summary", "N/A")
            
            bench_numbers_field = fields.get(BENCH_NUMBER_FIELD)
            if bench_numbers_field and isinstance(bench_numbers_field, list):
                bench_number = ', '.join(
                    [bn.get('value', 'N/A') for bn in bench_numbers_field if bn.get('value')]
                )
            else:
                bench_number = "N/A"

            issue_type = fields.get("issuetype", {}).get("name", "N/A")
            child_issue_type = child_fields.get("issuetype", {}).get("name", "N/A")

            assignee_obj = child_fields.get("assignee")
            child_assignee = assignee_obj.get("displayName") if assignee_obj else "Unassigned"

            reporter_obj = child_fields.get("reporter")
            child_reporter = reporter_obj.get("displayName") if reporter_obj else "N/A"

            output.append({
                "Jira ID": key,
                "Summary": summary,
                "KPIT TC ID": kpit_tc_id,
                "Bench Number" : bench_number,
                "Test Pack Name": test_pack_name,
                "Parent Status": parent_status,
                "Parent Issue Type": issue_type,
                "Number of Child Tickets": num_children,
                "Child ticket": child_key,
                "Child Status": child_status,
                "Child Summary": child_summary,
                "Child Assignee": child_assignee,
                "Child Reporter": child_reporter,
                "Child Issue Type": child_issue_type
            })

            print(f"| {key} | {summary} | {kpit_tc_id} | {bench_number} | {test_pack_name} | Parent Status: {parent_status} | {issue_type} | Child: {child_key} | Child Status: {child_status} | Child Summary: {child_summary} | Child Assignee: {child_assignee} | Child Reporter: {child_reporter} | {child_issue_type} |")

    else:
        output.append({
            "Jira ID": key,
            "Summary": summary,
            "KPIT TC ID": kpit_tc_id,
            "Test Pack Name": test_pack_name,
            "Parent Status": parent_status,
            "Number of Child Tickets": num_children,
            "Child ticket": "N/A",
            "Child Status": "N/A",
            "Child Summary": "N/A",
            "Child Assignee": "N/A",
            "Child Reporter": "N/A"
        })
        print(f"{key} | {summary} | {kpit_tc_id} | {test_pack_name} | Parent Status: {parent_status} | No child tickets")

# ====== SAVE TO EXCEL ======
df = pd.DataFrame(output)
df.to_excel("jira_test_cases_with_children_and_status_TS_Rework_Required_L1.xlsx", index=False)

print("✅ Excel file 'jira_test_cases_with_children_and_status_TS_Rework_Required_L1.xlsx' created successfully.")
