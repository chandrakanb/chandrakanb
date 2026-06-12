from jira import JIRA
import sys
import pandas as pd
import re

excel_path = "failure_tickets.xlsx"
sheet_name = sys.argv[1]

jira_server = "https://kpithondajapan.atlassian.net"
username = "chandrakanb@kpit.com"
with open("api_token.txt", "r") as f:
    api_token = f.read().strip()

print("🔐 Connecting to JIRA...")
jira = JIRA(server=jira_server, basic_auth=(username, api_token))
print("✅ Connected to JIRA")

print(f"📄 Reading Excel: {excel_path} (sheet: {sheet_name})...")
df = pd.read_excel(excel_path, sheet_name=sheet_name, dtype=str, keep_default_na=False).fillna("NA")
print(f"✅ Loaded {len(df)} tickets")

output_data = []

# Match "10th June 2025" format
date_pattern = re.compile(r'\d{1,2}(st|nd|rd|th)?\s+(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4}', re.IGNORECASE)

def extract_date_result(text):
    date_match = date_pattern.search(text)
    date_str = date_match.group() if date_match else "NA"
    result = "Pass" if "pass" in text.lower() else "Fail" if "fail" in text.lower() else "NA"
    return date_str, result

print("🔄 Processing each ticket...")

for index, row in df.iterrows():
    try:
        ticket_id = row['Jira ID']
        test_id = row.get('Test Script ID', 'NA')
        print(f"🔍 Fetching comments for {ticket_id}...")
        issue = jira.issue(ticket_id)
        comments = jira.comments(issue)[-3:]  # Last 3 comments

        record = {"Jira ID": ticket_id, "Test Script ID": test_id}
        pass_count = 0
        total_count = len(comments)

        for i, comment in enumerate(reversed(comments), 1):  # oldest first
            date, result = extract_date_result(comment.body)
            print(f"   ➤ Comment {i}: Date='{date}' | Result='{result}'")
            record[f"Date{i}"] = date
            record[f"Result{i}"] = result
            if result == "Pass":
                pass_count += 1

        # Fill missing comments with NA
        for j in range(len(comments)+1, 4):
            record[f"Date{j}"] = "NA"
            record[f"Result{j}"] = "NA"

        record["Pass Count"] = f"{pass_count}/{total_count}"
        print(f"✅ Processed {ticket_id}: Pass Count = {record['Pass Count']}")
        output_data.append(record)

    except Exception as e:
        print(f"❌ Failed on {row['Jira ID']}: {e}")

# Write to Excel
print("💾 Writing results to Excel...")
out_df = pd.DataFrame(output_data)
out_df.to_excel("last_3_comments_summary.xlsx", index=False)
print("✅ Excel file 'last_3_comments_summary.xlsx' generated successfully.")
