from jira import JIRA
import sys
import pandas as pd
import os

# --- JIRA CREDENTIALS ---
excel_path = "delete_last_comment.xlsx"
sheet_name = "Sheet1" 
jira_server = "https://kpithondajapan.atlassian.net"
username = "chandrakanb@kpit.com"
api_token = os.getenv("JIRA_API_TOKEN")
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

# --- READ EXCEL FILE ---
df = pd.read_excel(excel_path, sheet_name=sheet_name, dtype=str, keep_default_na=False).fillna("NA")

# --- PROCESS EACH TICKET ---
for index, row in df.iterrows():
    try:
        ticket_id = row['Jira ID']
        issue = jira.issue(ticket_id)
        comments = jira.comments(issue)

        if comments:
            last_comment = comments[-1]
            comment_id = last_comment.id
            delete_url = f"{jira_server}/rest/api/2/issue/{ticket_id}/comment/{comment_id}"

            response = jira._session.delete(delete_url)

            if response.status_code == 204:
                print(f"✅ Deleted last comment (ID: {comment_id}) on {ticket_id}")
            else:
                print(f"❌ Failed to delete comment on {ticket_id}: {response.status_code} - {response.text}")
        else:
            print(f"ℹ️ No comments found on {ticket_id}")
    except Exception as e:
        print(f"❌ Failed on {ticket_id}: {e}")
        