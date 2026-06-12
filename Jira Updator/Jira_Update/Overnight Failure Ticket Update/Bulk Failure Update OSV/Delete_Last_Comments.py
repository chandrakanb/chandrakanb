from jira import JIRA
import sys
import pandas as pd

# --- JIRA CREDENTIALS ---
excel_path = "OVN_Analysis.xlsx"
sheet_name = sys.argv[1]  
jira_server = "https://kpithondajapan.atlassian.net"
username = "chandrakanb@kpit.com"
with open("api_token.txt", "r") as f:
    api_token = f.read().strip()

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
        