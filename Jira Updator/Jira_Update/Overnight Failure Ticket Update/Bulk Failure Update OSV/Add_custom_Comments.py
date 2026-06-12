from jira import JIRA
import pandas as pd
import sys

# --- Config ---
excel_path = "failure_tickets.xlsx"
sheet_name = sys.argv[1]  # Get sheet name from command-line arg
custom_comment = "3/3"

jira_server = "https://kpithondajapan.atlassian.net"
username = "chandrakanb@kpit.com"
with open("api_token.txt", "r") as f:
    api_token = f.read().strip()

# --- Connect to JIRA ---
print("🔐 Connecting to JIRA...")
jira = JIRA(server=jira_server, basic_auth=(username, api_token))
print("✅ Connected to JIRA")

# --- Read Excel ---
print(f"📄 Reading Excel file: {excel_path}")
df = pd.read_excel(excel_path, sheet_name=sheet_name, dtype=str, keep_default_na=False).fillna("NA")
print(f"📋 Found {len(df)} Jira IDs")

# --- Post Comment ---
for index, row in df.iterrows():
    try:
        ticket_id = row['Jira ID']
        print(f"💬 Adding comment to {ticket_id}...")
        jira.add_comment(ticket_id, custom_comment)
        print(f"✅ Commented on {ticket_id}")
    except Exception as e:
        print(f"❌ Failed on {ticket_id}: {e}")

print("🏁 Done posting comments.")
