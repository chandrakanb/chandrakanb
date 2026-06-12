
import pandas as pd
from jira import JIRA

# --- CONFIG ---
excel_path = "tickets.xlsx"
jira_server = "https://kpithondajapan.atlassian.net"
username = "chandrakanb@kpit.com"
JIRA_API_TOKEN = os.getenv('JIRA_API_TOKEN')

# --- CONNECT TO JIRA ---
jira = JIRA(server=jira_server, basic_auth=(username, JIRA_API_TOKEN))

# --- READ EXCEL FILE ---
df = pd.read_excel(excel_path)

# --- PROCESS EACH TICKET ---
for index, row in df.iterrows():
    try:
        ticket_id = row['Jira ID']
        bench_number = row['Bench Number']
        test_pack_name = row['Test Pack Name']

        issue = jira.issue(ticket_id)

        fields = {}

        # Update multi-checkbox: Bench Number
        #if not pd.isna(bench_number):
            #fields['customfield_10505'] = [{'value': bench_number}]

        # Update text field: Test Pack Name
        if not pd.isna(test_pack_name):
            fields['customfield_11042'] = test_pack_name

        issue.update(fields=fields)

        print(f"[SUCCESS] Updated: {ticket_id}")

    except Exception as e:
        print(f"[ERROR] Could not process {row['Jira ID']}: {e}")
