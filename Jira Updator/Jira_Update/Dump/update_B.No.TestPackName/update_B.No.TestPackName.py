
import pandas as pd
from jira import JIRA

# --- CONFIG ---
excel_path = "tickets.xlsx"
jira_server = "https://kpithondajapan.atlassian.net"
username = "chandrakanb@kpit.com"
password = "ATATT3xFfGF0Iu6kCatn1bDOePjQr2PkgumxzQc_qtUxj54ea9lhQc0gzlXMz4ZYdY08r8aOTFVFW1mdRESB0dSQwBm4xsSlUPIbNUEI7TdsyLw_SV-AJf0cyN79JpfO-Y4lvyXhWm8cXFEGlEKDWsXgxyF_DGkgE7XcEcA-XHpgG4g06doAFGM=6038CD2D"

# --- CONNECT TO JIRA ---
jira = JIRA(server=jira_server, basic_auth=(username, password))

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
            fields['customfield_10502'] = test_pack_name

        issue.update(fields=fields)

        print(f"[SUCCESS] Updated: {ticket_id}")

    except Exception as e:
        print(f"[ERROR] Could not process {row['Jira ID']}: {e}")
