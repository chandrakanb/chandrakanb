import pandas as pd
from jira import JIRA
from datetime import datetime

# --- CONFIG ---
excel_path = "tickets.xlsx"
jira_server = "https://kpithondajapan.atlassian.net"
username = "chandrakanb@kpit.com"
password = "ATATT3xFfGF00OVb1BHWgz1klgLPAPkq8W8nFtC2j6k_ZYhvBxAWBFwEnCRk18fPD1ctA4TWpWJ79K4s4OPXJOiskD0GInEa14A0ex8puHBIQRLaGOsmNkVRY4ij9rKJr1q2JgK9D4zR-nK0rpl4QevUi5_g2PWiSM5HB7qEZAOMUaWUirk1vD4=54745E1F"

# --- CONNECT TO JIRA ---
jira = JIRA(server=jira_server, basic_auth=(username, password))

# --- READ EXCEL FILE ---
df = pd.read_excel(excel_path)

# --- PROCESS EACH TICKET ---
for index, row in df.iterrows():
    try:
        ticket_id = row['Jira ID']
        status = row['Status']
        issue = jira.issue(ticket_id)

        # Update fields if status is Failure
        if status == "Failure":
            fields = {
                'customfield_10851': [{'value': row['customfield_10851']}],  # Failure Category
                'customfield_10617': row['customfield_10617'],              # Sub-Category
                'customfield_10614': row['customfield_10614'],              # Root Cause
                'customfield_10618': row['customfield_10618']               # Fixation
            }
            issue.update(fields=fields)

        # Add formatted comment (table)
        if status == "Failure":
            comment = (
                f"+----------------------+----------------------+\n"
                f"| Date                 | Failure Category     |\n"
                f"+----------------------+----------------------+\n"
                f"| {row['Date']}        | {row['customfield_10851']}         |\n"
                f"+----------------------+----------------------+\n"
                f"| Sub-Category         | Root Cause           |\n"
                f"+----------------------+----------------------+\n"
                f"| {row['customfield_10617']}           | {row['customfield_10614']}       |\n"
                f"+----------------------+----------------------+\n"
                f"| Fixation             |\n"
                f"+----------------------+\n"
                f"| {row['customfield_10618']}           |\n"
                f"+----------------------+\n"
            )
        else:
            comment = (
                f"+----------------------+----------------------+\n"
                f"| Date                 | Status               |\n"
                f"+----------------------+----------------------+\n"
                f"| {row['Date']}        | {status}             |\n"
                f"+----------------------+----------------------+\n"
            )

        jira.add_comment(issue, comment)
        print(f"Updated ticket: {ticket_id}")

    except Exception as e:
        print(f"Error processing {row['Jira ID']}: {e}")
