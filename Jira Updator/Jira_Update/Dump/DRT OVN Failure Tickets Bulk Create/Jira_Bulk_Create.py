import pandas as pd
from jira import JIRA

# Configuration
JIRA_URL = 'https://kpithondajapan.atlassian.net'
USERNAME = 'chandrakanb@kpit.com'
API_TOKEN = 'ATATT3xFfGF00OVb1BHWgz1klgLPAPkq8W8nFtC2j6k_ZYhvBxAWBFwEnCRk18fPD1ctA4TWpWJ79K4s4OPXJOiskD0GInEa14A0ex8puHBIQRLaGOsmNkVRY4ij9rKJr1q2JgK9D4zR-nK0rpl4QevUi5_g2PWiSM5HB7qEZAOMUaWUirk1vD4=54745E1F'
EXCEL_PATH = 'create_issues.xlsx'
SHEET_NAME = 'Sheet1'

# Connect to Jira
jira = JIRA(server=JIRA_URL, basic_auth=(USERNAME, API_TOKEN))

# Read Excel
df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)

# Create issues
for index, row in df.iterrows():
    try:
        issue_dict = {
            'project': {'key': 'DRT'},
            'summary': row.get('Summary') or 'No summary provided',
            'issuetype': {'name': row['Work type'].strip()},
            'labels': str(row.get('Labels')).split(",") if pd.notna(row.get('Labels')) else [],
            'assignee': {'name': row['Assignee']} if pd.notna(row.get('Assignee')) else None,
            
            'customfield_10505': [{'value': row.get('Bench Number')}] if pd.notna(row.get('Bench Number')) else None,
            'customfield_10502': row.get('Test Pack Name'),
            'customfield_10612': row.get('HeatCount'),
            'customfield_10616': [{'value': row.get('TE Name')}] if pd.notna(row.get('TE Name')) else None,
            'customfield_10744': [{'value': row.get('TL Name')}] if pd.notna(row.get('TL Name')) else None,
        }

        # Optional parent for sub-task
        if pd.notna(row.get('Parent')):
            issue_dict['parent'] = {'key': row['Parent']}

        new_issue = jira.create_issue(fields=issue_dict)
        print(f"✅ Created: {new_issue.key}")

        # Transition to In_Progress if specified
        if row.get('Status') == 'In_Progress':
            transitions = jira.transitions(new_issue)
            progress_id = next((t['id'] for t in transitions if t['name'] == 'In Progress'), None)
            if progress_id:
                jira.transition_issue(new_issue, progress_id)
                print(f"🔄 Transitioned {new_issue.key} to 'In Progress'")

    except Exception as e:
        print(f"❌ Failed at row {index + 2}: {e}")

print("\n✔️ All done.")
