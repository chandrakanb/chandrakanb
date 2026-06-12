import pandas as pd
from jira import JIRA

# Config
JIRA_URL = 'https://kpithondajapan.atlassian.net'
USERNAME = 'chandrakanb@kpit.com'
API_TOKEN = 'ATATT3xFfGF00OVb1BHWgz1klgLPAPkq8W8nFtC2j6k_ZYhvBxAWBFwEnCRk18fPD1ctA4TWpWJ79K4s4OPXJOiskD0GInEa14A0ex8puHBIQRLaGOsmNkVRY4ij9rKJr1q2JgK9D4zR-nK0rpl4QevUi5_g2PWiSM5HB7qEZAOMUaWUirk1vD4=54745E1F'
EXCEL_PATH = 'labels.xlsx'        # Path to Excel file
SHEET_NAME = 'B6_DEV_nwly_crtd'             # Update this to match your sheet

# Connect to Jira
jira = JIRA(server=JIRA_URL, basic_auth=(USERNAME, API_TOKEN))

# Load specified sheet
df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)

# Process each row
for index, row in df.iterrows():
    issue_key = row['Issue key']
    label_to_add = str(row['Labels'])

    try:
        issue = jira.issue(issue_key)
        current_labels = set(issue.fields.labels)

        if label_to_add in current_labels:
            print(f"{issue_key}: Label '{label_to_add}' already exists. Skipping.")
        else:
            updated_labels = list(current_labels | {label_to_add})
            issue.update(fields={'labels': updated_labels})
            print(f"{issue_key}: ✅ Added label '{label_to_add}'")
    except Exception as e:
        print(f"{issue_key}: ❌ Failed to update – {e}")

print("\n✔️ All done.")
