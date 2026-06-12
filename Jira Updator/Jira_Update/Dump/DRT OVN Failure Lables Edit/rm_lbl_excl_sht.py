import pandas as pd
from jira import JIRA

# Config
JIRA_URL = 'https://kpithondajapan.atlassian.net'
USERNAME = 'chandrakanb@kpit.com'
API_TOKEN = 'ATATT3xFfGF0vyxGCLQlxTSmVDAT1x_thX3ZupBKc49TOecSEClW46sYV6DtFvPXD45KUVPKM0yMkMfHOh05pBkRKtxkP68bSlCnXbWuMaGByr3hMDvivDw8cIhn91vC5osOSxwOckMoU-_csav18Z5M8tKinqJy-IfNArVvsBx4OJPgdIDAN-I=B748E4FD'
EXCEL_PATH = 'labels.xlsx'  # Replace with your actual path
SHEET_NAME = 'b7_dev'       # Replace with the actual sheet name

# Connect to Jira
jira = JIRA(server=JIRA_URL, basic_auth=(USERNAME, API_TOKEN))

# Load the Excel sheet
df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)

# Process each row to remove labels
for index, row in df.iterrows():
    issue_key = row['Issue key']
    label_to_remove = "Rel_2025-12-24"

    try:
        issue = jira.issue(issue_key)
        current_labels = set(issue.fields.labels)

        if label_to_remove not in current_labels:
            print(f"{issue_key}: Label '{label_to_remove}' not present. Skipping.")
        else:
            updated_labels = list(current_labels - {label_to_remove})
            issue.update(fields={'labels': updated_labels})
            print(f"{issue_key}: ✅ Removed label '{label_to_remove}'")
    except Exception as e:
        print(f"{issue_key}: ❌ Error removing label – {e}")

print("\n✔️ Label removal completed.")
