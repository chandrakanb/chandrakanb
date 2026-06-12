from redminelib import Redmine
import pandas as pd

REDMINE_URL = 'https://kap.kpit.com/redmine'
USERNAME = 'chandrakanb'
PASSWORD = 'Penguin@1008'

excel_path = "Data.xlsx"
sheet_name = "Sheet1"

# --- READ EXCEL FILE ---
df = pd.read_excel(excel_path, sheet_name=sheet_name, dtype=str, keep_default_na=False).fillna("NA")
for index, row in df.iterrows():
    try:
        ISSUE_ID = row['ISSUE ID']
        redmine = Redmine(REDMINE_URL, username=USERNAME, password=PASSWORD)

        # Update only the required custom field
        redmine.issue.update(
            ISSUE_ID,
            custom_fields=[
                {'id': 1215, 'value': 'DRT_Automation_Bench06'}
            ]
        )

        print(f"Issue #{ISSUE_ID} Bench Number updated successfully!")

    except Exception as e:
        print(f"Error: {e}")
