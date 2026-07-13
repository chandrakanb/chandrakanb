from redminelib import Redmine
import pandas as pd

REDMINE_URL = 'https://kap.kpit.com/redmine'
USERNAME = 'chandrakanb'
PASSWORD = 'Babbaa@240804'

excel_path = "Data.xlsx"
sheet_name = "Sheet1"

# The ID for Kunal Das from your HTML <option value="14563">
ASSIGNED_TO_ID = 14563 

# --- READ EXCEL FILE ---
df = pd.read_excel(excel_path, sheet_name=sheet_name, dtype=str, keep_default_na=False).fillna("NA")

# Initialize Redmine client once
redmine = Redmine(REDMINE_URL, username=USERNAME, password=PASSWORD)

for index, row in df.iterrows():
    try:
        ISSUE_ID = row['ISSUE ID']

        # Update ONLY the Assigned To field
        redmine.issue.update(
            ISSUE_ID,
            assigned_to_id=ASSIGNED_TO_ID
        )

        print(f"Issue #{ISSUE_ID} Assigned To updated successfully!")

    except Exception as e:
        print(f"Error updating Issue #{row.get('ISSUE ID', 'Unknown')}: {e}")

# This is AI generated code, please refer KPIT AI Policy before using this in your projects