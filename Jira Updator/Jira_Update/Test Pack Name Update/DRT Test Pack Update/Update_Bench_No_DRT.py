import os
import sys
import pandas as pd
from jira import JIRA
import logging
from datetime import datetime

# --- CONFIG ---
excel_path = "Test_Pack_Update_DRT.xlsx"
sheet_name = "Sheet1"
jira_server = "https://kpithondajapan.atlassian.net"
username = "chandrakanb@kpit.com"

# Load API token
JIRA_API_TOKEN = os.getenv("JIRA_API_TOKEN")
if not JIRA_API_TOKEN:
    if os.path.isfile("JIRA_API_TOKEN.txt"):
        with open("JIRA_API_TOKEN.txt", "r") as f:
            JIRA_API_TOKEN = f.read().strip()
    else:
        raise SystemExit("❌ API token not found!")

# Connect to Jira
jira = JIRA(server=jira_server, basic_auth=(username, JIRA_API_TOKEN))

# Logging
os.makedirs("logs", exist_ok=True)
timestamp = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
log_file_path = os.path.join("logs", f"{timestamp}.log")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[logging.FileHandler(log_file_path, encoding='utf-8'),
              logging.StreamHandler(sys.stdout)]
)
log = logging.getLogger()

# Read Excel
df = pd.read_excel(excel_path, sheet_name=sheet_name, dtype=str).fillna("")

# --- PROCESS EACH TICKET ---
for index, row in df.iterrows():
    try:
        ticket = row["Jira ID"]
        new_values = row["Script Group"].replace(" ", "")  # remove extra spaces

        # Convert comma separated → list
        new_values_list = [{"value": v} for v in new_values.split(",") if v]

        issue = jira.issue(ticket)
        log.info(f"⏳ Updating: {ticket}")

        # 1️⃣ Remove previous values
        issue.update(fields={"customfield_10505": []})

        # 2️⃣ Add new values
        issue.update(fields={"customfield_10505": new_values_list})

        log.info(f"✔ Updated → {ticket}: {new_values}\n")

    except Exception as e:
        log.error(f"❌ Failed {ticket} → {e}\n")

log.info("🎉 ALL TICKETS UPDATED SUCCESSFULLY")
