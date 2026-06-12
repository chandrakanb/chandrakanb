import os
import sys
import pandas as pd
from jira import JIRA
import logging
from datetime import datetime

# --- CONFIG ---
excel_path = "Test_Pack_Update_DRT.xlsx"
new_excel_path = "Test_Pack_Update_DRT_Updated.xlsx"
sheet_name = "Sheet1"
jira_server = "https://kpithondajapan.atlassian.net"
username = "chandrakanb@kpit.com"

# Read API token
JIRA_API_TOKEN = os.getenv("JIRA_API_TOKEN")
if not JIRA_API_TOKEN and os.path.isfile("JIRA_API_TOKEN.txt"):
    with open("JIRA_API_TOKEN.txt", "r") as f:
        JIRA_API_TOKEN = f.read().strip()

if not JIRA_API_TOKEN:
    raise SystemExit("❌ API token not found")

# --- CONNECT TO JIRA ---
jira = JIRA(server=jira_server, basic_auth=(username, JIRA_API_TOKEN))

# --- LOGGING ---
os.makedirs("logs", exist_ok=True)
timestamp = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
log_file_path = f"logs/{timestamp}.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[
        logging.FileHandler(log_file_path, encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
log = logging.getLogger()

# --- READ EXCEL ---
df = pd.read_excel(excel_path, sheet_name=sheet_name, dtype=str).fillna("")

# Create Bench Number column if not available
if "Bench Number" not in df.columns:
    df["Bench Number"] = ""

# --- PROCESS EACH TICKET ---
for index, row in df.iterrows():
    try:
        ticket_id = row["Jira ID"]
        test_script_id = row["Test Script ID"]

        issue = jira.issue(ticket_id)
        log.info(f"⏳ Processing '{ticket_id}: {test_script_id}'...")

        # Extract only value from customfield_10505
        raw = getattr(issue.fields, "customfield_10505", "")
        bench_no = ""

        if isinstance(raw, list) and raw:  # multi-select
            bench_no = ", ".join([item.value for item in raw if hasattr(item, "value")])
        elif hasattr(raw, "value"):        # single select
            bench_no = raw.value
        elif isinstance(raw, dict) and "value" in raw:
            bench_no = raw["value"]
        else:
            bench_no = str(raw)

        # Save to DataFrame
        df.at[index, "Bench Number"] = bench_no
        print(bench_no)
        log.info(f"  📌 Bench Number = {bench_no}")

    except Exception as e:
        log.error(f"[ERROR] Could not process {ticket_id}: '{test_script_id}' → {e}")

# --- SAVE NEW EXCEL ---
df.to_excel(new_excel_path, sheet_name=sheet_name, index=False)
log.info(f"📁 New Excel generated successfully → {new_excel_path}")
print(f"\n✔ Output Excel Generated: {new_excel_path}")
