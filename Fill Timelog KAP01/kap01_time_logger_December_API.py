from redminelib import Redmine
from datetime import datetime, timedelta
import logging
import os
import sys

# --- Configure logging ---
os.makedirs("logs", exist_ok=True)
timestamp = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
log_file_path = os.path.join("logs", f"{timestamp}.log")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[
        logging.FileHandler(log_file_path, encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
log = logging.getLogger()

base_url = "https://kap01.kpit.com/kap"
api_key = "6c9f23629c0c5aa8b881b9a3066ead07f0ebfd9d"

tasks = [
    {"url": "https://kap01.kpit.com/kap/issues/486738/time_entries/new", "hours": 4, "comment": "Automation Test Execution", "work_category": "TEST AUTOMATION", "work_type": "WORK"},
    {"url": "https://kap01.kpit.com/kap/issues/486741/time_entries/new", "hours": 2, "comment": "Test Script Development", "work_category": "TEST AUTOMATION", "work_type": "WORK"},
    {"url": "https://kap01.kpit.com/kap/issues/486742/time_entries/new", "hours": 2, "comment": "Test Script Development Internal Rework", "work_category": "TEST AUTOMATION", "work_type": "INTERNAL REWORK"},
    {"url": "https://kap01.kpit.com/kap/issues/486740/time_entries/new", "hours": 1, "comment": "Automation-Meeting/Trainings", "work_category": "TEAM MEETING", "work_type": ""}
]

dates = [datetime.today() - timedelta(days=i) for i in range(10)]
date_strings = [d.strftime("%Y-%m-%d") for d in dates]

def get_issue_id_from_url(url):
    parts = url.strip("/").split("/")
    for i, p in enumerate(parts):
        if p == "issues" and i + 1 < len(parts):
            try:
                return int(parts[i + 1])
            except ValueError:
                return None
    return None

if __name__ == "__main__":
    print("Select one or more dates from the list below (comma-separated indices):")
    for i, d in enumerate(date_strings):
        print(f"{i + 1}: {d}")

    choices = input("Enter your choices (e.g., 1,3,5): ")
    selected_indices = [int(x.strip()) - 1 for x in choices.split(",") if x.strip().isdigit()]
    selected_dates = [date_strings[i] for i in selected_indices if 0 <= i < len(date_strings)]

    redmine = Redmine(base_url, key=api_key)

    for spent_on in selected_dates:
        print(f"Logging time for date: {spent_on}")
        for task in tasks:
            issue_id = get_issue_id_from_url(task["url"])
            if not issue_id:
                print(f"Skipping task, could not parse issue id from {task['url']}")
                continue

            custom_fields = [
                {"id": 46, "value": task["work_category"]},
                {"id": 48, "value": "India"},
                {"id": 49, "value": "Pune"}
            ]
            if task["work_category"] != "TEAM MEETING" and task["work_type"]:
                custom_fields.append({"id": 47, "value": task["work_type"]})

            try:
                entry = redmine.time_entry.create(
                    issue_id=issue_id,
                    spent_on=spent_on,
                    hours=task["hours"],
                    comments=task["comment"],
                    custom_fields=custom_fields
                )
                log.info(f"  OK: Issue {issue_id}, {task['hours']}h, entry id {entry.id}")
            except Exception as e:
                log.error(f"  FAIL: Issue {issue_id}, {task['hours']}h – {e}")
