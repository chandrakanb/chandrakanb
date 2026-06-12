import time
import os
import getpass
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from datetime import datetime, timedelta

tasks = [
    {"url": "https://kap01.kpit.com/kap/issues/486738/time_entries/new", "hours": "4", "comment": "Automation Test Execution", "work_category": "TEST AUTOMATION", "work_type": "WORK"},
    {"url": "https://kap01.kpit.com/kap/issues/486741/time_entries/new", "hours": "2", "comment": "Test Script Development", "work_category": "TEST AUTOMATION", "work_type": "WORK"},
    {"url": "https://kap01.kpit.com/kap/issues/486742/time_entries/new", "hours": "2", "comment": "Test Script Development Internal Rework", "work_category": "TEST AUTOMATION", "work_type": "INTERNAL REWORK"},
    {"url": "https://kap01.kpit.com/kap/issues/486740/time_entries/new", "hours": "1", "comment": "Automation-Meeting/Trainings", "work_category": "TEAM MEETING", "work_type": ""}
]

dates = [(datetime.today() - timedelta(days=i)).strftime("%m/%d/%Y") for i in range(10)]

script_dir = os.path.dirname(os.path.abspath(__file__))
chromedriver_path = os.path.join(script_dir, "chromedriver.exe")
api_log_file = os.path.join(script_dir, "api_logs.txt")

def write_api_log(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(api_log_file, "a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] {message}\n")

def capture_network_logs(driver):
    try:
        logs = driver.get_log("performance")
    except Exception:
        return

    for entry in logs:
        try:
            log = json.loads(entry["message"])["message"]
        except Exception:
            continue

        method = log.get("method")
        params = log.get("params", {})

        if method == "Network.requestWillBeSent":
            request = params.get("request", {})
            url = request.get("url", "")
            http_method = request.get("method", "")
            if "timelog_sync" in url or "time_entries" in url or "kap/timelog" in url:
                write_api_log(f"REQUEST {http_method} {url}")

        if method == "Network.responseReceived":
            response = params.get("response", {})
            url = response.get("url", "")
            status = response.get("status", "")
            mime = response.get("mimeType", "")
            if "timelog_sync" in url or "time_entries" in url or "kap/timelog" in url or "json" in str(mime).lower():
                write_api_log(f"RESPONSE {status} {url} ({mime})")

def log_time(selected_date, driver, is_first_date):
    try:
        for task_index, task in enumerate(tasks):
            if is_first_date and task_index == 0:
                time.sleep(10)

            time.sleep(1)
            driver.get(task["url"])
            time.sleep(1)
            driver.find_element(By.ID, "time_entry_spent_on").clear()
            time.sleep(1)
            driver.find_element(By.ID, "time_entry_spent_on").send_keys(selected_date)
            time.sleep(1)
            driver.find_element(By.ID, "time_entry_hours").send_keys(task["hours"])
            time.sleep(1)
            driver.find_element(By.ID, "time_entry_comments").send_keys(task["comment"])
            time.sleep(1)
            driver.find_element(By.ID, "time_entry_custom_field_values_46").send_keys(task["work_category"])
            if task["work_category"] != "TEAM MEETING":
                time.sleep(2)
                driver.find_element(By.ID, "time_entry_custom_field_values_47").send_keys(task["work_type"])
            time.sleep(4)
            driver.find_element(By.XPATH, '//*[@id="new_time_entry"]/input[3]').click()
            time.sleep(8)
            capture_network_logs(driver)

        print(f"✅ Time log submitted for {selected_date}")
    except Exception as e:
        print(f"❌ Error on {selected_date}: {e}")

if __name__ == "__main__":
    print("Select one or more dates from the list below (comma-separated indices):")
    for i, date in enumerate(dates):
        print(f"{i + 1}: {date}")

    choices = input("Enter your choices (e.g., 1,3,5): ")
    selected_indices = [int(x.strip()) - 1 for x in choices.split(",") if x.strip().isdigit()]
    selected_dates = [dates[i] for i in selected_indices if 0 <= i < len(dates)]

    username = "chandrakanb@kpit.com"
    password = "Beastmonk@1008"

    service = Service(chromedriver_path)

    options = Options()
    options.set_capability("goog:loggingPrefs", {"performance": "ALL"})

    driver = webdriver.Chrome(service=service, options=options)

    try:
        driver.get(tasks[0]['url'])
        time.sleep(5)
        driver.find_element(By.XPATH, '//*[@id="content"]/div[1]/a').click()
        time.sleep(5)
        driver.find_element(By.XPATH, '//*[@id="userNameInput"]').send_keys(username)
        driver.find_element(By.XPATH, '//*[@id="passwordInput"]').send_keys(password)
        driver.find_element(By.XPATH, '//*[@id="submitButton"]').click()
        time.sleep(10)
        capture_network_logs(driver)

        for idx, date in enumerate(selected_dates):
            is_first = (idx == 0)
            log_time(date, driver, is_first)

    finally:
        capture_network_logs(driver)
        driver.quit()
