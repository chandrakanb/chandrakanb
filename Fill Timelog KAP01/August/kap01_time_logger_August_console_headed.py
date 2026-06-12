import time
import os
import getpass
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from datetime import datetime, timedelta

# === TASK CONFIGURATION ===
tasks = [
    {"url": "https://kap01.kpit.com/kap/issues/268885/time_entries/new", "hours": "4", "comment": "Automation Test Execution"},
    {"url": "https://kap01.kpit.com/kap/issues/268887/time_entries/new", "hours": "2", "comment": "Test Script Development"},
    {"url": "https://kap01.kpit.com/kap/issues/268888/time_entries/new", "hours": "2", "comment": "Test Script Development Internal Rework"},
    {"url": "https://kap01.kpit.com/kap/issues/268886/time_entries/new", "hours": "1", "comment": "Automation-Meeting/Trainings"},
]

# === DATE OPTIONS (Last 7 days) ===
dates = [(datetime.today() - timedelta(days=i)).strftime("%m/%d/%Y") for i in range(10)]

# === PATH SETUP ===
script_dir = os.path.dirname(os.path.abspath(__file__))
chromedriver_path = os.path.join(script_dir, "chromedriver.exe")

# === FUNCTION TO FILL TIME ENTRIES ===
def log_time(selected_date, driver, is_first_date):
    try:
        for task_index, task in enumerate(tasks):
            if is_first_date and task_index == 0:
                time.sleep(10)  # Delay only for the first task of the first date

            driver.get(task["url"])
            driver.find_element(By.ID, "time_entry_spent_on").clear()
            driver.find_element(By.ID, "time_entry_spent_on").send_keys(selected_date)
            driver.find_element(By.ID, "time_entry_hours").send_keys(task["hours"])
            driver.find_element(By.ID, "time_entry_comments").send_keys(task["comment"])
            time.sleep(2)
            driver.find_element(By.XPATH, '//*[@id="new_time_entry"]/input[3]').click()
            time.sleep(5)

        print(f"✅ Time log submitted for {selected_date}")
    except Exception as e:
        print(f"❌ Error on {selected_date}: {e}")

# === MAIN EXECUTION ===
if __name__ == "__main__":
    print("Select one or more dates from the list below (comma-separated indices):")
    for i, date in enumerate(dates):
        print(f"{i + 1}: {date}")

    choices = input("Enter your choices (e.g., 1,3,5): ")
    selected_indices = [int(x.strip()) - 1 for x in choices.split(",") if x.strip().isdigit()]
    selected_dates = [dates[i] for i in selected_indices if 0 <= i < len(dates)]

	# username & password entry on console 
    #username = input("Enter your KPIT email: ")
    #password = getpass.getpass("Enter your KPIT password: ")
	
	# hardcoded credentials
    username = "chandrakanb@kpit.com"
    password = "Sanyasi@1008"

    service = Service(chromedriver_path)
    driver = webdriver.Chrome(service=service)

    try:
        driver.get(tasks[0]['url'])
        time.sleep(5)
        driver.find_element(By.XPATH, '//*[@id="content"]/div[1]/a').click()
        time.sleep(5)
        driver.find_element(By.XPATH, '//*[@id="userNameInput"]').send_keys(username)
        driver.find_element(By.XPATH, '//*[@id="passwordInput"]').send_keys(password)
        driver.find_element(By.XPATH, '//*[@id="submitButton"]').click()
        time.sleep(10)

        for idx, date in enumerate(selected_dates):
            is_first = (idx == 0)
            log_time(date, driver, is_first)

    finally:
        driver.quit()
