import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from datetime import datetime, timedelta

# === TASK CONFIGURATION ===
tasks = [
    {"url": "https://kap01.kpit.com/kap/issues/106917/time_entries/new", "hours": "4", "comment": "Automation Test Execution"},
    {"url": "https://kap01.kpit.com/kap/issues/106921/time_entries/new", "hours": "2", "comment": "Test Script Development"},
    {"url": "https://kap01.kpit.com/kap/issues/106922/time_entries/new", "hours": "2", "comment": "Test Script Development Internal Rework"},
    {"url": "https://kap01.kpit.com/kap/issues/106920/time_entries/new", "hours": "1", "comment": "Automation-Meeting/Trainings"},
]

# === DATE OPTIONS (Last 7 days) ===
dates = [(datetime.today() - timedelta(days=i)).strftime("%m/%d/%Y") for i in range(7)]

# === PATH SETUP ===
script_dir = os.path.dirname(os.path.abspath(__file__))
chromedriver_path = os.path.join(script_dir, "chromedriver.exe")

# === FUNCTION TO FILL TIME ENTRIES ===
def log_time(selected_date, driver):
    try:
        for task in tasks:
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

    username = "chandrakanb@kpit.com"
    password = "Vairagya@108"

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

        for date in selected_dates:
            log_time(date, driver)

    finally:
        driver.quit()
