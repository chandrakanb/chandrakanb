import tkinter as tk
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from datetime import datetime, timedelta
import time
import os
import getpass

# === TASK CONFIGURATION ===
tasks = [
    {"url": "https://kap01.kpit.com/kap/issues/106917/time_entries/new", "hours": "4", "comment": "Automation Test Execution"},
    {"url": "https://kap01.kpit.com/kap/issues/106921/time_entries/new", "hours": "2", "comment": "Test Script Development"},
    {"url": "https://kap01.kpit.com/kap/issues/106922/time_entries/new", "hours": "2", "comment": "Test Script Development Internal Rework"},
    {"url": "https://kap01.kpit.com/kap/issues/106920/time_entries/new", "hours": "1", "comment": "Automation-Meeting/Trainings"},
]

# === PATH SETUP ===
script_dir = os.path.dirname(os.path.abspath(__file__))
chromedriver_path = os.path.join(script_dir, "chromedriver.exe")

# === FUNCTION TO FILL TIME ENTRIES ===
def log_time(selected_date, username, password):
    service = Service(chromedriver_path)
    driver = webdriver.Chrome(service=service)

    try:
        # Step 1: Open login page
        driver.get(tasks[0]['url'])
        time.sleep(5)

        # Step 2: Click "Login with SSO"
        driver.find_element(By.XPATH, '//*[@id="content"]/div[1]/a').click()
        time.sleep(5)

        # Step 3: Enter credentials
        driver.find_element(By.XPATH, '//*[@id="userNameInput"]').send_keys(username)
        driver.find_element(By.XPATH, '//*[@id="passwordInput"]').send_keys(password)
        driver.find_element(By.XPATH, '//*[@id="submitButton"]').click()
        time.sleep(10)

        # Step 4: Fill all time logs
        for task in tasks:
            driver.get(task["url"])
            driver.find_element(By.ID, "time_entry_spent_on").clear()
            driver.find_element(By.ID, "time_entry_spent_on").send_keys(selected_date)
            driver.find_element(By.ID, "time_entry_hours").send_keys(task["hours"])
            driver.find_element(By.ID, "time_entry_comments").send_keys(task["comment"])
            time.sleep(2)
            driver.find_element(By.XPATH, '//*[@id="new_time_entry"]/input[3]').click()
            time.sleep(5)

        messagebox.showinfo("Done", "✅ All time entries submitted.")
    except Exception as e:
        messagebox.showerror("Error", f"Something went wrong:\n{e}")
    finally:
        driver.quit()

# === GUI SETUP ===
def start_gui():
    root = tk.Tk()
    root.title("KAP01 Time Logger")
    root.geometry("300x180")

    tk.Label(root, text="Select Date (MM/DD/YYYY):").pack(pady=10)

    # Last 7 days in MM/DD/YYYY format
    dates = [(datetime.today() - timedelta(days=i)).strftime("%m/%d/%Y") for i in range(7)]
    selected_date = tk.StringVar(value=dates[0])

    tk.OptionMenu(root, selected_date, *dates).pack()

    def on_submit():
        root.withdraw()
        
        # static entry user id & password
        username = "chandrakanb@kpit.com"
        password = "Vairagya@108"

        # dynamic entry user id & password
        # username = input("Enter your KPIT email ID: ")
        # password = getpass.getpass("Enter your KPIT password: ")

        log_time(selected_date.get(), username, password)
        
        root.quit()
        
    tk.Button(root, text="Submit Time Log", command=on_submit).pack(pady=20)
    root.mainloop()

start_gui()
