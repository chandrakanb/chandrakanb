import os
import json
import requests
import pandas as pd
from requests.auth import HTTPBasicAuth
import sys

base_url = "https://jira.honda-ivi.com/rest/atm/1.0/"
user = "kpt_jorawar_lu"
pw = "Iamlj!2024"

auth = HTTPBasicAuth(user, pw)

if len(sys.argv) != 5:
    print("Usage: python script.py <slavedata> <testRunKey> <bench_no>")
    sys.exit(1)

slavedata = sys.argv[1]
testRunKey = sys.argv[2]
bench_no = sys.argv[3]
build_para = sys.argv[4]

print("Build Parameter:", build_para)

notification_excel_file = os.path.join(slavedata, "DRT_Execution_Report.xlsx")
input_excel_file = os.path.join(slavedata, build_para + "_DRT_Execution.xlsx")
json_file_path = os.path.join(slavedata, "data.json")

# Code 1
try:
    df_notification = pd.read_excel(notification_excel_file)
    df_traceability = pd.read_excel(input_excel_file)

    if 'Status' in df_notification.columns:
        status_mapping = df_notification.set_index('Test ID')['Status'].to_dict()
        df_traceability['Status'] = df_traceability['TestCase_ID'].map(status_mapping)
        df_traceability.to_excel(input_excel_file, index=False)
        print("Mapped KITE Execution status with DRT_Execution (Traceability)")

        # Code 2
        try:
            with open(json_file_path, 'r') as json_file:
                json_data = json.load(json_file)
        except FileNotFoundError:
            json_data = []

        try:
            excel_data = pd.read_excel(input_excel_file, engine='openpyxl')

            if 'Status' in excel_data.columns:
                excel_data.rename(columns={'Status': 'Status'}, inplace=True)

            selected_bench = "Bench" + str(bench_no) 
            filtered_data = excel_data[(excel_data['Benchwise_Data'] == selected_bench) & (excel_data['Status'].isin(['Pass', 'Fail']))]

            test_cases = dict(zip(filtered_data['ZephyrKey'], filtered_data['Status']))

            json_data = []

            for test_cycle_key, status in test_cases.items():
                json_entry = {
                    'testCaseKey': test_cycle_key,
                    'status': status
                }
                json_data.append(json_entry)

            with open(json_file_path, 'w') as json_file:
                json.dump(json_data, json_file, indent=4)

            print("JSON data updated and saved to:", json_file_path)

            # Code 3
            for test_case in json_data:
                testCaseKey = test_case['testCaseKey']
                status = test_case['status']

                create_test_run_url = f"{base_url}testrun/{testRunKey}/testcase/{testCaseKey}/testresult"
                test_run_payload = {
                    "status": status
                }

                print("Updating Test Result for Test Case Key:", testCaseKey)
                print("Payload:", test_run_payload)

                create_test_run_response = requests.put(create_test_run_url, headers={"Content-Type": "application/json"}, auth=auth, json=test_run_payload)

        except Exception as e:
            print("An error occurred in Code 2:", e)

except Exception as e:
    print("An error occurred in Code 1:", e)
    
    