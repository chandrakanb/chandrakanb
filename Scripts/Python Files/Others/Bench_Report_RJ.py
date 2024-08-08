import os
import re
import bs4
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import pandas as pd
import xml.etree.ElementTree as ET
import requests
import time
# import numpy as np


def read_html_file(execution_report_details):
    html_file = os.path.join(execution_report_details['execution_report_path'], execution_report_details['html_file'])
    df_summary = pd.read_html(html_file)
    df_summary = df_summary[0]
    df_summary = df_summary.filter(df_summary.columns[:2])
    infile = open(html_file)
    lines = infile.readlines()
    FileLines = "".join(lines)
    soup = BeautifulSoup(FileLines, 'html.parser')
    reportframe_add = soup.find('div', attrs={'id':'reportframe'})
    content_src = None
    for content in reportframe_add.contents:
        if isinstance(content, bs4.Tag):
            content_src= content['src']
            break
    content_src = content_src.replace("../","")
    content_src_paths = content_src.split("/")
    curr_path = execution_report_details['execution_report_path']
    curr_path = os.path.normpath(curr_path + os.sep + os.pardir)
    '''for i in content_src_paths:
        curr_path = os.path.join(curr_path, i)
    df_details_report = pd.read_html(curr_path)'''
    '''infile = open(curr_path)
    lines = infile.readlines()
    FileLines = "".join(lines)
    soup = BeautifulSoup(FileLines, 'html.parser')'''
    
    curr_path = os.path.join(curr_path,'XMLReport','Segments','XMLReportSegment_0.xml')
    
    tree = ET.parse(curr_path)
    root = tree.getroot()
    mainlist = list()
    for tcs_node in tree.iter("TestCases"):
        for tc_node in tcs_node.iter("TestCase"):
            tc_id = tc_node.find("testCaseID").text
            testCaseObjective = tc_node.find("testCaseObjective").text
            start_time = tc_node.find("startTime").text
            video_logger_link = tc_node.find("VideoLogLink").text
            result = tc_node.find("result").text
            total_time = tc_node.find("totalExecutionTime").text
            end_time = tc_node.find("endTime").text
            print(tc_id)
            
            maindict = {"Test ID": tc_id,"Test Objective":testCaseObjective,"Start Time":start_time,"Video Logger Link":video_logger_link,"Status":result,"Total Time":total_time,"End Time":end_time}
            mainlist.append(maindict)
            
    df_detail_report = pd.DataFrame(mainlist)

    df_detail_report['file_path'] = curr_path
    return df_summary, df_detail_report


def iterate_tuple(dir_tuple):
    for item_tulpe in dir_tuple:
        if isinstance(item_tulpe, tuple):
            iterate_tuple(item_tulpe)
        elif isinstance(item_tulpe, list):
            for html_file in item_tulpe:
                if html_file == "HTMLReport.html":
                    execution_report_path = dir_tuple[0]
                    execution_report_name = execution_report_path.replace(os.getcwd(), "")

                    sub_folder_list = execution_report_name.split('\\')
                    test_pack = Iteration_count = None
                    for each_folder in sub_folder_list:
                        if "ITERATION" in each_folder.upper():
                            Iteration_count = each_folder
                            x = re.search("Iteration_[0-9]+", each_folder)
                            Iteration_count = x.group()
                            test_pack = each_folder.replace("_{}".format(Iteration_count),"")
                    Bench_name = sub_folder_list[1]
                    execution_report_name = sub_folder_list[2]
                    return {"execution_report_name": execution_report_name,
                            "execution_report_path": execution_report_path, "Bench_name":Bench_name,
                            "Iteration_count": Iteration_count,
                            "Test_Pack": test_pack,
                            "html_file": html_file}
    return


def add_additional_columns(df):
    for col_ in ['Bench_name', 'execution_report_name', 'Test_Pack', 'Iteration_count']:
        df[col_] = execution_report_details[col_]
    return df


def generate_Detailed_Report_table(df):
    df_1 = df.filter(["Test_Pack", "Test ID"])
    df['Unique_col'] = df['Test_Pack'] + df['Test ID']
    df['Total Time'] = [datetime.strptime(i, "%H:%M:%S") for i in df['Total Time']]
    # df.to_excel("working_dump.xlsx")
    df_2 = pd.pivot_table(df, index=["Test_Pack", "Test ID"], columns='Status', aggfunc='count', values='Unique_col')
    #df_2.to_excel("working_report.xlsx")
    #df['Total Time'] = [datetime.strptime(i, "%H:%M:%S") for i in df['Total Time']]
    #df_3 = pd.pivot_table(df, index=["Test_Pack", "Test ID"], columns='Total Time', aggfunc=np.sum, values='Unique_col')
    return df_2


df_summary_main = pd.DataFrame([])
df_detail_report_main = pd.DataFrame([])

for dir_tuple in os.walk(os.getcwd()):
    if iterate_tuple(dir_tuple):
        execution_report_details = iterate_tuple(dir_tuple)
        try:
            df_summary, df_detail_report = read_html_file(execution_report_details)
            df_summary = df_summary.T
            df_summary.columns = df_summary.loc['Summary'].to_list()
            df_summary = add_additional_columns(df_summary)
            df_detail_report = add_additional_columns(df_detail_report)
            df_summary_main = pd.concat([df_summary_main,df_summary],axis=0) if df_summary_main.size > 0 else df_summary
            df_detail_report_main = pd.concat([df_detail_report_main,df_detail_report],axis=0) if df_detail_report_main.size > 0 else df_detail_report
        except ValueError:
            print(execution_report_details)

df_summary_main = df_summary_main.loc[df_summary_main['Start Time'] != 'Start Time']
df_detail_report_main['Start Time_1'] = [datetime.strptime(i, "%d %b %Y %H:%M:%S") for i in df_detail_report_main['Start Time']]
df_detail_report_main['Time_1'] = [datetime.strptime(i, "%H:%M:%S") for i in df_detail_report_main['Total Time']]
End_Time = list()
for x,y in zip(df_detail_report_main['Start Time_1'], df_detail_report_main['Time_1']):
    __day__ = 0 if y.year == 1900 else y.day
    End_Time.append(x + timedelta(days=__day__, hours=y.hour, minutes=y.minute, seconds=y.second))
df_detail_report_main['End Time'] = End_Time

"""df_detail_report_main = df_detail_report_main.filter(["Test ID", "Test Objective", "Start Time", "End Time",
                                                      "Total Time", "Status", "Bench_name", "execution_report_name",
                                                      "Test_Pack", "Iteration_count"])"""

df_detail_report_main = df_detail_report_main.filter(["Test_Pack","Test ID", "Status", "Total Time", "Bench_name",
                                                    "Iteration_count", "file_path"])

df_detail_report_main_table = df_detail_report_main.filter(["Test_Pack","Test ID", "Status", "Total Time", "Bench_name",])

# df_detail_report_main_table = generate_Detailed_Report_table(df_detail_report_main)
df_detail_report_main_table['Remark Failure Reason'] = ""

writer = pd.ExcelWriter('Execution_Report.xlsx', engine='xlsxwriter')
#df_summary_main.to_excel(writer, sheet_name='Summary', index=False)
#df_detail_report_main.to_excel(writer, sheet_name='Detailed Report dump', index = False)
df_detail_report_main_table.to_excel(writer, sheet_name='Detailed Report', index = False)
writer.close()


time.sleep(10)

# Read the Excel file
excel_file = "Execution_Report.xlsx"
df = pd.read_excel(excel_file)

# Apply filter on specific columns and insert new columns
filtered_dfs = []

filtered_df_post = df[df['Test ID'].str.contains('Post')]
filtered_df_post.insert(0, 'Script Type', 'Postcondition')
filtered_dfs.append(filtered_df_post)

# Filter and insert 'Script Type' column for 'KPIT_TC' tests
filtered_df_kpit = df[df['Test ID'].str.contains('KPIT_TC')]
filtered_df_kpit.insert(0, 'Script Type', 'Test Script')
filtered_dfs.append(filtered_df_kpit)

# Filter and insert 'Script Type' column for 'Precon' tests
filtered_df_precon = df[df['Test ID'].str.contains('Precon')]
filtered_df_precon.insert(0, 'Script Type', 'Precondition')
filtered_dfs.append(filtered_df_precon)

# Concatenate the filtered DataFrames
filtered_df = pd.concat(filtered_dfs)

# Create pivot table in pandas
pivot_table_df = pd.pivot_table(filtered_df, index='Script Type', columns='Status', aggfunc='size', fill_value=0)

# Save the pivot table to a new Excel file using ExcelWriter
pivot_table_excel_file = "pivot_table.xlsx"
with pd.ExcelWriter(pivot_table_excel_file, engine='openpyxl') as writer:
    pivot_table_df.to_excel(writer)
print("\nPivot table saved to:", pivot_table_excel_file)
print(pivot_table_df)

# Function to send message to Teams
def send_message_to_teams(teams_webhook_url, message):
    try:
        # Prepare the request headers
        headers = {'Content-Type': 'application/json'}

        # Prepare the payload with the message
        payload = {
            "text": message
        }

        # Send the request
        response = requests.post(teams_webhook_url, headers=headers, json=payload)

        # Check for success
        if response.status_code == 200:
            print("Message sent successfully to Microsoft Teams!")
        else:
            print(f"Failed to send message. Status code: {response.status_code}, Response: {response.text}")

    except Exception as e:
        print(f"An error occurred: {e}")

# Replace 'YOUR_TEAMS_WEBHOOK_URL' with your actual Teams webhook URL
teams_webhook_url = 'https://kpitc.webhook.office.com/webhookb2/9269d31d-1feb-4fdc-9969-0ab2da2adcc7@3539451e-b46e-4a26-a242-ff61502855c7/IncomingWebhook/d06a3aa25c674c12ae21c2cf796762f9/85ba79cd-deb2-408d-8c24-664ce7ccb367'

# Convert pivot table DataFrame to string
message = pivot_table_df.to_html()

print(type(message))

# Call the function to send message to Teams
#send_message_to_teams(teams_webhook_url, message)