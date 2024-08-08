import smtplib
import os
import sys
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime, timedelta
 
build_type=sys.argv[2]
excel_file_path = f"{sys.argv[1]}"+f"{build_type}/"       #path after clone on master
pass_word=sys.argv[3]
Build_number=sys.argv[4]
gitlabpath=sys.argv[5]
 
 
today = datetime.now()
yesterday = today - timedelta(days=1)
current_date = today.strftime("%Y-%m-%d")
 
excelPath = gitlabpath + build_type +"_DRT_Execution.xlsx"
 
total_ts_df=pd.read_excel(excelPath)
results_t_t={}
final_res=""
 
def get_bench_total_TS(df, bench_name):
    bench_data = df[df['Benchwise_Data'] == bench_name]
    total_TS=len(bench_data)
    results_t_t[bench_name] = {total_TS}
    return results_t_t
 
unique_bench_name_t=total_ts_df['Benchwise_Data'].unique()
results_t_t={bench: get_bench_total_TS(total_ts_df, bench) for bench in unique_bench_name_t}
 
bench_owner_data = [["Bench1", 'Sahil Bhalekar'], ["Bench2", "Vitthal Panhalkar"], ["Bench3", "Rohan Jagtap"],
        ["Bench4", "Rajashri Bhamare"], ["Bench5", "Srushti Potadar"], ["Bench6", 'Yukta Rane'],
        ["Bench7","Srushti Potadar"],["Bench8","Rohan Jagtap"],["Bench9","Chandrakant Bhalekar"],
        ["Bench10","Vaishnavi Kore"]]
 
def create_combine_excel(excel_file_path,date):
    try:
        list_dirs = ['Bench_Results', 'Auto_Flashing_Results']
        for p in list_dirs:
            combined_df = pd.concat(
                [pd.read_excel(os.path.join(excel_file_path + p + f"/{date}/", f)).assign(**{'Bench Name': os.path.splitext(f)[0]})
                 for f in os.listdir(excel_file_path + p + f"/{date}/")
                 if f.endswith(".xlsx") and f.startswith("Bench")], ignore_index=True)
            combined_df.to_excel(f'{excel_file_path + p + f"/{date}/"}combined_excel.xlsx', index=False)
        print(f"Combined Excel Created... for {date}")
    except FileNotFoundError as e:
        print(f"File not found: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")
 
create_combine_excel(excel_file_path,current_date)
folder_path=f"{excel_file_path}Bench_Results/"
entries=os.listdir(folder_path)
d=[(entry, os.path.getctime(os.path.join(folder_path, entry))) for entry in entries if os.path.isdir(os.path.join(folder_path, entry))]
sorted_directories = sorted(d, key=lambda x: x[1])
previous_folder_name = sorted_directories[-2][0]
create_combine_excel(excel_file_path,previous_folder_name)
 
excel_file = f"{excel_file_path}Bench_Results/{current_date}/combined_excel.xlsx"
excel1 = f"{excel_file_path}Auto_Flashing_Results/{current_date}/combined_excel.xlsx"
excel_file_yes_b= f"{excel_file_path}Bench_Results/{previous_folder_name}/combined_excel.xlsx"
excel_file_yes_a= f"{excel_file_path}Auto_Flashing_Results/{previous_folder_name}/combined_excel.xlsx"
 
try:
    df = pd.read_excel(excel_file)
    df1 = pd.read_excel(excel1)
    df_y_b = pd.read_excel(excel_file_yes_b)
    df_y_a = pd.read_excel(excel_file_yes_a)
except FileNotFoundError as e:
    print(f"File not found: {e}")
    sys.exit(1)
except Exception as e:
    print(f"An error occurred while reading the excel file: {e}")
    sys.exit(1)
 
unique_bench_name = df['Bench Name'].unique()
results_t_b = {}
unique_bench_name_y=df_y_b['Bench Name'].unique()
results_y_b={}
 
for bench in unique_bench_name:
    kpittc_df = df[(df['Test ID'].str.startswith('KPIT_TC')) & (df['Bench Name'] == bench)]
    kpittc_count = len(kpittc_df)
    pass_count = len(kpittc_df[kpittc_df['Status'] == 'Pass'])
    fail_count = len(kpittc_df[kpittc_df['Status'] == 'Fail'])
    verify_manully = len(kpittc_df[kpittc_df['Status'] == 'VERIFYMANUALLY'])
    percentage = pass_count / kpittc_count * 100 if kpittc_count != 0 else 0
    for item in bench_owner_data:
        if item[0]==bench:
            results_t_b[item[0],item[1]] = {
        'kpittc_count': kpittc_count,
        'pass_count': pass_count,
        'fail_count': fail_count,
        'verify_manully': verify_manully,
        'percentage': percentage
    }
 
for bench in unique_bench_name_y:
    kpittc_df_y_b = df_y_b[(df_y_b['Test ID'].str.startswith('KPIT_TC')) & (df_y_b['Bench Name'] == bench)]
    kpittc_count_y_b = len(kpittc_df_y_b)
    pass_count_y_b = len(kpittc_df_y_b[kpittc_df_y_b['Status'] == 'Pass'])
    fail_count_y_b = len(kpittc_df_y_b[kpittc_df_y_b['Status'] == 'Fail'])
    verify_manully_y_b = len(kpittc_df_y_b[kpittc_df_y_b['Status'] == 'VERIFYMANUALLY'])
    percentage_y_b = pass_count_y_b / kpittc_count_y_b * 100 if kpittc_count_y_b != 0 else 0
    results_y_b[bench] = {
        'kpittc_count_y_b': kpittc_count_y_b,
        'pass_count_y_b': pass_count_y_b,
        'fail_count_y_b': fail_count_y_b,
        'verify_manully_y_b': verify_manully_y_b,
        'percentage_y_b': percentage_y_b
    }
 
total_scripts_executed = sum(stats['kpittc_count'] for stats in results_t_b.values())
total_passed = sum(stats['pass_count'] for stats in results_t_b.values())
total_failed = sum(stats['fail_count'] for stats in results_t_b.values())
total_manuly = sum(stats['verify_manully'] for stats in results_t_b.values())
overall_percentage = total_passed / total_scripts_executed * 100 if total_scripts_executed != 0 else 0
 
def get_bench_data(df, bench_name):
    bench_data = df[df['Bench Name'] == bench_name]
    results = {
        'TestCase': list(bench_data['TestCase']),
        'Status': list(bench_data['Status']),
        'Failure_Description': list(bench_data['Failure_Description']),
        'Failure_Category': list(bench_data['Failure_Category'])
    }
    return results
 
bench_names = df1['Bench Name'].unique()
results1 = {bench: get_bench_data(df1, bench) for bench in bench_names}
 
def send_email(subject, recipient_email, body):
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    username = 'abc@gmail.com'
    password = 'abc@123'
    sender_email = username
 
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Cc'] = cc_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'html'))
 
    recipient_email=msg['To'].split(",")+msg['Cc'].split(",")
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(username, password)
        server.sendmail(sender_email, recipient_email, msg.as_string())
        print("Email sent successfully!")
    except smtplib.SMTPAuthenticationError as e:
        print("Failed to authenticate. Please check your email and password.")
        print(e)
    except smtplib.SMTPException as e:
        print(f"SMTP error occurred: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        server.quit()
 
if __name__ == "__main__":
    subject = f"[DRT] Overall Execution Report {today.strftime('%Y-%m-%d')}"
    table_style = "width: 70%; border-collapse: collapse; margin-top: 1px;"
    th_td_style = "border: 2px solid #ddd; padding: 2px;"
    body = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <title>[DRT] Overall Execution Summary & Benchwise Details</title>
    <style>
        body {{ font-family: Arial, sans-serif; color: #333; background-color: #f9f9f9; }}
        .container {{ max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #ccc; border-radius: 10px; background-color: #fff; box-shadow: 0px 0px 10px rgba(0, 0, 0, 0.1); }}
        h1 {{ text-align: center; color: #4CAF50; }}
        .rectangular-border {{ border: 2px solid #ddd; border-radius: 8px; padding: 20px; margin-bottom: 20px; }}
        .center{{ text-align: center; }}
        h4{{ font-weight: bold; }}
        table {{ {table_style} }}
        th, td {{ {th_td_style} }}
        th {{ background-color: #4CAF50; color: white; }}
    </style>
</head>
<body>"""
    body+=f"""
    <div class="container">
        <p>Please Find Below Overall Execution Summary & Benchwise Details.</p>
        <p><b>Build:</b> {Build_number}</p>
        <h4>Bench-wise Auto-Flashing Details :</h4>
        <table>
            <tr><th>Bench Name</th><th>TestCase</th><th>Status</th><th>Failure Description</th><th>Failure Category</th><th>Final Result</th></tr>"""
    for bench, data in results1.items():
        #if str(data['Status'][0]).strip("[]',")=='Fail' or str(data['Status'][1]).strip("[]',")=='Fail' or str(data['Status'][2]).strip("[]',")=='Fail':
        #    final_res="Fail"
        final_res = "Fail" if any(str(data['Status'][i]).strip("[]',") == 'Fail' for i in range(3)) else "Pass"
 
        body += f"""
            <tr><td rowspan="3">{bench}</td><td>{str(data['TestCase'][0]).strip("[]',")}</td><td>{str(data['Status'][0]).strip("[]',")}</td><td>{str(data['Failure_Description'][0]).strip("[]',")}</td><td align="center">{str(data['Failure_Category'][0]).strip("[]',")}</td><td rowspan="3">{final_res}</td></tr><tr><td>{str(data['TestCase'][1]).strip("[]',")}</td><td>{str(data['Status'][1]).strip("[]',")}</td><td>{str(data['Failure_Description'][1]).strip("[]',")}</td><td align="center">{str(data['Failure_Category'][1]).strip("[]',")}</td></tr><tr><td>{str(data['TestCase'][2]).strip("[]',")}</td><td>{str(data['Status'][2]).strip("[]',")}</td><td>{str(data['Failure_Description'][2]).strip("[]',")}</td><td align="center">{str(data['Failure_Category'][2]).strip("[]',")}</td></tr>"""
    body += f"""
        </table>
        <h4>Overall Execution Status :</h4>
        <table>
            <tr><th>Overall Execution Count</th><th>Count</th></tr>
            <tr><td>Total Scripts Executed</td><td align="center">{total_scripts_executed}</td></tr>
            <tr><td>Scripts Passed</td><td align="center">{total_passed}</td></tr>
            <tr><td>Scripts Failed</td><td align="center">{total_failed}</td></tr>
            <tr><td>Scripts Verify Manually</td><td align="center">{total_manuly}</td></tr>
            <tr><td>Overall Passing Percentage</td><td align="center">{overall_percentage:.2f}%</td></tr>
        </table>
        <h4>Bench-wise Execution Details :</h4>
        <table>
            <tr><td colspan="3"></td><td colspan="6";align="center">Yesterday</td><td colspan="6";align="center">Today</td></tr>
            <tr><th>Bench Name</th><th>Bench Owner</th><th>Auto Flashing Status</th><th>Total Scripts</th><th>Total Scripts Executed</th><th>Pass</th><th>Fail</th><th>Verify Manually</th><th>Passing Percentage</th><th>Total Scripts</th><th>Total Scripts Executed</th><th>Pass</th><th>Fail</th><th>Verify Manually</th><th>Passing Percentage</th></tr>"""
    #for bench, stats in all_bench_data:
    t_s_y=0
    t_s_e_y=0
    t_s_e_t=0
    for bench_today,stats_today in results_t_b.items():
        for bench_yesterday,stats_yesterday in results_y_b.items():
            for bench_total,stats_total in results_t_t.items():
                for bench_auto,stats_AF in results1.items():
                    if bench_today[0]==bench_yesterday and bench_today[0]==bench_auto and bench_today[0]==bench_total:
                        failure_style=""
                        if stats_today['fail_count']>stats_yesterday['fail_count_y_b']:
                            failure_style="#f44336"
                        elif stats_today['fail_count']<stats_yesterday['fail_count_y_b']:
                            failure_style="#4CAF50"
                        else:
                            failure_style="#fff"
                        failure_style_t=""
                        if int(str(stats_total[bench_today[0]]).strip("{}"))>stats_yesterday['kpittc_count_y_b']:
                            failure_style_t="#f44336"
                        else:
                            failure_style_t="#fff"
                       
                        t_s_y=t_s_y+int(str(stats_total[bench_today[0]]).strip("{}"))
                        t_s_e_y=t_s_e_y+int(stats_yesterday['kpittc_count_y_b'])
                        t_s_e_t=t_s_e_t+int(stats_today['kpittc_count'])
                        final_res = "Fail" if any(str(stats_AF['Status'][i]).strip("[]',") == 'Fail' for i in range(3)) else "Pass"
                        body += f"""
                        <tr><td>{bench_today[0]}</td><td align="center">{bench_today[1]}</td><td align="center">{final_res}</td><td align="center">{str(stats_total[bench_today[0]]).strip("{}")}</td><td align="center";bgcolor={failure_style_t}>{stats_yesterday['kpittc_count_y_b']}</td><td align="center">{stats_yesterday['pass_count_y_b']}</td><td align="center">{stats_yesterday['fail_count_y_b']}</td><td align="center">{stats_yesterday['verify_manully_y_b']}</td><td align="center">{stats_yesterday['percentage_y_b']:.2f}%</td><td align="center">{str(stats_total[bench_today[0]]).strip("{}")}</td><td align="center">{stats_today['kpittc_count']}</td><td align="center">{stats_today['pass_count']}</td>
                        <td align="center";bgcolor={failure_style}>{stats_today['fail_count']}</td>
                        <td align="center">{stats_today['verify_manully']}</td><td align="center">{stats_today['percentage']:.2f}%</td></tr>"""
    body += f"""
        <tr><td colspan="3"></td><td align="center">{t_s_y}</td><td align="center">{t_s_e_y}</td><td colspan="4"></td><td align="center">{t_s_y}</td><td align="center">{t_s_e_t}</td><td colspan="4"></td></tr>
        </table>
        <p><i>**This is a system-generated e-mail; please do not reply.**<i></p>
    </div>
</body>
</html>"""
 
    recipient_email='abc@gmail.com'
    cc_email=""
    send_email(subject, recipient_email, body)