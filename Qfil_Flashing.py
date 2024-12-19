import sys
import time
import serial
import serial.tools.list_ports
import os
import subprocess
from datetime import datetime

def detect_qfil_port():
    ports = serial.tools.list_ports.comports()
    for port in ports:
        if "Qualcomm" in port.description or "QDLoader" in port.description:
            return port.device
    return None

def strip_qfil_port():
    qfil_port = detect_qfil_port()
    if qfil_port is not None:
        print(f"QFIL port: {qfil_port}")
        port = qfil_port.strip("COM")
        return port
    else:
        sys.exit("QFIL port not found.")

def run_adb_command(command):
    try:
        result = subprocess.run(command, shell=True, capture_output=True, text=True, check=True)
        return result.stdout, result.stderr, result.returncode
    except subprocess.CalledProcessError as e:
        return "", e.stderr, e.returncode

def enable_adb_edl():
    adb_path = "D:/KITE/KITE/ExternalLib/adb/"
    os.chdir(adb_path)
    devices_output, devices_error, devices_code = run_adb_command("adb devices")
    if devices_code != 0:
        print(f"Error: {devices_error}")
        sys.exit()
    lines = devices_output.strip().splitlines()
    d = ""
    for i in lines:
        if 'device' in i:
            device_id = i.split()[0]
            if 7 <= len(device_id) <= 8:  # Check if its length is 7-8 characters
                d = device_id
                break
    if d == "":
        print("ICB Device Not Found")
        sys.exit()
    print(f"adb {d}")
    if len(lines) > 1 and lines[1].strip():
        subprocess.run(f"adb -s {d} reboot edl")

def launch_qfil(command):
    result_path = os.path.dirname(os.path.abspath(__file__))
    QFIL_Logs = os.path.join(result_path, "QFIL_Logs")
    os.makedirs(QFIL_Logs, exist_ok=True)

    Log_file_name = datetime.now().strftime("%Y_%m_%d_%I_%M_%p") + ".txt"
    Log_file_path = os.path.join(QFIL_Logs, Log_file_name)

    os.chdir("C:/Program Files (x86)/Qualcomm/QPST/bin/")

    try:
        result = subprocess.check_output(command, stderr=subprocess.STDOUT)
        output = result.decode('utf-8')  # Decode binary to string
        
        # Write logs to file
        with open(Log_file_path, 'w') as f:
            f.write(output)

        # Check for errors in output
        if "Download Fail" in output:
            with open(os.path.join(result_path, "flashing_result.txt"), 'w') as f:
                f.write('Fail')
            print("Error While Flashing. Please Check.")
        else:
            with open(os.path.join(result_path, "flashing_result.txt"), 'w') as f:
                f.write('Pass')
            print("Flashing Completed Successfully.")
    except subprocess.CalledProcessError as e:
        error_message = e.output.decode('utf-8') if e.output else "Unknown Error"
        with open(os.path.join(result_path, "flashing_result.txt"), 'w') as f:
            f.write('Fail')
        print(f"Error executing QFIL command: {error_message}")

if len(sys.argv) != 2:
    sys.exit("Usage: python_script.py <ufs_path>")
else:
    UFSPATH = sys.argv[1]
    #enable_adb_edl()
    time.sleep(4)
    port = strip_qfil_port()
    command = (
        fr'QFIL.exe -Mode=3 -DOWNLOADFLAT -COM={port} '
        fr'-SEARCHPATH="{UFSPATH}" '
        fr'-RawProgram="rawprogram_unsparse0.xml,rawprogram1.xml,rawprogram2.xml,rawprogram3.xml,rawprogram4.xml,rawprogram5.xml" '
        fr'-Patch="patch0.xml,patch1.xml,patch2.xml,patch2.xml,patch3.xml,patch4.xml,patch5.xml"'
    )
    launch_qfil(command)
