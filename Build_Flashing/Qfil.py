import sys
import time
import serial
import serial.tools.list_ports
import os
import subprocess

def detect_qfil_port():
    ports = serial.tools.list_ports.comports()
    for port in ports:
        if "Qualcomm" in port.description or "QDLoader" in port.description:
            return port.device
    return None
def print_qfil_port():
    qfil_port = detect_qfil_port()
    if qfil_port:
        print(f"QFIL port: {qfil_port}")
        return qfil_port
    else:
        print("QFIL port not found.")

def run_adb_command(command):
   try:
       result = subprocess.run(command, shell=True, capture_output=True, text=True, check=True)
       return result.stdout, result.stderr, result.returncode
   except subprocess.CalledProcessError as e:
       return "", e.stderr, e.returncode

def enable_adb_edl():
    os.chdir("D:/KITE/KITE/ExternalLib/adb/")
    devices_output, devices_error, devices_code = run_adb_command("adb devices")
    if devices_code !=0:
        print(f"Error : {devices_error}")
        sys.exit()
    lines = devices_output.strip().splitlines()
    d = ""
    for i in lines:
        l = len(i)
        if l == 15:
            d = i[:8]
            break
    if d == "":
        print("ICB Device Not Found")
        sys.exit()
    print(f"adb {d}")
    if len(lines) > 1 and lines[1].strip():
        subprocess.run(f"adb -s {d} reboot edl")

def launch_qfil(command):
    os.chdir("C:/Program Files (x86)/Qualcomm/QPST/bin/")
    #os.chdir("C:/ProgramData/Microsoft/Windows/Start Menu/Programs/QPST")
    # result=subprocess.run(command)
    # print(result)
    a=subprocess.check_output(command, stderr=subprocess.STDOUT)
    converted=a.decode('utf-8')
    os.chdir("D:/Build_Flashing/")
    if "Download Fail" in converted:
        with open('./QFIL_Result.txt', 'w') as f:
            f.write('Fail')
        print("Error While Flashing Please Check")
    else:
        with open('./QFIL_Result.txt', 'w') as f:
            f.write('Pass')

enable_adb_edl()
time.sleep(4)
port=print_qfil_port()
port=port.strip("COM")
command = fr'QFIL.exe -Mode=3 -DOWNLOADFLAT -COM={port} -SEARCHPATH="D:\Build_Flashing\Build" -RawProgram="rawprogram_unsparse0.xml,rawprogram1.xml,rawprogram2.xml,rawprogram3.xml,rawprogram4.xml,rawprogram5.xml" -Patch="patch0.xml,patch1.xml,patch2.xml,patch2.xml,patch3.xml,patch4.xml,patch5.xml"'
launch_qfil(command)