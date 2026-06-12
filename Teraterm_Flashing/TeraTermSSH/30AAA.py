import sys
import serial
import time
import serial.tools.list_ports

def find_silicon_labs_port():
    for port in serial.tools.list_ports.comports():
        print(port)
        if "Silicon Labs CP210x USB to UART Bridge" in port.description:
            return port.device  # Return COM port name (e.g., COM9)
    print("Check COM Port")
    sys.exit()
    return None

def launch_connection(port):
    try:
        ssh_client = serial.Serial(f"COM{port}", 115200, timeout=1)  # Replace with your COM port and baud rate
        print("Serial connection established")
        return ssh_client
    except serial.SerialException as e:
        print("Error opening serial connection: ", e)

def send_command(ssh_client, command):
    ssh_client.write(b"\r" + command.encode() + b"\n")
    time.sleep(0.1)  # delay for 0.1 seconds

def final(sw_variant,hw_variant):
    port = find_silicon_labs_port()
    port = port[3:]
    print(port)
    ssh_client = launch_connection(port)
    try:
        commands = [command.upper() for command in ["MCU_DEBUG00", "MCU_DEBUG01"]]
        for command in commands:
            send_command(ssh_client, command)
            time.sleep(1)
        commands = [command.lower() for command in ["switch command"]]
        for command in commands:
            send_command(ssh_client, command)
            time.sleep(1)
        for i in range(5):  # Send the hard variant code 4 times
            send_command(ssh_client, sw_variant)
            time.sleep(1)
        for i in range(5):  # Send the hard variant code 4 times
            send_command(ssh_client, hw_variant)
            time.sleep(1)
    finally:
        ssh_client.close()
        print("Serial connection closed")
sw_command= "MCU_DEBUG14333041414101"
hw_command= "MCU_DEBUG18010107"
final(sw_command,hw_command)