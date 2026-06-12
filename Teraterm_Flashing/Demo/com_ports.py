import serial.tools.list_ports

def print_all_com_ports():
    ports = serial.tools.list_ports.comports()
    if ports:
        print("Available COM Ports:")
        for port in ports:
            print(f"Port: {port.device}, Description: {port.description}, HWID: {port.hwid}")
    else:
        print("No COM ports available.")

# Example usage
if __name__ == "__main__":
    print_all_com_ports()
