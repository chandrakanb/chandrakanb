import serial.tools.list_ports

def find_silicon_labs_port():
    ports = serial.tools.list_ports.comports()
    for port in ports:
        if "Silicon Labs CP210x USB to UART Bridge" in port.description:
            return port.device  # Return the COM port name (e.g., COM9)
    return None  # Return None if not found

# Example usage
if __name__ == "__main__":
    silicon_labs_port = find_silicon_labs_port()
    if silicon_labs_port:
        print(f"Silicon Labs CP210x USB to UART Bridge found on: {silicon_labs_port}")
    else:
        print("Silicon Labs CP210x USB to UART Bridge not found.")
