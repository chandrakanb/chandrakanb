# This is AI generated code, please refer KPIT AI Policy before using this in your projects
import serial
import time

# Serial port settings (adjust to your setup)
PORT = "COM9"  # <-----  MAKE SURE THIS IS CORRECT!
BAUD_RATE = 9600
TIMEOUT = 1

def get_time_from_mcu():
    """Sends a command to the MCU and reads the time response."""
    try:
        ser = serial.Serial(PORT, BAUD_RATE, timeout=TIMEOUT)
        time.sleep(2) # Allow the serial connection to initialize

        ser.write(b"GET_TIME\n")  # Send the command (must be bytes)
        time.sleep(0.1)  #Allow time for the MCU to respond.
        response = ser.readline().decode('utf-8').strip()  # Read the response
        ser.close()

        if response:
            print(f"MCU Time: {response}")
            return response # You'll need to parse this string
        else:
            print("No response from MCU.")
            return None

    except serial.SerialException as e:
        print(f"Serial communication error: {e}")
        print(f"Please check the following:")
        print(f"- Is the device connected?")
        print(f"- Is the port name correct (check Device Manager on Windows)?")
        print(f"- Is the port being used by another program?")
        print(f"- Are the drivers installed correctly?")
        return None
    except Exception as e:
        print(f"An error occurred: {e}")
        return None

if __name__ == "__main__":
    time_str = get_time_from_mcu()
    if time_str:
        # Parse time_str into hours, minutes, seconds (example)
        try:
            hours, minutes, seconds = map(int, time_str.split(':'))
            print(f"Parsed time: {hours:02d}:{minutes:02d}:{seconds:02d}")
        except ValueError:
            print("Error parsing time string.  Check the format.")