import pandas as pd
import subprocess
import re
import datetime
from adbutils import adb
import sys
import os

def get_datetime_from_adb(device_id, output_path):
    """
    Gets date and time from an ADB-connected device.

    Args:
        device_id (str): The ID of the ADB device.

    Returns:
        tuple: A tuple containing (Day, Month, Year, Hour, Minute, Time Period)
               or None if there was an error.
    """
    try:
        # Execute adb shell command to get date and time
        device = adb.device(serial=device_id)
        result = device.shell("date")
        # Use regex to parse the output
        match = re.match(r"(\w+) (\w+) (\d+) (\d{2}):(\d{2}):(\d{2}) (\w+) (\d{4})", result)
        if match:
            date = match.group(3)
            month_name = match.group(2)
            year = match.group(8)
            hour_int = int(match.group(4))  # Convert hour to integer
            hour_str = match.group(4)
            minute = match.group(5)
            # time_period = match.group(7)

            # Convert month name to number
            month_dict = {
                "Jan": "01", "Feb": "02", "Mar": "03", "Apr": "04",
                "May": "05", "Jun": "06", "Jul": "07", "Aug": "08",
                "Sep": "09", "Oct": "10", "Nov": "11", "Dec": "12"
            }
            month = month_dict.get(month_name)
            # Adjust hour for AM/PM
            if hour_int > 12:
                time_period = "PM"
                hour_str = hour_int - 12
                
                if hour_str < 10:
                   hour_str = (f"0{hour_str}")
                   print(hour_str)
            elif hour_int == 12:
                time_period = "PM"  # Noon
            elif hour_int == 0:
                time_period = "AM"
                hour_str = 12  # Midnight
            else:
                time_period = "AM"

            return (date, month, year, hour_str, minute, time_period)
        else:
            print(f"Error: Could not parse date/time from ADB output:\n{result}")
            return None
    except adbutils.exceptions.AdbCommandFailed as e:
        print(f"Error executing ADB command: {e}")
        return None
    except adbutils.exceptions.DeviceNotConnected as e:
        print(f"Device not connected or not authorized: {e}")
        return None
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        return None

if __name__ == "__main__":
    import sys

    if len(sys.argv) != 2:
        print("Usage: python script.py <device_id>")
        sys.exit(1)

    device_id = sys.argv[1]
    output_path = "D:/chandrakanb/chandrakanb/Warm boot"

    datetime_info = get_datetime_from_adb(device_id, output_path)

    if datetime_info:
        date, month, year, hour, minute, time_period = datetime_info
        
        data = {'Sr. No.': [1, 2, 3, 4, 5, 6],
		'Key Name': ['ICBDate', 'ICBMonth', 'ICBYear', 'ICBHour', 'ICBMinute', 'ICBTimePeriod'],
		'Value': [date, month, year, hour, minute, time_period]}

        df = pd.DataFrame(data)

        # Use os.path.join for cross-platform compatibility
        filepath = os.path.join(output_path, "output2.xlsx")

        # Check if the directory exists, and create it if it doesn't
        directory = os.path.dirname(filepath)  # Extract the directory part
        if not os.path.exists(directory):
            os.makedirs(directory)  # Create the directory (and any parent directories if needed)

        df.to_excel(filepath, index=False)
        print(f"Successfully created Excel file at: {filepath}")
        
        print(f"Day: {date}")
        print(f"Month: {month}")
        print(f"Year: {year}")
        print(f"Hour: {hour}")
        print(f"Minute: {minute}")
        print(f"Time Period: {time_period}")