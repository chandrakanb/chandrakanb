import pandas as pd
from datetime import datetime
import sys
import os

def create_excel_with_datetime(output_path):
    """
    Creates an Excel file containing current date and time information.

    Args:
        output_path (str, optional): The path to save the Excel file.
                                     Defaults to "output.xlsx" in the current directory.
    """
    try:
        now = datetime.now()
        year = now.strftime('%Y')
        month = now.strftime('%m')
        date = now.strftime('%d')
        hour = now.strftime("%I")
        time_period = now.strftime("%p")
        minute = now.strftime("%M")

        data = {'Sr. No.': [1, 2, 3, 4, 5, 6],
                'Key Name': ['Date', 'Month', 'Year', 'Hour', 'Minute', 'TimePeriod'],
                'Value': [date, month, year, hour, minute, time_period]}

        df = pd.DataFrame(data)

        # Use os.path.join for cross-platform compatibility
        filepath = os.path.join(output_path, "output.xlsx")

        # Check if the directory exists, and create it if it doesn't
        directory = os.path.dirname(filepath)  # Extract the directory part
        if not os.path.exists(directory):
            os.makedirs(directory)  # Create the directory (and any parent directories if needed)

        df.to_excel(filepath, index=False)
        print(f"Successfully created Excel file at: {filepath}")

    except FileNotFoundError as e:
        print(f"Error: The specified directory was not found: {e}")
    except PermissionError as e:
        print(f"Error: Permission denied to write to the directory: {e}")
    except IOError as e:
        print(f"Error: An error occurred while writing to the file: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    # Get the output path from command-line arguments, if provided
    create_excel_with_datetime("D:/chandrakanb/chandrakanb/Warm boot")
    #create_excel_with_datetime("D:/KITE/KITE/KITE_DATA/")