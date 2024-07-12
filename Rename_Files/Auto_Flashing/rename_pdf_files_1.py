import os
import time
import re

# Define the current working directory where folders were previously renamed
directory = os.getcwd()

# Function to rename PDF files
def rename_pdf_files(directory):
    while True:
        try:
            for folder_name in os.listdir(directory):
                folder_path = os.path.join(directory, folder_name)
                if os.path.isdir(folder_path):
                    # Rename DetailedReport.pdf
                    detailed_report_old_name = "DetailedReport"
                    rename_pdf_by_pattern(folder_path, detailed_report_old_name, folder_name)
                    
                    # Rename DashboardReport.pdf
                    dashboard_report_old_name = "DashboardReport"
                    rename_pdf_by_pattern(folder_path, dashboard_report_old_name, folder_name)
            
            # If no error occurs, break out of the loop
            break
        
        except FileNotFoundError as e:
            time.sleep(5)  # Wait for 5 seconds before retrying

# Function to rename PDF files by pattern
def rename_pdf_by_pattern(folder_path, file_prefix, folder_name):
    for file_name in os.listdir(folder_path):
        match = re.match(r"Bench09_Auto_Flashing_Reports_(\d{1,2}[a-zA-Z]{2}_\w+)", folder_name)
        if file_name.startswith(file_prefix):
            if match:
                date_part = match.group(1)
                new_file_name = f"Bench09_Auto_Flashing_{date_part}_{file_prefix}.pdf"
                old_file_path = os.path.join(folder_path, file_name)
                new_file_path = os.path.join(folder_path, new_file_name)
                os.rename(old_file_path, new_file_path)
                print(f"Renamed '{file_name}' in '{folder_name}' to '{new_file_name}'")

# Call the function
if __name__ == "__main__":
    rename_pdf_files(directory)
