import os

# Define the current working directory where folders were previously renamed
directory = os.getcwd()

# Function to rename PDF files
def rename_pdf_files(directory):
    for folder_name in os.listdir(directory):
        folder_path = os.path.join(directory, folder_name)
        if os.path.isdir(folder_path):
            # Look for files ending with DetailedReport.pdf
            for file_name in os.listdir(folder_path):
                if file_name.endswith("DetailedReport.pdf"):
                    old_file_path = os.path.join(folder_path, file_name)
                    new_file_name = f"{folder_name}_DetailedReport.pdf"
                    new_file_path = os.path.join(folder_path, new_file_name)
                    os.rename(old_file_path, new_file_path)
                    print(f"Renamed '{file_name}' in '{folder_name}' to '{new_file_name}'")
                
                # Look for files ending with DashboardReport.pdf
                elif file_name.endswith("DashboardReport.pdf"):
                    old_file_path = os.path.join(folder_path, file_name)
                    new_file_name = f"{folder_name}_DashboardReport.pdf"
                    new_file_path = os.path.join(folder_path, new_file_name)
                    os.rename(old_file_path, new_file_path)
                    print(f"Renamed '{file_name}' in '{folder_name}' to '{new_file_name}'")

# Call the function
if __name__ == "__main__":
    rename_pdf_files(directory)
