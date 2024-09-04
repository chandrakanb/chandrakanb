import os
import webbrowser

# Function to open a file in Edge browser
def open_in_edge(file_path):
    # Microsoft Edge browser path; adjust the path if necessary
    edge_path = "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe %s"
    webbrowser.get(edge_path).open(file_path)

# Get the current directory
directory = os.getcwd()

# Loop through all subfolders in the current directory
for subfolder in os.listdir(directory):
    subfolder_path = os.path.join(directory, subfolder)
    
    # Check if it's a directory and its name starts with a digit
    if os.path.isdir(subfolder_path) and subfolder[0].isdigit():
        # Define the path to the HTMLReport folder and HTMLReport.html file
        html_report_folder = os.path.join(subfolder_path, "HTMLReport")
        html_report_file = os.path.join(html_report_folder, "HTMLReport.html")
        
        # Check if HTMLReport folder and HTMLReport.html file exist
        if os.path.exists(html_report_file):
            print(f"Opening {html_report_file}")
            open_in_edge(html_report_file)
        else:
            print(f"HTMLReport.html does not exist in {html_report_folder}.")
