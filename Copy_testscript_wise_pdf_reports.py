import os
import shutil

def copy_pdf_reports(source_root, destination_root):
    """
    Copies PDF files from folders starting with "B" within the source root,
    specifically from the "XMLReport" "PdfReports" subdirectory,
    into new folders within the destination root, named after the "B" name 
    with "_Iteration_1" removed.

    Args:
        source_root (str): The root directory to search for "B" directories.
        destination_root (str): The root directory where the new folders and PDF files will be copied.
    """

    for folder_name in os.listdir(source_root):
        if folder_name.startswith("B"):  # Corrected condition
            source_path = os.path.join(source_root, folder_name, "XMLReport", "PdfReports")
            
            # Check if the PdfReports directory exists
            if os.path.exists(source_path):
                # Create destination folder name
                destination_folder_name = folder_name.replace(" B", "") #remove leading space
                destination_folder_name = destination_folder_name.replace("_Iteration_1", "")
                destination_path = os.path.join(destination_root, destination_folder_name)
                
                # Create the destination folder if it doesn't exist
                os.makedirs(destination_path, exist_ok=True)

                # Copy PDF files
                for filename in os.listdir(source_path):
                    if filename.endswith(".pdf"):
                        source_file_path = os.path.join(source_path, filename)
                        destination_file_path = os.path.join(destination_path, filename)
                        shutil.copy2(source_file_path, destination_file_path) #copy2 preserves metadata
                print(f"Copied PDF reports from {folder_name} to {destination_folder_name}")
            else:
                print(f"PdfReports directory not found in {folder_name}")

# Example Usage (replace with your actual paths)
source_directory = "."  # Current directory
destination_directory = "PDF REPORTS"

copy_pdf_reports(source_directory, destination_directory)