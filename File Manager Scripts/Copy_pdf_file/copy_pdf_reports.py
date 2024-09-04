import os
import shutil

def create_root_pdf_reports(script_dir):
    # Define the path for the root PdfReports directory
    root_pdf_reports_path = os.path.join(script_dir, 'PdfReports')

    # Create the root PdfReports directory if it doesn't exist
    if not os.path.exists(root_pdf_reports_path):
        os.makedirs(root_pdf_reports_path)
        print(f"Created root directory: {root_pdf_reports_path}")

    return root_pdf_reports_path

def copy_pdf_reports():
    # Get the directory where this script is located
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # Create the PdfReports directory in the root of the repository
    root_pdf_reports_path = create_root_pdf_reports(script_dir)

    # Iterate through each folder in the repository
    for root, dirs, files in os.walk(script_dir):
        # Skip the root PdfReports directory to avoid copying into itself
        if root == root_pdf_reports_path:
            continue
        
        # Define the path to check for the \XMLReport\PdfReports directory
        xml_pdf_reports_path = os.path.join(root, 'XMLReport', 'PdfReports')
        
        # Check if the \XMLReport\PdfReports directory exists
        if os.path.exists(xml_pdf_reports_path):
            # Copy the contents of \XMLReport\PdfReports to the root PdfReports directory
            for item in os.listdir(xml_pdf_reports_path):
                source_path = os.path.join(xml_pdf_reports_path, item)
                destination_path = os.path.join(root_pdf_reports_path, os.path.relpath(root, script_dir), item)

                # Ensure the subdirectory structure is maintained
                os.makedirs(os.path.dirname(destination_path), exist_ok=True)

                if os.path.isdir(source_path):
                    shutil.copytree(source_path, destination_path, dirs_exist_ok=True)
                else:
                    shutil.copy2(source_path, destination_path)
                
                print(f"Copied {source_path} to {destination_path}")

if __name__ == "__main__":
    copy_pdf_reports()
