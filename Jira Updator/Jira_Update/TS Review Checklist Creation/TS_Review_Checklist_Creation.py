import os, pandas as pd, openpyxl

input_file = os.path.join("input.xlsx")
data_file = os.path.join("data.xlsx")
output_files = "output_files"

# Create the output_files folder if it doesn't exist
if not os.path.exists(output_files):
    os.makedirs(output_files)
    print(f'Created output_files folder: "{output_files}"\n')  # Informative message

# === READ EXCEL ===
try:
    df = pd.read_excel(input_file)
except FileNotFoundError:
    print(f"Error: Input file not found at {input_file}\n")
    exit()  # Exit the script if the input file is missing

for _, row in df.iterrows():
    try:
        project_name = row["Project Name"]
        author = row["Author"]
        review_date = row["Review Date"]
        test_script_id = row["Test Script ID"]
        reviewer = row["Reviewer"]
        review_type = row["Review Type"]
        
        output_files_file = os.path.join(output_files, f"DRT_TS_Review_Checklist_{test_script_id}_Self.xlsx")
        output_files_file_name = f"DRT_TS_Review_Checklist_{test_script_id}_Self.xlsx"
        
        print(f"Project Name: {project_name}")
        print(f"Author: {author}")
        print(f"Review Date: {review_date}")
        print(f"Test Script ID: {test_script_id}")
        print(f"Reviewer: {reviewer}")
        print(f"Review Type: {review_type}")
        
        try:
            # Load the workbook
            workbook = openpyxl.load_workbook(data_file)

            # Select the active sheet
            sheet = workbook.active

            # Edit the cells
            sheet["C3"] = project_name  # Use strings for cell addresses
            sheet["C4"] = author
            sheet["C5"] = review_date
            sheet["C6"] = test_script_id
            sheet["C7"] = reviewer
            sheet["C8"] = review_type
            
            # Save the workbook
            workbook.save(output_files_file)

            print(f"Created {output_files_file_name} successfully.\n")

        except FileNotFoundError:
            print(f"Error: Data file not found at {data_file}\n")
            exit()  # Exit the script if the data file is missing
        except Exception as e:
            print(f"Error processing data file: {e}\n") # more specific error message

    except Exception as e:
        print(f"Error processing row: {e}\n")  # More descriptive error message