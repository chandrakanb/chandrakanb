import git
import os
import shutil
from datetime import datetime, timedelta
import sys

def copy_files(src_dir, dest_dir):
    if not os.path.exists(src_dir):
        raise FileNotFoundError("Source directory does not exist")
    if not os.path.exists(dest_dir):
        os.makedirs(dest_dir)
    files = [f for f in os.listdir(src_dir) if os.path.isfile(os.path.join(src_dir, f))]
    for file in files:
        src_file = os.path.join(src_dir, file)
        dest_file = os.path.join(dest_dir, file)
        shutil.copy2(src_file, dest_file)
        #print(f"Copied {src_file} to {dest_file}")

def recreate_folder(dir):
    """Recreate the specified directory."""
    try:
        if os.path.exists(dir):
            if sys.platform.startswith('win'):
                os.system(f'rmdir /s /q "{dir}"')
            else:
                shutil.rmtree(dir)
            print(f"Folder '{dir}' deleted successfully.")
    except Exception as e:
        print("Error deleting folder:", e)

def clone(url, dir, token, branch):
    """Clone the Git repository."""
    url_with_credentials = url.replace("https://", f"https://oauth2:{token}@")
    try:
        git.Repo.clone_from(url_with_credentials, dir, branch=branch)
        print(f"Repo Cloned at {dir}")
    except git.GitCommandError as e:
        print(f"Error while cloning repo: {e}")

def create_folder_and_copy(dir, pdf_path, folder):
    """Create a folder and copy the Excel data into it."""
    try:
        print(dir)
        new_folder_path = os.path.join(dir, folder)
        os.makedirs(new_folder_path, exist_ok=True)
        print(f"folder Created for {new_folder_path}")
        copy_files(pdf_path, new_folder_path)
    except Exception as e:
        print(f"An error occurred while creating folder and copying data: {e}")

def git_push(dir, token, branch):
    """Commit and push changes to the Git repository."""
    try:
        repo = git.Repo(dir)
        origin = repo.remote(name='origin')
        repo.git.add(update=True)
        repo.git.add(A=True)
        repo.git.commit('-m', 'Bench Files Updated')
        origin.push(refspec=f"{branch}:{branch}")
        print("Pushed changes successfully.")
    except git.GitCommandError as e:
        print(f"Error while pushing changes: {e}")

if __name__ == '__main__':
    try:
        #url = "https://" + sys.argv[1]
        url = "https://gitlab.kpit.com/lumeshj/build.git"
        #branch = sys.argv[2]
        branch = "OvernightReport_Backup"
        #dir = sys.argv[1]  # "D:/git_excel_demo/"
        dir = "D:/Jenkins_slave/Execution_PDF/"
        #token = sys.argv[4]
        token = "upEz9E84KGesbGaus7pv"
        bench_name = "Bench" + sys.argv[1]  # bench Name
        build_type = sys.argv[2] + "/"
        excel_path = sys.argv[3]

        current_date1 = datetime.now().strftime("%d_%m_%y")
        yesterday = datetime.now() - timedelta(days=1)
        yesterday = yesterday.strftime("%d_%m_%y")
        matching_folder_AF = next((dir_name for root, dirs, _ in os.walk(excel_path) for dir_name in dirs if
                                   dir_name.startswith(f'AF_ExecutionReport_{yesterday}')), None)
        if not matching_folder_AF:
            matching_folder_AF = next((dir_name for root, dirs, _ in os.walk(excel_path) for dir_name in dirs if
                                       dir_name.startswith(f'AF_ExecutionReport_{current_date1}')), None)
        matching_folder_DRT = next((dir_name for root, dirs, _ in os.walk(excel_path) for dir_name in dirs if
                                    dir_name.startswith(f'DRT_ExecutionReport_{yesterday}')), None)
        if not matching_folder_DRT:
            matching_folder_DRT = next((dir_name for root, dirs, _ in os.walk(excel_path) for dir_name in dirs if
                                        dir_name.startswith(f'DRT_ExecutionReport_{current_date1}')), None)
        if not matching_folder_AF and matching_folder_DRT:
            raise FileNotFoundError(f"No matching folder found for the current date: {current_date1}")

        current_date = datetime.now().strftime("%Y-%m-%d")
        pdf_path_execution = f"{excel_path}{matching_folder_DRT}\\XMLReports\\PdfReports\\"
        pdf_path_AF = f"{excel_path}{matching_folder_AF}\\XMLReports\\PdfReports\\"

        os.makedirs(dir, exist_ok=True)
        recreate_folder(dir)
        clone(url, dir, token, branch)
        os.makedirs(dir, exist_ok=True)
        if os.path.exists(pdf_path_execution):
            create_folder_and_copy(
                dir + "/OvernightReport_Backup/" + bench_name + "/" + current_date + "/" + build_type,
                pdf_path_execution, folder="Bench_Results")

        if os.path.exists(pdf_path_AF):
            create_folder_and_copy(
                dir + "/OvernightReport_Backup/" + bench_name + "/" + current_date + "/" + build_type, pdf_path_AF,
                folder="Auto_Flashing_Results")

        excel_path_execution = f"{excel_path}{matching_folder_DRT}\\DRT_Execution_Report.xlsx"
        if os.path.exists(excel_path_execution):
            shutil.copy2(excel_path_execution,
                         dir + "/OvernightReport_Backup/" + bench_name + "/" + current_date + "/" + build_type + "/Bench_Results/DRT_Execution_Report.xlsx")
            print(f"Excel Copied {excel_path_execution}")

        # data push
        git_push(dir, token, branch)

    except IndexError:
        print("Usage: python script_name.py <url> <branch> <dir> <token> <bench_name> <build_type> <excel_path>")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
