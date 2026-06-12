import subprocess
import os
import sys
import datetime

def run_adb_command(command):
    try:
        result = subprocess.run(command, shell=True, capture_output=True, text=True, check=True)
        return result.stdout, result.stderr, result.returncode
    except subprocess.CalledProcessError as e:
        return "", e.stderr, e.returncode

def check_adb_device():
    try:
        devices_output, devices_error, devices_code = run_adb_command("adb devices")
        if devices_code != 0:
            print(f"Error: {devices_error}")
            sys.exit()
        lines = devices_output.strip().splitlines()
        d = ""
        for i in lines:
            l = len(i)
            if l == 17:
                d = i[:8]
                return d
        if d == "":
            print("ICB Device Not Found")
            # Get the current directory to write the file there
            log_directory = os.getcwd()
            file_path = os.path.join(log_directory, 'adb_devices.txt')
            with open(file_path, 'w') as f:
                f.write('ICB Device Not Found Fail')
            print("Error While Capturing Logs")
            return d
    except FileNotFoundError:
        print("Error: adb command not found.  Ensure ADB is in your system's PATH.")
        sys.exit()


def execute_log_collection(argument):
    """
    Executes ADB commands to collect logs, creates a folder, and saves the logs.

    Args:
        argument (str): The argument to be included in the folder name.
    """

    # Get current date and time
    now = datetime.datetime.now()
    timestamp = now.strftime("%d_%m_%Y_%H_%M_%S")

    # Create folder name
    folder_name = f"{argument}_{timestamp}"
    folder_path = os.path.join("D:", "folder1", folder_name)

    # Create the folder
    try:
        os.makedirs(folder_path)
        print(f"Folder created: {folder_path}")
    except FileExistsError:
        print(f"Folder already exists: {folder_path}")
    except Exception as e:
        print(f"Error creating folder: {e}")
        return

    # Change directory to the created folder
    os.chdir(folder_path)

    # Execute ADB commands
    try:
        # Check ADB Device
        device_id = check_adb_device()
        if device_id == "":
            print("ADB device not found. Exiting.")
            return

        # Run ADB commands
        run_adb_command("adb shell rm -rf /sdcard/ICB_Log")
        run_adb_command("adb shell am broadcast -a com.honda.auto.action.EXPORT_LOGS --user 0")
        run_adb_command("adb pull /sdcard/ICB_Log")
        run_adb_command("adb bugreport")
        run_adb_command("adb logcat -d all >logcat.txt")

        print("Log collection completed successfully.")

    except Exception as e:
        print(f"Error during log collection: {e}")

# This is AI generated code, please refer KPIT AI Policy before using this in your projects

if __name__ == "__main__":
    if len(sys.argv) > 1:
        argument_value = sys.argv[1]
        execute_log_collection(argument_value)
    else:
        print("Error: Please provide an argument.")
        print("Usage: python log_collection.py <argument>")