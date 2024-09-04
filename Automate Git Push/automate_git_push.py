import os
import subprocess
import time

# Open Command Prompt in the current directory
subprocess.Popen("cmd.exe", cwd=os.getcwd())

# Wait for a few seconds for the Command Prompt to open
time.sleep(2)

# Trigger the VBScript to send keystrokes to the Command Prompt
subprocess.Popen(['cscript', '//nologo', 'send_keys_git_push.vbs'])
