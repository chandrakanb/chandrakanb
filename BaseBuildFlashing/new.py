import sys
import time
"""port = "COM9"
port = port[3:]
print(port)"""

"""port = "COM9"
port = port.strip("COM")
print(port)"""

"""port = None
port = port.strip("COM")
print(port)"""

"""if len(sys.argv) != 2:
        sys.exit("Usage: python_script.py <ufs_path>")
else:
    port = sys.argv[1]
    print(port)
    time.sleep(4)
    command = fr'COM={port}'
    print(command)"""


"""
port = None
port = port[3:]
print(port)
"""


"""
port = None
if port is not None:
    port = port.strip("COM")
else:
    port = "DefaultValue"
"""

"""
import os
QFIL_Logs = os.path.join(os.path.dirname(os.path.abspath(__file__)), "QFIL_Logs")
os.makedirs(QFIL_Logs, exist_ok=True)
print(f"Folder ready at: {QFIL_Logs}")
"""

import os
from datetime import datetime

QFIL_Logs = os.path.join(os.path.dirname(os.path.abspath(__file__)), "QFIL_Logs")
print(QFIL_Logs)
os.makedirs(QFIL_Logs, exist_ok=True)
Log_file_name = datetime.now().strftime("%Y_%m_%d_%I_%M_%p") + ".txt"
print(Log_file_name)
Log_file_path = os.path.join(QFIL_Logs, Log_file_name)
print(Log_file_path)
