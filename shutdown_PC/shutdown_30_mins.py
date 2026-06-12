import os
import time

# Set the delay in minutes
shutdown_delay = 30

# Notify the user
print(f"The laptop will shut down in {shutdown_delay} minutes.")

# Wait for the specified delay
time.sleep(shutdown_delay)

# Shutdown command (works on Windows)
os.system("shutdown /s /t 0")
