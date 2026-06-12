import os
import time

# PowerShell script to find and restart USB Mass Storage controller
powershell_script = '''
$usbStorageDevice = Get-PnpDevice -PresentOnly | Where-Object { $_.InstanceId -like '*USBSTOR*' }
if ($usbStorageDevice) {
    # Disable USB Mass Storage Device
    Disable-PnpDevice -InstanceId $usbStorageDevice.InstanceId -Confirm:$false
    Start-Sleep -Seconds 2

    # Re-enable USB Mass Storage Device
    Enable-PnpDevice -InstanceId $usbStorageDevice.InstanceId -Confirm:$false
    Write-Host "USB storage device reconnected successfully."
} else {
    Write-Host "USB storage device not found."
}
'''

# Save the PowerShell script to a temporary file
with open("reconnect_usb.ps1", "w") as file:
    file.write(powershell_script)

# Run the PowerShell script using os.system
os.system('powershell -ExecutionPolicy Bypass -File reconnect_usb.ps1')

# Clean up the temporary file
os.remove("reconnect_usb.ps1")
