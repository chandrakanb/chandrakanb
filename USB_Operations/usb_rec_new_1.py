import os
import subprocess

def reconnect_usb():
    try:
        # PowerShell command to get the USB storage device (adjust for different device types)
        get_usb_device_cmd = '''
        Get-PnpDevice | Where-Object { $_.FriendlyName -like "*USB Mass Storage Device*" -and $_.Status -eq "OK" } | Select-Object -ExpandProperty InstanceId
        '''
        
        # Run PowerShell command to get the USB device instance ID
        usb_instance_id = subprocess.run(
            ['powershell', '-Command', get_usb_device_cmd],
            capture_output=True, text=True
        ).stdout.strip()

        if not usb_instance_id:
            print("No USB mass storage devices found.")
            return

        # Escape the '&' character in InstanceId for PowerShell
        usb_instance_id_escaped = usb_instance_id.replace("&", "`&")

        # PowerShell command to disable the USB device
        disable_cmd = f'powershell Disable-PnpDevice -InstanceId "{usb_instance_id_escaped}" -Confirm:$false'
        os.system(disable_cmd)
        print("USB device disabled successfully.")

        # PowerShell command to re-enable the USB device
        enable_cmd = f'powershell Enable-PnpDevice -InstanceId "{usb_instance_id_escaped}" -Confirm:$false'
        os.system(enable_cmd)
        print("USB device re-enabled successfully.")

        # Simple PowerShell command to rescan hardware devices
        rescan_cmd = 'powershell "Get-PnpDevice -PresentOnly | Out-Null"'
        os.system(rescan_cmd)
        print("System hardware scan triggered.")

    except Exception as e:
        print(f"Error reconnecting USB device: {e}")

# Call the function to reconnect the USB
reconnect_usb()
