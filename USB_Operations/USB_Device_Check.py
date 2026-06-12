import os
import sys
import ctypes
import subprocess
import brainstem
from brainstem.result import Result
import time
import psutil  # Replaces wmi for system info
# import pyusb For USB operations (you may need to install pyusb)
from time import sleep
# Create an instance of a USBCSwitch module.
cswitch = brainstem.stem.USBCSwitch()
Port = 1
# Locate and connect to the first object you find on USB
def locateAndConnectFirstObject():
   result = cswitch.discoverAndConnect(brainstem.link.Spec.USB)
   if result != Result.NO_ERROR:
       print("Error %d encountered connecting to BrainStem Module.\n" % (result))
       sys.exit(1)
   else:
       print("Connected to BrainStem module.\n")
# Prep USBCSwitch for testing
def prepUSBSwitch():
   cswitch.usb.setPortDisable(0)
   cswitch.mux.setEnable(False)
   cswitch.mux.setChannel(0)
# Enable Port AND Mux
def enablePortAndMux():
   cswitch.usb.setPortEnable(0)
   cswitch.mux.setEnable(True)
# When finished with Switch using, deinitialize it
def deInitializeSwitch():
   cswitch.usb.setPortDisable(0)
   cswitch.mux.setEnable(False)
# Disconnect from Switch device  
def disconnectSwitch():
   cswitch.disconnect()
# Check if external USB is connected using psutil
def check_external_usb_connected():
   retry_count = 0
   max_retries = 10
   while retry_count < max_retries:
       time.sleep(10)
       external_usb_devices = []
       # psutil.disk_partitions() gets information about all partitions
       partitions = psutil.disk_partitions()
       for partition in partitions:
           if 'removable' in partition.opts:
               external_usb_devices.append(partition.device)
       if external_usb_devices:
           print("External USB device(s): Connected")
           break  # Exit the loop if USB device is found
       else:
           print(f"External USB device(s): Not Connected, Retry {retry_count + 1}/{max_retries}")
           deInitializeSwitch()
           disconnectSwitch()
           locateAndConnectFirstObject()
           prepUSBSwitch()
           enablePortAndMux()
           cswitch.mux.setChannel(Port)
           retry_count += 1
   if retry_count == max_retries:
       print("Max retries reached. USB device not found.")
# Check USB permission
def check_usb_permission():
   partitions = psutil.disk_partitions()
   for partition in partitions:
       if 'removable' in partition.opts:
           drive = partition.device
           # Check read/write permission
           if os.access(drive, os.R_OK) and os.access(drive, os.W_OK):
               print(f"Drive: {drive}")
               print("Read/Write permission: Granted")
               print("USB device is ready to use")
           else:
               print(f"Drive: {drive}")
               print("Read/Write permission: Denied")
def start():
   locateAndConnectFirstObject()
   prepUSBSwitch()
   enablePortAndMux()
   cswitch.mux.setChannel(Port)
   check_external_usb_connected()
   check_usb_permission()
start()
