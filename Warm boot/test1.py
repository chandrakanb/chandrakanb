from adbutils import adb

device = adb.device()  # Connects to the first available device

result = device.shell("date")
print(result)