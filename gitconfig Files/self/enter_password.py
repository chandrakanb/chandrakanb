import time
import pyautogui

# The string to type after waiting
text_to_type = "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"

# Wait for 10 seconds
print("Waiting for 10 seconds before typing...")
time.sleep(10)

# Simulate typing the text
pyautogui.typewrite(text_to_type)

print("Text has been typed.")
