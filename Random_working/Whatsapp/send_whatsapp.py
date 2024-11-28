"""import pywhatkit as pwk
# import pywhatkit.core as core

try:
    # core.check_connection()  # Test internet connectivity
    # print("Internet check passed.")
    phone_number = "+919923604110"
    message = "Test msg."
    pwk.sendwhatmsg_instantly(phone_number, message)
except Exception as e:
    print("Error occurred:", e)
"""


import pywhatkit as pwk

try:
    pwk.sendwhatmsg(
        phone_no="+919923604110", 
        message="Hello, this is a test message!", 
        time_hour=6, 
        time_min=21, 
        wait_time=20, 
        tab_close=True, 
        close_time=5
    )
except Exception as e:
    print(f"An error occurred: {e}")
