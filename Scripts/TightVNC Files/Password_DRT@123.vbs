Set WshShell = WScript.CreateObject("WScript.Shell")

' Wait for 5 seconds
WScript.Sleep 2000

' Set the password
Password = "DRT@123"

' Paste the password
WshShell.SendKeys Password

' Press Enter
WshShell.SendKeys "{ENTER}"
