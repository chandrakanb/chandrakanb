Set WshShell = WScript.CreateObject("WScript.Shell")

' Display message for waiting for the program to start
WScript.Echo "Waiting for the program to start..."
' WScript.Sleep 5000 ' Wait for the program to start (5 seconds)

' Activate the GlobalProtect window
WScript.Echo "Activating the GlobalProtect window..."
WshShell.AppActivate "GlobalProtect"
WScript.Sleep 1000 ' Wait for the dialog box to be activated

' Display message for Tabbing to the Portal field
WScript.Echo "Tabbing to the Portal field..."
WshShell.SendKeys "{TAB}"
WScript.Sleep 500 ' Wait for half a second

' Display message for entering the Portal
WScript.Echo "Entering Portal..."
WshShell.SendKeys "connect.kpit.com"
WScript.Sleep 500 ' Wait for half a second

' Display message for Tabbing to the ENTER
WScript.Echo "Tabbing to the ENTER..."
WshShell.SendKeys "{TAB}"
WScript.Sleep 500 ' Wait for half a second

' Display message for pressing Enter to Portal setting
WScript.Echo "Pressing Enter to Portal setting..."
WshShell.SendKeys "{ENTER}"
WScript.Sleep 2000 ' Wait for half a second

' Display message for entering username
WScript.Echo "Entering username..."
WshShell.SendKeys "CHANDRAKANB"
WScript.Sleep 500 ' Wait for half a second

' Display message for Tabbing to the password field
WScript.Echo "Tabbing to the password field..."
WshShell.SendKeys "{TAB}"
WScript.Sleep 500 ' Wait for half a second

' Display message for entering password
WScript.Echo "Entering password..."
WshShell.SendKeys "IMDEV@69"
WScript.Sleep 500 ' Wait for half a second

' Display message for Tabbing to the ENTER
WScript.Echo "Tabbing to the ENTER..."
WshShell.SendKeys "{TAB}"
WScript.Sleep 500 ' Wait for half a second

' Display message for pressing Enter to submit credentials
WScript.Echo "Pressing Enter to submit credentials..."
WshShell.SendKeys "{ENTER}"
