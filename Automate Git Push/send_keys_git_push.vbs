Set WshShell = WScript.CreateObject("WScript.Shell")

' Activate Command Prompt window
WshShell.AppActivate "Command Prompt"
WScript.Sleep 5000 ' Wait for a second

' Enter the "git push" command
WshShell.SendKeys "git push origin reppy"
WScript.Sleep 500 ' Wait for half a second
WshShell.SendKeys "{ENTER}"

' Wait for 2 seconds before entering credentials
WScript.Sleep 5000

' Enter username
WshShell.SendKeys "chandrakanb"
WScript.Sleep 500
WshShell.SendKeys "{ENTER}"

' Wait a little before entering the token
WScript.Sleep 5000

' Enter the GitHub token
WshShell.SendKeys "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"
WScript.Sleep 500
WshShell.SendKeys "{ENTER}"
