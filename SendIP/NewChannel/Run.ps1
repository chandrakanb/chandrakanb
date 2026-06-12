$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
Start-Process -FilePath "$scriptDir\run.bat" -WindowStyle Hidden
