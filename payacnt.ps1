Add-Type -AssemblyName System.Windows.Forms
Start-Sleep -m 500

notepad.exe
Start-Sleep -m 1000

[System.Windows.Forms.SendKeys]::SendWait("test{ENTER}")