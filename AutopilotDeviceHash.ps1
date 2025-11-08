Set-ExecutionPolicy Bypass -Scope Process -Force
Install-Script -Name Get-WindowsAutopilotInfo -Force
Get-WindowsAutopilotInfo -OutputFile C:\DeviceHash.csv
