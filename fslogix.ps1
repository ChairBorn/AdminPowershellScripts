write-host "Configuring FSLogix"
$fileServer="env:FILESERVER"
$profileShare="\\$($fileServer)\fslogix-profiles"

$user="localhost\env:USER"
$secret={env:KEY}
# Registry Settings for FSLogix Profile Container
New-Item -Path "HKLM:\SOFTWARE" -Name "FSLogix" -ErrorAction Ignore
New-Item -Path "HKLM:\SOFTWARE\FSLogix" -Name "Profiles" -ErrorAction Ignore
New-ItemProperty -Path "HKLM:\SOFTWARE\FSLogix\Profiles" -Name "Enabled" -Value 1 -force
New-ItemProperty -Path "HKLM:\SOFTWARE\FSLogix\Profiles" -Name "VHDLocations" -Value $profileShare -force
New-ItemProperty -Path "HKLM:\SOFTWARE\FSLogix\Profiles" -Name "ConcurrentUserSessions" -Value 1 -force
New-ItemProperty -Path "HKLM:\SOFTWARE\FSLogix\Profiles" -Name "DeleteLocalProfileWhenVHDShouldApply" -Value 1 -force
New-ItemProperty -Path "HKLM:\SOFTWARE\FSLogix\Profiles" -Name "FlipFlopProfileDirectoryName" -Value 1 -force
New-ItemProperty -Path "HKLM:\SOFTWARE\FSLogix\Profiles" -Name "IsDynamic" -Value 1 -force
New-ItemProperty -Path "HKLM:\SOFTWARE\FSLogix\Profiles" -Name "KeepLocalDir" -Value 0 -force
New-ItemProperty -Path "HKLM:\SOFTWARE\FSLogix\Profiles" -Name "ProfileType" -Value 0 -force
New-ItemProperty -Path "HKLM:\SOFTWARE\FSLogix\Profiles" -Name "SizeInMBs" -Value 40000 -force
New-ItemProperty -Path "HKLM:\SOFTWARE\FSLogix\Profiles" -Name "VolumeType" -Value "VHDX" -force
New-ItemProperty -Path "HKLM:\SOFTWARE\FSLogix\Profiles" -Name "AccessNetworkAsComputerObject" -Value 1 -force

# Include credentials in the profile
New-Item -Path "HKLM:\Software\Policies\Microsoft" -Name "AzureADAccount" -ErrorAction Ignore
New-ItemProperty -Path "HKLM:\Software\Policies\Microsoft\AzureADAccount" -Name "LoadCredKeyFromProfile" -Value 1 -force

# Store credentials to access the storage account
cmdkey.exe /add:$fileServer /user:$($user) /pass:$($secret)
# Disable Windows Defender Credential Guard (only needed for Windows 11 22H2)
New-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa" -Name "LsaCfgFlags" -Value 0 -force

write-host "The script has finished."
