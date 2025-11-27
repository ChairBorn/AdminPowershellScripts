Write-Host "Configuring FSLogix..."

# Storage account + share info
$fileServer = "blackstonefslogixstorage.file.core.windows.net"
$profileShare = "\\$fileServer\fslogix-profiles"

# Correct user format for Azure Files SMB key auth
$user = "Azure\blackstonefslogixstorage"
# $user = "localhost\blackstonefslogixstorage"
# Storage Account Key
$secret = "$env:KEY"

Write-Host "Authenticating to Azure File Share..."

# Store credentials
cmdkey.exe /add:$fileServer /user:$user /pass:$secret | Out-Null

# Attempt to mount the share (non-persistent)
net use $profileShare /user:$user $secret /persistent:no | Out-Null

Write-Host "Checking FSLogix share accessibility..."
if (!(Test-Path $profileShare)) {
    Write-Host "ERROR: Cannot access FSLogix share: $profileShare"
    Write-Host "Check networking (port 445), firewall, private endpoints, and NTFS permissions."
    exit 1
}

Write-Host "Share is accessible. Applying FSLogix configuration..."

# FSLogix Registry Settings
$fslogixPath = "HKLM:\SOFTWARE\FSLogix\Profiles"
New-Item -Path "HKLM:\SOFTWARE\FSLogix" -ErrorAction Ignore | Out-Null
New-Item -Path $fslogixPath -ErrorAction Ignore | Out-Null

Set-ItemProperty -Path $fslogixPath -Name Enabled -Type DWord -Value 1
Set-ItemProperty -Path $fslogixPath -Name VHDLocations -Type MultiString -Value $profileShare
Set-ItemProperty -Path $fslogixPath -Name ConcurrentUserSessions -Type DWord -Value 1
Set-ItemProperty -Path $fslogixPath -Name DeleteLocalProfileWhenVHDShouldApply -Type DWord -Value 1
Set-ItemProperty -Path $fslogixPath -Name FlipFlopProfileDirectoryName -Type DWord -Value 1
Set-ItemProperty -Path $fslogixPath -Name IsDynamic -Type DWord -Value 1
Set-ItemProperty -Path $fslogixPath -Name KeepLocalDir -Type DWord -Value 0
Set-ItemProperty -Path $fslogixPath -Name ProfileType -Type DWord -Value 0
Set-ItemProperty -Path $fslogixPath -Name SizeInMBs -Type DWord -Value 40000
Set-ItemProperty -Path $fslogixPath -Name VolumeType -Value "VHDX"
Set-ItemProperty -Path $fslogixPath -Name AccessNetworkAsComputerObject -Type DWord -Value 1

# Azure AD cached credential load (SSO improvement)
New-Item -Path "HKLM:\Software\Policies\Microsoft\AzureADAccount" -ErrorAction Ignore | Out-Null
Set-ItemProperty -Path "HKLM:\Software\Policies\Microsoft\AzureADAccount" -Name LoadCredKeyFromProfile -Type DWord -Value 1

# Disable Credential Guard (needed for AVD + FSLogix sometimes)
Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa" -Name LsaCfgFlags -Type DWord -Value 0

# Set time zone + report status
tzutil /s "Pacific Standard Time"
if ($LASTEXITCODE -eq 0) {
    Write-Host "Time zone successfully set to Pacific Standard Time."
} else {
    Write-Host "WARNING: Time zone may not have been set correctly. Exit code: $LASTEXITCODE"
}

Write-Host "FSLogix has been successfully configured."
