# 1. Connect to Exchange Online (if not already connected)
Connect-ExchangeOnline

# --- CONFIGURATION VARIABLES ---
# The email of the source calendar (The one being shared)
$CalendarOwner = "source@example.com" 

# The email of the user having connection issues
$UserWithIssue = "user@example.com"

# The permission level required (Usually 'Editor' for DocketCalendar integrations, or 'Reviewer' for read-only)
$AccessLevel = "Editor" 
# -------------------------------

Write-Host "Resetting permissions for $UserWithIssue on $CalendarOwner's calendar..." -ForegroundColor Cyan

# 2. Remove the existing permission (Un-share)
Try {
    Remove-MailboxFolderPermission -Identity "$($CalendarOwner):\Calendar" -User $UserWithIssue -ErrorAction Stop -Confirm:$false
    Write-Host "Successfully removed existing permissions." -ForegroundColor Green
}
Catch {
    Write-Host "No existing permissions found, or error removing them. Proceeding to add..." -ForegroundColor Yellow
}

# 3. Add the permission back (Re-share)
Try {
    Add-MailboxFolderPermission -Identity "$($CalendarOwner):\Calendar" -User $UserWithIssue -AccessRights $AccessLevel -ErrorAction Stop
    Write-Host "Successfully re-added permissions with level: $AccessLevel" -ForegroundColor Green
}
Catch {
    Write-Host "Error adding permissions: $($_.Exception.Message)" -ForegroundColor Red
}

# 4. Verify the new result
Write-Host "Current Permissions:" -ForegroundColor Cyan
Get-MailboxFolderPermission -Identity "$($CalendarOwner):\Calendar" -User $UserWithIssue