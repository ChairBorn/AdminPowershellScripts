$AllUsers = Get-MgUser -All -Property Id, UserPrincipalName, DisplayName
$Results = @()

ForEach ($User in $AllUsers) {
    # Retrieve all methods for this specific user
    $Methods = Get-MgUserAuthenticationMethod -UserId $User.Id
    
    # Create a custom object for the report
    $UserObject = [PSCustomObject]@{
        User    = $User.UserPrincipalName
        Name    = $User.DisplayName
        # Join all method types into a single comma-separated string
        Methods = ($Methods.AdditionalProperties["@odata.type"] -replace "#microsoft.graph.","" -join ", ")
    }
    $Results += $UserObject
}

# Export to CSV
$Results | Export-Csv -Path "C:\Temp\Detailed_MFA_Report.csv" -NoTypeInformation
Write-Host "Detailed report exported."