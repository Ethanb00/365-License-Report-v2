# Check for existing Microsoft Graph connection
$mgContext = Get-MgContext

# If no existing connection, prompt for connection
if ($null -eq $mgContext ) {
    # Connect to Microsoft Graph
    Write-Host "`nMicrosoft Graph connection required."
    Connect-MgGraph -NoWelcome -Scopes "User.Read.All", "Group.Read.All", "Directory.Read.All", "Reports.Read.All", "AuditLog.Read.All", "Organization.Read.All"
    Write-Host "`nConnected to Microsoft Graph as: $((Get-MgContext).Account)" 
}
# If existing connection, confirm to stay connected or reconnect
else {
    Write-Host "`nExisting Microsoft Graph connection found. Logged in user: $($mgContext.Account)"
    # Prompt to stay connected or reconnect
    $stayConnected = Read-Host "`nDo you want to stay connected as this user? (Y/N)"
    do {
        $ans = $stayConnected.Trim().ToUpper()
        if ($ans -eq 'Y') {
            # Stay connected
            Write-Host "`nStaying connected as: $($mgContext.Account)" -ForegroundColor Green
            break
        }
        elseif ($ans -eq 'N') {
            # Disconnect and reconnect
            Write-Host "`nDisconnecting and reconnecting to Microsoft Graph."
            Disconnect-MgGraph | Out-Null
            Connect-MgGraph -NoWelcome -Scopes "User.Read.All", "Group.Read.All", "Directory.Read.All", "Reports.Read.All", "AuditLog.Read.All", "Organization.Read.All"
            Write-Host "`nConnected to Microsoft Graph as: $((Get-MgContext).Account)" -ForegroundColor Green
            break
        }
        else {
            # Invalid input, prompt again
            $stayConnected = Read-Host "`nInvalid entry. Please enter Y or N:"
        }
    } while ($true)
}