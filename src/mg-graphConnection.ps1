<#
.SYNOPSIS
    Microsoft Graph API authentication and connection management

.DESCRIPTION
    Handles authentication to Microsoft Graph API with required permissions.
    - Checks for existing connection and offers to reuse or reconnect
    - Requests necessary delegated permissions for license reporting
    - Supports interactive authentication flow
    
.NOTES
    Required Scopes:
    - User.Read.All: Read all user profiles and license assignments
    - Group.Read.All: Read group information for group-based licenses  
    - Directory.Read.All: Read directory objects
    - Reports.Read.All: Access usage reports
    - AuditLog.Read.All: Read sign-in activity (requires Entra ID P1/P2)
    - Organization.Read.All: Read organization details and domains
#>

# ============================================================================
# CHECK FOR EXISTING CONNECTION
# ============================================================================
# Verify if user is already authenticated to Microsoft Graph

$mgContext = Get-MgContext

# ============================================================================
# AUTHENTICATION LOGIC
# ============================================================================

# If no existing connection, prompt for connection
if ($null -eq $mgContext ) {
    Write-Host "`nMicrosoft Graph connection required." -ForegroundColor Cyan
    Write-Host "You will be prompted to sign in with an account that has:" -ForegroundColor Yellow
    Write-Host "  - Global Admin or Reports Reader role" -ForegroundColor Yellow
    Write-Host "  - Access to the target Microsoft 365 tenant`n" -ForegroundColor Yellow
    
    # Connect with all required scopes
    Connect-MgGraph -NoWelcome -ContextScope Process -Scopes "User.Read.All", "Group.Read.All", "Directory.Read.All", "Reports.Read.All", "AuditLog.Read.All", "Organization.Read.All"
    
    Write-Host "`nSuccessfully connected to Microsoft Graph" -ForegroundColor Green
    Write-Host "Logged in as: $((Get-MgContext).Account)" -ForegroundColor Cyan
}
# If existing connection, confirm to stay connected or reconnect
else {
    Write-Host "`nExisting Microsoft Graph connection found." -ForegroundColor Green
    Write-Host "Logged in as: $($mgContext.Account)" -ForegroundColor Cyan
    
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
            Write-Host "`nDisconnecting and reconnecting to Microsoft Graph..." -ForegroundColor Yellow
            Disconnect-MgGraph | Out-Null
            Connect-MgGraph -NoWelcome -ContextScope Process -Scopes "User.Read.All", "Group.Read.All", "Directory.Read.All", "Reports.Read.All", "AuditLog.Read.All", "Organization.Read.All"
            Write-Host "`nSuccessfully reconnected as: $((Get-MgContext).Account)" -ForegroundColor Green
            break
        }
        else {
            # Invalid input, prompt again
            $stayConnected = Read-Host "`nInvalid entry. Please enter Y or N:"
        }
    } while ($true)
}