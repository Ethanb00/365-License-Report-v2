# =========================================================
# collectUserDetails.ps1
# =========================================================

param(
    [Parameter(Mandatory = $true)]
    [string]$GlobalWorkingPath
)

Write-Host "--- Collecting User Profile Details from Graph ---" -ForegroundColor Cyan
Write-Host "Output directory: $GlobalWorkingPath" -ForegroundColor Yellow

# Define the output file name
$UserDetailFileName = "user_details.csv"
$UserDetailFilePath = Join-Path -Path $GlobalWorkingPath -ChildPath $UserDetailFileName

# --- 1. Define the required properties ---
# These are the properties needed for your final report
$PropertiesToRetrieve = @(
    'UserPrincipalName',
    'DisplayName',
    'Department',
    'JobTitle',
    'CreatedDateTime',
    'AccountEnabled',
    'SignInActivity' # Required to get LastSignInDateTime
)

$AllUsers = Get-MgUser -All -Property $PropertiesToRetrieve


# --- 2. Collect Data (Requires Microsoft Graph PowerShell Module) ---
try {
    Write-Host "Querying all tenant users..."
    # Replace Get-MgUser with your actual Microsoft Graph cmdlet if needed (e.g., Get-AzureADUser)
    # --- Build SKU mapping (from subscribed SKUs and local data/skus.csv) ---
    $skuIdToPart = @{}
    $skuMapByStringId = @{}
    $skuMapByGuid = @{}
    try {
        $subscribed = Get-MgSubscribedSku -ErrorAction SilentlyContinue
        foreach ($s in $subscribed) {
            if ($s.SkuId -and $s.SkuPartNumber) { $skuIdToPart[$s.SkuId] = $s.SkuPartNumber }
        }
    }
    catch {
        # ignore if cmdlet not available or permission denied
    }

    $skuCsvPath = Join-Path -Path (Split-Path -Parent $PSScriptRoot) -ChildPath 'data\skus.csv'
    if (Test-Path -Path $skuCsvPath) {
        try {
            $skuRows = Import-Csv -Path $skuCsvPath -ErrorAction SilentlyContinue
            foreach ($r in $skuRows) {
                if ($r.String_Id -and -not $skuMapByStringId.ContainsKey($r.String_Id)) { $skuMapByStringId[$r.String_Id] = $r.Product_Display_Name }
                if ($r.GUID -and -not $skuMapByGuid.ContainsKey($r.GUID)) { $skuMapByGuid[$r.GUID] = $r.Product_Display_Name }
            }
        }
        catch {}
    }

    # Merge custom SKU names
    $customSkuCsvPath = Join-Path -Path (Split-Path -Parent $PSScriptRoot) -ChildPath 'data\CustomSkuNames.csv'
    if (Test-Path -Path $customSkuCsvPath) {
        try {
            $customRows = Import-Csv -Path $customSkuCsvPath -ErrorAction SilentlyContinue
            foreach ($cr in $customRows) {
                if ($cr.Id -and $cr.FriendlyName -and -not $skuMapByGuid.ContainsKey($cr.Id) -and -not $skuMapByStringId.ContainsKey($cr.Id)) {
                    $skuMapByGuid[$cr.Id] = $cr.FriendlyName
                }
            }
        }
        catch {}
    }

    $ExportData = foreach ($u in $AllUsers) {
        # Fetch license details ONCE
        $ld = try { Get-MgUserLicenseDetail -UserId $u.Id -ErrorAction SilentlyContinue } catch { @() }
        
        # 1. AssignedSkus
        $parts = $ld | ForEach-Object {
            $sid = $_.SkuId
            if ($skuIdToPart.ContainsKey($sid)) { $skuIdToPart[$sid] } else { $sid }
        } | Where-Object { $_ } | Select-Object -Unique
        $assignedSkus = if ($parts -and $parts.Count -gt 0) { $parts -join ';' } else { 'None' }

        # 2. AssignedFriendlyNames
        $names = $ld | ForEach-Object {
            $sid = $_.SkuId
            if ($skuMapByGuid.ContainsKey($sid)) { $skuMapByGuid[$sid] }
            elseif ($skuIdToPart.ContainsKey($sid) -and $skuMapByStringId.ContainsKey($skuIdToPart[$sid])) { $skuMapByStringId[$skuIdToPart[$sid]] }
            elseif ($skuIdToPart.ContainsKey($sid)) { $skuIdToPart[$sid] }
            else { $sid }
        } | Where-Object { $_ } | Select-Object -Unique
        $assignedFriendlyNames = if ($names -and $names.Count -gt 0) { $names -join ';' } else { 'None' }

        # 3. AssignmentDetails
        $detailStrings = $ld | ForEach-Object {
            $sid = $_.SkuId
            $part = if ($skuIdToPart.ContainsKey($sid)) { $skuIdToPart[$sid] } else { $sid }
            # AppliesTo logic
            $applies = if ($_.AppliesTo) { $_.AppliesTo } elseif ($_.ServicePlans -and ($_.ServicePlans | Where-Object { $_.AppliesTo }) ) { ($_.ServicePlans | Where-Object { $_.AppliesTo })[0].AppliesTo } else { 'User' }
            "$part ($applies)"
        } | Where-Object { $_ } | Select-Object -Unique
        $assignmentDetails = if ($detailStrings -and $detailStrings.Count -gt 0) { $detailStrings -join ';' } else { 'None' }

        # 4. Last Sign In
        $lastSignIn = if ($u.SignInActivity -and $u.SignInActivity.LastSignInDateTime) { $u.SignInActivity.LastSignInDateTime } else { "N/A" }

        # 5. Account Status
        $accountStatus = if ($u.AccountEnabled -eq $true) { "Enabled" } else { "Disabled" }

        [PSCustomObject]@{
            'UPN'                   = $u.UserPrincipalName
            'Display Name'          = $u.DisplayName
            'Department'            = $u.Department
            'Title'                 = $u.JobTitle
            'Account Created Time'  = $u.CreatedDateTime
            'Last Sign-in Date'     = $lastSignIn
            'Account Status'        = $accountStatus
            'AssignedSkus'          = $assignedSkus
            'AssignedFriendlyNames' = $assignedFriendlyNames
            'AssignmentDetails'     = $assignmentDetails
        }
    }
    $ExportData | Export-Csv -Path $UserDetailFilePath -NoTypeInformation -Encoding UTF8
    Write-Host "`n✅ Successfully saved user profile details to: $UserDetailFileName" -ForegroundColor Green

}
catch {
    Write-Host "`n❌ Error querying Microsoft Graph for user details: $($_.Exception.Message)" -ForegroundColor Red
    # Return $false or throw error if data is critical
}