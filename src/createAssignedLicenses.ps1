param(
    [Parameter(Mandatory = $true)]
    [string]$GlobalWorkingPath
)

Write-Host "--- Creating AssignedLicenses.csv (normalized user->sku assignments) ---" -ForegroundColor Cyan

$outFile = Join-Path -Path $GlobalWorkingPath -ChildPath 'AssignedLicenses.csv'
$summaryFile = Join-Path -Path $GlobalWorkingPath -ChildPath 'AssignedLicenses_Summary.csv'

function Safe-ImportCsv($path) { if (Test-Path -Path $path) { Import-Csv -Path $path } else { @() } }

# Build helper maps: SkuId -> SkuPartNumber and friendly names from local skus.csv if present
$skuIdToPart = @{}
$skuMapByGuid = @{}
$skuMapByStringId = @{}
try {
    $subscribed = Get-MgSubscribedSku -ErrorAction SilentlyContinue
    foreach ($s in $subscribed) {
        if ($s.SkuId -and $s.SkuPartNumber) { $skuIdToPart[$s.SkuId] = $s.SkuPartNumber }
    }
}
catch { }

$skuCsvPath = Join-Path -Path (Split-Path -Parent $PSScriptRoot) -ChildPath 'data\skus.csv'
if (Test-Path -Path $skuCsvPath) {
    try {
        $skuRows = Import-Csv -Path $skuCsvPath -ErrorAction SilentlyContinue
        foreach ($r in $skuRows) {
            if ($r.GUID -and -not $skuMapByGuid.ContainsKey($r.GUID)) { $skuMapByGuid[$r.GUID] = $r.Product_Display_Name }
            if ($r.String_Id -and -not $skuMapByStringId.ContainsKey($r.String_Id)) { $skuMapByStringId[$r.String_Id] = $r.Product_Display_Name }
        }
    }
    catch { }
}

# Load custom friendly names from data\CustomSkuNames.csv
$customSkuCsvPath = Join-Path -Path (Split-Path -Parent $PSScriptRoot) -ChildPath 'data\CustomSkuNames.csv'
if (Test-Path -Path $customSkuCsvPath) {
    try {
        $customRows = Import-Csv -Path $customSkuCsvPath -ErrorAction SilentlyContinue
        foreach ($cr in $customRows) {
            # Merge into GUID and StringId maps
            if ($cr.Id -and $cr.FriendlyName -and -not $skuMapByGuid.ContainsKey($cr.Id) -and -not $skuMapByStringId.ContainsKey($cr.Id)) {
                $skuMapByGuid[$cr.Id] = $cr.FriendlyName
            }
        }
    }
    catch { }
}

Write-Host "Querying users (Id/UPN/DisplayName, CreatedDateTime, AccountEnabled, SignInActivity)..." -ForegroundColor Yellow
try {
    $users = Get-MgUser -All -Property Id, UserPrincipalName, DisplayName, CreatedDateTime, AccountEnabled, SignInActivity, licenseAssignmentStates
}
catch {
    Write-Host "Failed to query users via Graph: $_" -ForegroundColor Yellow
    $users = @()
}

# We'll produce one consolidated row per user. Also build a flat list for the per-SKU summary.
$outRows = @()
$flatAssignments = @()
foreach ($u in $users) {
    $uid = $u.Id
    $upn = $u.UserPrincipalName
    $display = $u.DisplayName
    $acctCreated = $u.CreatedDateTime
    $acctEnabled = if ($u.AccountEnabled -eq $true) { 'Enabled' } else { 'Disabled' }
    $lastSignIn = 'N/A'
    try { if ($u.SignInActivity -and $u.SignInActivity.LastSignInDateTime) { $lastSignIn = $u.SignInActivity.LastSignInDateTime } } catch {}

    try {
        $ld = Get-MgUserLicenseDetail -UserId $uid -ErrorAction SilentlyContinue
    }
    catch {
        $ld = @()
    }

    $states = $u.licenseAssignmentStates

    $parts = @()
    $friendlyNames = @()
    $detailStrings = @()
    if ($ld -and $ld.Count -gt 0) {
        foreach ($entry in $ld) {
            $sid = $entry.SkuId
            $part = ''
            if ($skuIdToPart.ContainsKey($sid)) { $part = $skuIdToPart[$sid] }

            $friendly = ''
            if ($skuMapByGuid.ContainsKey($sid)) { $friendly = $skuMapByGuid[$sid] }
            elseif ($part -and $skuMapByStringId.ContainsKey($part)) { $friendly = $skuMapByStringId[$part] }
            elseif ($part) { $friendly = $part }
            else { $friendly = $sid }

            $state = $states | Where-Object { $_.SkuId -eq $sid }
            $assignedByGroup = $false
            if ($state) { $assignedByGroup = $state.AssignedByGroup }
            $applies = if ($assignedByGroup) { 'Group' } else { 'Direct' }

            $parts += $part
            $friendlyNames += $friendly
            $detailStrings += ("$applies")

            # add to flat assignments for summary
            $flatAssignments += [PSCustomObject]@{ SkuPartNumber = $part; FriendlyName = $friendly; UPN = $upn }
        }
    }

    # Use a clear separator ' | ' between multiple values
    $partsJoined = if ($parts.Count -gt 0) { ($parts | Where-Object { $_ -and $_ -ne '' } | Select-Object -Unique) -join ' | ' } else { 'None' }
    $friendlyJoined = if ($friendlyNames.Count -gt 0) { ($friendlyNames | Where-Object { $_ -and $_ -ne '' } | Select-Object -Unique) -join ' | ' } else { 'None' }
    $detailsJoined = if ($detailStrings.Count -gt 0) { ($detailStrings | Select-Object -Unique) -join ' | ' } else { 'None' }

    $outRows += [PSCustomObject]@{
        UPN                   = $upn
        DisplayName           = $display
        UserId                = $uid
        AccountStatus         = $acctEnabled
        AccountCreatedTime    = $acctCreated
        LastSignInDate        = $lastSignIn
        AssignedSkus          = $partsJoined
        AssignedFriendlyNames = $friendlyJoined
        AssignmentDetails     = $detailsJoined
    }
}

# Diagnostics and write normalized assigned licenses CSV
Write-Host "GlobalWorkingPath: $GlobalWorkingPath" -ForegroundColor Cyan
if (-not (Test-Path -Path $GlobalWorkingPath)) {
    Write-Host "Output directory does not exist; creating: $GlobalWorkingPath" -ForegroundColor Yellow
    try { New-Item -ItemType Directory -Path $GlobalWorkingPath -Force | Out-Null } catch { Write-Host "Failed to create output directory: $_" -ForegroundColor Red }
}

Write-Host "User count retrieved: $($users.Count)" -ForegroundColor Cyan
Write-Host "Assigned rows to write (users): $($outRows.Count)" -ForegroundColor Cyan

try {
    $parentDir = Split-Path -Path $outFile -Parent
    if (-not (Test-Path -Path $parentDir)) { New-Item -ItemType Directory -Path $parentDir -Force | Out-Null }
    if ($outRows.Count -gt 0) {
        $outRows | Export-Csv -Path $outFile -NoTypeInformation -Encoding UTF8 -Force
    }
    else {
        # write header-only empty file using consolidated column names
        $empty = [PSCustomObject]@{
            UPN = ''; DisplayName = ''; UserId = ''; AccountStatus = ''; AccountCreatedTime = ''; LastSignInDate = ''; AssignedSkus = ''; AssignedFriendlyNames = ''; AssignmentDetails = ''
        }
        $empty | Export-Csv -Path $outFile -NoTypeInformation -Encoding UTF8 -Force
    }
    Write-Host "Wrote assigned licenses to: $outFile" -ForegroundColor Green
}
catch {
    Write-Host "Failed to write AssignedLicenses.csv: $_" -ForegroundColor Red
}

# Create a per-sku summary file (counts) from the flat assignment list
try {
    if ($flatAssignments -and $flatAssignments.Count -gt 0) {
        $summary = $flatAssignments | Group-Object -Property @{Expression = { if ($_.SkuPartNumber) { $_.SkuPartNumber } else { 'Unknown' } } } | ForEach-Object {
            $key = $_.Name
            $count = $_.Count
            $fn = ($_.Group | Select-Object -First 1).FriendlyName
            [PSCustomObject]@{ SkuKey = $key; FriendlyName = $fn; AssignedCount = $count }
        }
        if ($summary) { $summary | Export-Csv -Path $summaryFile -NoTypeInformation -Encoding UTF8 -Force }
        Write-Host "Wrote assigned licenses summary to: $summaryFile" -ForegroundColor Green
    }
    else {
        # Ensure file exists with header
        $emptySummary = [PSCustomObject]@{ SkuKey = ''; FriendlyName = ''; AssignedCount = 0 }
        $emptySummary | Export-Csv -Path $summaryFile -NoTypeInformation -Encoding UTF8 -Force
        Write-Host "Wrote empty assigned licenses summary to: $summaryFile" -ForegroundColor Yellow
    }
}
catch {
    Write-Host "Failed to write AssignedLicenses_Summary.csv: $_" -ForegroundColor Yellow
}

Write-Host "Assigned license collection complete." -ForegroundColor Cyan
