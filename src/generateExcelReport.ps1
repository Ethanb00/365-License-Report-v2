param(
    [Parameter(Mandatory = $true)]
    [string]$GlobalWorkingPath,
    [string]$SubscribedCsvPath = $(Join-Path -Path $GlobalWorkingPath -ChildPath 'SubscribedSKUs.csv'),
    [string]$RenewalCsvPath = $(Join-Path -Path $GlobalWorkingPath -ChildPath 'LicenseRenewalData.csv'),
    [string]$AssignedCsvPath = $(Join-Path -Path $GlobalWorkingPath -ChildPath 'AssignedLicenses.csv')
)

function Safe-ImportCsv($path) {
    if (Test-Path -Path $path) { Import-Csv -Path $path } else { @() }
}

# Company Details
$parentPath = Split-Path -Path $GlobalWorkingPath -Parent
$companyName = Split-Path -Path $parentPath -Leaf
$generated = (Get-Date).ToString('MM-dd-yyyy HH:mm:ss')

# Load Data
$skus = Safe-ImportCsv -path $SubscribedCsvPath
$renewals = Safe-ImportCsv -path $RenewalCsvPath
$assigned = Safe-ImportCsv -path $AssignedCsvPath

# Load Pricing Data
$ClientRoot = Split-Path -Path $GlobalWorkingPath -Parent
$PricingCsvPath = Join-Path -Path $ClientRoot -ChildPath 'ClientPricing.csv'
$pricingTable = @{}
if (Test-Path $PricingCsvPath) {
    $pRows = Import-Csv -Path $PricingCsvPath
    foreach ($r in $pRows) {
        if ($r.SkuPartNumber -and $r.Cost) {
            $pricingTable[$r.SkuPartNumber] = [decimal]$r.Cost
        }
    }
}

$today = (Get-Date).Date
$excelPath = Join-Path -Path $GlobalWorkingPath -ChildPath 'LicenseReport.xlsx'

# --- 1. OVERVIEW DATA ---
$totalConsumed = 0
$totalPrepaid = 0
$estMonthlyCost = 0
foreach ($s in $skus) {
    $cost = if ($pricingTable.ContainsKey($s.SkuPartNumber)) { $pricingTable[$s.SkuPartNumber] } else { 0 }
    $units = try { [int]$s.EnabledPrepaidUnits } catch { 0 }
    $consumed = try { [int]$s.ConsumedUnits } catch { 0 }
    $totalConsumed += $consumed
    $totalPrepaid += $units
    $estMonthlyCost += ($units * $cost)
}

# Active Paid Users calculation
$paidUsersSet = @{}
foreach ($u in $assigned) {
    if ($u.AssignedSkus -and $u.AssignedSkus -ne 'None') {
        $parts = $u.AssignedSkus -split ' \| '
        foreach ($p in $parts) {
            if ($pricingTable.ContainsKey($p.Trim()) -and $pricingTable[$p.Trim()] -gt 0) {
                $paidUsersSet[$u.UPN] = $true
                break
            }
        }
    }
}

# Next Paid Renewal logic
$paidRenewals = @()
foreach ($rw in $renewals) {
    $skuPart = if ($rw.PSObject.Properties.Name -contains 'SkuPartNumber') { $rw.SkuPartNumber } elseif ($rw.PSObject.Properties.Name -contains 'SkuPart') { $rw.SkuPart } else { $rw.SkuPartNumber }
    $rawDate = if ($rw.PSObject.Properties.Name -contains 'RenewalDate') { $rw.RenewalDate } elseif ($rw.PSObject.Properties.Name -contains 'Renewal date') { $rw.'Renewal date' } else { $null }
    if ($skuPart -and $rawDate -and $pricingTable.ContainsKey($skuPart) -and $pricingTable[$skuPart] -gt 0) {
        try {
            $parsed = [datetime]::Parse($rawDate)
            if ($parsed -ge $today) {
                $paidRenewals += [PSCustomObject]@{
                    SkuPartNumber = $skuPart
                    RenewalDate   = $parsed
                }
            }
        }
        catch {}
    }
}
$nextRenewalDate = if ($paidRenewals.Count -gt 0) { ($paidRenewals | Sort-Object RenewalDate)[0].RenewalDate } else { $null }

$overviewData = @(
    [PSCustomObject]@{ Metric = 'Company Name'; Value = $companyName }
    [PSCustomObject]@{ Metric = 'Report Generated'; Value = $generated }
    [PSCustomObject]@{ Metric = ''; Value = '' }
    [PSCustomObject]@{ Metric = 'Est. Monthly Cost'; Value = ('{0:C2}' -f $estMonthlyCost) }
    [PSCustomObject]@{ Metric = 'Est. Annual Cost'; Value = ('{0:C2}' -f ($estMonthlyCost * 12)) }
    [PSCustomObject]@{ Metric = 'Active Paid Users'; Value = $paidUsersSet.Count }
    [PSCustomObject]@{ Metric = 'Next Paid License Renewal'; Value = if ($nextRenewalDate) { $nextRenewalDate.ToString('MM-dd-yyyy') } else { 'N/A' } }
    [PSCustomObject]@{ Metric = ''; Value = '' }
    [PSCustomObject]@{ Metric = 'Total Consumed Licenses'; Value = $totalConsumed }
    [PSCustomObject]@{ Metric = 'Total Prepaid Licenses'; Value = $totalPrepaid }
)

# --- 2. LICENSING DATA ---
$licensingData = foreach ($r in $skus) {
    $cost = if ($pricingTable.ContainsKey($r.SkuPartNumber)) { $pricingTable[$r.SkuPartNumber] } else { 0 }
    $units = try { [int]$r.EnabledPrepaidUnits } catch { 0 }
    [PSCustomObject]@{
        'License Name'    = $r.FriendlyName
        'MSFT SKU Name'   = $r.SkuPartNumber
        'Sku ID'          = $r.SkuId
        'Assigned'        = $r.ConsumedUnits
        'Available'       = $r.EnabledPrepaidUnits
        'Unit Cost (Mo)'  = $cost
        'Unit Cost (Yr)'  = ($cost * 12)
        'Total Cost (Mo)' = ($units * $cost)
        'Total Cost (Yr)' = ($units * $cost * 12)
    }
}

# --- 3. ASSIGNMENTS DATA ---
$assignmentsData = foreach ($a in $assigned) {
    if ($a.AssignedFriendlyNames -and $a.AssignedFriendlyNames -ne 'None') {
        $userCost = 0
        if ($a.AssignedSkus -and $a.AssignedSkus -ne 'None') {
            $parts = $a.AssignedSkus -split ' \| '
            foreach ($p in $parts) {
                if ($pricingTable.ContainsKey($p.Trim())) { $userCost += $pricingTable[$p.Trim()] }
            }
        }

        [PSCustomObject]@{
            'Display Name'    = $a.DisplayName
            'Email'           = $a.UPN
            'Account Status'  = $a.AccountStatus
            'Account Created' = $a.AccountCreatedTime
            'Last Sign-in'    = $a.LastSignInDate
            'Monthly Cost'    = $userCost
            'Licenses'        = $a.AssignedFriendlyNames
            'Assigned Via'    = $a.AssignmentDetails
        }
    }
}

# --- 4. UNLICENSED USERS ---
$unlicensedData = foreach ($a in $assigned) {
    if (-not $a.AssignedFriendlyNames -or $a.AssignedFriendlyNames -eq 'None') {
        [PSCustomObject]@{
            'Display Name'    = $a.DisplayName
            'Email'           = $a.UPN
            'Account Status'  = $a.AccountStatus
            'Account Created' = $a.AccountCreatedTime
            'Last Sign-in'    = $a.LastSignInDate
        }
    }
}

# --- 5. RENEWALS DATA ---
$renewalsData = foreach ($rw in $renewals) {
    # Re-evaluate friendly name for renewals
    $skuPart = if ($rw.PSObject.Properties.Name -contains 'SkuPartNumber') { $rw.SkuPartNumber } elseif ($rw.PSObject.Properties.Name -contains 'SkuPart') { $rw.SkuPart } else { $rw.SkuPartNumber }
    $friendly = $null
    if ($skus) {
        $match = $skus | Where-Object { $_.SkuPartNumber -eq $skuPart } | Select-Object -First 1
        if ($match) { $friendly = $match.FriendlyName }
    }
    if (-not $friendly) { $friendly = $skuPart }

    [PSCustomObject]@{
        'License Name'             = $friendly
        'MSFT SKU Name'            = $skuPart
        'Commerce Subscription ID' = if ($rw.PSObject.Properties.Name -contains 'CommerceSubscriptionId') { $rw.CommerceSubscriptionId } else { '' }
        'Subscription ID'          = if ($rw.PSObject.Properties.Name -contains 'SubscriptionId') { $rw.SubscriptionId } else { '' }
        'Trial/Paid'               = if ($rw.PSObject.Properties.Name -contains 'IsTrial') { $rw.IsTrial } else { '' }
        'Status'                   = if ($rw.PSObject.Properties.Name -contains 'SubscriptionStatus') { $rw.SubscriptionStatus } else { '' }
        'Total Licenses'           = if ($rw.PSObject.Properties.Name -contains 'TotalLicenses') { $rw.TotalLicenses } else { '' }
        'Renewal Date'             = if ($rw.PSObject.Properties.Name -contains 'RenewalDate') { $rw.RenewalDate } elseif ($rw.PSObject.Properties.Name -contains 'Renewal date') { $rw.'Renewal date' } else { '' }
        'Cost (Mo)'                = if ($pricingTable.ContainsKey($skuPart)) { $pricingTable[$skuPart] } else { 0 }
        'Notes'                    = if ($rw.PSObject.Properties.Name -contains 'Notes') { $rw.Notes } else { '' }
    }
}

# --- GENERATE EXCEL WORKBOOK ---
if (Test-Path $excelPath) { Remove-Item $excelPath }

$overviewData    | Export-Excel -Path $excelPath -WorksheetName 'Overview' -AutoSize -BoldTopRow
$licensingData   | Export-Excel -Path $excelPath -WorksheetName 'Licensing' -AutoSize -BoldTopRow -AutoFilter -TableStyle Light9
$assignmentsData | Export-Excel -Path $excelPath -WorksheetName 'Assignments' -AutoSize -BoldTopRow -AutoFilter -TableStyle Light9
$unlicensedData  | Export-Excel -Path $excelPath -WorksheetName 'Unlicensed Users' -AutoSize -BoldTopRow -AutoFilter -TableStyle Light9
$renewalsData    | Export-Excel -Path $excelPath -WorksheetName 'All Renewals' -AutoSize -BoldTopRow -AutoFilter -TableStyle Light9

# Additional formatting for currency columns
$pkg = Open-ExcelPackage -Path $excelPath
$licensingSheet = $pkg.Workbook.Worksheets['Licensing']
# Columns F, G, H, I (6, 7, 8, 9) are costs
Set-ExcelColumn -Worksheet $licensingSheet -Column 6 -NumberFormat 'Currency'
Set-ExcelColumn -Worksheet $licensingSheet -Column 7 -NumberFormat 'Currency'
Set-ExcelColumn -Worksheet $licensingSheet -Column 8 -NumberFormat 'Currency'
Set-ExcelColumn -Worksheet $licensingSheet -Column 9 -NumberFormat 'Currency'

$assignmentsSheet = $pkg.Workbook.Worksheets['Assignments']
# Column F (6) is monthly cost
Set-ExcelColumn -Worksheet $assignmentsSheet -Column 6 -NumberFormat 'Currency'

$renewalsSheet = $pkg.Workbook.Worksheets['All Renewals']
# Column D (4) is cost
Set-ExcelColumn -Worksheet $renewalsSheet -Column 4 -NumberFormat 'Currency'

Close-ExcelPackage $pkg

Write-Host "Wrote Excel report to: $excelPath" -ForegroundColor Cyan
