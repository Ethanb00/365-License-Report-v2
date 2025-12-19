param(
    [Parameter(Mandatory = $true)]
    [string]$GlobalWorkingPath
)

# Fetch license SKUs
[array]$Skus = Get-MgSubscribedSku
$SkuReport = [System.Collections.Generic.List[Object]]::new()

foreach ($Sku in $Skus) {
    $DataLine = [PSCustomObject][Ordered]@{
        SkuPartNumber = $Sku.SkuPartNumber
        SkuId         = $Sku.SkuId
        ActiveUnits   = $Sku.PrepaidUnits.Enabled
        WarningUnits  = $Sku.PrepaidUnits.Warning
        ConsumedUnits = $Sku.ConsumedUnits
    }
    $SkuReport.Add($DataLine)
}

# Fetch renewal data
$Uri = "https://graph.microsoft.com/v1.0/directory/subscriptions"
try {
    [array]$SkuData = Invoke-MgGraphRequest -Uri $Uri -Method Get
}
catch {
    Write-Error "Failed to retrieve subscription data from Graph: $_"
    return
}

# Build renewal hash table with duplicate handling
$SkuHash = @{}
$DuplicateSkus = @()

foreach ($Sku in $SkuData.Value) {
    if ($SkuHash.ContainsKey($Sku.SkuId)) {
        $DuplicateSkus += $Sku.SkuId
    }
    else {
        # Extract additional details using actual Graph API property names
        $isTrial = if ($Sku.isTrial) { "Trial" } else { "Paid" }
        
        $SkuHash[$Sku.SkuId] = [PSCustomObject]@{
            RenewalDate            = $Sku.nextLifecycleDateTime
            CommerceSubscriptionId = $Sku.commerceSubscriptionId
            SubscriptionId         = $Sku.id
            IsTrial                = $isTrial
            Status                 = $Sku.status
            TotalLicenses          = $Sku.totalLicenses
        }
    }
}


if ($DuplicateSkus.Count -gt 0) {
    Write-Warning "Duplicate SkuIds detected in renewal data: $(( $DuplicateSkus | Sort-Object | Get-Unique ) -join ', ')"
}

# Enrich report with renewal info
foreach ($R in $SkuReport) {
    if ($SkuHash.ContainsKey($R.SkuId)) {
        $details = $SkuHash[$R.SkuId]
        $SkuRenewalDate = $details.RenewalDate
        $R | Add-Member -NotePropertyName "Renewal date" -NotePropertyValue $SkuRenewalDate -Force
        $R | Add-Member -NotePropertyName "CommerceSubscriptionId" -NotePropertyValue $details.CommerceSubscriptionId -Force
        $R | Add-Member -NotePropertyName "SubscriptionId" -NotePropertyValue $details.SubscriptionId -Force
        $R | Add-Member -NotePropertyName "IsTrial" -NotePropertyValue $details.IsTrial -Force
        $R | Add-Member -NotePropertyName "SubscriptionStatus" -NotePropertyValue $details.Status -Force
        $R | Add-Member -NotePropertyName "TotalLicenses" -NotePropertyValue $details.TotalLicenses -Force

        if ($SkuRenewalDate) {
            $DaysToRenew = ($SkuRenewalDate - (Get-Date)).Days
            $R | Add-Member -NotePropertyName "Days to renewal" -NotePropertyValue $DaysToRenew -Force

            $Status = if ($DaysToRenew -lt 0) { "Expired" } elseif ($DaysToRenew -le 30) { "Expiring Soon" } else { "Active" }
            $R | Add-Member -NotePropertyName "Status" -NotePropertyValue $Status -Force
        }
        else {
            $R | Add-Member -NotePropertyName "Days to renewal" -NotePropertyValue "Unknown" -Force
            $R | Add-Member -NotePropertyName "Status" -NotePropertyValue "Unknown" -Force
        }
    }
    else {
        Write-Warning "No renewal metadata found for SKU: $($R.SkuPartNumber) [$($R.SkuId)]"
        $R | Add-Member -NotePropertyName "Renewal date" -NotePropertyValue "Unknown" -Force
        $R | Add-Member -NotePropertyName "CommerceSubscriptionId" -NotePropertyValue "Unknown" -Force
        $R | Add-Member -NotePropertyName "SubscriptionId" -NotePropertyValue "Unknown" -Force
        $R | Add-Member -NotePropertyName "IsTrial" -NotePropertyValue "Unknown" -Force
        $R | Add-Member -NotePropertyName "SubscriptionStatus" -NotePropertyValue "Unknown" -Force
        $R | Add-Member -NotePropertyName "TotalLicenses" -NotePropertyValue "Unknown" -Force
        $R | Add-Member -NotePropertyName "Days to renewal" -NotePropertyValue "Unknown" -Force
        $R | Add-Member -NotePropertyName "Status" -NotePropertyValue "Unknown" -Force
    }
}

# Load pricing reference
$PricingPath = Join-Path $GlobalWorkingPath 'pricing.csv'
if (-not (Test-Path $PricingPath)) {
    Write-Warning "Pricing reference file not found: $PricingPath"
}
else {
    $PricingData = Import-Csv -Path $PricingPath
    $PricingLookup = @{}
    foreach ($Item in $PricingData) {
        if (-not $PricingLookup.ContainsKey($Item.SkuPartNumber)) {
            $PricingLookup[$Item.SkuPartNumber] = $Item.DisplayName
        }
    }

    # Enrich SkuReport with DisplayName
    $UnmatchedSkus = @()
    foreach ($R in $SkuReport) {
        $DisplayName = $PricingLookup[$R.SkuPartNumber]
        if ($DisplayName) {
            $R | Add-Member -NotePropertyName "DisplayName" -NotePropertyValue $DisplayName -Force
        }
        else {
            $UnmatchedSkus += $R.SkuPartNumber
            $R | Add-Member -NotePropertyName "DisplayName" -NotePropertyValue "Unknown" -Force
        }
    }

    if ($UnmatchedSkus.Count -gt 0) {
        Write-Warning "The following SKUs were not found in pricing.csv: $($UnmatchedSkus | Sort-Object | Get-Unique) -join ', '
"
    }
}

# Reorder properties to place DisplayName first
$OrderedReport = $SkuReport | ForEach-Object {
    [PSCustomObject][Ordered]@{
        DisplayName            = $_.DisplayName
        SkuPartNumber          = $_.SkuPartNumber
        SkuId                  = $_.SkuId
        CommerceSubscriptionId = $_.CommerceSubscriptionId
        SubscriptionId         = $_.SubscriptionId
        IsTrial                = $_.IsTrial
        SubscriptionStatus     = $_.SubscriptionStatus
        TotalLicenses          = $_.TotalLicenses
        ActiveUnits            = $_.ActiveUnits
        WarningUnits           = $_.WarningUnits
        ConsumedUnits          = $_.ConsumedUnits
        'Renewal date'         = $_.'Renewal date'
        'Days to renewal'      = $_.'Days to renewal'
        Status                 = $_.Status
    }
}

# Export to CSV with DisplayName as first column
$LicenseRenewalData = Join-Path $GlobalWorkingPath "LicenseRenewalData.csv"
$OrderedReport | Export-Csv -Path $LicenseRenewalData -NoTypeInformation -Encoding UTF8

# Display summary
$OrderedReport | Format-Table DisplayName, SkuPartNumber, ActiveUnits, WarningUnits, ConsumedUnits, "Renewal date", "Days to renewal", Status -AutoSize
Write-Host "`nLicense renewal data saved to: $LicenseRenewalData" -ForegroundColor Cyan