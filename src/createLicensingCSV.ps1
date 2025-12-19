param(
    [Parameter(Mandatory = $true)]
    [string]$GlobalWorkingPath
)

# Load SKU friendly-name mapping from data/skus.csv (parent folder of `src`)
$skuCsvPath = Join-Path -Path (Split-Path -Parent $PSScriptRoot) -ChildPath 'data\skus.csv'
if (Test-Path -Path $skuCsvPath) {
    $skuRows = Import-Csv -Path $skuCsvPath
    $skuMapByStringId = @{}
    $skuMapByGuid = @{}
    foreach ($r in $skuRows) {
        if ($r.String_Id -and -not $skuMapByStringId.ContainsKey($r.String_Id)) {
            $skuMapByStringId[$r.String_Id] = $r.Product_Display_Name
        }
        if ($r.GUID -and -not $skuMapByGuid.ContainsKey($r.GUID)) {
            $skuMapByGuid[$r.GUID] = $r.Product_Display_Name
        }
    }
}
else {
    $skuMapByStringId = @{}
    $skuMapByGuid = @{}
}

# Load renewal data (if any) from the same output folder
$renewalCsvPath = Join-Path -Path $GlobalWorkingPath -ChildPath 'LicenseRenewalData.csv'
if (Test-Path -Path $renewalCsvPath) {
    $renewalRows = Import-Csv -Path $renewalCsvPath
    $renewalMapBySkuId = @{}
    $renewalMapByPart = @{}
    foreach ($rr in $renewalRows) {
        # Normalize common column names from different renewal CSV producers
        $normSkuId = $null
        $normPart = $null
        $normDate = $null
        $normNotes = $null

        if ($rr.PSObject.Properties.Name -contains 'SkuId') { $normSkuId = $rr.SkuId }
        elseif ($rr.PSObject.Properties.Name -contains 'Sku_Id') { $normSkuId = $rr.Sku_Id }
        elseif ($rr.PSObject.Properties.Name -contains 'SkuId ') { $normSkuId = $rr.'SkuId ' }

        if ($rr.PSObject.Properties.Name -contains 'SkuPartNumber') { $normPart = $rr.SkuPartNumber }
        elseif ($rr.PSObject.Properties.Name -contains 'SkuPart') { $normPart = $rr.SkuPart }

        if ($rr.PSObject.Properties.Name -contains 'RenewalDate') { $normDate = $rr.RenewalDate }
        elseif ($rr.PSObject.Properties.Name -contains 'Renewal date') { $normDate = $rr.'Renewal date' }
        elseif ($rr.PSObject.Properties.Name -contains 'Renewal_Date') { $normDate = $rr.Renewal_Date }

        if ($rr.PSObject.Properties.Name -contains 'Notes') { $normNotes = $rr.Notes }
        elseif ($rr.PSObject.Properties.Name -contains 'RenewalNotes') { $normNotes = $rr.RenewalNotes }

        $entry = [PSCustomObject]@{
            SkuId         = $normSkuId
            SkuPartNumber = $normPart
            RenewalDate   = $normDate
            Notes         = $normNotes
        }

        if ($entry.SkuId -and -not $renewalMapBySkuId.ContainsKey($entry.SkuId)) { $renewalMapBySkuId[$entry.SkuId] = $entry }
        if ($entry.SkuPartNumber -and -not $renewalMapByPart.ContainsKey($entry.SkuPartNumber)) { $renewalMapByPart[$entry.SkuPartNumber] = $entry }
    }
}
else {
    $renewalMapBySkuId = @{}
    $renewalMapByPart = @{}
}

Get-MgSubscribedSku | ForEach-Object {
    $skuId = $_.SkuId
    $skuPartNumber = $_.SkuPartNumber
    $accountName = $_.AccountName
    $accountId = $_.AccountId
    $capabilityStatus = $_.CapabilityStatus
    $consumedUnits = $_.ConsumedUnits
    $enabledPrepaidUnits = $_.PrepaidUnits.Enabled
    $lockedOutPrepaidUnits = $_.PrepaidUnits.LockedOut
    $suspendedPrepaidUnits = $_.PrepaidUnits.Suspended
    $warningPrepaidUnits = $_.PrepaidUnits.Warning
    $servicePlanIds = ($_.ServicePlans | ForEach-Object { $_.ServicePlanId }) -join ";"
    $servicePlanNames = ($_.ServicePlans | ForEach-Object { $_.ServicePlanName }) -join ";"
    $subscriptionIds = ($_.ServicePlans | ForEach-Object { $_.ProvisioningStatus }) -join ";"

    # Load custom friendly names from data\CustomSkuNames.csv
    $customSkuCsvPath = Join-Path -Path (Split-Path -Parent $PSScriptRoot) -ChildPath 'data\CustomSkuNames.csv'
    $customSkuMap = @{}
    if (Test-Path -Path $customSkuCsvPath) {
        try {
            $customRows = Import-Csv -Path $customSkuCsvPath
            foreach ($cr in $customRows) {
                if ($cr.Id -and $cr.FriendlyName) { $customSkuMap[$cr.Id] = $cr.FriendlyName }
            }
        }
        catch {}
    }

    # Try to resolve friendly display name from SKUs CSV (by SkuPartNumber/String_Id or SkuId/GUID)
    $friendlyName = ''
    if ($skuMapByStringId.ContainsKey($skuPartNumber)) {
        $friendlyName = $skuMapByStringId[$skuPartNumber]
    }
    elseif ($skuMapByGuid.ContainsKey($skuId)) {
        $friendlyName = $skuMapByGuid[$skuId]
    }
    elseif ($customSkuMap.ContainsKey($skuId)) {
        $friendlyName = $customSkuMap[$skuId]
    }
    elseif ($customSkuMap.ContainsKey($skuPartNumber)) {
        $friendlyName = $customSkuMap[$skuPartNumber]
    }

    # If no friendly name found, prompt the user to enter one (shows SkuId and SkuPartNumber)
    if ([string]::IsNullOrWhiteSpace($friendlyName)) {
        $prompt = "No friendly name found for SkuId: $skuId`nSkuPartNumber: $skuPartNumber`nEnter friendly name (or press Enter to leave blank):"
        try {
            $entered = Read-Host -Prompt $prompt
        }
        catch {
            # In some non-interactive contexts Read-Host can fail; fallback to empty
            $entered = ''
        }
        if (-not [string]::IsNullOrWhiteSpace($entered)) {
            $friendlyName = $entered.Trim()
            # Save to CustomSkuNames.csv
            $newEntry = [PSCustomObject]@{
                Id           = $skuId
                FriendlyName = $friendlyName
            }
            $newEntry | Export-Csv -Path $customSkuCsvPath -Append -NoTypeInformation -Encoding utf8
        }
    }

    # Determine renewal info (prefer SkuId, then SkuPartNumber). Compute into local variables to avoid parser issues.
    $mergedRenewalDate = ''
    $mergedRenewalNotes = ''
    if ($renewalMapBySkuId.ContainsKey($skuId)) {
        $mergedRenewalDate = $renewalMapBySkuId[$skuId].RenewalDate
        if ($renewalMapBySkuId[$skuId].PSObject.Properties.Name -contains 'Notes' -and $renewalMapBySkuId[$skuId].Notes) { $mergedRenewalNotes = $renewalMapBySkuId[$skuId].Notes }
        elseif ($renewalMapBySkuId[$skuId].PSObject.Properties.Name -contains 'RenewalNotes' -and $renewalMapBySkuId[$skuId].RenewalNotes) { $mergedRenewalNotes = $renewalMapBySkuId[$skuId].RenewalNotes }
    }
    elseif ($renewalMapByPart.ContainsKey($skuPartNumber)) {
        $mergedRenewalDate = $renewalMapByPart[$skuPartNumber].RenewalDate
        if ($renewalMapByPart[$skuPartNumber].PSObject.Properties.Name -contains 'Notes' -and $renewalMapByPart[$skuPartNumber].Notes) { $mergedRenewalNotes = $renewalMapByPart[$skuPartNumber].Notes }
        elseif ($renewalMapByPart[$skuPartNumber].PSObject.Properties.Name -contains 'RenewalNotes' -and $renewalMapByPart[$skuPartNumber].RenewalNotes) { $mergedRenewalNotes = $renewalMapByPart[$skuPartNumber].RenewalNotes }
    }

    $outputObject = [PSCustomObject]@{
        SkuId                 = $skuId
        SkuPartNumber         = $skuPartNumber
        AccountName           = $accountName
        AccountId             = $accountId
        CapabilityStatus      = $capabilityStatus
        ConsumedUnits         = $consumedUnits
        EnabledPrepaidUnits   = $enabledPrepaidUnits
        LockedOutPrepaidUnits = $lockedOutPrepaidUnits
        SuspendedPrepaidUnits = $suspendedPrepaidUnits
        WarningPrepaidUnits   = $warningPrepaidUnits
        ServicePlanIds        = $servicePlanIds
        ServicePlanNames      = $servicePlanNames
        SubscriptionIds       = $subscriptionIds
        FriendlyName          = $friendlyName
        RenewalDate           = $mergedRenewalDate
        RenewalNotes          = $mergedRenewalNotes
    }

    $outputObject
} | Export-Csv -Path (Join-Path -Path $GlobalWorkingPath -ChildPath "SubscribedSKUs.csv") -NoTypeInformation -Encoding utf8