<#
.SYNOPSIS
    Downloads and maintains the official Microsoft SKU catalog

.DESCRIPTION
    - Downloads the latest official Microsoft 365 SKU list from Microsoft
    - Updates data\skus.csv with current product and service plan identifiers
    - Cleans up CustomSkuNames.csv by removing entries that are now in official list
    - Reduces manual SKU name management workload
    
.NOTES
    Source: Microsoft's official Product names and service plan identifiers for licensing
    Updated: Each time the script runs
    Impact: Ensures reports use the latest official SKU names
#>

# ============================================================================
# DOWNLOAD OFFICIAL MICROSOFT SKU CATALOG
# ============================================================================
# Fetch the latest product and service plan identifiers from Microsoft

$CSVurl = "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv"
$CSVpath = "data\skus.csv"
$CustomCSVpath = "data\CustomSkuNames.csv"

Write-Host "Downloading official Microsoft SKU list..." -ForegroundColor Cyan
Invoke-WebRequest -Uri $CSVurl -OutFile $CSVpath

Write-Host "Successfully downloaded SKU catalog to: $CSVpath" -ForegroundColor Green

# ============================================================================
# CLEANUP CUSTOM SKU NAMES
# ============================================================================
# Remove custom SKU entries that are now in the official Microsoft list
# This reduces duplication and simplifies maintenance

if (Test-Path -Path $CustomCSVpath) {
    Write-Host "Cleaning up CustomSkuNames.csv..." -ForegroundColor Cyan
    
    # Load official SKUs
    $officialSkus = Import-Csv -Path $CSVpath
    $customSkus = Import-Csv -Path $CustomCSVpath

    # Build a lookup table of all official SKU identifiers
    $officialIds = @{}
    foreach ($row in $officialSkus) {
        if ($row.GUID) { $officialIds[$row.GUID] = $true }
        if ($row.String_Id) { $officialIds[$row.String_Id] = $true }
    }

    # Remove any custom entries that now exist officially
    $cleanedCustom = $customSkus | Where-Object { -not $officialIds.ContainsKey($_.Id) }

    # If entries were removed, export the cleaned list
    if ($cleanedCustom.Count -lt $customSkus.Count) {
        $cleanedCustom | Export-Csv -Path $CustomCSVpath -NoTypeInformation -Encoding utf8
        $removedCount = $customSkus.Count - $cleanedCustom.Count
        Write-Host "Removed $removedCount custom SKU entries (now in official list)" -ForegroundColor Green
    } else {
        Write-Host "No duplicate custom SKUs to clean up" -ForegroundColor Gray
    }
}