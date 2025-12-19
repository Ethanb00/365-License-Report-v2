$CSVurl = "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv"
$CSVpath = "data\skus.csv"
$CustomCSVpath = "data\CustomSkuNames.csv"

Write-Host "Downloading official Microsoft SKU list..."
Invoke-WebRequest -Uri $CSVurl -OutFile $CSVpath

if (Test-Path -Path $CustomCSVpath) {
    Write-Host "Cleaning up CustomSkuNames.csv..."
    $officialSkus = Import-Csv -Path $CSVpath
    $customSkus = Import-Csv -Path $CustomCSVpath

    $officialIds = @{}
    foreach ($row in $officialSkus) {
        if ($row.GUID) { $officialIds[$row.GUID] = $true }
        if ($row.String_Id) { $officialIds[$row.String_Id] = $true }
    }

    $cleanedCustom = $customSkus | Where-Object { -not $officialIds.ContainsKey($_.Id) }

    if ($cleanedCustom.Count -lt $customSkus.Count) {
        $cleanedCustom | Export-Csv -Path $CustomCSVpath -NoTypeInformation -Encoding utf8
        Write-Host "Removed $($customSkus.Count - $cleanedCustom.Count) entries from $CustomCSVpath because they are now in the official list." -ForegroundColor Green
    }
}