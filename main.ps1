param(
    [switch]$SkipDataCollection,
    [string]$ClientName,
    [switch]$SkipPricing
)

# Initialize script environment
. ".\src\initialization.ps1"

if ($SkipDataCollection) {
    # --- Skip data collection mode: reuse existing CSVs ---
    Write-Host "`n=== SKIP DATA COLLECTION MODE ===" -ForegroundColor Cyan
    Write-Host "Reusing existing CSVs for report generation.`n" -ForegroundColor Cyan

    $OutputRoot = Join-Path -Path $PSScriptRoot -ChildPath "Output\Clients"

    if ($ClientName) {
        $clientPath = Join-Path -Path $OutputRoot -ChildPath $ClientName
        if (-not (Test-Path $clientPath)) {
            Write-Error "Client folder not found: $clientPath"
            exit 1
        }
    }
    else {
        # Prompt user to select a client folder
        $clients = Get-ChildItem -Path $OutputRoot -Directory | Sort-Object Name
        if ($clients.Count -eq 0) {
            Write-Error "No client folders found in $OutputRoot"
            exit 1
        }
        Write-Host "Available clients:"
        for ($i = 0; $i -lt $clients.Count; $i++) {
            Write-Host "  [$i] $($clients[$i].Name)"
        }
        $selection = Read-Host "Select client number"
        $clientPath = $clients[[int]$selection].FullName
    }

    # Find the most recent dated folder
    $datedFolders = Get-ChildItem -Path $clientPath -Directory | Where-Object { $_.Name -match '^\d{4}-\d{2}-\d{2}$' } | Sort-Object Name -Descending
    if ($datedFolders.Count -eq 0) {
        Write-Error "No dated folders found in $clientPath"
        exit 1
    }
    $GlobalWorkingPath = $datedFolders[0].FullName
    Write-Host "Using existing data from: $GlobalWorkingPath`n" -ForegroundColor Green

}
else {
    # --- Normal mode: collect all data ---

    # Redownload SKU csv file
    . ".\src\getSKUcsv.ps1"

    # Connect to Microsoft Graph
    . ".\src\mg-graphConnection.ps1"

    # Select or create client folder
    . ".\src\clientFolderScaffolding.ps1"

    # Collect User Details
    . ".\src\collectUserDetails.ps1" -GlobalWorkingPath $GlobalWorkingPath

    # Create a CSV listing all licenses in the tenant and their details
    . ".\src\GetRenewalDetail.ps1" -GlobalWorkingPath $GlobalWorkingPath
    . ".\src\createLicensingCSV.ps1" -GlobalWorkingPath $GlobalWorkingPath

    . ".\src\createAssignedLicenses.ps1" -GlobalWorkingPath $GlobalWorkingPath
}

# Manage Pricing (GUI) - skip if requested
if (-not $SkipPricing) {
    . ".\src\managePricing.ps1" -GlobalWorkingPath $GlobalWorkingPath
}

# Create HTML Report
. ".\src\generateHTMLReport.ps1" -GlobalWorkingPath $GlobalWorkingPath

# Create Excel Report
. ".\src\generateExcelReport.ps1" -GlobalWorkingPath $GlobalWorkingPath

Write-Host "`n=== Report generation complete ===" -ForegroundColor Green
