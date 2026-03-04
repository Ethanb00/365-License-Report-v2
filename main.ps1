<#
.SYNOPSIS
    Microsoft 365 License Report Generator - Main Orchestration Script

.DESCRIPTION
    Generates comprehensive license reports for Microsoft 365 tenants including:
    - User license assignments and consumption data
    - Subscription renewal dates and cost analysis
    - Interactive HTML dashboard and Excel workbook
    - Inactive user identification (30+ days without sign-in)
    
    Supports multi-client management with persistent pricing and historical data.

.PARAMETER SkipDataCollection
    Bypasses Microsoft Graph data collection and regenerates reports using existing CSV files.
    Useful for quickly regenerating reports after pricing changes or when working offline.
    
    Example: .\main.ps1 -SkipDataCollection

.PARAMETER ClientName
    Specifies the client name to use when running in SkipDataCollection mode.
    If omitted, the script will prompt for interactive client selection.
    
    Example: .\main.ps1 -SkipDataCollection -ClientName "Contoso Ltd"

.PARAMETER SkipPricing
    Bypasses the pricing management GUI and uses existing pricing data from ClientPricing.csv.
    Required for automated/scheduled execution or when pricing is already configured.
    
    Example: .\main.ps1 -SkipPricing

.EXAMPLE
    .\main.ps1
    Standard execution: Collects all data, opens pricing GUI, generates reports

.EXAMPLE
    .\main.ps1 -SkipDataCollection -ClientName "Contoso" -SkipPricing
    Quick regeneration: Uses existing data and pricing for specified client

.EXAMPLE
    .\main.ps1 -SkipPricing
    Automated mode: Collects fresh data but skips pricing GUI (uses saved pricing)

.NOTES
    Author: Ethan Bennett
    Organization: Dataprise
    Version: 2.0
    Requires: PowerShell 7+ (or Windows PowerShell 5.1 minimum)
              Microsoft.Graph module
              Microsoft.PowerShell.SecretManagement module
              ImportExcel module
    
    For detailed documentation, see README.md
#>

param(
    [switch]$SkipDataCollection,
    [string]$ClientName,
    [switch]$SkipPricing
)

# ============================================================================
# INITIALIZATION
# ============================================================================
# Verify environment, install required modules, configure secret storage
. ".\src\initialization.ps1"

# ============================================================================
# AUTHENTICATION
# ============================================================================
# Connect to Microsoft Graph with required scopes (or verify existing connection)
. ".\src\mg-graphConnection.ps1"

# ============================================================================
# AUTHENTICATION
# ============================================================================
# Connect to Microsoft Graph with required scopes (or verify existing connection)
. ".\src\mg-graphConnection.ps1"

# ============================================================================
# EXECUTION MODE: SKIP DATA COLLECTION
# ============================================================================
# Use existing CSV files to regenerate reports without querying Microsoft Graph
# Useful for: quick iterations, pricing adjustments, offline work

if ($SkipDataCollection) {
    Write-Host "`n=== SKIP DATA COLLECTION MODE ===" -ForegroundColor Cyan
    Write-Host "Reusing existing CSVs for report generation.`n" -ForegroundColor Cyan

    $OutputRoot = Join-Path -Path $PSScriptRoot -ChildPath "Output\Clients"

    # Client folder selection logic
    if ($ClientName) {
        # Use specified client name
        $clientPath = Join-Path -Path $OutputRoot -ChildPath $ClientName
        if (-not (Test-Path $clientPath)) {
            Write-Error "Client folder not found: $clientPath"
            exit 1
        }
    }
    else {
        # Interactive client selection
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

    # Find the most recent dated folder (YYYY-MM-DD format)
    $datedFolders = Get-ChildItem -Path $clientPath -Directory | 
        Where-Object { $_.Name -match '^\d{4}-\d{2}-\d{2}$' } | 
        Sort-Object Name -Descending
    
    if ($datedFolders.Count -eq 0) {
        Write-Error "No dated folders found in $clientPath"
        exit 1
    }
    
    $GlobalWorkingPath = $datedFolders[0].FullName
    Write-Host "Using existing data from: $GlobalWorkingPath`n" -ForegroundColor Green

}
else {
    # ============================================================================
    # EXECUTION MODE: FULL DATA COLLECTION
    # ============================================================================
    # Query Microsoft Graph for all license and user data

    # Step 1: Download latest official Microsoft SKU catalog
    # Updates data\skus.csv and cleans up data\CustomSkuNames.csv
    . ".\src\getSKUcsv.ps1"

    # Step 2: Client folder scaffolding
    # Prompts user to select existing client or create new one
    # Creates dated subfolder (YYYY-MM-DD) for today's data
    # Sets $GlobalWorkingPath to the dated folder path
    . ".\src\clientFolderScaffolding.ps1"

    # Step 3: Collect user details from Microsoft Graph
    # Queries: UserPrincipalName, DisplayName, Department, JobTitle,
    #          AccountEnabled, CreatedDateTime, SignInActivity, License assignments
    # Output: user_details.csv
    . ".\src\collectUserDetails.ps1" -GlobalWorkingPath $GlobalWorkingPath

    # Step 4: Fetch subscription renewal dates and license metadata
    # Uses /directory/subscriptions endpoint for renewal tracking
    # Output: LicenseRenewalData.csv
    . ".\src\GetRenewalDetail.ps1" -GlobalWorkingPath $GlobalWorkingPath
    
    # Step 5: Create comprehensive SKU mapping CSV
    # Merges official SKU data with custom names and renewal info
    # Prompts for friendly names on unknown SKUs
    # Output: SubscribedSKUs.csv
    . ".\src\createLicensingCSV.ps1" -GlobalWorkingPath $GlobalWorkingPath

    # Step 6: Create normalized license assignment report
    # One row per user with all assigned licenses
    # Includes both direct and group-based assignments
    # Output: AssignedLicenses.csv, AssignedLicenses_Summary.csv
    . ".\src\createAssignedLicenses.ps1" -GlobalWorkingPath $GlobalWorkingPath

}

# ============================================================================
# PRICING MANAGEMENT
# ============================================================================
# Open GUI for entering/updating monthly cost per SKU (unless -SkipPricing specified)
# Pricing data is stored in ClientRoot\ClientPricing.csv (persists across daily runs)

if (-not $SkipPricing) {
    . ".\src\managePricing.ps1" -GlobalWorkingPath $GlobalWorkingPath
}

# ============================================================================
# REPORT GENERATION
# ============================================================================

# Generate Interactive HTML Dashboard
# Creates: LicenseReport.html with cost summaries, license tables, inactive user alerts
. ".\src\generateHTMLReport.ps1" -GlobalWorkingPath $GlobalWorkingPath

# Generate Excel Workbook
# Creates: LicenseReport.xlsx with formatted sheets for all data
. ".\src\generateExcelReport.ps1" -GlobalWorkingPath $GlobalWorkingPath

# ============================================================================
# COMPLETION
# ============================================================================
Write-Host "`n=== Report generation complete ===" -ForegroundColor Green
Write-Host "Reports saved to: $GlobalWorkingPath" -ForegroundColor Cyan
Write-Host "`nGenerated files:" -ForegroundColor Yellow
Write-Host "  - LicenseReport.html (Interactive Dashboard)" -ForegroundColor White
Write-Host "  - LicenseReport.xlsx (Excel Workbook)" -ForegroundColor White
Write-Host "  - AssignedLicenses.csv (User License Assignments)" -ForegroundColor White
Write-Host "  - LicenseRenewalData.csv (Subscription Renewals)" -ForegroundColor White
Write-Host "  - SubscribedSKUs.csv (SKU Details)" -ForegroundColor White
