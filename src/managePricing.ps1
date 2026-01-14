<#
.SYNOPSIS
    Interactive GUI for managing license pricing per SKU

.DESCRIPTION
    Displays a Windows Forms DataGridView allowing users to enter or update
    the monthly cost for each license SKU. Data is persisted in ClientRoot\ClientPricing.csv
    
    Features:
    - Auto-loads existing pricing from ClientPricing.csv
    - Displays current and historical SKUs
    - Read-only SKU and Friendly Name columns
    - Editable Monthly Cost column
    - Numeric validation and cleanup
    
.PARAMETER GlobalWorkingPath
    The dated output folder (e.g., Clients\ClientName\2026-01-14)
    Used to locate ClientRoot (parent folder) for pricing persistence

.OUTPUTS
    ClientPricing.csv - Saved in client root with columns:
    - SkuPartNumber
    - FriendlyName
    - Cost
    - Currency (USD)

.NOTES
    Pricing stored at client level, not date level, for persistence across runs
    Form requires Windows and .NET Framework (Windows only)
    TopmostWindow ensures visibility above other windows
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$GlobalWorkingPath
)

# ============================================================================
# LOAD WINDOWS FORMS ASSEMBLIES
# ============================================================================
# Required for GUI components

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ============================================================================
# DETERMINE PATHS AND LOAD DATA
# ============================================================================
# Pricing stored in client root (parent of dated folder) for persistence
# $GlobalWorkingPath = Clients\ClientName\YYYY-MM-DD
# $ClientRoot = Clients\ClientName

$ClientRoot = Split-Path -Path $GlobalWorkingPath -Parent
$PricingCsvPath = Join-Path -Path $ClientRoot -ChildPath 'ClientPricing.csv'
$SubscribedCsvPath = Join-Path -Path $GlobalWorkingPath -ChildPath 'SubscribedSKUs.csv'

# Load current SKUs from this run's data
$currentSkus = @()
if (Test-Path $SubscribedCsvPath) {
    $currentSkus = Import-Csv -Path $SubscribedCsvPath
}

# Load existing pricing data (persisted at client level)
$pricingData = @{}
if (Test-Path $PricingCsvPath) {
    $rows = Import-Csv -Path $PricingCsvPath
    foreach ($r in $rows) {
        $price = $r.Cost -replace '[^0-9.]', ''
        $data = @{ Cost = $price; FriendlyName = $r.FriendlyName }
        $pricingData[$r.SkuPartNumber] = $data
    }
}

# ============================================================================
# MERGE DATA FOR GUI DISPLAY
# ============================================================================
# Combine current SKUs and historical pricing into single list

$displayList = @()

# 1. Add all current SKUs from this run
foreach ($sku in $currentSkus) {
    $id = $sku.SkuPartNumber
    $name = $sku.FriendlyName
    $cost = "0.00"
    
    # Use existing price if available, otherwise default to 0.00
    if ($pricingData.ContainsKey($id)) {
        if ($pricingData[$id].Cost) { $cost = $pricingData[$id].Cost }
        # Prefer current friendly name, but fallback to saved if missing
        if (-not $name) { $name = $pricingData[$id].FriendlyName }
    }
    
    $displayList += [PSCustomObject]@{
        SkuPartNumber = $id
        FriendlyName  = $name
        Cost          = $cost
    }
}

# 2. Add any historical pricing entries no longer in current SKUs
# This allows users to see what they paid for discontinued licenses
foreach ($key in $pricingData.Keys) {
    $exists = $displayList | Where-Object { $_.SkuPartNumber -eq $key }
    if (-not $exists) {
        $displayList += [PSCustomObject]@{
            SkuPartNumber = $key
            FriendlyName  = $pricingData[$key].FriendlyName
            Cost          = $pricingData[$key].Cost
        }
    }
}

# ============================================================================
# CREATE PRICING MANAGEMENT FORM
# ============================================================================
# Windows Forms GUI for user input
$form = New-Object System.Windows.Forms.Form
$form.Text = "Manage License Pricing"
$form.Size = New-Object System.Drawing.Size(700, 600)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::White
$form.TopMost = $true
$form.BringToFront()

# ============================================================================
# FORM UI COMPONENTS
# ============================================================================

# Instructions label
$lblInfo = New-Object System.Windows.Forms.Label
$lblInfo.Text = "Enter the monthly cost for each license SKU below. Values are saved to: $PricingCsvPath"
$lblInfo.Location = New-Object System.Drawing.Point(10, 10)
$lblInfo.Size = New-Object System.Drawing.Size(660, 30)
$lblInfo.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Controls.Add($lblInfo)

# ============================================================================
# DATAGRIDVIEW: Pricing Data Entry
# ============================================================================

# DataGridView
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Size = New-Object System.Drawing.Size(660, 470)
$grid.Location = New-Object System.Drawing.Point(10, 45)
$grid.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
$grid.AllowUserToAddRows = $false
$grid.AutoSizeColumnsMode = "Fill"
$grid.BackgroundColor = [System.Drawing.Color]::WhiteSmoke
$grid.BorderStyle = "None"
$grid.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# Add Columns
$colId = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colId.Name = "SkuPartNumber"
$colId.HeaderText = "Sku Part Number"
$colId.ReadOnly = $true
$colId.FillWeight = 30
$grid.Columns.Add($colId) | Out-Null

$colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colName.Name = "FriendlyName"
$colName.HeaderText = "Friendly Name"
$colName.ReadOnly = $true
$colName.FillWeight = 40
$grid.Columns.Add($colName) | Out-Null

$colCost = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colCost.Name = "Cost"
$colCost.HeaderText = "Monthly Cost ($)"
$colCost.FillWeight = 20
$colCost.DefaultCellStyle.Alignment = "MiddleRight"
$grid.Columns.Add($colCost) | Out-Null

# Populate Grid
foreach ($item in $displayList) {
    if ($null -eq $item.Cost -or $item.Cost -eq '') { $item.Cost = "0.00" }
    $row = $item.SkuPartNumber, $item.FriendlyName, $item.Cost
    $grid.Rows.Add($row) | Out-Null
}

$form.Controls.Add($grid)

# ============================================================================
# SAVE BUTTON
# ============================================================================
# Validates and exports pricing data to CSV

$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Text = "Save Pricing"
$btnSave.Location = New-Object System.Drawing.Point(560, 525)
$btnSave.Size = New-Object System.Drawing.Size(110, 30)
$btnSave.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$btnSave.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212) # Windows Blue
$btnSave.ForeColor = [System.Drawing.Color]::White
$btnSave.FlatStyle = "Flat"
$btnSave.DialogResult = "OK"
$btnSave.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

# Click handler: Export pricing data
$btnSave.Add_Click({
        $finalList = @()
        foreach ($row in $grid.Rows) {
            if (-not $row.IsNewRow) {
                $costVal = $row.Cells["Cost"].Value
                # Basic cleanup/validation
                $costVal = [string]$costVal -replace '[^0-9.]', ''
                if (-not $costVal) { $costVal = "0.00" }
            
                $finalList += [PSCustomObject]@{
                    SkuPartNumber = $row.Cells["SkuPartNumber"].Value
                    FriendlyName  = $row.Cells["FriendlyName"].Value
                    Cost          = $costVal
                    Currency      = "USD"
                }
            }
        }
    
        $finalList | Export-Csv -Path $PricingCsvPath -NoTypeInformation
        Write-Host "Pricing saved to $PricingCsvPath" -ForegroundColor Green
        $form.Close()
    })

$form.Controls.Add($btnSave)

# ============================================================================
# DISPLAY FORM
# ============================================================================
# Show the pricing management window

Write-Host "`n=== Opening Pricing Management Window ===" -ForegroundColor Cyan
Write-Host "Configure monthly costs for each SKU in the window." -ForegroundColor Yellow
Write-Host "Click 'Save Pricing' when complete." -ForegroundColor Yellow
Write-Host "Pricing is saved to: $PricingCsvPath`n" -ForegroundColor Gray

$form.ShowDialog() | Out-Null
