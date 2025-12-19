param(
    [Parameter(Mandatory = $true)]
    [string]$GlobalWorkingPath
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Determine paths
# $GlobalWorkingPath points to the dated folder (e.g. Clients\ClientName\YYYY-MM-DD)
# We want the pricing file in the Client Root (e.g. Clients\ClientName) so it persists across days.
$ClientRoot = Split-Path -Path $GlobalWorkingPath -Parent
$PricingCsvPath = Join-Path -Path $ClientRoot -ChildPath 'ClientPricing.csv'
$SubscribedCsvPath = Join-Path -Path $GlobalWorkingPath -ChildPath 'SubscribedSKUs.csv'

# Load Subscribed SKUs from the current run
$currentSkus = @()
if (Test-Path $SubscribedCsvPath) {
    $currentSkus = Import-Csv -Path $SubscribedCsvPath
}

# Load Existing Pricing from the valid persistent location
$pricingData = @{}
if (Test-Path $PricingCsvPath) {
    $rows = Import-Csv -Path $PricingCsvPath
    foreach ($r in $rows) {
        $price = $r.Cost -replace '[^0-9.]', ''
        $data = @{ Cost = $price; FriendlyName = $r.FriendlyName }
        $pricingData[$r.SkuPartNumber] = $data
    }
}

# Merge Data for Display
$displayList = @()
# 1. Add all current SKUs
foreach ($sku in $currentSkus) {
    $id = $sku.SkuPartNumber
    $name = $sku.FriendlyName
    $cost = "0.00"
    
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

# 2. Add any pricing entries that are not in current SKUs (historical)
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

# Create Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Manage License Pricing"
$form.Size = New-Object System.Drawing.Size(700, 600)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::White

# Instructions
$lblInfo = New-Object System.Windows.Forms.Label
$lblInfo.Text = "Enter the monthly cost for each license SKU below. Values are saved to: $PricingCsvPath"
$lblInfo.Location = New-Object System.Drawing.Point(10, 10)
$lblInfo.Size = New-Object System.Drawing.Size(660, 30)
$lblInfo.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Controls.Add($lblInfo)

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

# Save Button
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

$form.ShowDialog() | Out-Null
