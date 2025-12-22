param(
  [Parameter(Mandatory = $true)]
  [string]$GlobalWorkingPath,
  [string]$SubscribedCsvPath = $(Join-Path -Path $GlobalWorkingPath -ChildPath 'SubscribedSKUs.csv'),
  [string]$RenewalCsvPath = $(Join-Path -Path $GlobalWorkingPath -ChildPath 'LicenseRenewalData.csv'),
  [switch]$OpenReport
)

function Safe-ImportCsv($path) {
  if (Test-Path -Path $path) { Import-Csv -Path $path } else { @() }
}


# Company Details
$parentPath = Split-Path -Path $GlobalWorkingPath -Parent
$companyName = Split-Path -Path $parentPath -Leaf

# Secure string helper (if not already present) and logo lookup using Logo.dev only for custom domains
function SecureStringToPlain([System.Security.SecureString]$ss) {
  if ($null -eq $ss) { return '' }
  $ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($ss)
  try { [Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr) } finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr) }
}

# Determine domain from signed-in Microsoft Graph user (prefer authoritative tenant info)
$domain = $null
try {
  $me = (Get-MgContext).Account
  $domain = $me.Split('@')[-1].ToLower()
  if ($domain -eq 'onmicrosoft.com' -or $domain -match '\.onmicrosoft\.com$') {
    $org = Get-MgOrganization -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($org.DefaultDomain) {
      $domain = $org.DefaultDomain.ToLower()
    }
  }
  $LogoUrl = ''
  # Only attempt to fetch logo for domains that are not the default onmicrosoft domain
  $secureApiKey = Get-Secret -Name 'LogoApiKey' -ErrorAction SilentlyContinue
  if ($secureApiKey) {
    $LogoToken = SecureStringToPlain $secureApiKey
    $LogoUrl = "https://img.logo.dev/$($domain)?token=$($LogoToken)"
  }
}
catch {
  Write-Host "Failed to determine domain from Microsoft Graph context: $_" -ForegroundColor Yellow
}

$skus = Safe-ImportCsv -path $SubscribedCsvPath
$renewals = Safe-ImportCsv -path $RenewalCsvPath
# Load normalized assigned licenses if present
$assignedPath = Join-Path -Path $GlobalWorkingPath -ChildPath 'AssignedLicenses.csv'
$assigned = Safe-ImportCsv -path $assignedPath

# Load Pricing Data
# Logic: Look in Client Root (parent of dated folder) for persistence
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


# Filter licensed and unlicensed users
$licensed = $assigned | Where-Object { $_.AssignedFriendlyNames -and $_.AssignedFriendlyNames -ne 'None' }
$unlicensed = $assigned | Where-Object { -not $_.AssignedFriendlyNames -or $_.AssignedFriendlyNames -eq 'None' }


$topSkus = $skus | Sort-Object @{Expression = { [int]($_.ConsumedUnits) }; Descending = $true } | Select-Object -First 10

$generated = (Get-Date).ToString('MMMM dd, yyyy')

# Process renewal data for summaries and enhanced table
$today = (Get-Date).Date
$renewalsParsed = @()
foreach ($rw in $renewals) {
  # Normalize renewal date and notes from various CSV column names
  $rawDate = $null
  if ($rw.PSObject.Properties.Name -contains 'RenewalDate') { $rawDate = $rw.RenewalDate }
  elseif ($rw.PSObject.Properties.Name -contains 'Renewal date') { $rawDate = $rw.'Renewal date' }
  elseif ($rw.PSObject.Properties.Name -contains 'Renewal_Date') { $rawDate = $rw.Renewal_Date }
  elseif ($rw.PSObject.Properties.Name -contains 'nextLifecycleDateTime') { $rawDate = $rw.nextLifecycleDateTime }

  $parsed = $null
  if ($rawDate) {
    try { $parsed = [datetime]::Parse($rawDate) } catch { $parsed = $null }
  }

  $notes = $null
  if ($rw.PSObject.Properties.Name -contains 'Notes') { $notes = $rw.Notes }
  elseif ($rw.PSObject.Properties.Name -contains 'RenewalNotes') { $notes = $rw.RenewalNotes }

  # Normalize SKU identifiers
  $skuPart = if ($rw.PSObject.Properties.Name -contains 'SkuPartNumber') { $rw.SkuPartNumber } elseif ($rw.PSObject.Properties.Name -contains 'SkuPart') { $rw.SkuPart } else { $rw.SkuPartNumber }
  $skuIdVal = if ($rw.PSObject.Properties.Name -contains 'SkuId') { $rw.SkuId } elseif ($rw.PSObject.Properties.Name -contains 'Sku_Id') { $rw.Sku_Id } else { $rw.SkuId }

  $renewalsParsed += [PSCustomObject]@{
    SkuId                  = $skuIdVal
    SkuPartNumber          = $skuPart
    RenewalDate            = $rawDate
    ParsedDate             = $parsed
    Notes                  = $notes
    CommerceSubscriptionId = if ($rw.PSObject.Properties.Name -contains 'CommerceSubscriptionId') { $rw.CommerceSubscriptionId } else { '' }
    SubscriptionId         = if ($rw.PSObject.Properties.Name -contains 'SubscriptionId') { $rw.SubscriptionId } else { '' }
    IsTrial                = if ($rw.PSObject.Properties.Name -contains 'IsTrial') { $rw.IsTrial } else { '' }
    SubscriptionStatus     = if ($rw.PSObject.Properties.Name -contains 'SubscriptionStatus') { $rw.SubscriptionStatus } else { '' }
    TotalLicenses          = if ($rw.PSObject.Properties.Name -contains 'TotalLicenses') { $rw.TotalLicenses } else { '' }
  }
}

# Enrich renewalsParsed with a FriendlyName by matching imported SKUs (robust comparison)
foreach ($r in $renewalsParsed) {
  $friendly = $null
  if ($skus) {
    if ($r.SkuId) {
      $match = $skus | Where-Object {
        $a = ($_.'SkuId' -as [string])
        $b = ($r.SkuId -as [string])
        if ($a -and $b) { $a.Trim().ToLower() -eq $b.Trim().ToLower() } else { $false }
      } | Select-Object -First 1
      if ($match) { $friendly = $match.FriendlyName }
    }
    if (-not $friendly -and $r.SkuPartNumber) {
      $match2 = $skus | Where-Object {
        $a = ($_.'SkuPartNumber' -as [string])
        $b = ($r.SkuPartNumber -as [string])
        if ($a -and $b) { $a.Trim().ToLower() -eq $b.Trim().ToLower() } else { $false }
      } | Select-Object -First 1
      if ($match2) { $friendly = $match2.FriendlyName }
    }
  }
  if (-not $friendly) { $friendly = if ($r.SkuPartNumber) { $r.SkuPartNumber } else { '' } }
  $r | Add-Member -NotePropertyName FriendlyName -NotePropertyValue $friendly -Force
}

$renewalsKnownCount = ($renewalsParsed | Where-Object { $_.ParsedDate -ne $null }).Count
$renewalsUpcoming90 = $renewalsParsed | Where-Object { $_.ParsedDate -ne $null -and $_.ParsedDate -ge $today -and $_.ParsedDate -le $today.AddDays(90) }

# Calculate Next Paid Renewal
$nextPaidRenewalDate = $null
$nextPaidRenewalName = ''
$nextPaidRenewalDays = $null

$paidRenewals = $renewalsParsed | Where-Object { 
  $_.ParsedDate -ne $null -and 
  $pricingTable.ContainsKey($_.SkuPartNumber) -and 
  $pricingTable[$_.SkuPartNumber] -gt 0 
} | Sort-Object ParsedDate

if ($paidRenewals.Count -gt 0) {
  $nextPaidRenewalDate = $paidRenewals[0].ParsedDate
  $nextPaidRenewalName = $paidRenewals[0].FriendlyName
  $nextPaidRenewalDays = ([int]([math]::Floor(($nextPaidRenewalDate - $today).TotalDays)))
}

# Calculate Total Paid Users
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
$totalPaidUsersCount = $paidUsersSet.Count
$renewalsUpcomingCount = $renewalsUpcoming90.Count
$nextRenewal = $renewalsParsed | Where-Object { $_.ParsedDate -ne $null -and $_.ParsedDate -ge $today } | Sort-Object ParsedDate | Select-Object -First 1
$displayNextRenewal = ''
if ($nextRenewal) {
  # Find all renewals on the same date (date-only comparison)
  $sameDateList = @()
  if ($renewalsParsed) {
    $sameDateList = $renewalsParsed | Where-Object { $_.ParsedDate -ne $null -and ([datetime]$_.ParsedDate).Date -eq ([datetime]$nextRenewal.ParsedDate).Date }
  }

  # If multiple licenses share the next renewal date, show a grouped message
  if ($sameDateList -and ($sameDateList.Count -gt 1)) {
    $dateOnly = ([datetime]$nextRenewal.ParsedDate).ToString('MM-dd-yyyy')
    $displayNextRenewal = "Multiple licenses expiring on $dateOnly ($($sameDateList.Count))"
  }
  else {
    # Single next renewal — resolve a friendly name robustly
    $nextName = $null
    if ($nextRenewal.PSObject.Properties.Name -contains 'FriendlyName' -and ($nextRenewal.FriendlyName -as [string]) -and $nextRenewal.FriendlyName.Trim() -ne '') {
      $nextName = $nextRenewal.FriendlyName.Trim()
    }
    if (-not $nextName -and $skus) {
      if ($nextRenewal.SkuId) {
        $match = $skus | Where-Object {
          $a = ($_.'SkuId' -as [string])
          $b = ($nextRenewal.SkuId -as [string])
          if ($a -and $b) { $a.Trim().ToLower() -eq $b.Trim().ToLower() } else { $false }
        } | Select-Object -First 1
        if ($match) { $nextName = $match.FriendlyName }
      }
      if (-not $nextName -and $nextRenewal.SkuPartNumber) {
        $match2 = $skus | Where-Object {
          $a = ($_.'SkuPartNumber' -as [string])
          $b = ($nextRenewal.SkuPartNumber -as [string])
          if ($a -and $b) { $a.Trim().ToLower() -eq $b.Trim().ToLower() } else { $false }
        } | Select-Object -First 1
        if ($match2) { $nextName = $match2.FriendlyName }
      }
      if (-not $nextName -and $nextRenewal.SkuPartNumber) {
        $match3 = $skus | Where-Object { ($_.FriendlyName -as [string]) -and ($_.FriendlyName.ToLower().Contains(($nextRenewal.SkuPartNumber -as [string]).Trim().ToLower())) } | Select-Object -First 1
        if ($match3) { $nextName = $match3.FriendlyName }
      }
    }

    if ($nextName -and ($nextName -as [string]) -and $nextName.Trim() -ne '') {
      $displayNextRenewal = "$nextName - $($nextRenewal.ParsedDate.ToString('MM-dd-yyyy'))"
    }
    else {
      $displayNextRenewal = $nextRenewal.ParsedDate.ToString('MM-dd-yyyy')
    }
  }
}


$css = @'
/* Premium Modern Theme */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

:root {
  --bg-gradient-start: #0f0f23;
  --bg-gradient-end: #1a1a3e;
  --card-bg: rgba(255,255,255,0.03);
  --card-border: rgba(255,255,255,0.08);
  --card-hover: rgba(255,255,255,0.06);
  --text-primary: #f8fafc;
  --text-secondary: #94a3b8;
  --text-muted: #64748b;
  --accent-gradient: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
  --accent-blue: #3b82f6;
  --accent-purple: #8b5cf6;
  --accent-emerald: #10b981;
  --accent-amber: #f59e0b;
  --accent-rose: #f43f5e;
  --glass-blur: blur(20px);
  --shadow-glow: 0 0 40px rgba(102,126,234,0.15);
}

* { box-sizing: border-box; }

body {
  font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
  background: linear-gradient(135deg, var(--bg-gradient-start) 0%, var(--bg-gradient-end) 100%);
  min-height: 100vh;
  color: var(--text-primary);
  margin: 0;
  padding: 32px;
  line-height: 1.6;
}

.container { max-width: 1400px; margin: 0 auto; }

.header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 32px;
  padding: 24px 32px;
  background: var(--card-bg);
  backdrop-filter: var(--glass-blur);
  border: 1px solid var(--card-border);
  border-radius: 20px;
  box-shadow: var(--shadow-glow);
}

.logo img.logo { 
  max-height: 60px; 
  margin-right: 20px;
  border-radius: 12px;
  filter: drop-shadow(0 4px 12px rgba(0,0,0,0.3));
}

.title {
  font-size: 28px;
  font-weight: 800;
  background: var(--accent-gradient);
  -webkit-background-clip: text;
  -webkit-text-fill-color: transparent;
  background-clip: text;
  letter-spacing: -0.02em;
}

.small { color: var(--text-secondary); font-size: 13px; font-weight: 500; }

/* Navigation */
.nav {
  display: flex;
  gap: 12px;
  margin: 0 0 32px 0;
  padding: 8px;
  background: rgba(255, 255, 255, 0.04);
  backdrop-filter: blur(12px);
  -webkit-backdrop-filter: blur(12px);
  border: 1px solid rgba(255, 255, 255, 0.1);
  border-radius: 20px;
  width: fit-content;
  box-shadow: 0 8px 32px rgba(0, 0, 0, 0.2);
}

.nav button {
  background: transparent;
  border: none;
  padding: 10px 20px;
  border-radius: 14px;
  cursor: pointer;
  font-family: inherit;
  font-size: 14px;
  font-weight: 600;
  color: var(--text-secondary);
  transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
  display: flex;
  align-items: center;
  gap: 8px;
}

.nav button span {
  font-size: 16px;
  opacity: 0.7;
  transition: transform 0.3s ease;
}

.nav button:hover {
  background: rgba(255, 255, 255, 0.08);
  color: var(--text-primary);
  transform: translateY(-1px);
}

.nav button:hover span {
  opacity: 1;
  transform: scale(1.1);
}

.nav button.active {
  background: var(--accent-gradient);
  color: #fff;
  box-shadow: 0 10px 20px rgba(102, 126, 234, 0.3);
}

.nav button.active span {
  opacity: 1;
}

/* Overview Cards */
.overview-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
  gap: 20px;
  margin-bottom: 32px;
}

.overview-card {
  background: rgba(255, 255, 255, 0.05);
  backdrop-filter: blur(16px);
  -webkit-backdrop-filter: blur(16px);
  padding: 24px 28px;
  border-radius: 24px;
  border: 1px solid rgba(255, 255, 255, 0.1);
  transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
  position: relative;
  display: flex;
  flex-direction: column;
  min-height: 180px;
  box-shadow: 0 10px 30px rgba(0, 0, 0, 0.1);
}

.card-header {
  display: flex;
  justify-content: space-between;
  align-items: flex-start;
  margin-bottom: 20px;
}

.card-label-group {
  display: flex;
  flex-direction: column;
  gap: 6px;
}

.overview-card:hover {
  transform: translateY(-6px);
  background: rgba(255, 255, 255, 0.08);
  border-color: rgba(255, 255, 255, 0.2);
  box-shadow: 0 20px 40px rgba(0, 0, 0, 0.3), 0 0 20px rgba(102, 126, 234, 0.2);
}

.overview-card .icon {
  font-size: 1.8rem;
  transition: transform 0.4s ease;
  line-height: 1;
}

.overview-card:hover .icon {
  transform: scale(1.2) rotate(5deg);
}

.overview-card .label {
  color: var(--text-muted);
  font-size: 0.7rem;
  font-weight: 700;
  text-transform: uppercase;
  letter-spacing: 0.1em;
  margin: 0;
}

.overview-card .value {
  font-size: 1.5rem;
  font-weight: 800;
  color: var(--text-primary);
  line-height: 1.2;
  letter-spacing: -0.01em;
  margin-top: auto;
  margin-bottom: 8px;
  white-space: nowrap;
}

.overview-card .sub-value {
  font-size: 0.8rem;
  color: var(--text-secondary);
  height: 1.2rem; /* Fixed height for sub-value to ensure alignment */
  margin: 0;
}

.overview-card .badge-inline {
  display: inline-block;
}
  letter-spacing: 0.05em;
}

.badge-green { background: rgba(16,185,129,0.15); color: var(--accent-emerald); border: 1px solid rgba(16,185,129,0.3); }
.badge-amber { background: rgba(245,158,11,0.15); color: var(--accent-amber); border: 1px solid rgba(245,158,11,0.3); }
.badge-red { background: rgba(244,63,94,0.15); color: var(--accent-rose); border: 1px solid rgba(244,63,94,0.3); }
.badge-neutral { background: rgba(148,163,184,0.1); color: var(--text-secondary); border: 1px solid rgba(148,163,184,0.2); }

/* Sections & Tables */
.section {
  margin-top: 28px;
  background: var(--card-bg);
  backdrop-filter: var(--glass-blur);
  border-radius: 20px;
  padding: 24px;
  border: 1px solid var(--card-border);
}

.section h3, .section h4 {
  color: var(--text-primary);
  font-weight: 700;
  margin: 0 0 16px 0;
}

table {
  width: 100%;
  border-collapse: collapse;
  margin-top: 12px;
}

th, td {
  padding: 14px 16px;
  text-align: left;
  border-bottom: 1px solid var(--card-border);
  font-size: 13px;
}

th {
  background: rgba(255,255,255,0.02);
  font-weight: 600;
  color: var(--text-secondary);
  text-transform: uppercase;
  font-size: 11px;
  letter-spacing: 0.05em;
  position: relative;
  cursor: pointer;
  user-select: none;
}

th:hover { background: rgba(255,255,255,0.04); }

th::after {
  content: '\2195';
  position: absolute;
  right: 12px;
  top: 50%;
  transform: translateY(-50%);
  opacity: 0.3;
  font-size: 10px;
}

th[data-order="asc"]::after { content: '\25B2'; opacity: 0.8; color: var(--accent-blue); }
th[data-order="desc"]::after { content: '\25BC'; opacity: 0.8; color: var(--accent-blue); }

td { color: var(--text-primary); }

tr:hover td { background: rgba(255,255,255,0.02); }

tfoot tr { border-top: 2px solid var(--card-border); }
tfoot td { font-weight: 700; color: var(--text-primary); }

/* Search inputs */
.table-search {
  background: rgba(255,255,255,0.05);
  border: 1px solid var(--card-border);
  border-radius: 10px;
  padding: 10px 16px;
  color: var(--text-primary);
  font-family: inherit;
  font-size: 14px;
  width: 280px;
  transition: all 0.2s ease;
}

.table-search:focus {
  outline: none;
  border-color: var(--accent-blue);
  background: rgba(255,255,255,0.08);
  box-shadow: 0 0 0 3px rgba(59,130,246,0.15);
}

.table-search::placeholder { color: var(--text-muted); }

/* Resizer */
.resizer {
  position: absolute;
  right: 0;
  top: 0;
  bottom: 0;
  width: 4px;
  cursor: col-resize;
  z-index: 10;
}

.resizer:hover, .resizing { background: var(--accent-blue); }

/* Details/Summary */
details { margin-top: 20px; }

summary {
  cursor: pointer;
  color: var(--text-secondary);
  font-weight: 600;
  padding: 12px 0;
  transition: color 0.2s ease;
}

summary:hover { color: var(--text-primary); }

summary h4 { display: inline; margin: 0; }

/* Footer */
.footer {
  margin-top: 48px;
  border-top: 1px solid var(--card-border);
  padding-top: 24px;
  text-align: center;
  color: var(--text-muted);
  font-size: 12px;
}

.footer a { color: var(--accent-blue); text-decoration: none; }
.footer a:hover { text-decoration: underline; }

/* Utilities */
.muted { color: var(--text-muted); }

/* Responsive */
@media (max-width: 800px) {
  body { padding: 16px; }
  .header { flex-direction: column; align-items: flex-start; gap: 16px; padding: 20px; }
  .nav { width: 100%; overflow-x: auto; }
  .overview-grid { grid-template-columns: 1fr; }
}
'@

function To-HtmlSafe($s) { if ($null -eq $s) { '' } else { [System.Net.WebUtility]::HtmlEncode([string]$s) } }

# Format numeric quantities for display: thousands separators, and use infinity '∞' when > 1000
function Format-Qty($v) {
  if ($null -eq $v -or $v -eq '') { return '' }
  $ok = $false
  try {
    $n = [int]$v
    $ok = $true
  }
  catch {
    $ok = $false
  }
  if (-not $ok) { return [string]$v }
  if ($n -gt 1000) { return '∞' }
  return ('{0:N0}' -f $n)
}

# Helper to format date strings to MM-dd-yyyy, stripping time
function Format-Date($d) {
  if ([string]::IsNullOrWhiteSpace($d)) { return '' }
  try {
    $dt = [datetime]$d
    return $dt.ToString('MMMM dd, yyyy')
  }
  catch {
    return $d
  }
}

# Totals and display values used in HTML
$totalSkuRows = 0
$displayTotalConsumed = ''
$displayTotalPrepaid = ''
try {
  $totalSkuRows = if ($skus) { $skus.Count } else { 0 }
  if ($totalSkuRows -gt 0) {
    $totalConsumed = ($skus | ForEach-Object { try { [int]($_.ConsumedUnits) } catch { 0 } } | Measure-Object -Sum).Sum
    $totalPrepaid = ($skus | ForEach-Object { try { [int]($_.EnabledPrepaidUnits) } catch { 0 } } | Measure-Object -Sum).Sum
    $displayTotalConsumed = Format-Qty $totalConsumed
    $displayTotalPrepaid = Format-Qty $totalPrepaid
    
    # Calculate Estimated Monthly Cost
    $estMonthlyCost = 0
    foreach ($s in $skus) {
      $cost = 0
      if ($pricingTable.ContainsKey($s.SkuPartNumber)) { $cost = $pricingTable[$s.SkuPartNumber] }
      $units = 0
      try { $units = [int]$s.EnabledPrepaidUnits } catch {}
      $estMonthlyCost += ($units * $cost)
    }
    $displayMonthlyCost = '$' + ('{0:N2}' -f $estMonthlyCost)
  }
}
catch {
  $totalSkuRows = 0
  $displayTotalConsumed = ''
  $displayTotalPrepaid = ''
  # $displayMonthlyCost = '$0.00' # Already initialized above, no need to re-set here
}

# Initialize accumulation variables to prevent duplicates when rerunning in same session
$totalTableCostMo = 0
$totalTableCostYr = 0
$paidRowsHtml = ""
$freeRowsHtml = ""
$freeRenewalsHtml = ""

foreach ($r in $skus) {
  $friendly = To-HtmlSafe $r.FriendlyName
  $cost = 0
  if ($pricingTable.ContainsKey($r.SkuPartNumber)) { $cost = $pricingTable[$r.SkuPartNumber] }
  $units = 0
  try { $units = [int]$r.EnabledPrepaidUnits } catch {}
  $lineTotal = $units * $cost
    
  $costYear = $cost * 12
  $lineTotalYear = $lineTotal * 12
  
  $totalTableCostMo += $lineTotal
  $totalTableCostYr += $lineTotalYear

  $costMoStr = '{0:N2}' -f $cost
  $costYrStr = '{0:N2}' -f $costYear
  $totalMoStr = '{0:N2}' -f $lineTotal
  $totalYrStr = '{0:N2}' -f $lineTotalYear

  if ($cost -gt 0) {
    $paidRowsHtml += "<tr>"
    $paidRowsHtml += "<td>$(To-HtmlSafe $r.FriendlyName)</td>"
    $paidRowsHtml += "<td>$(To-HtmlSafe $r.SkuPartNumber)</td>"
    $paidRowsHtml += "<td>$(To-HtmlSafe $r.SkuId)</td>"
    $paidRowsHtml += "<td>$(To-HtmlSafe (Format-Qty $r.ConsumedUnits))</td>"
    $paidRowsHtml += "<td>$(To-HtmlSafe (Format-Qty $r.EnabledPrepaidUnits))</td>"
    $paidRowsHtml += "<td>`$$costMoStr</td>"
    $paidRowsHtml += "<td>`$$costYrStr</td>"
    $paidRowsHtml += "<td>`$$totalMoStr</td>"
    $paidRowsHtml += "<td>`$$totalYrStr</td>"
    $paidRowsHtml += "</tr>`n"
  }
  else {
    $freeRowsHtml += "<tr>"
    $freeRowsHtml += "<td>$(To-HtmlSafe $r.FriendlyName)</td>"
    $freeRowsHtml += "<td>$(To-HtmlSafe $r.SkuPartNumber)</td>"
    $freeRowsHtml += "<td>$(To-HtmlSafe $r.SkuId)</td>"
    $freeRowsHtml += "<td>$(To-HtmlSafe (Format-Qty $r.ConsumedUnits))</td>"
    $freeRowsHtml += "<td>$(To-HtmlSafe (Format-Qty $r.EnabledPrepaidUnits))</td>"
    $freeRowsHtml += "<td>-</td><td>-</td><td>-</td><td>-</td>"
    $freeRowsHtml += "</tr>`n"
  }
}

$totalTableCostMoStr = '{0:N2}' -f $totalTableCostMo
$totalTableCostYrStr = '{0:N2}' -f $totalTableCostYr

$topHtml = ""
foreach ($t in $topSkus) {
  # Use single-quoted literal and concatenation to avoid needing to escape double-quotes
  $topHtml += '<div class="top-item"><div>' + (To-HtmlSafe $t.FriendlyName) + ' <span class="muted small">(' + (To-HtmlSafe $t.SkuPartNumber) + ')</span></div><div><strong>' + (To-HtmlSafe (Format-Qty $t.ConsumedUnits)) + '</strong></div></div>' + "`n"
}

$renewalsHtml = ""
if ($renewalsParsed.Count -gt 0) {
  # Add a search box and make the table identifiable for client-side scripting
  $renewalsHtml += '<div style="margin-bottom:8px"><input type="search" id="renewalsSearch" data-target="renewalsTable" class="table-search" placeholder="Search renewals..." /></div>'
  $renewalsHtml += '<table id="renewalsTable" class="sortable searchable"><thead><tr><th>License Name</th><th>Commerce Sub ID</th><th>Trial/Paid</th><th>Status</th><th>Total Licenses</th><th>Renewal Date</th><th>Days Until Renewal</th></tr></thead><tbody>' + "`n"
  foreach ($rw in ($renewalsParsed | Sort-Object ParsedDate)) {
    $daysUntil = ''
    if ($rw.ParsedDate) { $daysUntil = ([int]([math]::Floor(($rw.ParsedDate - $today).TotalDays))) }
    if ($daysUntil -ne '') { $daysDisplay = $daysUntil -lt 0 ? ("Past: $([math]::Abs($daysUntil))d") : ("$daysUntil d") } else { $daysDisplay = '' }
    $dateDisplay = ''
    if ($rw.ParsedDate) { $dateDisplay = $rw.ParsedDate.ToString('MMMM dd, yyyy') }
    
    $isPaid = $false
    # Check if SKU is paid
    if ($rw.SkuPartNumber -and $pricingTable.ContainsKey($rw.SkuPartNumber) -and $pricingTable[$rw.SkuPartNumber] -gt 0) {
      $isPaid = $true
    }

    $rowHtml = "<tr><td>$(To-HtmlSafe $rw.FriendlyName)</td><td>$(To-HtmlSafe $rw.CommerceSubscriptionId)</td><td>$(To-HtmlSafe $rw.IsTrial)</td><td>$(To-HtmlSafe $rw.SubscriptionStatus)</td><td>$(To-HtmlSafe $rw.TotalLicenses)</td><td>$(To-HtmlSafe $dateDisplay)</td><td>$(To-HtmlSafe $daysDisplay)</td></tr>`n"
    if ($isPaid) {
      $renewalsHtml += $rowHtml
    }
    else {
      $freeRenewalsHtml += $rowHtml
    }
  }
  $renewalsHtml += "</tbody></table>`n"
  
  if ($freeRenewalsHtml) {
    $renewalsHtml += '<details style="margin-top:16px"><summary><h4 style="margin:6px 0; display:inline">Free / Zero-Cost Renewals</h4></summary>'
    $renewalsHtml += '<table id="renewalsFreeTable" class="sortable"><thead><tr><th>License Name</th><th>Commerce Sub ID</th><th>Trial/Paid</th><th>Status</th><th>Total Licenses</th><th>Renewal Date</th><th>Days Until Renewal</th></tr></thead><tbody>' + "`n"
    $renewalsHtml += $freeRenewalsHtml
    $renewalsHtml += '</tbody></table></details>'
  }
}
else {
  $renewalsHtml = '<div class="small muted">No renewal data found at ' + (To-HtmlSafe $RenewalCsvPath) + '</div>'
}

# Build Assignments HTML (per-user rows and per-sku summary)
$assignmentsHtml = ""
if ($assigned -and $assigned.Count -gt 0) {
  if ($licensed -and $licensed.Count -gt 0) {
    $assignmentsHtml += '<div style="margin-bottom:8px"><input type="search" id="assignmentsSearch" data-target="assignmentsUserTable" class="table-search" placeholder="Search users/licenses" /></div>'
    $assignmentsHtml += '<h4 style="margin:6px 0">Per-user Assigned Licenses</h4>'
    $assignmentsHtml += '<table id="assignmentsUserTable" class="sortable searchable"><thead><tr><th>Display Name</th><th>Email</th><th>Account Status</th><th>Account Created</th><th>Last Sign-in</th><th>Monthly Cost</th><th>Licenses (Friendly)</th><th>Assigned Via</th><th>Comments</th></tr></thead><tbody>' + "`n"
    foreach ($a in ($licensed | Sort-Object UPN)) {
      $assignmentsHtml += "<tr>"
      $assignmentsHtml += "<td>$(To-HtmlSafe $a.DisplayName)</td>"
      $assignmentsHtml += "<td>$(To-HtmlSafe $a.UPN)</td>"
      $assignmentsHtml += "<td>$(To-HtmlSafe $a.AccountStatus)</td>"
      $assignmentsHtml += "<td>$(To-HtmlSafe (Format-Date $a.AccountCreatedTime))</td>"
      $assignmentsHtml += "<td>$(To-HtmlSafe (Format-Date $a.LastSignInDate))</td>"
      
      $userCost = 0
      if ($a.AssignedSkus -and $a.AssignedSkus -ne 'None') {
        $parts = $a.AssignedSkus -split ' \| '
        foreach ($p in $parts) {
          $pClean = $p.Trim()
          if ($pricingTable.ContainsKey($pClean)) {
            $userCost += $pricingTable[$pClean]
          }
        }
      }
      $userCostStr = '{0:N2}' -f $userCost
      $assignmentsHtml += "<td>`$$userCostStr</td>"

      $licensesHtml = (($a.AssignedFriendlyNames -split '\|') | ForEach-Object { To-HtmlSafe $_ }) -join '<br>'
      $assignmentsHtml += "<td>$licensesHtml</td>"
      $assignmentsHtml += "<td>$(To-HtmlSafe $a.AssignmentDetails)</td>"

      # Logic for Comments column
      $comment = ''
      # Check if disabled and consuming paid license
      if ($a.AccountStatus -match 'Disabled' -and $userCost -gt 0) {
        $comment = 'Disabled user consuming paid license'
      }
      elseif ($a.AccountStatus -eq 'Enabled' -and $userCost -gt 0) {
        # Check inactivity > 30 days
        $lastSi = $a.LastSignInDate
        if ($lastSi -eq 'Unknown') {
          # Do nothing, we don't know the status due to permissions
        }
        elseif ([string]::IsNullOrWhiteSpace($lastSi) -or $lastSi -eq 'N/A') {
          # Never signed in, check created date
          if ($a.AccountCreatedTime -as [DateTime]) {
            $created = [DateTime]$a.AccountCreatedTime
            if ((Get-Date).AddDays(-30) -gt $created) {
              $comment = 'Active user consuming paid license, never signed in'
            }
          }
        }
        else {
          # Check last sign in date
          if ($lastSi -as [DateTime]) {
            $siDate = [DateTime]$lastSi
            if ((Get-Date).AddDays(-30) -gt $siDate) {
              $comment = 'Active user consuming paid license, >30 days since login'
            }
          }
        }
      }
      $assignmentsHtml += "<td>$(To-HtmlSafe $comment)</td>"

      $assignmentsHtml += "</tr>`n"
    }
    $assignmentsHtml += '</tbody></table>'
  }

  # Unlicensed users in collapsed section
  if ($unlicensed -and $unlicensed.Count -gt 0) {
    $assignmentsHtml += '<details style="margin-top:16px"><summary><h4 style="margin:6px 0; display:inline">Unlicensed Users</h4></summary>'
    $assignmentsHtml += '<table id="unlicensedTable" class="sortable"><thead><tr><th>UPN</th><th>Display Name</th><th>Account Status</th><th>Account Created</th><th>Last Sign-in</th></tr></thead><tbody>' + "`n"
    foreach ($u in ($unlicensed | Sort-Object UPN)) {
      $assignmentsHtml += "<tr>"
      $assignmentsHtml += "<td>$(To-HtmlSafe $u.UPN)</td>"
      $assignmentsHtml += "<td>$(To-HtmlSafe $u.DisplayName)</td>"
      $assignmentsHtml += "<td>$(To-HtmlSafe $u.AccountStatus)</td>"
      $assignmentsHtml += "<td>$(To-HtmlSafe (Format-Date $u.AccountCreatedTime))</td>"
      $assignmentsHtml += "<td>$(To-HtmlSafe (Format-Date $u.LastSignInDate))</td>"
      $assignmentsHtml += "</tr>`n"
    }
    $assignmentsHtml += '</tbody></table></details>'
  }
}
else {
  $assignmentsHtml = '<div class="small muted">No assigned license data found at ' + (To-HtmlSafe $assignedPath) + '</div>'
}

$html = @"
<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>License Report</title>
  <style>$css</style>
</head>
<body>
  <div class="container">
    <div class="header">
      <div style="display:flex;align-items:center">
        <div class="logo">
          <!-- logo omitted if $LogoUrl is empty -->
          <img src="$LogoUrl" alt="$companyName Logo" class="logo" style="display:$(if ($LogoUrl) { 'inline-block' } else { 'none' })">
        </div>
        <div>
          <div class="title">Microsoft 365 License Report</div>
          <div class="small">$companyName</div>
        </div>
      </div>
    </div>

    <div class="nav" role="navigation">
      <button data-tab="overview" class="active"><span>📊</span> Overview</button>
      <button data-tab="assignments"><span>👥</span> Assignments</button>
      <button data-tab="licenses"><span>💳</span> Licenses</button>
      <button data-tab="renewals"><span>🔄</span> Renewals</button>
    </div>

    <div id="overview" class="tab-content">
      <div class="overview-grid">
        <div class="overview-card">
          <div class="card-header">
            <div class="card-label-group">
              <div class="label">Est. Monthly Cost</div>
            </div>
            <div class="icon">💰</div>
          </div>
          <div class="value">$displayMonthlyCost</div>
          <div class="sub-value">Based on active seats</div>
        </div>

        <div class="overview-card">
          <div class="card-header">
            <div class="card-label-group">
              <div class="label">Est. Annual Cost</div>
            </div>
            <div class="icon">📅</div>
          </div>
          <div class="value">`$$('{0:N2}' -f ($estMonthlyCost * 12))</div>
          <div class="sub-value">Projected next 12 months</div>
        </div>

        <div class="overview-card">
          <div class="card-header">
            <div class="card-label-group">
              <div class="label">Active Paid Users</div>
            </div>
            <div class="icon">👥</div>
          </div>
          <div class="value">$totalPaidUsersCount</div>
          <div class="sub-value">Users with paid licenses</div>
        </div>

        <div class="overview-card">
          <div class="card-header">
            <div class="card-label-group">
              <div class="label">Next Paid Renewal</div>
              <div class="badge-inline">
                $(if ($nextPaidRenewalDays -ne $null) {
                    $color = if ($nextPaidRenewalDays -le 30) { 'badge-red' } elseif ($nextPaidRenewalDays -le 90) { 'badge-amber' } else { 'badge-green' }
                    "<span class='badge $color'>In $nextPaidRenewalDays days</span>"
                } else {
                    "<span class='badge badge-neutral'>No upcoming renewals</span>"
                })
              </div>
            </div>
            <div class="icon">🔔</div>
          </div>
          <div class="value">$(if ($nextPaidRenewalDate) { $nextPaidRenewalDate.ToString('MMMM dd, yyyy') } else { 'N/A' })</div>
          <div class="sub-value">Upcoming renewal date</div>
        </div>

        <div class="overview-card">
          <div class="card-header">
            <div class="card-label-group">
              <div class="label">Report Generated</div>
            </div>
            <div class="icon">⚙️</div>
          </div>
          <div class="value">$generated</div>
          <div class="sub-value">Data snapshot</div>
        </div>
      </div>
    </div>

    <!-- Users tab removed: user/account metadata is shown in Assignments tab -->

    <div id="assignments" class="tab-content" style="display:none">
      <div class="section">
        <h3 style="margin:0 0 8px 0">Assignments</h3>
        $assignmentsHtml
      </div>
    </div>

    <div id="licenses" class="tab-content" style="display:none">
      <div class="section">
          <h3 style="margin:0 0 8px 0">All Subscribed SKUs</h3>
          <div style="margin-bottom:8px"><input type="search" id="licensesSearch" data-target="licensesTable" class="table-search" placeholder="Search licenses..." /></div>
          <table id="licensesTable" class="sortable searchable">
            <thead>
                <tr>
                    <th>License Name</th>
                    <th>MSFT SKU Name</th>
                    <th>Sku ID</th>
                    <th>Assigned</th>
                    <th>Available</th>
                    <th>Unit Cost (Mo)</th>
                    <th>Unit Cost (Yr)</th>
                    <th>Total Cost (Mo)</th>
                    <th>Total Cost (Yr)</th>
                </tr>
            </thead>
            <tbody>
            $paidRowsHtml
            </tbody>
            <tfoot>
              <tr style="font-weight:bold;background:#fafbfc">
                <td colspan="5" style="text-align:right">Totals:</td>
                <td></td>
                <td></td>
                <td>`$$totalTableCostMoStr</td>
                <td>`$$totalTableCostYrStr</td>
              </tr>
            </tfoot>
          </table>
          
          $(if ($freeRowsHtml) {
            '<details style="margin-top:16px"><summary><h4 style="margin:6px 0; display:inline">Free / Zero-Cost Licenses</h4></summary>' +
            '<table id="licensesFreeTable" class="sortable"><thead><tr><th>License Name</th><th>MSFT SKU Name</th><th>Sku ID</th><th>Assigned</th><th>Available</th><th>Unit Cost</th><th>Total Cost</th></tr></thead><tbody>' +
            $freeRowsHtml +
            '</tbody></table></details>'
          })
        </div>
    </div>

    <div id="renewals" class="tab-content" style="display:none">
      <div class="section">
        <h3 style="margin:0 0 8px 0">Renewal Data</h3>
        $renewalsHtml
      </div>
    </div>

    <div class="footer">
      <p>Created by Ethan Bennett</p>
      <p>Logos provided by <a href="https://logo.dev" target="_blank" style="color:inherit">Logo.dev</a></p>
    </div>
  </div>

  <script>
    function showTab(name) {
      document.querySelectorAll('.tab-content').forEach(function(el){ el.style.display = 'none' })
      var el = document.getElementById(name); if (el) { el.style.display = 'block' }
      document.querySelectorAll('.nav button').forEach(function(b){ b.classList.remove('active') })
      var btn = document.querySelector('.nav button[data-tab="' + name + '"]'); if (btn) { btn.classList.add('active') }
      try { history.replaceState(null, '', '#' + name) } catch(e){}
    }

    // Make tables sortable by clicking headers
    function makeSortable(table) {
      var thead = table.tHead
      if (!thead) return
      var headers = thead.rows[0].cells
      Array.from(headers).forEach(function(th, idx){
        // th.style.cursor = 'pointer' // handled in CSS
        th.addEventListener('click', function(e){
          // Ignore clicks on resizers
          if (e.target.classList.contains('resizer')) return
          var newOrder = th.dataset.order === 'asc' ? 'desc' : 'asc'
          // clear other headers
          Array.from(headers).forEach(function(h){ delete h.dataset.order })
          th.dataset.order = newOrder
          sortTable(table, idx, newOrder)
        })
      })
    }

    function makeResizable(table) {
      var thead = table.tHead
      if (!thead) return
      var ths = Array.from(thead.rows[0].cells)
      ths.forEach(function(th){
        var resizer = document.createElement('div')
        resizer.classList.add('resizer')
        th.appendChild(resizer)
        createResizableColumn(th, resizer)
      })
    }
    
    function createResizableColumn(th, resizer) {
      var x = 0; var w = 0
      var mouseDownHandler = function(e) {
        x = e.clientX
        var styles = window.getComputedStyle(th)
        w = parseInt(styles.width, 10)
        document.addEventListener('mousemove', mouseMoveHandler)
        document.addEventListener('mouseup', mouseUpHandler)
        resizer.classList.add('resizing')
      }
      var mouseMoveHandler = function(e) {
        var dx = e.clientX - x
        th.style.width = (w + dx) + 'px'
      }
      var mouseUpHandler = function() {
        document.removeEventListener('mousemove', mouseMoveHandler)
        document.removeEventListener('mouseup', mouseUpHandler)
        resizer.classList.remove('resizing')
      }
      resizer.addEventListener('mousedown', mouseDownHandler)
    }

    function sortTable(table, colIndex, order) {
      var tbody = table.tBodies[0]
      var rows = Array.from(tbody.rows)
      rows.sort(function(a,b){
        var aText = (a.cells[colIndex] && a.cells[colIndex].textContent) ? a.cells[colIndex].textContent.trim() : ''
        var bText = (b.cells[colIndex] && b.cells[colIndex].textContent) ? b.cells[colIndex].textContent.trim() : ''
        var aNum = parseFloat(aText.replace(/[^0-9\.-]/g,''))
        var bNum = parseFloat(bText.replace(/[^0-9\.-]/g,''))
        var bothNum = !isNaN(aNum) && !isNaN(bNum)
        var cmp = 0
        if (bothNum) { cmp = aNum - bNum } else { cmp = aText.localeCompare(bText, undefined, {numeric:true}) }
        return order === 'asc' ? cmp : -cmp
      })
      rows.forEach(function(r){ tbody.appendChild(r) })
    }

    // Wire up a simple search box to filter table rows
    function attachSearch(input) {
      var targetId = input.dataset.target
      if (!targetId) return
      var table = document.getElementById(targetId)
      if (!table) return
      var tbody = table.tBodies[0]
      input.addEventListener('input', function(){
        var q = input.value.trim().toLowerCase()
        Array.from(tbody.rows).forEach(function(row){
          var text = row.textContent.toLowerCase()
          row.style.display = q === '' ? '' : (text.indexOf(q) === -1 ? 'none' : '')
        })
      })
    }

    window.addEventListener('load', function(){
      var h = location.hash.replace('#','') || 'overview';
      showTab(h);
      document.querySelectorAll('.nav button').forEach(function(b){ b.addEventListener('click', function(){ showTab(this.getAttribute('data-tab')) }) })
      // initialize sortables
      document.querySelectorAll('table.sortable').forEach(function(t){ makeSortable(t); makeResizable(t) })
      // initialize searches
      document.querySelectorAll('.table-search').forEach(function(inp){ attachSearch(inp) })
    })
  </script>

</body>
</html>
"@

$outPath = Join-Path -Path $GlobalWorkingPath -ChildPath 'LicenseReport.html'
$html | Out-File -FilePath $outPath -Encoding utf8
Write-Host "Wrote report to: $outPath"
if ($OpenReport) { Start-Process -FilePath $outPath }