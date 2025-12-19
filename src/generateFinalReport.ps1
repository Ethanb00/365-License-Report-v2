param(
  [Parameter(Mandatory = $true)]
  [string]$GlobalWorkingPath
)

Function Get-LicenseCosts {
  # Function to calculate the annual costs of the licenses assigned to a user account  
  [cmdletbinding()]
  Param( [array]$Licenses )
  [int]$Costs = 0
  ForEach ($License in $Licenses) {
    Try {
      [string]$LicenseCost = $PricingHashTable[$License]
      # Convert monthly cost to cents (because some licenses cost sums like 16.40)
      [float]$LicenseCostCents = [float]$LicenseCost * 100
      If ($LicenseCostCents -gt 0) {
        # Compute annual cost for the license
        [float]$AnnualCost = $LicenseCostCents * 12
        # Add to the cumulative license costs
        $Costs = $Costs + ($AnnualCost)
        # Write-Host ("License {0} Cost {1} running total {2}" -f $License, $LicenseCost, $Costs)
      }
    }
    Catch {
      Write-Host ("Error finding license {0} in pricing table - please check" -f $License)
    }
  }
  # Return 
  Return ($Costs / 100)
} 

[datetime]$RunDate = Get-Date
[string]$ReportRunDate = Get-Date ($RunDate) -format 'dd-MMM-yyyy HH:mm'
$Version = "1.95"

# Default currency - can be overwritten by a value read into the $ImportSkus array
[string]$Currency = "USD"

# Connect to the Graph. This connection uses the delegated permissions and roles available to the signed-in user. The
# signed-in account must hold a role like Exchange administrator to access user and group details.
# See https://practical365.com/connect-microsoft-graph-powershell-sdk/ for information about connecting to the Graph.
# In a production environment, it's best to use a registered Entra ID app to connect (app-only mode) to avoid the need for
# the signed-in user to have any administrative roles, like Exchange administrator.
Connect-MgGraph -Scope "Directory.AccessAsUser.All, Directory.Read.All, AuditLog.Read.All" -NoWelcome

# This step depends on the availability of some CSV files generated to hold information about the product licenses used in the tenant and 
# the service plans in those licenses. See https://github.com/12Knocksinna/Office365itpros/blob/master/CreateCSVFilesForSKUsAndServicePlans.PS1 
# for code to generate the CSVs. After the files are created, you need to edit them to add the display names for the SKUs and plans.
# Build Hash of Skus for lookup so that we report user-friendly display names - you need to create these CSV files from SKU and service plan
# data in your tenant.

$clientsRoot = Join-Path $PSScriptRoot "Clients"
$OrgName = (Get-MgOrganization).DisplayName
# Sanitize org name for folder: remove dots, allow only letters, numbers, spaces, hyphens, ampersands
$SafeOrgName = $OrgName -replace '\.', '' -replace '[^a-zA-Z0-9 &\-]', ''
$OrgOutputPath = Join-Path $clientsRoot $SafeOrgName.Trim()
if (-not (Test-Path $OrgOutputPath)) {
  New-Item -Path $OrgOutputPath -ItemType Directory | Out-Null
}
# CSV paths relative to the client/org output directory
$SkuDataPath = "$GlobalWorkingPath\SkuDataComplete.csv"
$ServicePlanPath = "$GlobalWorkingPath\ServicePlans.csv"
$UnlicensedAccounts = 0

# Helper: Show pricing summary and confirm
function Confirm-PricingData {
  param(
    [string]$PricingCsvPath
  )
  $pricing = Import-Csv $PricingCsvPath
  Write-Host "`n💰 Pricing Data for this client:"
  $pricing | Format-Table SkuPartNumber, DisplayName, Price, Currency -AutoSize
  $choice = Read-Host "Is this pricing correct? (Y to continue, N to exit)"
  if ($choice -notmatch '^(Y|y)') {
    Write-Host "Please update the pricing CSV at: $PricingCsvPath and re-run the script."
    exit 1
  }
}

# Define this variable if you want to do cost center reporting based on a cost center stored in one of the
# 15 Exchange Online custom attributes synchronized to Entra ID. Use the Entra ID attribute (like extensionAttribute6) 
# name not the Exchange Online attribute name (CustomAttribute6) Set the variable to $null or don't define it at all 
# to ignore cost centers
#$CostCenterAttribute = "extensionAttribute6"

If ((Test-Path $skuDataPath) -eq $False) {
  Write-Host ("Can't find the product data file ({0}). Exiting..." -f $skuDataPath) ; break 
}
If ((Test-Path $servicePlanPath) -eq $False) {
  Write-Host ("Can't find the serivice plan data file ({0}). Exiting..." -f $servicePlanPath) ; break 
}

# Always confirm pricing if pricing file exists
$PricingPath = Join-Path $OrgOutputPath "Pricing.csv"
if ((Test-Path $PricingPath) -eq $False) {
  Write-Host ("Can't find the pricing file ({0}). Exiting..." -f $PricingPath) ; break
}
Confirm-PricingData -PricingCsvPath $PricingPath

$ImportSkus = Import-CSV $skuDataPath
$ImportPricing = Import-CSV $PricingPath
$ImportServicePlans = Import-CSV $servicePlanPath
# ...existing code...
$SkuHashTable = @{}
ForEach ($Line in $ImportSkus) {
  $key = [string]$Line.SkuId
  if (-not $SkuHashTable.ContainsKey($key)) {
    $SkuHashTable.Add($key, [string]$Line.DisplayName)
  }
}
$ServicePlanHashTable = @{}
ForEach ($Line2 in $ImportServicePlans) {
  $key2 = [string]$Line2.ServicePlanId
  if (-not $ServicePlanHashTable.ContainsKey($key2)) {
    $ServicePlanHashTable.Add($key2, [string]$Line2.ServicePlanDisplayName)
  }
}
# If pricing information is in the $ImportPricing array, we can add the information to the report. We prepare to do this
# by setting the $PricingInfoAvailable to $true and populating the $PricingHashTable
$PricingInfoAvailable = $false
if ($ImportPricing[0].Price) {
  $PricingInfoAvailable = $true
  $Global:PricingHashTable = @{}
  ForEach ($Line in $ImportPricing) {
    $PricingHashTable.Add([string]$Line.SkuPartNumber, [string]$Line.Price)
  }
  if ($ImportPricing[0].Currency) {
    [string]$Currency = ($ImportPricing[0].Currency)
  }
}

# Find tenant accounts - but filtered so that we only fetch those with licenses
Write-Host "Finding licensed user accounts..."
$Users = @()
try {
  $Users = Get-MgUser -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'"  `
    -ConsistencyLevel eventual -CountVariable Records -All -PageSize 999 `
    -Property id, displayName, userPrincipalName, country, department, assignedlicenses, OnPremisesExtensionAttributes, `
    licenseAssignmentStates, createdDateTime, jobTitle, signInActivity, companyName, accountenabled -ErrorAction Stop |  `
    Sort-Object DisplayName
}
catch {
  Write-Host "Warning: Retrying user query without sign-in activity (requires AuditLog.Read.All permission)..." -ForegroundColor Yellow
  $Users = Get-MgUser -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'"  `
    -ConsistencyLevel eventual -CountVariable Records -All -PageSize 999 `
    -Property id, displayName, userPrincipalName, country, department, assignedlicenses, OnPremisesExtensionAttributes, `
    licenseAssignmentStates, createdDateTime, jobTitle, companyName, accountenabled -ErrorAction Stop |  `
    Sort-Object DisplayName
}

If (!($Users)) { 
  Write-Host "No licensed user accounts found - exiting"; break 
}
Else { 
  Write-Host ("{0} Licensed user accounts found - now processing their license data..." -f $Users.Count) 
}

# These are the properties used to create analyses for.
[array]$Departments = $Users.Department | Sort-Object -Unique
[array]$Countries = $Users.Country | Sort-Object -Unique
[array]$CostCenters = $Users.OnPremisesExtensionAttributes.($CostCenterAttribute) | Sort-Object -Unique
[array]$Companies = $Users.CompanyName | Sort-Object -Unique

# Control whether to use the detailed license report information to generate a line-by-line
# report of license assignments to users. This report is useful to detect duplicate licenses and
# to help allocate license costs to operating units within an organization. Set the value to false
# if you don't want to generate the detailed report.
$DetailedCompanyAnalysis = $false

$OrgName = (Get-MgOrganization).DisplayName

# Current subscriptions in the tenant. We use this table to remove expired licenses from the calculation
[array]$CurrentSubscriptions = Get-MgSubscribedSku
$CurrentSubscriptionsHash = @{}
ForEach ($S in $CurrentSubscriptions) {
  $CurrentSubscriptionsHash.Add($S.SkuId, $S.SkuPartNumber) 
}

$DuplicateSKUsAccounts = 0; $DuplicateSKULicenses = 0; $LicenseErrorCount = 0
$Report = [System.Collections.Generic.List[Object]]::new()
$DetailedLicenseReport = [System.Collections.Generic.List[Object]]::new()
$i = 0
[float]$TotalUserLicenseCosts = 0
[float]$TotalBoughtLicenseCosts = 0

ForEach ($User in $Users) {
  $UnusedAccountWarning = "OK"; $i++; $UserCosts = 0
  $ErrorMsg = ""; $LastLicenseChange = ""
  Write-Host ("Processing account {0} {1}/{2}" -f $User.UserPrincipalName, $i, $Users.Count)
  If ([string]::IsNullOrWhiteSpace($User.licenseAssignmentStates) -eq $False) {
    # Only process account if it has some licenses
    [array]$LicenseInfo = $Null; [array]$DisabledPlans = $Null; 
    #  Find out if any of the user's licenses are assigned via group-based licensing
    [array]$GroupAssignments = $User.licenseAssignmentStates | `
      Where-Object { $null -ne $_.AssignedByGroup -and $_.State -eq "Active" }
    #  Find out if any of the user's licenses are assigned via group-based licensing have an error
    [array]$GroupErrorAssignments = $User.licenseAssignmentStates | `
      Where-Object { $Null -ne $_.AssignedByGroup -and $_.State -eq "Error" }
    [array]$GroupLicensing = $Null
    # Find out when the last license change was made
    If ([string]::IsNullOrWhiteSpace($User.licenseAssignmentStates.lastupdateddatetime) -eq $False) {
      $LastLicenseChange = Get-Date(($user.LicenseAssignmentStates.lastupdateddatetime | Measure-Object -Maximum).Maximum) -format g
    }
    # Figure out the details of group-based licensing assignments if any exist
    ForEach ($G in $GroupAssignments) {
      $GroupName = (Get-MgGroup -GroupId $G.AssignedByGroup).DisplayName
      $GroupProductName = $SkuHashTable[$G.SkuId]
      $GroupLicensing += ("{0} assigned from {1}" -f $GroupProductName, $GroupName)
    }
    ForEach ($G in $GroupErrorAssignments) {
      $GroupName = (Get-MgGroup -GroupId $G.AssignedByGroup).DisplayName
      $GroupProductName = $SkuHashTable[$G.SkuId]
      $ErrorMsg = $G.Error
      $LicenseErrorCount++
      $GroupLicensing += ("{0} assigned from {1} BUT ERROR {2}!" -f $GroupProductName, $GroupName, $ErrorMsg)
    }
    $GroupLicensingAssignments = $GroupLicensing -Join ", "

    #  Find out if any of the user's licenses are assigned via direct licensing
    [array]$DirectAssignments = $User.licenseAssignmentStates | `
      Where-Object { $null -eq $_.AssignedByGroup -and $_.State -eq "Active" }

    # Figure out details of direct assigned licenses
    [array]$UserLicenses = $User.AssignedLicenses
    ForEach ($License in $DirectAssignments) {
      If ($SkuHashTable.ContainsKey($License.SkuId) -eq $True) {
        # We found a match in the SKU hash table
        $LicenseInfo += $SkuHashTable.Item($License.SkuId) 
      }
      Else {
        # Nothing found in the SKU hash table, so output the SkuID
        $LicenseInfo += $License.SkuId
      }
    }

    # Report any disabled service plans in licenses
    $License = $UserLicenses | Where-Object { -not [string]::IsNullOrWhiteSpace($_.DisabledPlans) }
    # Check if disabled service plans in a license
    ForEach ($DisabledPlan in $License.DisabledPlans) {
      # Try and find what service plan is disabled
      If ($ServicePlanHashTable.ContainsKey($DisabledPlan) -eq $True) {
        # We found a match in the Service Plans hash table
        $DisabledPlans += $ServicePlanHashTable.Item($DisabledPlan) 
      }
      Else {
        # Nothing doing, so output the Service Plan ID
        $DisabledPlans += $DisabledPlan 
      }
    } # End ForEach disabled plans

    # Detect if any duplicate licenses are assigned (direct and group-based)
    # Build a list of assigned SKUs
    $SkuUserReport = [System.Collections.Generic.List[Object]]::new()
    ForEach ($S in $DirectAssignments) {
      If ($CurrentSubscriptionsHash[$S.SkuId]) {
        $ReportLine = [PSCustomObject][Ordered]@{ 
          User       = $User.Id
          Name       = $User.DisplayName 
          Sku        = $S.SkuId
          Method     = "Direct"  
          Country    = $User.Country
          Department = $User.Department
          Company    = $User.CompanyName
        }
        $SkuUserReport.Add($ReportLine)
      }
    }
    ForEach ($S in $GroupAssignments) {
      If ($CurrentSubscriptionsHash[$S.SkuId]) {
        $ReportLine = [PSCustomObject][Ordered]@{ 
          User       = $User.Id
          Name       = $User.DisplayName
          Sku        = $S.SkuId
          Method     = "Group" 
          Country    = $User.Country
          Department = $User.Department
          Company    = $User.CompanyName
        }
        $SkuUserReport.Add($ReportLine)
      }
    }

    # Check if any duplicates exist
    [array]$DuplicateSkus = $SkuUserReport | Group-Object Sku | `
      Where-Object { $_.Count -gt 1 } | Select-Object -ExpandProperty Name

    # If duplicates exist, resolve their SKU IDs into Product names and generate a warning for the report
    [string]$DuplicateWarningReport = "N/A"
    If ($DuplicateSkus) {
      [array]$DuplicateSkuNames = $Null
      $DuplicateSKUsAccounts++
      $DuplicateSKULicenses = $DuplicateSKULicenses + $DuplicateSKUs.Count
      ForEach ($DS in $DuplicateSkus) {
        $SkuName = $SkuHashTable[$DS]
        $DuplicateSkuNames += $SkuName
      }
      $DuplicateWarningReport = ("Warning: Duplicate licenses detected for: {0}" -f ($DuplicateSkuNames -join ", "))
    }
  }
  Else { 
    $UnlicensedAccounts++
  }
  # Figure out the last time the account signed in. This is important for detecting unused accounts
  $LastSignIn = $User.SignInActivity.LastSignInDateTime
  $LastNonInteractiveSignIn = $User.SignInActivity.LastNonInteractiveSignInDateTime

  if (-not $LastSignIn -and -not $LastNonInteractiveSignIn) {
    $DaysSinceLastSignIn = "Unknown"
    $DaysSinceLastSignInInt = $null
    $UnusedAccountWarning = ("Unknown last sign-in for account")
    $LastAccess = $null
  }
  else {
    # Get the newest date, if both dates contain values
    if ($LastSignIn -and $LastNonInteractiveSignIn) {
      if ($LastSignIn -gt $LastNonInteractiveSignIn) {
        $CompareDate = $LastSignIn
      }
      else {
        $CompareDate = $LastNonInteractiveSignIn
      }
    }
    elseif ($LastSignIn) {
      # Only $LastSignIn has a value
      $CompareDate = $LastSignIn
    }
    else {
      # Only $LastNonInteractiveSignIn has a value
      $CompareDate = $LastNonInteractiveSignIn
    }

    $DaysSinceLastSignInInt = ($RunDate - $CompareDate).Days
    $DaysSinceLastSignIn = "$DaysSinceLastSignInInt"
    $LastAccess = Get-Date($CompareDate) -format g
    if ($DaysSinceLastSignInInt -gt 60) {
      $UnusedAccountWarning = ("Account unused for {0} days - check!" -f $DaysSinceLastSignInInt)
    }
  }

  $AccountCreatedDate = $null
  If ($User.CreatedDateTime) {
    $AccountCreatedDate = Get-Date($User.CreatedDateTime) -format 'dd-MMM-yyyy HH:mm' 
  }

  # If cost center reporting is enabled, extract the cost center for the user
  [string]$CostCenter = $Null
  If ($CostCenterAttribute) {
    $CostCenter = $User.OnPremisesExtensionAttributes.($CostCenterAttribute)
  }

  # Report information
  [string]$DisabledPlans = $DisabledPlans -join ", " 
  [string]$LicenseInfo = $LicenseInfo -join (", ")

  If ($User.AccountEnabled -eq $False) {
    $AccountStatus = "Disabled" 
  }
  Else {
    $AccountStatus = "Enabled"
  }

  If ($PricingInfoAvailable) { 
    # Output report line with pricing info
    # Map SkuId to SkuPartNumber for pricing lookup
    $UserLicensePartNumbers = @()
    foreach ($skuId in $UserLicenses.SkuId) {
      $skuObj = $ImportSkus | Where-Object { $_.SkuId -eq $skuId }
      if ($skuObj) {
        $UserLicensePartNumbers += $skuObj.SkuPartNumber
      }
      else {
        $UserLicensePartNumbers += $skuId  # fallback, will not match pricing
      }
    }
    [float]$UserCosts = Get-LicenseCosts -Licenses $UserLicensePartNumbers
    $TotalUserLicenseCosts = $TotalUserLicenseCosts + $UserCosts
    $ReportLine = [PSCustomObject][Ordered]@{  
      User                       = $User.DisplayName
      UPN                        = $User.UserPrincipalName
      Country                    = $User.Country
      Department                 = $User.Department
      Title                      = $User.JobTitle
      Company                    = $User.companyName
      "Direct assigned licenses" = $LicenseInfo
      "Disabled Plans"           = $DisabledPlans.Trim() 
      "Group based licenses"     = $GroupLicensingAssignments
      "Annual License Costs"     = ("{0} {1}" -f $Currency, ($UserCosts.toString('F2')))
      "Last license change"      = $LastLicenseChange
      "Account created"          = $AccountCreatedDate
      "Last Signin"              = $LastAccess
      "Days since last signin"   = $DaysSinceLastSignIn
      "Duplicates detected"      = $DuplicateWarningReport
      Status                     = $UnusedAccountWarning
      "Account status"           = $AccountStatus
      UserCosts                  = $UserCosts  
      'Cost Center'              = $CostCenter
    }
  }
  Else { 
    # No pricing information
    $ReportLine = [PSCustomObject][Ordered]@{  
      User                       = $User.DisplayName
      UPN                        = $User.UserPrincipalName
      Country                    = $User.Country
      Department                 = $User.Department
      Title                      = $User.JobTitle
      Company                    = $User.companyName
      "Direct assigned licenses" = $LicenseInfo
      "Disabled Plans"           = $DisabledPlans.Trim() 
      "Group based licenses"     = $GroupLicensingAssignments
      "Last license change"      = $LastLicenseChange
      "Account created"          = $AccountCreatedDate
      "Last Signin"              = $LastAccess
      "Days since last signin"   = $DaysSinceLastSignIn
      "Duplicates detected"      = $DuplicateWarningReport
      Status                     = $UnusedAccountWarning
      "Account status"           = $AccountStatus
    }
  }  
  $Report.Add($ReportLine)

  # Populate the detailed license assignment report
  $SkuUserReport = $SkuUserReport | Sort-Object Sku -Unique
  ForEach ($Item in $SkuUserReport) {
    $SkuReportLine = [PSCustomObject][Ordered]@{  
      User       = $Item.User
      Name       = $Item.name
      Sku        = $Item.Sku
      SkuName    = ($SkuHashTable[$Item.Sku])
      Method     = $Item.Method
      Country    = $Item.Country
      Department = $Item.Department
      Company    = $Item.Company    
    }
    $DetailedLicenseReport.Add($SkuReportLine)
  }
} # End ForEach Users

$UnderusedAccounts = $Report | Where-Object { $_.Status -ne "OK" }
$PercentUnderusedAccounts = ($UnderUsedAccounts.Count / $Report.Count).toString("P")

# This code grabs the SKU summary for the tenant and uses the data to create a SKU summary usage segment for the HTML report

# --- Product License Distribution Table with Annualized Pricing ---
$SkuReport = [System.Collections.Generic.List[Object]]::new()
[array]$SkuSummary = Get-MgSubscribedSku | Select-Object SkuId, SkuPartNumber, ConsumedUnits, PrepaidUnits
$SkuSummary = $SkuSummary | Where-Object { $_.ConsumedUnits -ne 0 }
ForEach ($S in $SkuSummary) {
  $SkuDisplayName = $SkuHashtable[$S.SkuId]
  $skuPart = $S.SkuPartNumber
  $pricingRow = $ImportPricing | Where-Object { $_.SkuPartNumber -eq $skuPart }
  $monthly = 0
  if ($pricingRow) {
    $tryPrice = $pricingRow.Price
    if ($tryPrice -match '^[0-9.]+$') {
      $monthly = [float]$tryPrice
    }
  }
  $annual = $monthly * 12
  $unitsPurchased = $S.PrepaidUnits.Enabled
  $percentAssigned = if ($unitsPurchased -gt 0) {
    [math]::Round(($S.ConsumedUnits / $unitsPurchased) * 100, 1)
  }
  else {
    $null
  }
  $annualTotal = $annual * $unitsPurchased
  $SkuReportLine = [PSCustomObject][Ordered]@{
    "SKU Id"                                    = $S.SkuId
    "SKU Name"                                  = $SkuDisplayName
    "Units used"                                = $S.ConsumedUnits
    "Units purchased"                           = $unitsPurchased
    "Percent bought licenses assigned to users" = if ($unitsPurchased -gt 0) { "$percentAssigned%" } else { "N/A" }
    "Annual licensing cost"                     = if ($annual -gt 0) { ("{0} {1:N2}" -f $Currency, $annualTotal) } else { "N/A" }
  }
  $SkuReport.Add($SkuReportLine)
}
$SkuReport = $SkuReport | Sort-Object "Annual licensing cost" -Descending

If ($PricingInfoAvailable) {
  $AverageCostPerUser = ($TotalUserLicenseCosts / $Users.Count)
  $AverageCostPerUserOutput = ("{0} {1}" -f $Currency, ('{0:N2}' -f $AverageCostPerUser))
  $TotalUserLicenseCostsOutput = ("{0} {1}" -f $Currency, ('{0:N2}' -f $TotalUserLicenseCosts))
  $TotalBoughtLicenseCostsOutput = ("{0} {1}" -f $Currency, ('{0:N2}' -f $TotalBoughtLicenseCosts))
  if ($TotalBoughtLicenseCosts -gt 0) {
    $PercentBoughtLicensesUsed = [math]::Round(($TotalUserLicenseCosts / $TotalBoughtLicenseCosts) * 100, 1).ToString() + "%"
  }
  else {
    $PercentBoughtLicensesUsed = "N/A"
  }
}
else {
  $SkuReport = $SkuReport | Sort-Object "SKU Name" -Descending
}

# Generate the department analysis
$DepartmentReport = [System.Collections.Generic.List[Object]]::new()
ForEach ($Department in $Departments) {
  [array]$DepartmentRecords = $Report | Where-Object { $_.Department -eq $Department }
  $DepartmentReportLine = [PSCustomObject][Ordered]@{
    Department  = $Department
    Accounts    = $DepartmentRecords.count
    Costs       = ("{0} {1}" -f $Currency, ('{0:N2}' -f ($DepartmentRecords | Measure-Object UserCosts -Sum).Sum))
    AverageCost = ("{0} {1}" -f $Currency, ('{0:N2}' -f ($DepartmentRecords | Measure-Object UserCosts -Average).Average))
  } 
  $DepartmentReport.Add($DepartmentReportLine)
}
$DepartmentHTML = $DepartmentReport | ConvertTo-HTML -Fragment
# Anyone without a department?
[array]$NoDepartment = $Report | Where-Object { $null -eq $_.Department }
If ($NoDepartment) {
  $NoDepartmentCosts = ("{0} {1}" -f $Currency, ('{0:N2}' -f ($NoDepartment | Measure-Object UserCosts -Sum).Sum))
}
Else {
  $NoDepartmentCosts = "Zero"
}

# Generate the country analysis
$CountryReport = [System.Collections.Generic.List[Object]]::new()
ForEach ($Country in $Countries) {
  [array]$CountryRecords = $Report | Where-Object { $_.Country -eq $Country }
  $CountryReportLine = [PSCustomObject][Ordered]@{
    Country     = $Country
    Accounts    = $CountryRecords.count
    Costs       = ("{0} {1}" -f $Currency, ('{0:N2}' -f ($CountryRecords | Measure-Object UserCosts -Sum).Sum))
    AverageCost = ("{0} {1}" -f $Currency, ('{0:N2}' -f ($CountryRecords | Measure-Object UserCosts -Average).Average))
  } 
  $CountryReport.Add($CountryReportLine)
}
$CountryHTML = $CountryReport | ConvertTo-HTML -Fragment
# Anyone without a country?
[array]$NoCountry = $Report | Where-Object { $null -eq $_.Country }
If ($NoCountry) {
  $NoCountryCosts = ("{0} {1}" -f $Currency, ('{0:N2}' -f ($NoCountry | Measure-Object UserCosts -Sum).Sum))
}
Else {
  $NoCountryCosts = "Zero"
}

# Generate cost center analysis
If ($PricingInfoAvailable -and $null -ne $CostCenterAttribute) { 
  $CostCenterReport = [System.Collections.Generic.List[Object]]::new()
  ForEach ($CostCenter in $CostCenters) {
    [array]$CostCenterRecords = $Report | Where-Object { $_.'Cost Center' -eq $CostCenter }
    $CostCenterReportLine = [PSCustomObject][Ordered]@{
      'Cost Center' = $CostCenter
      Accounts      = $CostCenterRecords.count
      Costs         = ("{0} {1}" -f $Currency, ('{0:N2}' -f ($CostCenterRecords | Measure-Object UserCosts -Sum).Sum))
      AverageCost   = ("{0} {1}" -f $Currency, ('{0:N2}' -f ($CostCenterRecords | Measure-Object UserCosts -Average).Average))
    } 
    $CostCenterReport.Add($CostCenterReportLine)
  }
  $CostCenterHTML = $CostCenterReport | ConvertTo-HTML -Fragment
  # Anyone without a cost center?
  [array]$NoCostCenter = $Report | Where-Object { $null -eq $_.'Cost Center' }
  If ($NoCostCenter) {
    $NoCostCenterCosts = ("{0} {1}" -f $Currency, ('{0:N2}' -f ($NoCostCenter | Measure-Object UserCosts -Sum).Sum))
  }
  Else {
    $NoCostCenterCosts = "Zero"
  }
}

# Generate the company analysis
$CompanyReport = [System.Collections.Generic.List[Object]]::new()
ForEach ($Company in $Companies) {
  [array]$CompanyRecords = $Report | Where-Object { $_.Company -eq $Company }
  $CompanyReportLine = [PSCustomObject][Ordered]@{
    Company     = $Company
    Accounts    = $CompanyRecords.count
    Costs       = ("{0} {1}" -f $Currency, ('{0:N2}' -f ($CompanyRecords | Measure-Object UserCosts -Sum).Sum))
    AverageCost = ("{0} {1}" -f $Currency, ('{0:N2}' -f ($CompanyRecords | Measure-Object UserCosts -Average).Average))
  }
  $CompanyReport.Add($CompanyReportLine)
}
$CompanyHTML = $CompanyReport | ConvertTo-HTML -Fragment
# Anyone without an assigned company?
[array]$NoCompany = $Report | Where-Object { $null -eq $_.Company }
If ($NoCompany) {
  $NoCompanyCosts = ("{0} {1}" -f $Currency, ('{0:N2}' -f ($NoCompany | Measure-Object UserCosts -Sum).Sum))
}
Else {
  $NoCompanyCosts = "Zero"
}

$CompanyAnalysisHTML = $null
# Detailed company analysis - example of breaking down costs by SKU for each company
ForEach ($Company in $Companies) {
  [array]$CompanyAssignments = $DetailedLicenseReport | Where-Object { $_.Company -eq $Company }
  $CompanyAnalysisHTML = $CompanyAnalysisHTML + ("<h2>Company Analysis: Product Licenses for {0}</h2><p>" -f $Company)
  [array]$Skus = $CompanyAssignments.Sku | Sort-Object -Unique
      
  ForEach ($Sku in $Skus) {
    [float]$AnnualCost = 0; [float]$AnnualCostLicense = 0; $AnnualCostLicenseFormatted = $null
    $SkuHTMLFooter = $null; [float]$AnnualCost = $null; $AnnualCostLicense = $null
    $SkuHeader = ("<h3>{0}</h3>" -f $SkuHashTable[$Sku])
    $AssignedSkus = $CompanyAssignments | Where-Object { $_.Sku -eq $Sku } | Select-Object Sku, Name, SkuName, Country, Department, Company 
    If ($PricingInfoAvailable) {
      $LicenseCostSKU = $PricingHashTable[$Sku]
      [float]$LicenseCostCents = [float]$LicenseCostSKU * 100
      If ($LicenseCostCents -gt 0) {
        # Compute annual cost for the license
        [float]$AnnualCost = $LicenseCostCents * 12
        # Compute cost for this SKU assigned to this company
        $AnnualCostLicense = ($AnnualCost * $AssignedSkus.count) / 100
        $AnnualCostLicenseFormatted = ("{0} {1}" -f $Currency, ('{0:N2}' -f $AnnualCostLicense))
      }
      Else {
        $AnnualCostLicenseFormatted = ("{0} {1}" -f $Currency, ('{0:N2}' -f 0))
      }
    }
    # Report the set of people assigned this SKU      
    $AssignedSkusHTML = $AssignedSkus | ConvertTo-HTML -fragment
    $CompanySKUDetailHTML = $SkuHeader + "<p>" + $AssignedSkusHTML + "<p>"
    $SkuHTMLFooter = ("<p>Annual cost for {0} license(s): {1}</p>" -f $AssignedSKUs.count, $AnnualCostLicenseFormatted)
    $CompanyAnalysisHTML = $CompanyAnalysisHTML + "</p>" + $CompanySKUDetailHTML + $SkuHTMLFooter
  }
  $CompanyAnalysisHTML = "<p>" + $CompanyAnalysisHTML + "</p>" + $CompanyHTMLFooter
} 

# Inactive user accounts - never signed in or last sign-in > 60 days ago, and not disabled
$InactiveUserAccounts = $Report | Where-Object {
  (($_."Days since last signin" -ge 60) -or ($_.'Days since last signin' -eq "Unknown")) -and ($_."Account status" -ne "disabled")
}
$InactiveUserAccountsCost = ("{0} {1}" -f $Currency, ('{0:N2}' -f ($InactiveUserAccounts | Measure-Object UserCosts -Sum).Sum))
$DisabledUserAccounts = $Report | Where-Object { $_."Account status" -eq "disabled" }
$DisabledUserAccountsCost = ("{0} {1}" -f $Currency, ('{0:N2}' -f ($DisabledUserAccounts | Measure-Object UserCosts -Sum).Sum))

# Cost spans for license comparison
$LowCost = $AverageCostPerUser * 0.8
$MediumCost = $AverageCostPerUser

# Create the HTML report
$HtmlHead = "<html>
	  <style>
	  BODY{font-family: Arial; font-size: 8pt;}
	  H1{font-size: 22px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	  H2{font-size: 18px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	  H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	  TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
	  TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
	  TD{border: 1px solid #969595; padding: 5px; }
  TD.disabledaccount{background: #FFC0CB;}
  TD.cost-low{background: #d6f5d6;}
  TD.cost-mid{background: #cce6f7;}
  TD.cost-high{background: #e0e6ff;}
  TD.inactiveaccount{background: #FFB3B3;}
  TD.staleaccount{background: #FFE5B4;}
    TD.duplicatelicenses{background: #F8FF00}
	  </style>
	  <body>
           <div align=center>
           <p><h1>Microsoft 365 Licensing Report</h1></p>
           <p><h2><b>For the " + $Orgname + " tenant</b></h2></p>
           <p><h3>Generated: " + $ReportRunDate + "</h3></p></div>"

If ($PricingInfoAvailable) {
  $HtmlBody1 = $Report | Select-Object User, UPN, Country, Department, Title, Company, "Direct assigned licenses", "Disabled Plans", "Group based licenses", "Annual License Costs", "Last license change", "Account created", "Last Signin", "Days since last signin", "Duplicates detected", Status, "Account status" | ConvertTo-Html -Fragment
  # Create an attribute class to use, name it, and append to the XML table attributes
  # Wrap the HTML fragment in a root node to ensure valid XML parsing
  $HTMLBody1Wrapped = "<root>" + $HTMLBody1 + "</root>"
  [xml]$XML = "<html>$HtmlBody1</html>"
  $TableClass = $XML.CreateAttribute("class")
  $TableClass.Value = "State"
  $XML.html.table.Attributes.Append($TableClass) | Out-Null
  # Conditional formatting for the table rows.  
  # Find the column index for 'Days since last signin' from the header row
  $headerRow = $XML.html.table.tr[0]
  $daysSinceColIdx = -1

  if ($XML.html.table -and $XML.html.table.tr.Count -gt 0) {
    $headerRow = $XML.html.table.tr[0]

    for ($i = 0; $i -lt $headerRow.ChildNodes.Count; $i++) {
      if ($headerRow.ChildNodes[$i].InnerText -eq 'Days since last signin') {
        $daysSinceColIdx = $i
        break
      }
    }
  }
  else {
    Write-Warning "Header row not found in HTML table. Skipping column index detection."
  }

  ForEach ($TableRow in $XML.html.table.SelectNodes("tr")) {
    if ($daysSinceColIdx -ge 0) {
      $tdCount = $TableRow.td.Count
      $rawVal = $null
      if ($tdCount -gt $daysSinceColIdx) {
        $rawVal = $TableRow.td[$daysSinceColIdx].InnerXml
      }
    }
    # each TR becomes a member of class "tablerow"
    $TableRow.SetAttribute("class", "tablerow")
    # Skip header row
    if ($TableRow.td.Count -eq 0) { continue }
    $DaysSinceLastSignIn = $null
    $DaysSinceLastSignInInt = $null
    if ($daysSinceColIdx -ge 0 -and $TableRow.td.Count -gt $daysSinceColIdx) {
      $DaysSinceLastSignInRaw = $TableRow.td[$daysSinceColIdx].InnerXml
      $DaysSinceLastSignIn = $DaysSinceLastSignInRaw -replace '[^\d]', ''
      if ($DaysSinceLastSignIn -match '^-?\d+$') {
        try {
          $DaysSinceLastSignInInt = [int]$DaysSinceLastSignIn
        }
        catch {
          $DaysSinceLastSignInInt = $null
        }
      }
      else {
        $DaysSinceLastSignInInt = $null
      }
    }
    else {
      $DaysSinceLastSignIn = "Unknown"
      $DaysSinceLastSignInInt = $null
    }
    # Level of license cost (neutral color scale)
    Try {
      $UserUPN = $TableRow.td[1]
    }
    Catch { Continue }
    $Cost = $Report.Where{ $_.UPN -eq $UserUPN } | Select-Object -ExpandProperty UserCosts
    # Calculate min/max for neutral scale
    $userCostsArray = $Report | ForEach-Object { $_.UserCosts }
    $minCost = ($userCostsArray | Measure-Object -Minimum).Minimum
    $maxCost = ($userCostsArray | Measure-Object -Maximum).Maximum
    $costRange = [math]::Max($maxCost - $minCost, 1)
    # Assign class
    if ($Cost -le ($minCost + $costRange / 3)) {
      $TableRow.SelectNodes("td")[9].SetAttribute("class", "cost-low")
    }
    elseif ($Cost -ge ($maxCost - $costRange / 3)) {
      $TableRow.SelectNodes("td")[9].SetAttribute("class", "cost-high")
    }
    else {
      $TableRow.SelectNodes("td")[9].SetAttribute("class", "cost-mid")
    }
    # Highlight accounts that haven't signed in for >=60 days (pale red), or >=30 and <60 days (pale orange) using dynamic column index
    if (($TableRow.td) -and ($null -ne $DaysSinceLastSignInInt) -and ($daysSinceColIdx -ge 0)) {
      if ($DaysSinceLastSignInInt -ge 60) {
        $TableRow.SelectNodes("td")[$daysSinceColIdx].SetAttribute("class", "inactiveaccount")
      }
      elseif ($DaysSinceLastSignInInt -ge 30 -and $DaysSinceLastSignInInt -lt 60) {
        $TableRow.SelectNodes("td")[$daysSinceColIdx].SetAttribute("class", "staleaccount")
      }
    }
    # If duplicate licenses are detected
    If (($TableRow.td) -and ([string]$TableRow.td[14] -ne 'N/A')) {
      # tag the TD with the color for duplicate licenses
      # Write-Host "Detected duplicate licenses for $($TableRow.td[1])"
      $TableRow.SelectNodes("td")[14].SetAttribute("class", "duplicatelicenses")
    }
    # If row has the account status set to disabled
    If (($TableRow.td) -and ([string]$TableRow.td[16] -eq 'disabled')) {
      ## tag the TD with the color for a disabled account
      $TableRow.SelectNodes("td")[16].SetAttribute("class", "disabledaccount")
    }
  }
  # Wrap the output table with a div tag
  $HTMLBody1 = [string]::Format('<div class="tablediv">{0}</div>', $XML.OuterXml)
  $HtmlBody1 = $HTMLBody1 + "<p>Report created for: " + $OrgName + "</p><p>Created: " + $ReportRunDate + "<p>" 

  $HtmlBody2 = $SkuReport | Select-Object "SKU Id", "SKU Name", "Units used", "Units purchased", "Annual licensing cost" | ConvertTo-Html -Fragment
  $HtmlSkuSeparator = "<p><h2>Product License Distribution</h2></p>"

  # Load license renewal CSV
  $LicenseRenewalCsvPath = Join-Path $GlobalWorkingPath "LicenseRenewalData.csv"
  if (-not (Test-Path $LicenseRenewalCsvPath)) {
    Write-Warning "License renewal CSV not found: $LicenseRenewalCsvPath"
    $HtmlBody3 = "<p><strong>License renewal data not available.</strong></p>"
  }
  else {
    $LicenseRenewalData = Import-Csv -Path $LicenseRenewalCsvPath

    # Convert to HTML fragment
    $HtmlBody3Raw = $LicenseRenewalData | ConvertTo-Html -Fragment

    # Wrap in XML for styling
    [xml]$HtmlBody3Xml = "<html>$HtmlBody3Raw</html>"
    $TableClass3 = $HtmlBody3Xml.CreateAttribute("class")
    $TableClass3.Value = "LicenseRenewalTable"
    $HtmlBody3Xml.html.table.Attributes.Append($TableClass3) | Out-Null

    # Wrap in div
    $HtmlBody3 = [string]::Format('<div class="tablediv">{0}</div>', $HtmlBody3Xml.OuterXml)
    $HtmlBody3 = "<p><h2>License Renewal Summary</h2></p>" + $HtmlBody3
  }
 



  $HtmlTail = "<p></p>"
}
# Add first set of cost analysis if pricing information is available
If ($PricingInfoAvailable) {
  $HTMLTail = $HTMLTail + "<h2>Licensing Cost Analysis</h2>" +
  "<p>Total licensing cost for tenant:              " + $TotalBoughtLicenseCostsOutput + "</p>" +
  "<p>Total cost for assigned licenses:             " + $TotalUserLicenseCostsOutput + "</p>" +
  "<p>Percent bought licenses assigned to users:    " + $PercentBoughtLicensesUsed + "</p>" +
  "<p>Average licensing cost per user:              " + $AverageCostPerUserOutput + "</p>" +
  "<p><h2>License Costs by Country</h2></p>         " + $CountryHTML +
  "<p>License costs for users without a country:    " + $NoCountryCosts +
  "<p><h2>License Costs by Department</h2></p>      " + $DepartmentHTML +
  "<p>License costs for users without a department: " + $NoDepartmentCosts +
  "<p><h2>License Costs by Company</h2></p>         " + $CompanyHTML +
  "<p>License costs for users without a department: " + $NoCompanyCosts

  If ($DetailedCompanyAnalysis) {
    $HTMLTail = $HTMLTail + $CompanyAnalysisHTML
  }
}

# Add cost center information if we've been asked to generate it
If ($CostCenterAttribute) {
  $HTMLTail = $HtmlTail + "<h2>Cost Center Analysis</h2><p></p>" + $CostCenterHTML + "<p></p>" +
  "<p>License costs for users without a cost center:    " + $NoCostCenterCosts 
}

# Add the second part of the cost analysis if pricing information is available
If ($PricingInfoAvailable) {
  $HTMLTail = $HTMLTail +
  "<p><h2>Inactive User Accounts</h2></p>" +
  "<p>Number of inactive user accounts:             " + $InactiveUserAccounts.Count + "</p>" +
  "<p>Names of inactive accounts:                   " + ($InactiveUserAccounts.User -join ", ") + "</p>" +
  "<p>Cost of inactive user accounts:               " + $InactiveUserAccountsCost + "</p>" +
  "<p><h2>Disabled User Accounts</h2></p>" +
  "<p>Number of disabled accounts:                  " + $DisabledUserAccounts.Count + "</p>" +
  "<p>Names of disabled accounts:                   " + ($DisabledUserAccounts.User -join ", ") + "</p>" +
  "<p>Cost of disabled user accounts:               " + $DisabledUserAccountsCost + "</p>"
}

$HTMLTail = $HTMLTail +
"<p>-----------------------------------------------------------------------------------------------------------------------------</p>" +  
"<p>Number of licensed user accounts found:    " + $Report.Count + "</p>" +
"<p>Number of underused user accounts found:   " + $UnderUsedAccounts.Count + "</p>" +
"<p>Percent underused user accounts:           " + $PercentUnderusedAccounts + "</p>" +
"<p>Accounts detected with duplicate licenses: " + $DuplicateSKUsAccounts + "</p>" +
"<p>Count of duplicate licenses:               " + $DuplicateSKULicenses + "</p>" +
"<p>Count of errors:                           " + $LicenseErrorCount + "</p>" +
"<p>-----------------------------------------------------------------------------------------------------------------------------</p>"


# Build HTML report
$HTMLTail = $HTMLTail + "<p>Microsoft 365 Licensing Report<b> " + $Version + "</b></p>"
$HtmlReportFile = Join-Path $OrgOutputPath "Microsoft 365 Licensing Report.html"
$HtmlReport = $Htmlhead + $Htmlbody1 + $HtmlSkuSeparator + $HtmlBody2 + $HtmlBody3 + $Htmltail
$HtmlReport | Out-File $HtmlReportFile -Encoding UTF8

# Generate output report files
if ($DetailedCompanyAnalysis) {
  $DetailedExcelOutputFile = Join-Path $OrgOutputPath "Detailed Microsoft 365 Licensing Report.xlsx"
  if ($DetailedLicenseReport.Count -gt 0) {
    $DetailedLicenseReport | Export-Excel -Path $DetailedExcelOutputFile `
      -WorksheetName "Detailed Microsoft 365 Licensing" `
      -Title ("Detailed Microsoft 365 Licensing Report {0}" -f (Get-Date -format 'dd-MMM-yyyy')) `
      -TitleBold -TableName "DetailedMicrosoft365LicensingReport"
    Write-Host "Detailed report exported to: $DetailedExcelOutputFile"
  }
  else {
    Write-Warning "Detailed report skipped: no data available."
  }
}
else {
  $CSVOutputFile = Join-Path $OrgOutputPath "Microsoft 365 Licensing Report.CSV"
  $Report | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8
  Write-Host "Basic report exported to: $CSVOutputFile"
}

Write-Host ""
Write-Host "Microsoft 365 Licensing Report complete"
Write-Host "---------------------------------------"
Write-Host ""
Write-Host ("An HTML report is available in {0}" -f $HtmlReportFile)
If ($ExcelGenerated) {
  Write-Host ("An Excel report is available in {0}" -f $ExcelOutputFile)
}
Else {
  Write-Host ("A CSV report is available in {0}" -f $CSVOutputFile)
}