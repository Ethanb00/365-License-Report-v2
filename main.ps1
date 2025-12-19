# Initialize script environment
. ".\src\initialization.ps1"

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

# Manage Pricing (GUI)
. ".\src\managePricing.ps1" -GlobalWorkingPath $GlobalWorkingPath

# Create HTML Report
. ".\src\generateHTMLReport.ps1" -GlobalWorkingPath $GlobalWorkingPath