# Microsoft 365 License Report v2

A comprehensive PowerShell-based reporting tool for Microsoft 365 licenses. This tool generates detailed insights into license costs, user assignments, renewal dates, and license utilization with interactive HTML and Excel reports.

## Features

- **Interactive HTML Dashboard**: Modern, responsive dashboard displaying key metrics including monthly/annual costs, active paid users, and upcoming renewals
- **Excel Report Generation**: Automated Excel workbook creation with formatted worksheets for all license data
- **Pricing Management GUI**: User-friendly interface to manage per-SKU pricing for accurate cost estimation and budget forecasting
- **License Inactivity Detection**: Automatically flags users with paid licenses who haven't signed in for 30+ days, helping identify cost-saving opportunities
- **Custom SKU Name Management**: Define and persist friendly names for new, uncommon, or custom SKUs not yet in Microsoft's official catalog
- **Renewal Date Tracking**: Consolidated view of all license subscription renewal dates with warnings for upcoming expirations
- **Client Logo Integration**: Automatic logo fetching via Logo.dev API with support for custom logo URLs per client
- **Historical Data Preservation**: Organized folder structure maintains daily snapshots of license data for trend analysis
- **Skip Data Collection Mode**: Rapidly regenerate reports from existing data without re-querying Microsoft Graph
- **Multi-Client Support**: Easily manage and track licenses for multiple clients with isolated folder structures

## Prerequisites

### Required Permissions
- **Microsoft 365 Admin Access**: Global Admin or Reports Reader role
- **Entra ID P1 or P2 License**: Required for `SignInActivity` data (user last login dates)
  - Without this, the script will still function but "Days since last login" and inactivity detection features will show "Unknown"

### Software Requirements
- **PowerShell 7+** (Recommended for best performance and compatibility)
- **Windows PowerShell 5.1** (Minimum supported version)

### PowerShell Modules
The following modules are automatically installed and configured by the script:
- `Microsoft.Graph` - Microsoft Graph API access
- `Microsoft.PowerShell.SecretManagement` - Secure credential storage
- `Microsoft.PowerShell.SecretStore` - Secret vault implementation
- `ImportExcel` - Excel report generation

### Microsoft Graph API Scopes
The script requires the following delegated permissions:
- `User.Read.All` - Read all user profiles and license assignments
- `Organization.Read.All` - Read organization and domain information
- `Directory.Read.All` - Read directory objects
- `Reports.Read.All` - Read usage reports
- `AuditLog.Read.All` - Read sign-in activity data

### Optional Requirements
- **Logo.dev API Key** (Optional) - For automatic client logo fetching
  - Sign up at [logo.dev](https://logo.dev) for a free or paid API key
  - Without this, you can manually specify logo URLs per client

## Installation and Setup

### 1. Clone or Download the Repository
```powershell
# Clone the repository
git clone <repository-url>
cd 365-License-Report-v2

# Or download and extract the ZIP file
```

### 2. PowerShell Module Installation
The required PowerShell modules will be automatically installed when you first run the script. However, you can manually install them:

```powershell
# Install Microsoft Graph PowerShell SDK
Install-Module Microsoft.Graph -Scope CurrentUser -Force

# Install Secret Management modules for secure API key storage
Install-Module Microsoft.PowerShell.SecretManagement -Scope CurrentUser -Force
Install-Module Microsoft.PowerShell.SecretStore -Scope CurrentUser -Force

# Install ImportExcel for Excel report generation
Install-Module ImportExcel -Scope CurrentUser -Force
```

### 3. Configure Secret Store (Optional - for Logo.dev Integration)
To enable automatic logo fetching, configure your Logo.dev API key:

```powershell
# Register the secret vault (one-time setup)
Register-SecretVault -Name SecretStore -ModuleName Microsoft.PowerShell.SecretStore -DefaultVault

# Store your Logo.dev API key securely
Set-Secret -Name 'LogoApiKey' -Secret 'YOUR_LOGO_DEV_TOKEN'
```

**Note**: If you skip this step, the script will prompt you to enter the API key on first run, or you can run without logo fetching.

## Usage

### Basic Usage
Run the script from the project root directory:

```powershell
.\main.ps1
```

### Command-Line Parameters

#### Skip Data Collection Mode
Regenerate reports using existing CSV data without re-querying Microsoft Graph:

```powershell
# Interactive client selection
.\main.ps1 -SkipDataCollection

# Specify client by name
.\main.ps1 -SkipDataCollection -ClientName "Contoso Ltd"
```

**Use Cases**:
- Quickly regenerate reports after adjusting pricing data
- Create reports with different formatting without re-fetching user data
- Work offline with previously collected data

#### Skip Pricing GUI
Bypass the pricing management window and use existing pricing data:

```powershell
.\main.ps1 -SkipPricing
```

**Use Cases**:
- Automated/scheduled report generation
- When pricing data is already configured
- Running in non-interactive environments

#### Combined Parameters
```powershell
# Regenerate report for specific client without pricing GUI
.\main.ps1 -SkipDataCollection -ClientName "Contoso Ltd" -SkipPricing
```

### Workflow Overview

#### First Run (Full Data Collection)
1. **Initialization**: Checks for required modules and installs if missing
2. **Authentication**: Prompts for Microsoft Graph login with required scopes
3. **SKU Data Download**: Fetches latest official Microsoft SKU catalog
4. **Client Selection**: Choose existing client folder or create new one
5. **Data Collection**:
   - Retrieves all user profiles with license assignments
   - Fetches subscription renewal dates
   - Collects SKU details and consumption data
   - Gathers sign-in activity (if available)
6. **Pricing Configuration**: GUI opens for entering/updating monthly costs per SKU
7. **Report Generation**:
   - Creates interactive HTML dashboard
   - Generates Excel workbook with formatted sheets
   - Saves all CSVs for future reference

#### Subsequent Runs (Same Client)
- Script remembers your pricing data (stored in client folder)
- Auto-detects existing SKUs and pre-fills pricing
- Creates new dated folder for each run to maintain history
- Only prompts for pricing on new/changed SKUs

### Logo Management

The script automatically retrieves your client's logo from [Logo.dev](https://logo.dev) based on their primary domain. 

#### Changing the Logo

If the auto-fetched logo is incorrect or you want to use a custom image:

1. **Open the Generated Report**: Navigate to `Output\Clients\[ClientName]\[Date]\LicenseReport.html`
2. **Click the Logo**: Hover over the logo in the top-left corner and click (you'll see "Click to edit" hint)
3. **Enter New Logo URL**: 
   - Paste the URL of your desired logo image
   - Use the preview button to verify it displays correctly
4. **Save the Logo**:
   - Click "Save Logo" button in the modal
   - Download the generated `LogoUrl.txt` file
   - Move it to your client's root folder: `Output\Clients\[ClientName]\LogoUrl.txt`
5. **Future Reports**: All subsequent reports for this client will automatically use your custom logo

**Important Notes**:
- Logo URLs are stored per-client, allowing different logos for each client
- The logo URL must be publicly accessible
- Supported formats: PNG, JPG, SVG, GIF
- Recommended dimensions: 200x200px or larger (maintains aspect ratio)

### Custom SKU Names

When the script encounters a SKU not in Microsoft's official catalog, it will prompt you to enter a friendly name:

```
No friendly name found for SkuId: 1234-5678-90ab-cdef
SkuPartNumber: CUSTOM_SKU_NAME
Enter friendly name (or press Enter to leave blank): My Custom License
```

Custom names are saved to `data\CustomSkuNames.csv` and reused in future runs. When Microsoft adds official names for these SKUs, the script automatically cleans up duplicates.

### Pricing Management

**Location**: Pricing data is stored in `Output\Clients\[ClientName]\ClientPricing.csv`

**Why Client-Level Storage?**
- Persists across daily runs (not tied to a specific date folder)
- Each client can have different negotiated pricing
- Easily copy between similar clients

**Updating Pricing**:
1. Run the script normally - the pricing GUI will open
2. Edit the "Monthly Cost ($)" column for each SKU
3. Click "Save Pricing" button
4. Pricing is immediately applied to current and future reports

**Pricing GUI Features**:
- Auto-loads existing pricing data
- Displays all current SKUs plus historical SKUs (even if no longer active)
- Sortable and searchable grid
- Validates numeric input

### Output Structure

```
Output/
└── Clients/
    └── [ClientName]/
        ├── ClientPricing.csv          # Persistent pricing data
        ├── LogoUrl.txt                # Custom logo URL (optional)
        └── [YYYY-MM-DD]/              # Daily snapshot folder
            ├── LicenseReport.html     # Interactive HTML dashboard
            ├── LicenseReport.xlsx     # Excel workbook with all data
            ├── AssignedLicenses.csv   # User-to-license assignments
            ├── AssignedLicenses_Summary.csv  # License usage counts
            ├── LicenseRenewalData.csv # Renewal dates and subscription info
            ├── SubscribedSKUs.csv     # SKU details and consumption
            └── user_details.csv       # User profiles and activity
```


## Project Structure

```
365-License-Report-v2/
│
├── main.ps1                    # Main orchestration script (entry point)
├── README.md                   # This documentation file
│
├── data/                       # Reference data and mappings
│   ├── skus.csv               # Official Microsoft SKU catalog (auto-updated)
│   └── CustomSkuNames.csv     # User-defined friendly names for custom SKUs
│
├── src/                        # Core functionality modules
│   ├── initialization.ps1              # Environment setup and module installation
│   ├── mg-graphConnection.ps1          # Microsoft Graph authentication
│   ├── getSKUcsv.ps1                   # Downloads latest Microsoft SKU catalog
│   ├── clientFolderScaffolding.ps1     # Client folder selection/creation
│   ├── collectUserDetails.ps1          # Fetches user profiles and license data
│   ├── GetRenewalDetail.ps1            # Retrieves subscription renewal information
│   ├── createLicensingCSV.ps1          # Processes SKU data and creates SubscribedSKUs.csv
│   ├── createAssignedLicenses.ps1      # Creates normalized license assignment reports
│   ├── managePricing.ps1               # Pricing management GUI
│   ├── manageLogo.ps1                  # Logo URL storage and retrieval
│   ├── updateLogoUrl.ps1               # Logo URL update functionality
│   ├── generateHTMLReport.ps1          # Creates interactive HTML dashboard
│   ├── generateExcelReport.ps1         # Generates formatted Excel workbook
│   └── generateFinalReport.ps1         # Legacy report generator (deprecated)
│
└── Output/                     # Generated reports and data (created on first run)
    └── Clients/               # Multi-client support structure
        └── [ClientName]/      # Individual client folders
            ├── ClientPricing.csv          # Persistent pricing configuration
            ├── LogoUrl.txt                # Custom logo URL (optional)
            └── [YYYY-MM-DD]/              # Daily data snapshots
                ├── LicenseReport.html     # Interactive dashboard
                ├── LicenseReport.xlsx     # Excel report
                ├── AssignedLicenses.csv   # User assignments
                ├── AssignedLicenses_Summary.csv
                ├── LicenseRenewalData.csv
                ├── SubscribedSKUs.csv
                └── user_details.csv
```

### Key Components

#### main.ps1
The orchestration script that coordinates all modules. Supports command-line parameters for flexible execution modes.

#### src/initialization.ps1
- Verifies and installs required PowerShell modules
- Creates necessary directory structure
- Configures SecretStore vault
- Prompts for Logo.dev API key (if not configured)

#### src/collectUserDetails.ps1
- Queries all Microsoft 365 users via Microsoft Graph
- Retrieves user profiles, departments, job titles
- Fetches license assignments and sign-in activity
- Resolves SKU IDs to friendly names
- Handles both direct and group-based license assignments

#### src/managePricing.ps1
- Displays interactive Windows Forms GUI
- Loads existing pricing from ClientPricing.csv
- Merges current SKUs with historical SKUs
- Validates numeric input
- Persists pricing data to client folder

#### src/generateHTMLReport.ps1
- Creates modern, responsive HTML dashboard
- Calculates cost summaries (monthly/annual)
- Identifies inactive users (30+ days since login)
- Generates renewal calendar
- Embeds interactive JavaScript for filtering and sorting
- Supports custom logos per client

#### src/generateExcelReport.ps1
- Creates multi-sheet Excel workbook
- Applies professional formatting and styling
- Includes conditional formatting for inactive users
- Adds data validation and filtering
- Creates summary pivot tables

## Troubleshooting

### Common Issues

#### "SignInActivity data not available"
**Problem**: Users show "Unknown" for last sign-in dates  
**Solution**: Your tenant requires Entra ID P1 or P2 licensing. The script will continue to work, but without sign-in data. Inactivity detection will be unavailable.

#### "Failed to retrieve subscription data from Graph"
**Problem**: Error when fetching renewal dates  
**Solution**: 
- Ensure you have sufficient permissions (Reports.Read.All)
- Try running as a Global Administrator
- Some trial tenants may not have subscription endpoint access

#### Pricing GUI doesn't appear
**Problem**: Script completes but no pricing window opens  
**Solutions**:
- Check your taskbar - the window might be behind other applications
- Run with `-SkipPricing` parameter to bypass the GUI
- Ensure Windows Forms support is available (standard in Windows)

#### "Get-MgUser: Insufficient privileges"
**Problem**: Authentication succeeds but data collection fails  
**Solution**:
- Disconnect and reconnect: `Disconnect-MgGraph` then rerun script
- Verify you consented to all requested scopes during login
- Check that your account has Global Admin or Reports Reader role

#### Custom SKU names not persisting
**Problem**: Script re-prompts for SKU names on each run  
**Solution**:
- Verify `data\CustomSkuNames.csv` exists and is writable
- Check file permissions on the data directory
- Ensure the CSV format is correct (Id, FriendlyName columns)

#### Logo not displaying in report
**Problem**: Broken image in HTML report  
**Solutions**:
- Verify Logo.dev API key is correctly stored: `Get-Secret -Name 'LogoApiKey'`
- Check that the domain is not `*.onmicrosoft.com` (use custom domain)
- Try manually setting a logo URL using the in-report editor
- Ensure the logo URL is publicly accessible

### Performance Optimization

For tenants with many users (1000+):
- First run may take 5-15 minutes depending on tenant size
- Use `-SkipDataCollection` mode for quick report regeneration
- Consider scheduling the script during off-peak hours
- Sign-in activity queries are the slowest component

### Getting Help

If you encounter issues not covered here:
1. Check the PowerShell error output for specific error messages
2. Verify all prerequisites are met
3. Ensure you're running PowerShell 7+ for best compatibility
4. Review the Microsoft Graph API documentation for permission requirements

## Advanced Usage

### Automated/Scheduled Execution

Create a scheduled task to automatically generate reports:

```powershell
# Create a scheduled task script
$scriptPath = "C:\Path\To\365-License-Report-v2\main.ps1"
$params = "-SkipDataCollection -ClientName 'Contoso Ltd' -SkipPricing"

$action = New-ScheduledTaskAction -Execute "pwsh.exe" -Argument "-File `"$scriptPath`" $params"
$trigger = New-ScheduledTaskTrigger -Daily -At 6AM
Register-ScheduledTask -TaskName "M365 License Report" -Action $action -Trigger $trigger
```

### Batch Processing Multiple Clients

```powershell
# Process multiple clients sequentially
$clients = @("Client A", "Client B", "Client C")
foreach ($client in $clients) {
    .\main.ps1 -ClientName $client -SkipPricing
}
```

### Exporting Data for External Systems

All data is stored in CSV format for easy integration:

```powershell
# Import and analyze data programmatically
$licenses = Import-Csv ".\Output\Clients\Contoso\2026-01-14\AssignedLicenses.csv"
$inactive = $licenses | Where-Object { $_.LastSignInDate -lt (Get-Date).AddDays(-30) }
$inactive | Export-Csv ".\InactiveUsers.csv" -NoTypeInformation
```

## Security Considerations

- **API Keys**: Stored securely using PowerShell SecretManagement module
- **Credentials**: Microsoft Graph uses modern authentication (OAuth 2.0)
- **Data Storage**: All reports stored locally, not transmitted externally
- **Permissions**: Uses least-privilege approach with read-only scopes
- **Audit Trail**: Each run creates timestamped folder for compliance

## Contributing

Contributions are welcome! Please ensure:
- Code follows PowerShell best practices
- Comments are clear and descriptive
- New features include documentation updates
- Test with multiple tenant sizes and configurations

## Version History

### v2.0 (Current)
- Complete rewrite with modular architecture
- Added Excel report generation
- Improved HTML dashboard with modern UI
- Multi-client support with isolated data
- Enhanced error handling and logging
- Custom logo management
- Skip data collection mode for faster iterations

### v1.0 (Legacy)
- Initial HTML report generation
- Basic license data collection
- Simple pricing management

## Attribution
- **Author**: Ethan Bennett
- **Organization**: Dataprise
- **Logo Service**: [Logo.dev](https://logo.dev)
- **License**: Internal Use

## Support

For questions, issues, or feature requests, contact:
- **Internal Support**: Ethan Bennett
- **Documentation**: See README.md and inline code comments
