# Microsoft 365 License Report v2

A comprehensive PowerShell-based reporting tool for Microsoft 365 licenses, providing detailed insights into costs, seat assignments, and renewal dates with a beautiful HTML dashboard.

## Features

- **Consolidated Dashboard**: View high-level metrics (Monthly/Annual Cost, Active Paid Users, Next Renewal).
- **Pricing Management**: Integrated GUI to manage per-SKU pricing for accurate cost estimation.
- **License Inactivity Flagging**: Automatically identifies users consuming paid licenses who haven't signed in for 30+ days.
- **Persistent Custom SKU Names**: Save friendly names for new or uncommon SKUs locally until they are officially supported by Microsoft.
- **Renewal Tracking**: Consolidated view of license expiration dates for better budgeting and planning.
- **Logo Integration**: Automatically fetches client logos via Logo.dev (requires API key).

## Prerequisites

- **Global Admin or Reports Reader** access to Microsoft 365 tenant.
- **Entra ID P1 or P2 Licensing**: Required for the tenant to populate `SignInActivity` data (last login dates). Without this, the "Days since last login" and inactivity flagging features will not work.
- **PowerShell 7+** (Recommended)
- **PowerShell Modules**:
  - `Microsoft.Graph`
  - `Microsoft.PowerShell.SecretManagement`
  - `Microsoft.PowerShell.SecretStore`
- **Required Microsoft Graph Scopes**:
  - `User.Read.All`
  - `Organization.Read.All`
  - `Directory.Read.All`
  - `SubscribedSku.Read.All`
  - `Reports.Read.All`
  - `AuditLog.Read.All`

## Setup

1. **Install Modules**:
   ```powershell
   Install-Module Microsoft.Graph -Scope CurrentUser
   Install-Module Microsoft.PowerShell.SecretManagement -Scope CurrentUser
   Install-Module Microsoft.PowerShell.SecretStore -Scope CurrentUser
   ```

2. **Configure Secret Store**:
   The script uses a secure vault to store your Logo.dev API key.
   ```powershell
   Register-SecretVault -Name SecretStore -ModuleName Microsoft.PowerShell.SecretStore -DefaultVault
   Set-Secret -Name 'LogoApiKey' -Secret 'YOUR_LOGO_DEV_TOKEN'
   ```

## Usage

Run the main script from the project root:

```powershell
.\main.ps1
```

### Execution Flow:
1. **Module Check**: Verifies all dependencies are present.
2. **Authentication**: Prompts for Microsoft Graph login.
3. **Data Collection**: Fetches users, licenses, and renewal data.
4. **Pricing GUI**: A window will appear allowing you to enter or update the monthly cost for each SKU detected in the tenant.
5. **Report Generation**: An HTML report is generated and saved in the `Output\Clients\[ClientName]\[Date]` directory.

## File Structure

- `main.ps1`: The orchestrator script.
- `src/`: Core logic components:
  - `collectUserDetails.ps1`: Primary data gathering.
  - `managePricing.ps1`: GUI for cost management.
  - `generateHTMLReport.ps1`: Build the interactive dashboard.
  - `createLicensingCSV.ps1`: Processes SKU mappings.
- `data/`:
  - `skus.csv`: Official Microsoft SKU list (automatically updated).
  - `CustomSkuNames.csv`: User-defined friendly names for SKUs.
- `Output/`: Historically organized reports and CSV exports.

## Attribution
- Built by Ethan Bennett.
- Logos provided by [Logo.dev](https://logo.dev).
