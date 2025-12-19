# Check for the existence of the 'Clients' folder
$clientFolderPath = ".\Output\Clients"

if (Test-Path -Path $clientFolderPath -PathType Container) {
Write-Host "`nThe 'Clients' folder already exists at: $clientFolderPath" -ForegroundColor Green
} else {
    # The folder does not exist, so create it
    Write-Host "`nThe 'Clients' folder does not exist. Creating it now..." -ForegroundColor Yellow
    New-Item -Path $clientFolderPath -ItemType Directory | Out-Null
        Write-Host "`nCreation complete." -ForegroundColor Green
}

if (Test-Path -Path $clientFolderPath -PathType Container) {   
} 
else {
    Write-Host "`nFailed to create the 'Clients' folder at: $clientFolderPath" -ForegroundColor Red
}

if (-not (Get-Module -Name Microsoft.PowerShell.SecretManagement -ListAvailable)) {
    Write-Host "Module not found. Installing now..."
    Install-Module -Name Microsoft.PowerShell.SecretManagement -Force
} else {
    Write-Host "Module is already installed."
}
if (-not (Get-Module -Name Microsoft.PowerShell.SecretStore -ListAvailable)) {
    Write-Host "Module not found. Installing now..."
    Install-Module -Name Microsoft.PowerShell.SecretStore -Force
} else {
    Write-Host "Module is already installed."
}

# Register the Secret Store vault if not already registered
if (-not (Get-SecretVault -Name 'SecretStore' -ErrorAction SilentlyContinue)) {
    Write-Host "Registering Secret Store vault..."
    Register-SecretVault -Name 'SecretStore' -ModuleName 'Microsoft.PowerShell.SecretStore' -DefaultVault

} else {
    Write-Host "Secret Store vault is already registered."
}   
if (-not (Get-Secret -Name 'LogoApiKey' -ErrorAction SilentlyContinue)) {
    if ($env:CI) {
        Write-Host "CI environment detected; skipping LogoApiKey prompt." -ForegroundColor Yellow
    } else {
        try {
            $sec = Read-Host "Enter API key" -AsSecureString
            Set-Secret -Name 'LogoApiKey' -Secret $sec 
        } catch {
            Write-Host "Skipping LogoApiKey setup (no interactive input): $_" -ForegroundColor Yellow
        }
    }
} else {
    Write-Host "API key already stored in Secret Management."
}