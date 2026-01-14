<#
.SYNOPSIS
    Environment initialization and dependency management

.DESCRIPTION
    This script performs initial setup for the license reporting tool:
    - Creates required directory structure (Output\Clients)
    - Installs and verifies PowerShell modules
    - Configures SecretStore vault for secure API key storage
    - Prompts for Logo.dev API key if not already configured
    
    Runs automatically when main.ps1 is executed.
#>

# ============================================================================
# DIRECTORY STRUCTURE SETUP
# ============================================================================
# Ensure the Clients output folder exists for multi-client report storage

$clientFolderPath = ".\Output\Clients"

if (Test-Path -Path $clientFolderPath -PathType Container) {
    Write-Host "`nThe 'Clients' folder already exists at: $clientFolderPath" -ForegroundColor Green
} else {
    Write-Host "`nThe 'Clients' folder does not exist. Creating it now..." -ForegroundColor Yellow
    New-Item -Path $clientFolderPath -ItemType Directory | Out-Null
    Write-Host "`nCreation complete." -ForegroundColor Green
}

if (-not (Test-Path -Path $clientFolderPath -PathType Container)) {
    Write-Host "`nFailed to create the 'Clients' folder at: $clientFolderPath" -ForegroundColor Red
    exit 1
}

# ============================================================================
# MODULE INSTALLATION AND VERIFICATION
# ============================================================================
# Install required PowerShell modules if not already present
# All modules are installed to CurrentUser scope to avoid requiring elevation

# Microsoft.PowerShell.SecretManagement - Provides cross-platform secret storage abstraction
if (-not (Get-Module -Name Microsoft.PowerShell.SecretManagement -ListAvailable)) {
    Write-Host "Installing Microsoft.PowerShell.SecretManagement module..." -ForegroundColor Yellow
    Install-Module -Name Microsoft.PowerShell.SecretManagement -Force -Scope CurrentUser
    Write-Host "Installation complete." -ForegroundColor Green
} else {
    Write-Host "Microsoft.PowerShell.SecretManagement module is already installed." -ForegroundColor Green
}

# Microsoft.PowerShell.SecretStore - Vault implementation for SecretManagement
if (-not (Get-Module -Name Microsoft.PowerShell.SecretStore -ListAvailable)) {
    Write-Host "Installing Microsoft.PowerShell.SecretStore module..." -ForegroundColor Yellow
    Install-Module -Name Microsoft.PowerShell.SecretStore -Force -Scope CurrentUser
    Write-Host "Installation complete." -ForegroundColor Green
} else {
    Write-Host "Microsoft.PowerShell.SecretStore module is already installed." -ForegroundColor Green
}

# ImportExcel - Excel file generation without requiring Excel installation
if (-not (Get-Module -Name ImportExcel -ListAvailable)) {
    Write-Host "Installing ImportExcel module..." -ForegroundColor Yellow
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
    Write-Host "Installation complete." -ForegroundColor Green
} else {
    Write-Host "ImportExcel module is already installed." -ForegroundColor Green
}

# ============================================================================
# SECRET VAULT CONFIGURATION
# ============================================================================
# Configure secure storage for API keys and sensitive data
# SecretStore provides encrypted storage protected by the current user

# Register the Secret Store vault if not already registered
if (-not (Get-SecretVault -Name 'SecretStore' -ErrorAction SilentlyContinue)) {
    Write-Host "Registering Secret Store vault..." -ForegroundColor Yellow
    Register-SecretVault -Name 'SecretStore' -ModuleName 'Microsoft.PowerShell.SecretStore' -DefaultVault
    Write-Host "Secret Store vault registered successfully." -ForegroundColor Green
} else {
    Write-Host "Secret Store vault is already registered." -ForegroundColor Green
}

# ============================================================================
# LOGO.DEV API KEY SETUP
# ============================================================================
# Prompt for Logo.dev API key if not already stored
# This is optional - the script will work without it, but logos won't auto-fetch

if (-not (Get-Secret -Name 'LogoApiKey' -ErrorAction SilentlyContinue)) {
    if ($env:CI) {
        # Skip interactive prompt in CI/CD environments
        Write-Host "CI environment detected; skipping LogoApiKey prompt." -ForegroundColor Yellow
    } else {
        try {
            Write-Host "`nLogo.dev API key not found." -ForegroundColor Yellow
            Write-Host "Enter your Logo.dev API key to enable automatic logo fetching." -ForegroundColor Cyan
            Write-Host "(Press Enter to skip - you can manually set logos later)" -ForegroundColor Gray
            $sec = Read-Host "API key" -AsSecureString
            if ($sec.Length -gt 0) {
                Set-Secret -Name 'LogoApiKey' -Secret $sec
                Write-Host "API key stored securely." -ForegroundColor Green
            } else {
                Write-Host "Skipping Logo.dev setup. Logos will need to be manually configured." -ForegroundColor Yellow
            }
        } catch {
            Write-Host "Skipping LogoApiKey setup (no interactive input): $_" -ForegroundColor Yellow
        }
    }
} else {
    Write-Host "Logo.dev API key found in Secret Management." -ForegroundColor Green
}