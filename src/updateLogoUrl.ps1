param(
    [Parameter(Mandatory = $true)]
    [string]$ClientRoot,
    [Parameter(Mandatory = $true)]
    [string]$LogoUrl
)

<#
.SYNOPSIS
Saves a logo URL to a client's LogoUrl.txt file.
This script is designed to be called from the HTML report or as a standalone utility.
.PARAMETER ClientRoot
The client folder path (parent of dated folders)
.PARAMETER LogoUrl
The logo URL to save
#>

$LogoUrlPath = Join-Path -Path $ClientRoot -ChildPath 'LogoUrl.txt'

try {
    if (-not (Test-Path -Path $ClientRoot -PathType Container)) {
        Write-Error "Client folder not found: $ClientRoot"
        exit 1
    }
    
    $LogoUrl.Trim() | Out-File -FilePath $LogoUrlPath -Encoding utf8 -NoNewline -Force
    Write-Host "Logo URL saved successfully to: $LogoUrlPath" -ForegroundColor Green
    Write-Host "New logo URL: $LogoUrl" -ForegroundColor Green
    exit 0
}
catch {
    Write-Error "Failed to save logo URL: $_"
    exit 1
}
