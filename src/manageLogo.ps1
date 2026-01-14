param(
    [Parameter(Mandatory = $true)]
    [string]$ClientRoot,
    [string]$DefaultLogoUrl = ''
)

<#
.SYNOPSIS
Manages logo URL storage and retrieval for a client.
Stores logo URLs in ClientRoot\LogoUrl.txt for persistence across script runs.
.PARAMETER ClientRoot
The client folder path (parent of dated folders)
.PARAMETER DefaultLogoUrl
The default/auto-generated logo URL to use if no stored URL exists
#>

$LogoUrlPath = Join-Path -Path $ClientRoot -ChildPath 'LogoUrl.txt'

function Get-StoredLogoUrl {
    if (Test-Path -Path $LogoUrlPath) {
        $stored = Get-Content -Path $LogoUrlPath -Raw -ErrorAction SilentlyContinue
        if ($stored -and $stored.Trim()) {
            return $stored.Trim()
        }
    }
    return $DefaultLogoUrl
}

function Set-LogoUrl {
    param([string]$LogoUrl)
    
    if ($LogoUrl -and $LogoUrl.Trim()) {
        $LogoUrl.Trim() | Out-File -FilePath $LogoUrlPath -Encoding utf8 -NoNewline
        return $true
    }
    return $false
}

# Return the logo URL (stored or default)
Get-StoredLogoUrl
