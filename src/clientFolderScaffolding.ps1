<#
.SYNOPSIS
    Client folder selection and date-based folder creation

.DESCRIPTION
    Provides interactive menu for selecting existing client or creating new client folders.
    Creates dated subfolder (YYYY-MM-DD) for organizing daily reports and maintaining history.
    Sets global working path for report output.
    
.OUTPUTS
    Sets $GlobalWorkingPath variable with the dated folder path
    Example: Output\Clients\Contoso Ltd\2026-01-14
#>

# ============================================================================
# HELPER FUNCTION: SELECT OR CREATE CLIENT FOLDER
# ============================================================================
# Interactive menu for client folder selection with new folder creation option

function Select-Or-Create-Folder {

    # ========================================================================
    # Configuration: Root directory for client folders
    # ========================================================================
    # Set the root directory where the folders are located
    $RootDirectory = ".\Output\Clients"

    if (-not (Test-Path -Path $RootDirectory -PathType Container)) {
        Write-Host "Error: The specified root directory '$RootDirectory' does not exist." -ForegroundColor Red
        return
    }

    # ========================================================================
    # Step 1: Retrieve all existing client folders
    # ========================================================================
    # Cast result to [array] to ensure proper collection handling even with single folder
    [array]$Folders = Get-ChildItem -Path $RootDirectory -Directory | Select-Object -ExpandProperty Name
    $FolderCount = $Folders.Count

    # ========================================================================
    # Step 2: Interactive Menu Loop
    # ========================================================================
    # Display menu options and get user selection
    
    while ($true) {
        Write-Host ""
        Write-Host "====================================" -ForegroundColor Cyan
        Write-Host "   Select Client or Create New   " -ForegroundColor Cyan
        Write-Host "====================================" -ForegroundColor Cyan
        Write-Host ""
        
        # Display existing client folders
        for ($i = 0; $i -lt $FolderCount; $i++) {
            $Index = $i + 1
            Write-Host "  [$Index] $($Folders[$i])" 
        }

        # Add 'Create New' option
        $CreateNewOption = $FolderCount + 1
        Write-Host "------------------------------------"
        Write-Host "  [$CreateNewOption] Create NEW Client" -ForegroundColor Green
        Write-Host "------------------------------------"
        Write-Host ""

        # Get user input with dynamic prompt based on folder count
        if ($FolderCount -eq 0) {
            $PromptText = "Enter '1' to create a new client"
        } else {
            $PromptText = "Enter the number of your choice (1-$CreateNewOption)"
        }

        [string]$Selection = Read-Host $PromptText
        
        # ====================================================================
        # Process user selection
        # ====================================================================
        
        # Option 1: Create New Client Folder
        if ($Selection -eq $CreateNewOption) {
            [string]$NewFolderName = Read-Host "Enter the name for the new client folder"
            $NewFolderPath = Join-Path -Path $RootDirectory -ChildPath $NewFolderName

            if (Test-Path -Path $NewFolderPath -PathType Container) {
                Write-Host "Error: A folder named '$NewFolderName' already exists." -ForegroundColor Red
            } else {
                # Create the new folder
                New-Item -Path $NewFolderPath -ItemType Directory | Out-Null
                Write-Host ""
                Write-Host "Successfully created new client folder: '$NewFolderName'" -ForegroundColor Green
                
                # Set the selected path and exit the loop
                $SelectedPath = $NewFolderPath
                break
            }
        } 
        
        # Option 2: Select Existing Client Folder
        elseif ($Selection -ge 1 -and $Selection -le $FolderCount) {
            $Index = [int]$Selection - 1
            $SelectedFolderName = $Folders[$Index]
            $SelectedPath = Join-Path -Path $RootDirectory -ChildPath $SelectedFolderName
            
            Write-Host ""
            Write-Host "Selected client: '$SelectedFolderName'" -ForegroundColor Green
            break
        }
        
        # Option 3: Invalid input
        else {
            Write-Host "Invalid selection. Please enter a number between 1 and $CreateNewOption." -ForegroundColor Red
        }
    }
    
    # Return the selected/created client path



# ============================================================================
# HELPER FUNCTION: GET FINAL WORKING PATH WITH DATE SUBFOLDER
# ============================================================================
# Creates dated subfolder (YYYY-MM-DD) for organizing daily reports

function Get-Final-Working-Path {
    # Step 1: Get the user-selected/created project folder path
    $ProjectFolderPath = Select-Or-Create-Folder
    
    if (-not $ProjectFolderPath) {
        # If the user-selection function returned $null (e.g., due to an error)
        return $null 
    }
    
    # Step 2: Generate the date string for the new subfolder
    # Format: YYYY-MM-DD (allows sorting and easy date identification)
    $DateFolderName = Get-Date -Format "yyyy-MM-dd"
    $DateFolderPath = Join-Path -Path $ProjectFolderPath -ChildPath $DateFolderName
    
    Write-Host "`nCreating dated folder: '$DateFolderName'" -ForegroundColor Cyan

    # Step 3: Create/Reuse the date folder
    try {
        # -Force handles both creation and reuse scenarios
        $NewFolder = New-Item -Path $DateFolderPath -ItemType Directory -Force 
        
        Write-Host "Dated folder ready: $($NewFolder.FullName)" -ForegroundColor Green
        
        # Step 4: Return the absolute final path
        return $NewFolder.FullName
        
    } catch {
        Write-Host "Error creating dated folder: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

# ============================================================================
# SCRIPT EXECUTION: Initialize Global Working Path
# ============================================================================
# Capture the output of the function into a variable in the main session
$GlobalWorkingPath = Get-Final-Working-Path

if ($GlobalWorkingPath) {
    Write-Host "`n=======================================================" -ForegroundColor Magenta
    Write-Host "SUCCESS: Working path established:" -ForegroundColor Magenta
    Write-Host "$GlobalWorkingPath" -ForegroundColor Green
    Write-Host "=======================================================" -ForegroundColor Magenta
} else {
    Write-Host "`nError: Failed to establish working path" -ForegroundColor Red
}
