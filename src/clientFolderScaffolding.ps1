function Select-Or-Create-Folder {
    # --- Configuration ---
    # Set the root directory where the folders are located
    $RootDirectory = ".\Output\Clients" #

    if (-not (Test-Path -Path $RootDirectory -PathType Container)) {
        Write-Host "Error: The specified root directory '$RootDirectory' does not exist." -ForegroundColor Red
        return
    }

    # --- Step 1: Get Existing Folders ---
    # FIX: Cast the result to [array] to ensure it is always treated as a collection, 
    # even when only one folder is returned.
    [array]$Folders = Get-ChildItem -Path $RootDirectory -Directory | Select-Object -ExpandProperty Name
    $FolderCount = $Folders.Count

    # --- Step 2: Loop for Menu Display and Selection ---
    while ($true) {
        Write-Host ""
        Write-Host "====================================" -ForegroundColor Cyan
        Write-Host "      Select a Project Folder       " -ForegroundColor Cyan
        Write-Host "====================================" -ForegroundColor Cyan
        Write-Host ""
        
        # 1. List Existing Folders
        for ($i = 0; $i -lt $FolderCount; $i++) {
            $Index = $i + 1
            # The indexing here will now correctly grab the full folder name
            Write-Host "  $Index) $($Folders[$i])" 
        }

        # 2. Add 'Create New' Option
        $CreateNewOption = $FolderCount + 1
        Write-Host "------------------------------------"
        Write-Host "  $CreateNewOption) Create NEW Folder" -ForegroundColor Green
        Write-Host "------------------------------------"
        Write-Host ""

        # --- Step 3: Get User Input with Dynamic Prompt ---
        
        if ($FolderCount -eq 0) {
            # Case 1: Directory is empty (only option is 1)
            $PromptText = "Enter '1' to create a new folder"
        } else {
            # Case 2: Directory has folders (range is 1 to $CreateNewOption)
            $PromptText = "Enter the number of your choice (1-$CreateNewOption)"
        }

        [string]$Selection = Read-Host $PromptText
        # --- Step 4: Validate and Process Input ---
        
        # A. Create New Folder Option
        if ($Selection -eq $CreateNewOption) {
            [string]$NewFolderName = Read-Host "Enter the name for the new folder"
            $NewFolderPath = Join-Path -Path $RootDirectory -ChildPath $NewFolderName

            if (Test-Path -Path $NewFolderPath -PathType Container) {
                Write-Host "Error: A folder named '$NewFolderName' already exists." -ForegroundColor Red
            } else {
                # Create the new folder
                New-Item -Path $NewFolderPath -ItemType Directory | Out-Null
                Write-Host ""
                Write-Host "Successfully created new folder: '$NewFolderName'" -ForegroundColor Green
                
                # Set the selected path and exit the loop
                $SelectedPath = $NewFolderPath
                break
            }
        } 
        
        # B. Select Existing Folder Option
        elseif ($Selection -ge 1 -and $Selection -le $FolderCount) {
            $Index = [int]$Selection - 1
            $SelectedFolderName = $Folders[$Index]
            $SelectedPath = Join-Path -Path $RootDirectory -ChildPath $SelectedFolderName
            
            Write-Host ""
            Write-Host "You selected folder: '$SelectedFolderName'" -ForegroundColor Yellow
            break # Exit the loop after a valid selection
        } 
        
        # C. Invalid Input
        else {
            Write-Host "Invalid selection. Please enter a number between 1 and $CreateNewOption." -ForegroundColor Red
        }
    }
    
    # --- Step 5: Return the Result ---
    return $SelectedPath
}




function Get-Final-Working-Path {
    # 1. Get the user-selected/created project folder path
    $ProjectFolderPath = Select-Or-Create-Folder
    
    if (-not $ProjectFolderPath) {
        # If the user-selection function returned $null (e.g., due to an error)
        return $null 
    }
    
    # 2. Generate the date string for the new folder name
    $DateFolderName = Get-Date -Format "yyyy-MM-dd"
    $DateFolderPath = Join-Path -Path $ProjectFolderPath -ChildPath $DateFolderName
    
    Write-Host "`nAttempting to create/update date subfolder: '$DateFolderName'" -ForegroundColor Yellow

    # 3. Create/Reuse the date folder using -Force
    try {
        # -Force handles the overwrite/reuse requirement
        $NewFolder = New-Item -Path $DateFolderPath -ItemType Directory -Force 
        
        Write-Host "Successfully ensured date folder exists at: $($NewFolder.FullName)" -ForegroundColor Green
        
        # 4. Return the absolute final path
        return $NewFolder.FullName
        
    } catch {
        Write-Host "Error creating date subfolder: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}


# ====================================================================
# SCRIPT EXECUTION
# ====================================================================

# Capture the output of the function into a variable in the main session
$GlobalWorkingPath = Get-Final-Working-Path

if ($GlobalWorkingPath) {
    Write-Host "`n=======================================================" -ForegroundColor Magenta
    Write-Host "SUCCESS: Your final working path is:" -ForegroundColor Magenta
    Write-Host "$GlobalWorkingPath" -ForegroundColor Magenta
    Write-Host "=======================================================" -ForegroundColor Magenta
    
    # Example usage: You can now use $GlobalWorkingPath anywhere else in your script
    # Set-Location -Path $GlobalWorkingPath
    # Copy-Item -Path "C:\Source\File.txt" -Destination $GlobalWorkingPath
} else {
    Write-Host "`nScript terminated or failed to set a final working path." -ForegroundColor Red
}
