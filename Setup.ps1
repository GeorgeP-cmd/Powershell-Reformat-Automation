# This program automates the installations of DSS included software from USB drive 
# ("D:\New Computer Setup\") ctrl+f replace
# 
# INCLUDED = Chrome, Firefox, FoxIt, DellCommandUpdate, Office365
# PREREQUISITE = Run as PowerShell admin: **set-executionpolicy remotesigned**
# Reverse: set-executionpolicy restricted
# -----------------------------------------------------------------------------
# Function to check for installer and execute installation
# -----------------------------------------------------------------------------
function Install-Software {
    param(
        [Parameter(Mandatory=$true)]
        [string]$SoftwareName,
        
        [Parameter(Mandatory=$true)]
        [string[]]$InstallerPaths
    )

    Write-Host "--- Checking for **$SoftwareName** installer ---"
    $installerFound = $false

    foreach ($path in $InstallerPaths) {
        if (Test-Path $path) {
            Write-Host "**Installer found** for $SoftwareName at $path. Starting installation..."

            # Determine the execution command based on file extension
            $extension = [System.IO.Path]::GetExtension($path).ToLower()
            $command = ""
            
            switch ($extension) {
                ".msi" {
                    $command = "msiexec.exe"
                }
                ".exe" {
                    
                    $command = $path
                    
                }
                default {
                    Write-Warning "Unsupported installer type: $extension for $SoftwareName. Skipping installation."
                    continue
                }
            }

            try {
                # Use Start-Process for execution and -Wait for synchronous install
                Start-Process -FilePath $command -ErrorAction Stop
                Write-Host "**$SoftwareName installed successfully.**"
                $installerFound = $true
                break # Exit the loop once the installer is found and run
            }
            catch {
                Write-Error "Installation of $SoftwareName failed: $($_.Exception.Message)"
            }
        }
    }
    

    if (-not $installerFound) {
        Write-Warning "**$SoftwareName installer not found** in any specified location."
    }
    Write-Host "----------------------------------------------------`n"
}

# -----------------------------------------------------------------------------
# Dell Command Update Installation Logic
# -----------------------------------------------------------------------------
$dellCommandUpdatePaths = @(
    "D:\New Computer Setup\Dell-Command-Update-Application_8DGG4_WIN_4.4.0_A00.EXE",
    "E:\New Computer Setup\Dell-Command-Update-Application_8DGG4_WIN_4.4.0_A00.EXE"
)

# Prompt the user
$installDellCommandUpdate = Read-Host "Would you like to install Dell Command Update? (y/n)"

if ($installDellCommandUpdate -match '^[yY]') {
    Install-Software -SoftwareName "Dell Command Update" -InstallerPaths $dellCommandUpdatePaths
} else {
    Write-Host "Skipping Dell Command Update installation.`n"
}

# -----------------------------------------------------------------------------
# Additional Software Installations
# -----------------------------------------------------------------------------

## 7zip
$7zipPaths = @(
    "D:\New Computer Setup\7z2404-x64.exe",
    "E:\New Computer Setup\7z2404-x64.exe"
)
Install-Software -SoftwareName "7Zip" -InstallerPaths $7zipPaths

## Microsoft Office
$officePaths = @(
    "D:\New Computer Setup\OfficeSetup.exe",
    "E:\New Computer Setup\OfficeSetup.exe"
)
Install-Software -SoftwareName "Microsoft Office 365" -InstallerPaths $officePaths

## Mozilla Firefox
$firefoxPaths = @(
    "D:\New Computer Setup\Firefox Installer.exe",
    "E:\New Computer Setup\Firefox Installer.exe"
)
Install-Software -SoftwareName "Mozilla Firefox" -InstallerPaths $firefoxPaths

## Google Chrome
# Note: Enterprise MSI is preferred for silent deployment, otherwise /S for .exe.
$chromePaths = @(
    "D:\New Computer Setup\ChromeSetup.exe",
    "E:\New Computer Setup\ChromeSetup.exe"
)
Install-Software -SoftwareName "Google Chrome" -InstallerPaths $chromePaths

## FoxIt PDF Reader
$foxitPaths = @(
    "D:\New Computer Setup\FoxitPDFReader20242_enu_Setup_Prom.exe",
    "E:\New Computer Setup\FoxitPDFReader20242_enu_Setup_Prom.exe"
)
Install-Software -SoftwareName "FoxIt PDF Reader" -InstallerPaths $foxitPaths

# -----------------------------------------------------------------------------
# Company & Location Client Installer Selector
# -----------------------------------------------------------------------------
function Install-CompanyClient {
    Write-Host "--- Company Client Setup ---"
    
    # Define where to look for the "Companies" folders
    # We check D: and E: for a folder named "Companies" (Change this name if needed)
    $possibleRoots = @("D:\New Computer Setup\VSA X Install", "E:\New Computer Setup\VSA X Install") 
    $rootPath = $null

    foreach ($path in $possibleRoots) {
        if (Test-Path $path) {
            $rootPath = $path
            break
        }
    }

    if (-not $rootPath) {
        Write-Warning "Could not find a 'Companies' folder on D: or E:. Skipping."
        return
    }

    # Get list of Companies (Directories in the root)
    $companies = Get-ChildItem -Path $rootPath -Directory

    if ($companies.Count -eq 0) {
        Write-Warning "No company folders found in $rootPath."
        return
    }

    # Display Numbered List of Companies
    Write-Host "Found the following companies:"
    for ($i = 0; $i -lt $companies.Count; $i++) {
        Write-Host "[$($i + 1)] $($companies[$i].Name)"
    }

    # Prompt User for Company
    $validSelection = $false
    $selectedCompany = $null
    while (-not $validSelection) {
        $userinput = Read-Host "Enter the number for the desired Company"
        if ($userinput -match '^\d+$' -and [int]$userinput -gt 0 -and [int]$userinput -le $companies.Count) {
            $selectedCompany = $companies[[int]$userinput - 1]
            $validSelection = $true
        } else {
            Write-Warning "Invalid selection. Please try again."
        }
    }

    # Check for Locations (Sub-directories inside the selected company)
    $locations = Get-ChildItem -Path $selectedCompany.FullName -Directory
    $finalPath = $selectedCompany.FullName

    if ($locations.Count -gt 0) {
        Write-Host "`nMultiple locations found for $($selectedCompany.Name):"
        for ($i = 0; $i -lt $locations.Count; $i++) {
            Write-Host "[$($i + 1)] $($locations[$i].Name)"
        }

        $validLocSelection = $false
        while (-not $validLocSelection) {
            $locInput = Read-Host "Enter the number for the desired Location"
            if ($locInput -match '^\d+$' -and [int]$locInput -gt 0 -and [int]$locInput -le $locations.Count) {
                $finalPath = $locations[[int]$locInput - 1].FullName
                $validLocSelection = $true
            } else {
                Write-Warning "Invalid location selection."
            }
        }
    }

    # Find and Open the Web Shortcut (.url or .lnk)
    Write-Host "`nSearching for shortcuts in: $finalPath" -ForegroundColor Gray
    
    # Open the Grid View Window
    $shortcut = Get-ChildItem -Path "$finalPath\*" -Include *.url, *.lnk -File | 
                Select-Object Name, FullName | 
                Out-GridView -Title "Select installer for $($selectedCompany.Name)" -OutputMode Single

    # Check if the user actually picked something
    if ($shortcut) {
        Write-Host "Launching $($shortcut.Name)..." -ForegroundColor Cyan
        
        # Launch the file
        Invoke-Item -Path $shortcut.FullName
    }
    else {
        Write-Warning "Selection cancelled by user."
    }
}

# Execute the function
Install-CompanyClient

# -----------------------------------------------------------------------------
# Continue with Windows Updates as needed
# -----------------------------------------------------------------------------

## Windows Updates
Write-Host "--- Starting Windows Update Check ---"
if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
    Write-Host "PSWindowsUpdate module not found. Installing now..."
    # Note: Install-Module requires administrator rights
    try {
        Install-Module -Name PSWindowsUpdate -Force -Scope CurrentUser -ErrorAction Stop
        Write-Host "PSWindowsUpdate module installed successfully."
    }
    catch {
        Write-Warning "Could not install PSWindowsUpdate module. Skipping Windows Updates."
    }
}

if (Get-Module -ListAvailable -Name PSWindowsUpdate) {
    Import-Module PSWindowsUpdate
    Write-Host "Running Windows Updates. This may take a while..."
    # Install all available updates and auto-reboot if necessary
    Install-WindowsUpdate -AcceptAll -AutoReboot
    Write-Host "**Windows Update process complete.**"
}
Write-Host "-------------------------------------"
