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
                    $arguments = "/i `"$path`" /qn /norestart"
                }
                ".exe" {
                    # NOTE: Silent install arguments for EXEs can vary widely.
                    # Common ones are /s, /S, /quiet, /q.
                    # The following is a general guess for a silent install.
                    $command = $path
                    $arguments = "" 
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