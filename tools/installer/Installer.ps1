<#
.SYNOPSIS
    Installs SecurePassGenerator from local files or downloads from GitHub (supports both public and private repositories).

.DESCRIPTION
    This script provides two installation methods:
    - Offline: Installs from local script file or zip file
    - Online: Downloads from GitHub, extracts it, and creates a desktop shortcut
    
    For online installation, you can choose between the latest stable release or pre-release versions.
    
    The script includes three hash verification modes for security:
    - Default Mode: Verifies hash if available, continues with warning if not
    - Strict Mode: Installation fails if the hash file is missing or the hash doesn't match
    - Disabled Mode: Skips hash verification entirely
    
    The script supports both private and public GitHub repositories:
    - For private repositories: Add your GitHub token to the $GitHubToken variable
    - For public repositories: Leave the $GitHubToken variable empty
    
    The script automatically detects whether a token is provided and adjusts its authentication method accordingly.

.PARAMETER InstallType
    Specifies the installation type: "Offline" or "Online"

.PARAMETER ReleaseType
    Specifies which type of release to download when using online installation: "Latest" (default) or "PreRelease"

.PARAMETER DisableHashVerification
    When specified, disables SHA256 hash verification of downloaded or local zip files.
    By default, hash verification is enabled for security.

.PARAMETER StrictHashVerification
    When specified, requires that a hash file must be present for installation to proceed.
    If hash verification is enabled and this parameter is specified, installation will fail
    if no hash file is found. By default, installation will continue with a warning if no
    hash file is found.

.EXAMPLE
    .\Installer.ps1 -InstallType Online
    Installs the latest stable release from GitHub

.EXAMPLE
    .\Installer.ps1 -InstallType Online -ReleaseType PreRelease
    Installs the most recent pre-release version from GitHub

.EXAMPLE
    .\Installer.ps1 -InstallType Offline
    Installs from local files in the same directory

.EXAMPLE
    .\Installer.ps1 -InstallType Online -DisableHashVerification
    Installs the latest stable release from GitHub without verifying the file hash

.EXAMPLE
    .\Installer.ps1 -InstallType Online -StrictHashVerification
    Installs the latest stable release from GitHub with strict hash verification (fails if no hash file)

.NOTES
    File Name      : Installer.ps1
    Author         : onlyalex1984
    Prerequisite   : PowerShell 5.1 or later
    Version        : 1.1
    
.LINK
    https://github.com/onlyalex1984
#>

# Script parameters
param (
    [Parameter(Mandatory=$true)]
    [ValidateSet("Offline", "Online")]
    [string]$InstallType,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("Latest", "PreRelease")]
    [string]$ReleaseType = "Latest",
    
    [Parameter(Mandatory=$false)]
    [switch]$DisableHashVerification,
    
    [Parameter(Mandatory=$false)]
    [switch]$StrictHashVerification
)

# Application configuration
# Using ShortcutName for both the shortcut and app name for consistency
$ExeName = "SecurePassGenerator.ps1"
$ShortcutName = "SecurePassGenerator"
$ShortcutDescription = "Generate secure passwords with a modern GUI"
$IconName = "securepassgenerator.ico"  # Icon filename to use for the shortcut

# GitHub configuration for online installation
$GitHubUsername = "onlyalex1984"
$RepositoryName = "securepassgenerator-powershell"
$GitHubToken = ""  # Leave empty for public repositories, add token for private repositories
$AssetPattern = "SecurePassGenerator-PS-*.zip"  # Pattern to match asset names
$HashFilePattern = "$AssetPattern.sha256"  # Pattern to match hash files (used for documentation and clarity)

# Hash verification configuration
$EnableHashVerification = -not $DisableHashVerification  # Enable by default, disable if parameter is specified
$StrictHashMode = $StrictHashVerification  # Disabled by default, enable if parameter is specified

# Installation paths
$AppDataPath = [System.Environment]::GetFolderPath('ApplicationData')
$InstallDir = Join-Path -Path $AppDataPath -ChildPath "securepassgenerator-ps"
$DesktopPath = [System.Environment]::GetFolderPath('Desktop')
$LogPath = Join-Path -Path $InstallDir -ChildPath "install_log.txt"
$TempZipPath = Join-Path -Path $env:TEMP -ChildPath "SecurePassGenerator-PS.zip"  # Temporary name, will be updated

# Create log function
function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARNING", "ERROR")]
        [string]$Level = "INFO"
    )
    
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$Timestamp] [$Level] $Message"
    
    # Ensure log directory exists
    if (-not (Test-Path -Path (Split-Path -Path $LogPath -Parent))) {
        New-Item -Path (Split-Path -Path $LogPath -Parent) -ItemType Directory -Force | Out-Null
    }
    
    # Write to log file
    Add-Content -Path $LogPath -Value $LogMessage
    
    # Also write to console (using system default colors)
    switch ($Level) {
        "INFO" { Write-Host $LogMessage }
        "WARNING" { Write-Host $LogMessage }
        "ERROR" { Write-Host $LogMessage }
    }
}

# Function to unblock files downloaded from the internet
function Unblock-FileIfNeeded {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Path
    )
    
    try {
        if (Test-Path -Path $Path) {
            if ((Get-Item -Path $Path) -is [System.IO.FileInfo]) {
                # If it's a file, unblock it directly
                Write-Log "Unblocking file: $Path" -Level "INFO"
                Unblock-File -Path $Path -ErrorAction SilentlyContinue
                return $true
            }
            elseif ((Get-Item -Path $Path) -is [System.IO.DirectoryInfo]) {
                # If it's a directory, unblock all PowerShell scripts (.ps1) and any .exe files recursively
                $FilesToUnblock = @()
                $ExeFiles = Get-ChildItem -Path $Path -Filter "*.exe" -Recurse
                $PsFiles = Get-ChildItem -Path $Path -Filter "*.ps1" -Recurse
                
                $FilesToUnblock += $ExeFiles
                $FilesToUnblock += $PsFiles
                
                if ($FilesToUnblock.Count -gt 0) {
                    Write-Log "Unblocking $($FilesToUnblock.Count) files in: $Path" -Level "INFO"
                    foreach ($File in $FilesToUnblock) {
                        Unblock-File -Path $File.FullName -ErrorAction SilentlyContinue
                    }
                }
                return $true
            }
        }
        else {
            Write-Log "Path not found for unblocking: $Path" -Level "WARNING"
            return $false
        }
    }
    catch {
        Write-Log "Error unblocking file(s): $_" -Level "WARNING"
        return $false
    }
}

# Function to handle unblocking at different stages of the installation process
function Unblock-InstalledFiles {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateSet("BeforeCopy", "AfterCopy", "AfterExtract", "AfterMove")]
        [string]$Stage,
        
        [Parameter(Mandatory = $true)]
        [string]$InstallDirectory,
        
        [Parameter(Mandatory = $false)]
        [string]$SpecificFile = $null,
        
        [Parameter(Mandatory = $false)]
        [string]$SourcePath = $null
    )
    
    Write-Log "Unblocking files at stage: $Stage" -Level "INFO"
    
    switch ($Stage) {
        "BeforeCopy" {
            # Unblock source file before copying
            if ($SourcePath) {
                Write-Log "Unblocking source file before copying: $SourcePath" -Level "INFO"
                Unblock-FileIfNeeded -Path $SourcePath
            }
        }
        "AfterCopy" {
            # Unblock all files in the installation directory after copying
            Unblock-FileIfNeeded -Path $InstallDirectory
        }
        "AfterExtract" {
            # Unblock all files after extracting from zip
            Unblock-FileIfNeeded -Path $InstallDirectory
        }
        "AfterMove" {
            # Unblock specific file after moving from subdirectory
            if ($SpecificFile) {
                $TargetPath = Join-Path -Path $InstallDirectory -ChildPath $SpecificFile
                Unblock-FileIfNeeded -Path $TargetPath
            }
        }
    }
}

# Function to create desktop shortcut
function New-Shortcut {
    param (
        [Parameter(Mandatory = $true)]
        [string]$TargetPath,
        
        [Parameter(Mandatory = $true)]
        [string]$ShortcutPath,
        
        [Parameter(Mandatory = $false)]
        [string]$Description = "SecurePassGenerator",
        
        [Parameter(Mandatory = $false)]
        [string]$Arguments = "",
        
        [Parameter(Mandatory = $false)]
        [string]$WorkingDirectory = "",
        
        [Parameter(Mandatory = $false)]
        [string]$IconLocation = ""
    )
    
    try {
        $WshShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut($ShortcutPath)
        $Shortcut.TargetPath = $TargetPath
        $Shortcut.Description = $Description
        
        if (-not [string]::IsNullOrEmpty($Arguments)) {
            $Shortcut.Arguments = $Arguments
        }
        
        if (-not [string]::IsNullOrEmpty($WorkingDirectory)) {
            $Shortcut.WorkingDirectory = $WorkingDirectory
        }
        
        # Set icon location
        if (-not [string]::IsNullOrEmpty($IconLocation) -and (Test-Path -Path $IconLocation)) {
            $Shortcut.IconLocation = $IconLocation
            Write-Log "Using custom icon: $IconLocation" -Level "INFO"
        }
        # Set PowerShell icon for .ps1 files if no custom icon
        elseif ($TargetPath.EndsWith(".ps1")) {
            $Shortcut.IconLocation = "powershell.exe,0"
        }
        
        $Shortcut.Save()
        Write-Log "Created shortcut at $ShortcutPath" -Level "INFO"
        return $true
    }
    catch {
        Write-Log "Failed to create shortcut: $_" -Level "ERROR"
        return $false
    }
}

# Function to copy icon file if it exists
function Copy-IconFile {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SourceDir,
        
        [Parameter(Mandatory = $true)]
        [string]$DestinationDir
    )
    
    $IconPath = Join-Path -Path $SourceDir -ChildPath $IconName
    $DestIconPath = Join-Path -Path $DestinationDir -ChildPath $IconName
    
    if (Test-Path -Path $IconPath) {
        try {
            Write-Log "Found icon file: $IconPath" -Level "INFO"
            Copy-Item -Path $IconPath -Destination $DestIconPath -Force
            Write-Log "Copied icon file to: $DestIconPath" -Level "INFO"
            return $DestIconPath
        }
        catch {
            Write-Log "Failed to copy icon file: $_" -Level "WARNING"
            return $null
        }
    }
    else {
        Write-Log "No icon file found at: $IconPath" -Level "INFO"
        return $null
    }
}

# Function for offline installation
function Install-OfflineApp {
    Write-Log "Starting offline installation" -Level "INFO"
    
    try {
        # First, prioritize finding the script file directly
        Write-Log "Looking for script file in the current directory..." -Level "INFO"
        $SourceScript = Join-Path -Path $PSScriptRoot -ChildPath $ExeName
        
        # Also look for any PowerShell script in the current directory
        if (-not (Test-Path -Path $SourceScript)) {
            $ScriptFiles = Get-ChildItem -Path $PSScriptRoot -Filter "*.ps1"
            if ($ScriptFiles.Count -gt 0) {
                $SourceScript = $ScriptFiles[0].FullName
                Write-Log "Found script: $($ScriptFiles[0].Name)" -Level "INFO"
            }
        }
        
        # If script found, use it directly (no hash verification needed)
        if (Test-Path -Path $SourceScript) {
            Write-Log "Found script file: $SourceScript" -Level "INFO"
            $SourceType = "ps1"
            $SourcePath = $SourceScript
            
            # Unblock the script file before copying
            Unblock-InstalledFiles -Stage "BeforeCopy" -InstallDirectory $InstallDir -SourcePath $SourcePath
        }
                # If no script file found, look for a zip file
        else {
            Write-Log "No script file found, looking for zip file..." -Level "INFO"
            $ZipFiles = Get-ChildItem -Path $PSScriptRoot -Filter $AssetPattern
            $SourceZip = if ($ZipFiles.Count -gt 0) { $ZipFiles[0].FullName } else { $null }
            
            if (Test-Path -Path $SourceZip) {
                Write-Log "Found zip file: $SourceZip" -Level "INFO"
                $SourceType = "zip"
                $SourcePath = $SourceZip
            }
            else {
                Write-Log "Error: No installation files (script or zip) found in the current directory" -Level "ERROR"
                return $false
            }
        }
        
        # Create destination directory if it doesn't exist
        if (!(Test-Path -Path $InstallDir)) {
            Write-Log "Creating directory: $InstallDir" -Level "INFO"
            New-Item -Path $InstallDir -ItemType Directory -Force | Out-Null
        }
        
        # Process based on source type
        if ($SourceType -eq "ps1") {
            # Copy script and any supporting files to destination
            Write-Log "Copying $ExeName to $InstallDir" -Level "INFO"
            Copy-Item -Path $SourcePath -Destination $InstallDir -Force
            
            # Copy any supporting module files if they exist
            $SourceDir = Split-Path -Parent $SourcePath
            $ModuleFiles = Get-ChildItem -Path $SourceDir -Filter "*.psm1" -ErrorAction SilentlyContinue
            if ($ModuleFiles) {
                Write-Log "Copying supporting module files" -Level "INFO"
                foreach ($Module in $ModuleFiles) {
                    Copy-Item -Path $Module.FullName -Destination $InstallDir -Force
                    Write-Log "Copied $($Module.Name)" -Level "INFO"
                }
            }
            
            # Copy LICENSE file if it exists
            $LicenseFile = Get-ChildItem -Path $SourceDir -Filter "LICENSE*" -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($LicenseFile) {
                Write-Log "Copying LICENSE file" -Level "INFO"
                Copy-Item -Path $LicenseFile.FullName -Destination $InstallDir -Force
                Write-Log "Copied $($LicenseFile.Name)" -Level "INFO"
            }
            
            # Unblock the copied script and module files
            Unblock-InstalledFiles -Stage "AfterCopy" -InstallDirectory $InstallDir
            
            # Copy icon file if it exists
            $IconPath = Copy-IconFile -SourceDir $PSScriptRoot -DestinationDir $InstallDir
            
            # Create shortcut
            $ScriptPath = Join-Path -Path $InstallDir -ChildPath $ExeName
            $ShortcutPath = Join-Path -Path $DesktopPath -ChildPath "$ShortcutName.lnk"
            
            # Create PowerShell shortcut to execute the script
            $PowerShellExe = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
            $Arguments = "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ScriptPath`""
            
            # Only pass IconLocation if an icon file was found
            if ($IconPath -and (Test-Path -Path $IconPath)) {
                $ShortcutCreated = New-Shortcut -TargetPath $PowerShellExe -ShortcutPath $ShortcutPath -Description $ShortcutDescription -Arguments $Arguments -WorkingDirectory $InstallDir -IconLocation $IconPath
            } else {
                $ShortcutCreated = New-Shortcut -TargetPath $PowerShellExe -ShortcutPath $ShortcutPath -Description $ShortcutDescription -Arguments $Arguments -WorkingDirectory $InstallDir
            }
            
            if ($ShortcutCreated) {
                Write-Log "Offline installation completed successfully" -Level "INFO"
                return $true
            }
            else {
                Write-Log "Failed to create shortcut" -Level "ERROR"
                return $false
            }
        }
        elseif ($SourceType -eq "zip") {
            # Log hash verification status for zip file installation
            if ($EnableHashVerification) {
                if ($StrictHashMode) {
                    Write-Log "Hash verification enabled with strict mode for zip file (will fail if hash file is missing)" -Level "INFO"
                } else {
                    Write-Log "Hash verification enabled for zip file (default)" -Level "INFO"
                }
            }
            else {
                Write-Log "Hash verification disabled via command-line parameter" -Level "INFO"
            }
            
            # Verify hash if enabled and hash file exists (only applies to zip files)
            if ($EnableHashVerification) {
                # For offline installation, we check if a hash file exists for the zip file
                $HashFilePath = "$SourcePath.sha256"
                
                # Also look for hash files using the pattern in case the hash file doesn't exactly match the zip name
                if (-not (Test-Path -Path $HashFilePath)) {
                    $HashFiles = Get-ChildItem -Path $PSScriptRoot -Filter $HashFilePattern
                    if ($HashFiles.Count -gt 0) {
                        $HashFilePath = $HashFiles[0].FullName
                        Write-Log "Found hash file using pattern: $HashFilePath" -Level "INFO"
                    }
                }
                
                if (Test-Path -Path $HashFilePath) {
                    Write-Log "Found hash file: $HashFilePath" -Level "INFO"
                    
                    try {
                        # Read the expected hash from the file
                        $ExpectedHash = Get-Content -Path $HashFilePath -Raw
                        $ExpectedHash = $ExpectedHash.Trim()
                        
                        # Extract just the hash part if the file contains filename too
                        if ($ExpectedHash -match "^([a-fA-F0-9]{64})") {
                            $ExpectedHash = $Matches[1]
                        }
                        
                        # Store both lowercase and uppercase versions for comparison and display
                        $ExpectedHashUpper = $ExpectedHash.ToUpper()
                        
                        # Calculate the actual hash of the zip file
                        $ActualHash = Get-FileHash -Path $SourcePath -Algorithm SHA256
                        
                        # Compare the hashes (case-insensitive)
                        if ($ActualHash.Hash -eq $ExpectedHash -or $ActualHash.Hash.ToLower() -eq $ExpectedHash.ToLower()) {
                            Write-Log "Expected SHA256 hash: $ExpectedHashUpper" -Level "INFO"
                            Write-Log "Actual SHA256 hash: $($ActualHash.Hash)" -Level "INFO"
                            Write-Log "Hash verification successful" -Level "INFO"
                        }
                        else {
                            Write-Log "Hash verification failed! The zip file may be corrupted or tampered with." -Level "ERROR"
                            Write-Log "Expected: $ExpectedHashUpper" -Level "ERROR"
                            Write-Log "Actual: $($ActualHash.Hash)" -Level "ERROR"
                            return $false
                        }
                    }
                    catch {
                        Write-Log "Error during hash verification: $_" -Level "WARNING"
                        Write-Log "Continuing installation without hash verification" -Level "WARNING"
                    }
                }
                else {
                    if ($StrictHashMode) {
                        Write-Log "No hash file found at $HashFilePath and strict hash verification is enabled" -Level "ERROR"
                        Write-Log "Installation cannot proceed without hash verification in strict mode" -Level "ERROR"
                        return $false
                    } else {
                        Write-Log "No hash file found at $HashFilePath, skipping verification" -Level "WARNING"
                    }
                }
            }
            else {
                Write-Log "Hash verification disabled, skipping" -Level "INFO"
            }
            
            # Extract zip file to destination
            Write-Log "Extracting zip file to $InstallDir" -Level "INFO"
            
            # Check if destination directory exists and has content
            if (Test-Path -Path $InstallDir) {
                $ExistingFiles = Get-ChildItem -Path $InstallDir
                if ($ExistingFiles.Count -gt 0) {
                    Write-Log "Destination directory not empty, clearing contents..." -Level "INFO"
                    Get-ChildItem -Path $InstallDir -Force | Remove-Item -Recurse -Force
                }
            }
            
            # Extract the zip file
            Expand-Archive -Path $SourcePath -DestinationPath $InstallDir -Force
            Write-Log "Successfully extracted zip file" -Level "INFO"
            
            # Unblock all script files in the installation directory
            Unblock-InstalledFiles -Stage "AfterExtract" -InstallDirectory $InstallDir
            
            # Reorganize files - move script and LICENSE from subdirectory to main directory
            Write-Log "Reorganizing extracted files" -Level "INFO"
            
            # Find subdirectories in the installation directory
            $SubDirs = Get-ChildItem -Path $InstallDir -Directory
            
            if ($SubDirs.Count -gt 0) {
                # Assume the first subdirectory contains our files
                $ExtractedDir = $SubDirs[0].FullName
                Write-Log "Found extracted subdirectory: $ExtractedDir" -Level "INFO"
                
                # Find the script and LICENSE files
                $ScriptFile = Get-ChildItem -Path $ExtractedDir -Filter "*.ps1" | Select-Object -First 1
                $LicenseFile = Get-ChildItem -Path $ExtractedDir -Filter "LICENSE*" | Select-Object -First 1
                
                if ($ScriptFile) {
                    # Move the script to the main directory
                    $TargetScriptPath = Join-Path -Path $InstallDir -ChildPath $ScriptFile.Name
                    Copy-Item -Path $ScriptFile.FullName -Destination $TargetScriptPath -Force
                    Write-Log "Moved script to main directory: $TargetScriptPath" -Level "INFO"
                }
                else {
                    Write-Log "No script file found in subdirectory" -Level "WARNING"
                }
                
                if ($LicenseFile) {
                    # Move the LICENSE file to the main directory
                    $TargetLicensePath = Join-Path -Path $InstallDir -ChildPath $LicenseFile.Name
                    Copy-Item -Path $LicenseFile.FullName -Destination $TargetLicensePath -Force
                    Write-Log "Moved LICENSE file to main directory: $TargetLicensePath" -Level "INFO"
                }
                else {
                    Write-Log "No LICENSE file found in subdirectory" -Level "WARNING"
                }
                
                # Remove the subdirectory and its contents
                Remove-Item -Path $ExtractedDir -Recurse -Force
                Write-Log "Removed subdirectory after moving necessary files" -Level "INFO"
                
                # Unblock the moved script files
                if ($ScriptFile) {
                    Unblock-InstalledFiles -Stage "AfterMove" -InstallDirectory $InstallDir -SpecificFile $ScriptFile.Name
                }
            }
            
            # Copy icon file if it exists
            $IconPath = Copy-IconFile -SourceDir $PSScriptRoot -DestinationDir $InstallDir
            
            # Create shortcut
            $ScriptFiles = Get-ChildItem -Path $InstallDir -Filter "*.ps1" -Recurse
            
            if ($ScriptFiles.Count -eq 0) {
                Write-Log "No PowerShell script found in the extracted files" -Level "ERROR"
                return $false
            }
            
            $MainScript = $ScriptFiles | Where-Object { $_.Name -like "SecurePassGenerator*.ps1" } | Select-Object -First 1
            if ($null -eq $MainScript) {
                $MainScript = $ScriptFiles | Select-Object -First 1
                Write-Log "Using first found script: $($MainScript.FullName)" -Level "WARNING"
            }
            else {
                Write-Log "Found main script: $($MainScript.FullName)" -Level "INFO"
            }
            
            $ShortcutPath = Join-Path -Path $DesktopPath -ChildPath "$ShortcutName.lnk"
            
            # Create PowerShell shortcut to execute the script
            $PowerShellExe = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
            $Arguments = "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$($MainScript.FullName)`""
            $WorkingDirectory = Split-Path -Parent $MainScript.FullName
            
            # Only pass IconLocation if an icon file was found
            if ($IconPath -and (Test-Path -Path $IconPath)) {
                $ShortcutCreated = New-Shortcut -TargetPath $PowerShellExe -ShortcutPath $ShortcutPath -Description $ShortcutDescription -Arguments $Arguments -WorkingDirectory $WorkingDirectory -IconLocation $IconPath
            } else {
                $ShortcutCreated = New-Shortcut -TargetPath $PowerShellExe -ShortcutPath $ShortcutPath -Description $ShortcutDescription -Arguments $Arguments -WorkingDirectory $WorkingDirectory
            }
            
            if ($ShortcutCreated) {
                Write-Log "Offline installation completed successfully" -Level "INFO"
                return $true
            }
            else {
                Write-Log "Failed to create shortcut" -Level "ERROR"
                return $false
            }
        }
        
        return $false
    }
    catch {
        Write-Log "Error during offline installation: $_" -Level "ERROR"
        return $false
    }
}

# Function for online installation
function Install-OnlineApp {
    Write-Log "Starting SecurePassGenerator online installation" -Level "INFO"
    
    # Log hash verification status for online installation
    if ($EnableHashVerification) {
        if ($StrictHashMode) {
            Write-Log "Hash verification enabled with strict mode (will fail if hash file is missing)" -Level "INFO"
        } else {
            Write-Log "Hash verification enabled (default)" -Level "INFO"
        }
    }
    else {
        Write-Log "Hash verification disabled via command-line parameter" -Level "INFO"
    }
    
    # Create installation directory if it doesn't exist
    if (-not (Test-Path -Path $InstallDir)) {
        try {
            New-Item -Path $InstallDir -ItemType Directory -Force | Out-Null
            Write-Log "Created installation directory: $InstallDir" -Level "INFO"
        }
        catch {
            Write-Log "Failed to create installation directory: $_" -Level "ERROR"
            return $false
        }
    }
    
    # Step 1: Get the latest release information
    Write-Log "Fetching latest release information from GitHub API" -Level "INFO"
    try {
        # Initialize headers
        $Headers = @{
            "Accept" = "application/vnd.github+json"
            "X-GitHub-Api-Version" = "2022-11-28"
        }
        
        # Add authorization header if token is provided
        if (-not [string]::IsNullOrWhiteSpace($GitHubToken)) {
            Write-Log "Using GitHub token for authentication" -Level "INFO"
            $Headers["Authorization"] = "token $GitHubToken"
        } else {
            Write-Log "No GitHub token provided, accessing as public repository" -Level "INFO"
        }
        
        # Determine which API endpoint to use based on ReleaseType
        if ($ReleaseType -eq "Latest") {
            $ReleaseUrl = "https://api.github.com/repos/$GitHubUsername/$RepositoryName/releases/latest"
            Write-Log "Using latest release API URL: $ReleaseUrl" -Level "INFO"
            
            $ReleaseInfo = Invoke-RestMethod -Uri $ReleaseUrl -Headers $Headers -Method Get
            Write-Log "Successfully retrieved latest release information" -Level "INFO"
        }
        else {
            # For pre-releases, we need to get all releases and filter
            $ReleasesUrl = "https://api.github.com/repos/$GitHubUsername/$RepositoryName/releases"
            Write-Log "Using all releases API URL to find pre-releases: $ReleasesUrl" -Level "INFO"
            
            $AllReleases = Invoke-RestMethod -Uri $ReleasesUrl -Headers $Headers -Method Get
            
            # Filter for pre-releases
            $PreReleases = $AllReleases | Where-Object { $_.prerelease -eq $true }
            
            if ($PreReleases.Count -eq 0) {
                Write-Log "No pre-releases found, falling back to latest release" -Level "WARNING"
                $ReleaseInfo = $AllReleases | Where-Object { $_.prerelease -eq $false } | Select-Object -First 1
                
                if ($null -eq $ReleaseInfo) {
                    Write-Log "No releases found at all" -Level "ERROR"
                    return $false
                }
            }
            else {
                # Get the most recent pre-release
                $ReleaseInfo = $PreReleases | Select-Object -First 1
                Write-Log "Found pre-release: $($ReleaseInfo.name)" -Level "INFO"
            }
        }
        
        # Find the asset matching our pattern
        $Assets = $ReleaseInfo.assets | Where-Object { $_.name -like $AssetPattern }
        
        if ($Assets.Count -eq 0) {
            Write-Log "No assets matching pattern '$AssetPattern' found in the latest release" -Level "ERROR"
            return $false
        }
        
        # Use the latest version if multiple assets match
        $Asset = $Assets | Sort-Object -Property name -Descending | Select-Object -First 1
        $AssetName = $Asset.name
        Write-Log "Selected asset: $AssetName" -Level "INFO"
        
        # Update the temp zip path with the actual asset name
        $TempZipPath = Join-Path -Path $env:TEMP -ChildPath $AssetName
        
        $AssetId = $Asset.id
        Write-Log "Found asset ID: $AssetId" -Level "INFO"
        
        # Step 2: Download the asset
        $DownloadUrl = "https://api.github.com/repos/$GitHubUsername/$RepositoryName/releases/assets/$AssetId"
        Write-Log "Downloading asset from: $DownloadUrl" -Level "INFO"
        
        # Set security protocol to TLS 1.2
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
        
        # Primary download method using WebClient (faster)
        try {
            Write-Log "Downloading asset using WebClient (primary method)" -Level "INFO"
            
            $WebClient = New-Object System.Net.WebClient
            $WebClient.Headers.Add("Accept", "application/octet-stream")
            $WebClient.Headers.Add("User-Agent", "PowerShell Script")
            
            # Add authorization header if token is provided
            if (-not [string]::IsNullOrWhiteSpace($GitHubToken)) {
                $WebClient.Headers.Add("Authorization", "token $GitHubToken")
            }
            
            # Download the file
            $WebClient.DownloadFile($DownloadUrl, $TempZipPath)
            
            if (Test-Path -Path $TempZipPath) {
                $FileSize = (Get-Item -Path $TempZipPath).Length
                Write-Log "Successfully downloaded asset to $TempZipPath (Size: $FileSize bytes)" -Level "INFO"
            } else {
                Write-Log "Download appeared to succeed but file not found at $TempZipPath" -Level "ERROR"
                throw "File not found after download"
            }
        }
        catch {
            Write-Log "Primary download method failed: $_" -Level "WARNING"
            
            # Try alternative download method using Invoke-WebRequest
            try {
                Write-Log "Attempting alternative download method with Invoke-WebRequest..." -Level "INFO"
                
                # Initialize download headers
                $DownloadHeaders = @{
                    "Accept" = "application/octet-stream"
                    "User-Agent" = "PowerShell Script"
                }
                
                # Add authorization header if token is provided
                if (-not [string]::IsNullOrWhiteSpace($GitHubToken)) {
                    $DownloadHeaders["Authorization"] = "token $GitHubToken"
                }
                
                # Use Invoke-WebRequest as fallback
                Invoke-WebRequest -Uri $DownloadUrl -Headers $DownloadHeaders -OutFile $TempZipPath -Method Get
                
                if (Test-Path -Path $TempZipPath) {
                    $FileSize = (Get-Item -Path $TempZipPath).Length
                    Write-Log "Alternative download succeeded to $TempZipPath (Size: $FileSize bytes)" -Level "INFO"
                } else {
                    Write-Log "Alternative download appeared to succeed but file not found" -Level "ERROR"
                    return $false
                }
            }
            catch {
                Write-Log "All download methods failed: $_" -Level "ERROR"
                
                # Try one more approach with direct browser download URL if available
                try {
                    Write-Log "Attempting direct browser download URL approach..." -Level "INFO"
                    
                    # Get the browser download URL from the asset
                    $BrowserDownloadUrl = $Asset.browser_download_url
                    if ($BrowserDownloadUrl) {
                        Write-Log "Using browser_download_url: $BrowserDownloadUrl" -Level "INFO"
                        
                        $DirectWebClient = New-Object System.Net.WebClient
                        $DirectWebClient.Headers.Add("User-Agent", "PowerShell Script")
                        
                        # Add authorization header if token is provided
                        if (-not [string]::IsNullOrWhiteSpace($GitHubToken)) {
                            $DirectWebClient.Headers.Add("Authorization", "token $GitHubToken")
                        }
                        
                        $DirectWebClient.DownloadFile($BrowserDownloadUrl, $TempZipPath)
                        
                        if (Test-Path -Path $TempZipPath) {
                            $FileSize = (Get-Item -Path $TempZipPath).Length
                            Write-Log "Direct URL download succeeded (Size: $FileSize bytes)" -Level "INFO"
                        } else {
                            Write-Log "Direct URL download failed - file not found" -Level "ERROR"
                            return $false
                        }
                    } else {
                        Write-Log "No browser_download_url available" -Level "ERROR"
                        return $false
                    }
                }
                catch {
                    Write-Log "All download approaches failed: $_" -Level "ERROR"
                    return $false
                }
                finally {
                    if ($DirectWebClient) {
                        $DirectWebClient.Dispose()
                    }
                }
            }
        }
        finally {
            if ($WebClient) {
                $WebClient.Dispose()
            }
        }
        
        # Step 2.5: Verify file hash if enabled
        if ($EnableHashVerification) {
            Write-Log "Hash verification enabled, looking for hash file" -Level "INFO"
            
            # Find the hash file matching our asset
            # First try to find hash file with exact asset name
            $HashAssets = $ReleaseInfo.assets | Where-Object { $_.name -like "$AssetName.sha256" }
            
            # If not found, try using the pattern
            if ($HashAssets.Count -eq 0) {
                $HashAssets = $ReleaseInfo.assets | Where-Object { $_.name -like $HashFilePattern }
                if ($HashAssets.Count -gt 0) {
                    Write-Log "Found hash file using pattern: $($HashAssets[0].name)" -Level "INFO"
                }
            }
            
            if ($HashAssets.Count -eq 0) {
                if ($StrictHashMode) {
                    Write-Log "No hash file found for asset '$AssetName' and strict hash verification is enabled" -Level "ERROR"
                    Write-Log "Installation cannot proceed without hash verification in strict mode" -Level "ERROR"
                    return $false
                } else {
                    Write-Log "No hash file found for asset '$AssetName', skipping verification" -Level "WARNING"
                }
            }
            else {
                # Get the hash file
                $HashAsset = $HashAssets | Select-Object -First 1
                $HashAssetId = $HashAsset.id
                $HashDownloadUrl = "https://api.github.com/repos/$GitHubUsername/$RepositoryName/releases/assets/$HashAssetId"
                $TempHashPath = Join-Path -Path $env:TEMP -ChildPath "$AssetName.sha256"
                
                Write-Log "Downloading hash file from: $HashDownloadUrl" -Level "INFO"
                
                try {
                    # Download the hash file
                    $WebClient = New-Object System.Net.WebClient
                    $WebClient.Headers.Add("Accept", "application/octet-stream")
                    $WebClient.Headers.Add("User-Agent", "PowerShell Script")
                    
                    # Add authorization header if token is provided
                    if (-not [string]::IsNullOrWhiteSpace($GitHubToken)) {
                        $WebClient.Headers.Add("Authorization", "token $GitHubToken")
                    }
                    
                    $WebClient.DownloadFile($HashDownloadUrl, $TempHashPath)
                    
                    if (Test-Path -Path $TempHashPath) {
                        # Read the expected hash from the file
                        $ExpectedHash = Get-Content -Path $TempHashPath -Raw
                        $ExpectedHash = $ExpectedHash.Trim()
                        
                        # Extract just the hash part if the file contains filename too
                        if ($ExpectedHash -match "^([a-fA-F0-9]{64})") {
                            $ExpectedHash = $Matches[1]
                        }
                        
                        # Store both lowercase and uppercase versions for comparison and display
                        $ExpectedHashUpper = $ExpectedHash.ToUpper()
                        
                        # Calculate the actual hash of the downloaded file
                        $ActualHash = Get-FileHash -Path $TempZipPath -Algorithm SHA256
                        
                        # Compare the hashes (case-insensitive)
                        if ($ActualHash.Hash -eq $ExpectedHash -or $ActualHash.Hash.ToLower() -eq $ExpectedHash.ToLower()) {
                            Write-Log "Expected SHA256 hash: $ExpectedHashUpper" -Level "INFO"
                            Write-Log "Actual SHA256 hash: $($ActualHash.Hash)" -Level "INFO"
                            Write-Log "Hash verification successful" -Level "INFO"
                        }
                        else {
                            Write-Log "Hash verification failed! The downloaded file may be corrupted or tampered with." -Level "ERROR"
                            Write-Log "Expected: $ExpectedHashUpper" -Level "ERROR"
                            Write-Log "Actual: $($ActualHash.Hash)" -Level "ERROR"
                            return $false
                        }
                    }
                    else {
                        Write-Log "Failed to download hash file, skipping verification" -Level "WARNING"
                    }
                }
                catch {
                    Write-Log "Error during hash verification: $_" -Level "WARNING"
                    Write-Log "Continuing installation without hash verification" -Level "WARNING"
                }
                finally {
                    if ($WebClient) {
                        $WebClient.Dispose()
                    }
                    
                    # Clean up the hash file
                    if (Test-Path -Path $TempHashPath) {
                        Remove-Item -Path $TempHashPath -Force -ErrorAction SilentlyContinue
                    }
                }
            }
        }
        else {
            Write-Log "Hash verification disabled, skipping" -Level "INFO"
        }
        
        # Step 3: Extract the zip file
        Write-Log "Extracting zip file to $InstallDir" -Level "INFO"
        try {
            Expand-Archive -Path $TempZipPath -DestinationPath $InstallDir -Force
            Write-Log "Successfully extracted zip file" -Level "INFO"
            
            # Unblock all script files in the installation directory
            Unblock-InstalledFiles -Stage "AfterExtract" -InstallDirectory $InstallDir
            
            # Step 3.1: Reorganize files - move script and LICENSE from subdirectory to main directory
            Write-Log "Reorganizing extracted files" -Level "INFO"
            
            # Find subdirectories in the installation directory
            $SubDirs = Get-ChildItem -Path $InstallDir -Directory
            
            if ($SubDirs.Count -gt 0) {
                # Assume the first subdirectory contains our files
                $ExtractedDir = $SubDirs[0].FullName
                Write-Log "Found extracted subdirectory: $ExtractedDir" -Level "INFO"
                
                # Find the script and LICENSE files
                $ScriptFile = Get-ChildItem -Path $ExtractedDir -Filter "*.ps1" | Select-Object -First 1
                $LicenseFile = Get-ChildItem -Path $ExtractedDir -Filter "LICENSE*" | Select-Object -First 1
                
                if ($ScriptFile) {
                    # Move the script to the main directory
                    $TargetScriptPath = Join-Path -Path $InstallDir -ChildPath $ScriptFile.Name
                    Copy-Item -Path $ScriptFile.FullName -Destination $TargetScriptPath -Force
                    Write-Log "Moved script to main directory: $TargetScriptPath" -Level "INFO"
                }
                else {
                    Write-Log "No script file found in subdirectory" -Level "WARNING"
                }
                
                if ($LicenseFile) {
                    # Move the LICENSE file to the main directory
                    $TargetLicensePath = Join-Path -Path $InstallDir -ChildPath $LicenseFile.Name
                    Copy-Item -Path $LicenseFile.FullName -Destination $TargetLicensePath -Force
                    Write-Log "Moved LICENSE file to main directory: $TargetLicensePath" -Level "INFO"
                }
                else {
                    Write-Log "No LICENSE file found in subdirectory" -Level "WARNING"
                }
                
                # Remove the subdirectory and its contents
                Remove-Item -Path $ExtractedDir -Recurse -Force
                Write-Log "Removed subdirectory after moving necessary files" -Level "INFO"
            }
            else {
                Write-Log "No subdirectories found in the installation directory" -Level "WARNING"
            }
        }
        catch {
            Write-Log "Failed to extract or reorganize files: $_" -Level "ERROR"
            return $false
        }
        
        # Copy icon file if it exists
        $IconPath = Copy-IconFile -SourceDir $PSScriptRoot -DestinationDir $InstallDir
        
        # Step 4: Create desktop shortcut
        Write-Log "Looking for PowerShell script in $InstallDir" -Level "INFO"
        $ScriptFiles = Get-ChildItem -Path $InstallDir -Filter "*.ps1" -Recurse
        
        if ($ScriptFiles.Count -eq 0) {
            Write-Log "No PowerShell script found in the extracted files" -Level "ERROR"
            return $false
        }
        
        $MainScript = $ScriptFiles | Where-Object { $_.Name -like "SecurePassGenerator*.ps1" } | Select-Object -First 1
        if ($null -eq $MainScript) {
            $MainScript = $ScriptFiles | Select-Object -First 1
            Write-Log "Using first found script: $($MainScript.FullName)" -Level "WARNING"
        }
        else {
            Write-Log "Found main script: $($MainScript.FullName)" -Level "INFO"
        }
        
        $ShortcutPath = Join-Path -Path $DesktopPath -ChildPath "SecurePassGenerator.lnk"
        
        # Create PowerShell shortcut to execute the script
        $PowerShellExe = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
        $Arguments = "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$($MainScript.FullName)`""
        $WorkingDirectory = Split-Path -Parent $MainScript.FullName
        
        # Only pass IconLocation if an icon file was found
        if ($IconPath -and (Test-Path -Path $IconPath)) {
            $ShortcutCreated = New-Shortcut -TargetPath $PowerShellExe -ShortcutPath $ShortcutPath -Description $ShortcutDescription -Arguments $Arguments -WorkingDirectory $WorkingDirectory -IconLocation $IconPath
        } else {
            $ShortcutCreated = New-Shortcut -TargetPath $PowerShellExe -ShortcutPath $ShortcutPath -Description $ShortcutDescription -Arguments $Arguments -WorkingDirectory $WorkingDirectory
        }
        
        if (-not $ShortcutCreated) {
            Write-Log "Failed to create desktop shortcut" -Level "ERROR"
            return $false
        }
        
        # Step 5: Clean up
        Write-Log "Cleaning up temporary files" -Level "INFO"
        try {
            Remove-Item -Path $TempZipPath -Force
            Write-Log "Temporary files removed" -Level "INFO"
        }
        catch {
            Write-Log "Failed to clean up temporary files: $_" -Level "WARNING"
            # Continue despite cleanup failure
        }
        
        Write-Log "Installation completed successfully" -Level "INFO"
        return $true
    }
    catch {
        Write-Log "An error occurred during installation: $_" -Level "ERROR"
        return $false
    }
}

# Initialize log file
Write-Log "======================================================" -Level "INFO"
Write-Log "Installer started with InstallType: $InstallType, ReleaseType: $ReleaseType" -Level "INFO"
Write-Log "Script location: $PSScriptRoot" -Level "INFO"
Write-Log "User: $env:USERNAME" -Level "INFO"
Write-Log "Computer: $env:COMPUTERNAME" -Level "INFO"
Write-Log "======================================================" -Level "INFO"

# Execute installation based on type
$Success = $false

switch ($InstallType) {
    "Offline" {
        $Success = Install-OfflineApp
    }
    "Online" {
        $Success = Install-OnlineApp
    }
}

# Display final message
if ($Success) {
    Write-Host "`nSecurePassGenerator has been successfully installed."
    Write-Host "A shortcut has been created on your desktop."
    exit 0
}
else {
    Write-Host "`nInstallation failed. Please check the log file for details:"
    Write-Host $LogPath
    exit 1
}
