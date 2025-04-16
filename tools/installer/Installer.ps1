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

.EXAMPLE
    .\Installer.ps1 -InstallType DirectAPI -ReleaseType Latest
    Installs the latest stable release using the GitHub API directly (proxy-friendly method)

.EXAMPLE
    .\Installer.ps1 -InstallType DirectAPI -ReleaseType PreRelease
    Installs the latest pre-release version using the GitHub API directly (proxy-friendly method)

.NOTES
    File Name      : Installer.ps1
    Author         : onlyalex1984
    Prerequisite   : PowerShell 5.1 or later
    Version        : 1.0
    Copyright      : (C) 2025 onlyalex1984
    License        : GPL v3 - Full license text available in the project root directory
    
.LINK
    https://github.com/onlyalex1984
#>

#region Script Parameters
param (
    [Parameter(Mandatory=$false)]
    [ValidateSet("Offline", "Online", "DirectAPI")]
    [string]$InstallType,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("Latest", "PreRelease")]
    [string]$ReleaseType = "Latest",
    
    [Parameter(Mandatory=$false)]
    [switch]$DisableHashVerification,
    
    [Parameter(Mandatory=$false)]
    [switch]$StrictHashVerification
)

# Script configuration
$script:Version = "1.0"
#endregion

#region Configuration
# Initialize configuration
function Initialize-InstallerConfig {
    [CmdletBinding()]
    param()
    
    # Application configuration
    $config = @{
        # Application details
        ScriptName = "SecurePassGenerator.ps1"
        ShortcutName = "SecurePassGenerator"
        ShortcutDescription = "Generate secure passwords with a modern GUI"
        IconName = "SecurePassGenerator.ico"  # Icon filename to use for the shortcut
        
        # GitHub configuration for online installation
        GitHubUsername = "onlyalex1984"
        RepositoryName = "securepassgenerator-powershell"
        GitHubToken = ""  # Leave empty for public repositories, add token for private repositories
        AssetPattern = "SecurePassGenerator-PS-*.zip"  # Pattern to match asset names
        HashFilePattern = "SecurePassGenerator-PS-*.zip.sha256"  # Pattern to match hash files
        
        # Hash verification configuration
        # Disable hash verification for DirectAPI installation method or if DisableHashVerification is specified
        EnableHashVerification = ($InstallType -ne "DirectAPI") -and (-not $DisableHashVerification)
        StrictHashMode = $StrictHashVerification  # Disabled by default, enable if parameter is specified
        
        # Installation paths
        AppDataPath = [System.Environment]::GetFolderPath('ApplicationData')
        InstallDir = Join-Path -Path ([System.Environment]::GetFolderPath('ApplicationData')) -ChildPath "securepassgenerator-ps"
        DesktopPath = [System.Environment]::GetFolderPath('Desktop')
        LogPath = Join-Path -Path (Join-Path -Path ([System.Environment]::GetFolderPath('ApplicationData')) -ChildPath "securepassgenerator-ps") -ChildPath "install_log.txt"
        TempZipPath = Join-Path -Path $env:TEMP -ChildPath "SecurePassGenerator-PS.zip"
        
        # File management configuration
        FileManagement = @{
            # Required files (must be included in both online and offline installations)
            RequiredFiles = @(
                @{Pattern = "SecurePassGenerator.ps1"; Type = "Script"},  # Main script
                @{Pattern = "SecurePassGenerator.ico"; Type = "Icon"},    # Icon file
                @{Pattern = "LICENSE"; Type = "License"}                  # License file
            )
            # Optional files (include if available)
            OptionalFiles = @(
                @{Pattern = "README.md"; Type = "Readme"},                # Documentation
                @{Pattern = "CHANGELOG.md"; Type = "Changelog"}           # Change history
            )
            # Files to exclude (never include these files)
            ExcludeFiles = @(
                "Installer.ps1",            # Installer script
                "*.tmp",                    # Temporary files
                "*.bak"                     # Backup files
            )
        }
    }
    
    return $config
}

# Initialize the configuration
$Config = Initialize-InstallerConfig
#endregion

#region Logging Functions
function Write-Log {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
        [string]$Level = "INFO",
        
        [Parameter(Mandatory = $false)]
        [bool]$ShowInConsole = $false
    )

    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$Timestamp] [$Level] $Message"

    # Ensure log directory exists
    if (-not (Test-Path -Path (Split-Path -Path $Config.LogPath -Parent))) {
        New-Item -Path (Split-Path -Path $Config.LogPath -Parent) -ItemType Directory -Force | Out-Null
    }

    # Always write to log file regardless of level or ShowInConsole setting
    Add-Content -Path $Config.LogPath -Value $LogMessage

    # Only write to console if ShowInConsole is true or if it's an ERROR or WARNING
    if ($ShowInConsole -or $Level -eq "ERROR" -or $Level -eq "WARNING") {
        # Write to console with color-coding
        switch ($Level) {
            "INFO" { Write-Host $LogMessage }
            "WARNING" { Write-Host $LogMessage -ForegroundColor Yellow }
            "ERROR" { Write-Host $LogMessage -ForegroundColor Red }
            "SUCCESS" { Write-Host $LogMessage -ForegroundColor Green }
        }
    }
}
#endregion

#region UI Helper Functions
function Show-InstallerHelp {
    [CmdletBinding()]
    param()
    
    Write-Host ""
    Write-Host "+----------------------------------------------------------+"
    Write-Host "|             SecurePassGenerator Installer                |"
    Write-Host "|                                                          |"
    Write-Host "|  Version: $script:Version                                            |"
    Write-Host "+----------------------------------------------------------+"
    Write-Host ""
    Write-Host "This script requires parameters to run. For an interactive menu,"
    Write-Host "please use the RUN_ME_FIRST.bat file in the same directory."
    Write-Host ""
    Write-Host "Available parameters:"
    Write-Host ""
    Write-Host "  -InstallType <Offline|Online|DirectAPI>    (Required)"
    Write-Host "      Specifies whether to install from local files, download from GitHub,"
    Write-Host "      or use the GitHub API directly"
    Write-Host ""
    Write-Host "  -ReleaseType <Latest|PreRelease> (Optional, default: Latest)"
    Write-Host "      For online installation, specifies which release type to download"
    Write-Host ""
    Write-Host "  -DisableHashVerification         (Optional switch)"
    Write-Host "      Disables SHA256 hash verification of downloaded or local zip files"
    Write-Host ""
    Write-Host "  -StrictHashVerification          (Optional switch)"
    Write-Host "      Requires that a hash file must be present for installation to proceed"
    Write-Host ""
    Write-Host "Examples:"
    Write-Host ""
    Write-Host "  .\Installer.ps1 -InstallType Offline"
    Write-Host "  .\Installer.ps1 -InstallType Online -ReleaseType Latest"
    Write-Host "  .\Installer.ps1 -InstallType Online -ReleaseType PreRelease"
    Write-Host "  .\Installer.ps1 -InstallType Online -DisableHashVerification"
    Write-Host "  .\Installer.ps1 -InstallType Online -StrictHashVerification"
    Write-Host "  .\Installer.ps1 -InstallType DirectAPI -ReleaseType Latest"
    Write-Host "  .\Installer.ps1 -InstallType DirectAPI -ReleaseType PreRelease"
    Write-Host ""
    Write-Host "For an interactive menu, use:"
    Write-Host ""
    Write-Host "  .\RUN_ME_FIRST.bat"
    Write-Host ""
    exit
}

function Show-ProgressStep {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [int]$StepNumber,
        
        [Parameter(Mandatory = $true)]
        [int]$TotalSteps,
        
        [Parameter(Mandatory = $true)]
        [string]$StepDescription
    )
    
    Write-Host "[$StepNumber/$TotalSteps] $StepDescription" -ForegroundColor Cyan
}

function Show-StepResult {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("OK", "ERROR", "WARNING")]
        [string]$Status = "OK"
    )
    
    Write-Host "      $Message" -ForegroundColor Gray
    
    switch ($Status) {
        "OK" { 
            Write-Host "      [OK]" -ForegroundColor Green
        }
        "ERROR" { 
            Write-Host "      [ERROR]" -ForegroundColor Red
        }
        "WARNING" { 
            Write-Host "      [WARNING]" -ForegroundColor Yellow
        }
    }
    
    Write-Host ""
}

function Show-SecurityVerification {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string]$Hash = "",
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("OK", "ERROR", "SKIPPED", "DISABLED")]
        [string]$Status = "OK"
    )
    
    Write-Host "> Security: " -NoNewline -ForegroundColor Cyan
    
    switch ($Status) {
        "OK" {
            Write-Host "Verifying file integrity" -ForegroundColor Gray
            Write-Host "      SHA256: $Hash"
            Write-Host "      [OK]" -ForegroundColor Green
        }
        "ERROR" {
            Write-Host "Verification failed" -ForegroundColor Gray
            Write-Host "      SHA256: $Hash"
            Write-Host "      [ERROR]" -ForegroundColor Red
        }
        "SKIPPED" {
            Write-Host "Hash verification skipped (no hash file found)" -ForegroundColor Yellow
        }
        "DISABLED" {
            Write-Host "Hash verification disabled" -ForegroundColor Yellow
        }
    }
    
    Write-Host ""
}

function Show-InstallerBanner {
    [CmdletBinding()]
    param()
    
    Write-Host ""
    Write-Host "+----------------------------------------------------------+"
    Write-Host "|             SecurePassGenerator Installer                |"
    Write-Host "|                                                          |"
    Write-Host "|  Version: $script:Version                                            |"
    Write-Host "+----------------------------------------------------------+"
    Write-Host ""
}

function Show-InstallationDetails {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$InstallType,
        
        [Parameter(Mandatory = $false)]
        [string]$ReleaseType = "Latest"
    )
    
    if ($InstallType -eq "Online" -or $InstallType -eq "DirectAPI") {
        Write-Host "> Installation Type: Online ($ReleaseType Release)"
        if ($InstallType -eq "DirectAPI") {
            Write-Host "> Method: Direct API (proxy friendly)"
        }
    } else {
        Write-Host "> Installation Type: Offline"
        Write-Host "> Source: Local Files"
    }
    
    Write-Host "> Destination: $($Config.InstallDir)"
    Write-Host ""
}

function Show-CompletionMessage {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [bool]$Success,
        
        [Parameter(Mandatory = $false)]
        [string]$ErrorMessage = "",
        
        [Parameter(Mandatory = $false)]
        [string]$ScriptPath = ""
    )
    
    if ($Success) {
        Write-Host ""
        Write-Host "[SUCCESS] Installation Complete" -ForegroundColor Green
        Write-Host ""
        Write-Host "SecurePassGenerator has been successfully installed."
        Write-Host "A shortcut has been created on your desktop."
        
        if (-not [string]::IsNullOrEmpty($ScriptPath)) {
            Write-Host ""
            Write-Host "Launch the application from your desktop or run:"
            Write-Host "$ScriptPath" -ForegroundColor Cyan
        }
    }
    else {
        Write-Host ""
        Write-Host "[ERROR] Installation Failed" -ForegroundColor Red
        Write-Host ""
        
        if (-not [string]::IsNullOrEmpty($ErrorMessage)) {
            Write-Host "Error: $ErrorMessage" -ForegroundColor Red
            Write-Host ""
        }
        
        Write-Host "For troubleshooting, check the log file:"
        Write-Host "$($Config.LogPath)" -ForegroundColor Cyan
    }
}
#endregion

#region File Operation Functions
function Get-AllowedFiles {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,
        
        [Parameter(Mandatory = $false)]
        [switch]$IncludeDirectories
    )
    
    Write-Log "Filtering files from: $SourcePath" -Level "INFO"
    
    try {
        $AllFiles = Get-ChildItem -Path $SourcePath -Recurse -File
        $AllowedFiles = @()
        $ExcludedFiles = @()
        
        # Process required files
        foreach ($FilePattern in $Config.FileManagement.RequiredFiles) {
            $MatchingFiles = $AllFiles | Where-Object { $_.Name -like $FilePattern }
            
            foreach ($File in $MatchingFiles) {
                if ($AllowedFiles -notcontains $File) {
                    $AllowedFiles += $File
                    Write-Log "Required file matched: $($File.Name)" -Level "INFO"
                }
            }
        }
        
        # Process optional files
        foreach ($FilePattern in $Config.FileManagement.OptionalFiles) {
            $MatchingFiles = $AllFiles | Where-Object { $_.Name -like $FilePattern }
            
            foreach ($File in $MatchingFiles) {
                if ($AllowedFiles -notcontains $File) {
                    $AllowedFiles += $File
                    Write-Log "Optional file matched: $($File.Name)" -Level "INFO"
                }
            }
        }
        
        # Process excluded files (remove any that were matched)
        foreach ($FilePattern in $Config.FileManagement.ExcludeFiles) {
            $MatchingFiles = $AllowedFiles | Where-Object { $_.Name -like $FilePattern }
            
            foreach ($File in $MatchingFiles) {
                $AllowedFiles = $AllowedFiles | Where-Object { $_ -ne $File }
                $ExcludedFiles += $File
                Write-Log "Excluded file: $($File.Name)" -Level "INFO"
            }
        }
        
        # Include directories if requested
        if ($IncludeDirectories) {
            $AllDirs = Get-ChildItem -Path $SourcePath -Recurse -Directory
            $AllowedFiles += $AllDirs
        }
        
        Write-Log "Found $($AllowedFiles.Count) allowed files and $($ExcludedFiles.Count) excluded files" -Level "INFO"
        return $AllowedFiles
    }
    catch {
        Write-Log "Error filtering files: $_" -Level "ERROR"
        return @()
    }
}

function Unblock-FileIfNeeded {
    [CmdletBinding()]
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

function Unblock-InstalledFiles {
    [CmdletBinding()]
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

#region Shortcut Functions
function New-Shortcut {
    [CmdletBinding()]
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
        # Set PowerShell icon if no custom icon is found or if target is PowerShell
        else {
            $Shortcut.IconLocation = "powershell.exe,0"
            Write-Log "Using default PowerShell icon" -Level "INFO"
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
#endregion

#region Network Functions
function Test-InternetConnectivity {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string]$ServiceUrl = "https://github.com"
    )
    
    Write-Log "Testing internet connectivity..." -Level "INFO"
    
    try {
        # Extract the hostname from the URL
        $uri = [System.Uri]$ServiceUrl
        $hostname = $uri.Host
        
        Write-Log "Checking connectivity to: $hostname" -Level "INFO" -ShowInConsole $false
        
        # Use a direct HTTP request approach (proxy-friendly)
        try {
            # Create a WebRequest with default proxy settings
            $request = [System.Net.WebRequest]::Create($ServiceUrl)
            $request.Method = "GET"
            $request.Timeout = 10000 # 10 seconds timeout
            $request.UserAgent = "PowerShell-Installer"
            
            # Set proxy and credentials
            $request.Proxy = [System.Net.WebRequest]::DefaultWebProxy
            $request.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
            
            Write-Log "Attempting HTTP GET request to: $ServiceUrl" -Level "INFO" -ShowInConsole $false
            
            try {
                # Get the response
                $response = $request.GetResponse()
                Write-Log "Successfully connected to $ServiceUrl (HTTP Status: $($response.StatusCode))" -Level "INFO" -ShowInConsole $false
                return $true
            }
            finally {
                # Dispose of the response if it exists
                if ($null -ne $response) {
                    Write-Log "Disposing HTTP response" -Level "INFO"
                    $response.Dispose()
                }
            }
        }
        catch {
            Write-Log "HTTP request failed: $_" -Level "WARNING"
            
            # Try a fallback approach with Invoke-WebRequest
            try {
                Write-Log "Attempting fallback with Invoke-WebRequest..." -Level "INFO" -ShowInConsole $false
                $webRequestParams = @{
                    Uri = $ServiceUrl
                    Method = "GET"
                    UseBasicParsing = $true
                    TimeoutSec = 10
                    ErrorAction = "Stop"
                }
                
                $null = Invoke-WebRequest @webRequestParams
                Write-Log "Fallback request succeeded" -Level "INFO" -ShowInConsole $false
                return $true
            }
            catch {
                Write-Log "Fallback request failed: $_" -Level "WARNING"
                
                # As a last resort, try a ping test
                # This might not work in corporate environments but we'll try anyway
                try {
                    Write-Log "Attempting ping test as last resort..." -Level "INFO" -ShowInConsole $false
                    $pingResult = Test-Connection -ComputerName $hostname -Count 1 -Quiet
                    
                    if ($pingResult) {
                        Write-Log "Ping test successful" -Level "INFO" -ShowInConsole $false
                        return $true
                    }
                    else {
                        Write-Log "Ping test failed" -Level "WARNING"
                    }
                }
                catch {
                    Write-Log "Ping test error: $_" -Level "WARNING"
                }
                
                return $false
            }
        }
    }
    catch {
        Write-Log "Error testing internet connectivity: $_" -Level "ERROR"
        return $false
    }
}
#endregion

#region Hash Verification Functions
function Test-FileHash {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$FilePath,
        
        [Parameter(Mandatory = $false)]
        [string]$ExpectedHash = "",
        
        [Parameter(Mandatory = $false)]
        [string]$HashFilePath = "",
        
        [Parameter(Mandatory = $false)]
        [bool]$StrictMode = $false
    )
    
    Write-Log "Starting hash verification for file: $FilePath" -Level "INFO" -ShowInConsole $false
    Write-Log "Hash verification mode: StrictMode=$StrictMode" -Level "INFO"
    
    # If hash verification is disabled, return success
    if (-not $Config.EnableHashVerification) {
        Write-Log "Hash verification disabled in configuration, skipping" -Level "INFO" -ShowInConsole $false
        Show-SecurityVerification -Status "DISABLED"
        return @{
            Success = $true
            Status = "DISABLED"
            ActualHash = ""
            ExpectedHash = ""
        }
    }
    
    # Check if file exists
    if (-not (Test-Path -Path $FilePath)) {
        Write-Log "File not found for hash verification: $FilePath" -Level "ERROR"
        return @{
            Success = $false
            Status = "ERROR"
            ActualHash = ""
            ExpectedHash = ""
            ErrorMessage = "File not found: $FilePath"
        }
    } else {
        $FileInfo = Get-Item -Path $FilePath
        $FileSize = [Math]::Round($FileInfo.Length / 1KB, 2)
        Write-Log "File found: $FilePath (Size: $FileSize KB)" -Level "INFO"
    }
    
    # If hash file path is provided, try to read the expected hash from it
    if (-not [string]::IsNullOrEmpty($HashFilePath) -and (Test-Path -Path $HashFilePath)) {
        Write-Log "Hash file found: $HashFilePath" -Level "INFO" -ShowInConsole $false
        try {
            $HashFileContent = Get-Content -Path $HashFilePath -Raw
            Write-Log "Hash file content (raw): $HashFileContent" -Level "INFO"
            
            $ExpectedHash = $HashFileContent.Trim()
            
            # Extract just the hash part if the file contains filename too
            if ($ExpectedHash -match "^([a-fA-F0-9]{64})") {
                $ExpectedHash = $Matches[1]
                Write-Log "Extracted hash from file: $ExpectedHash" -Level "INFO"
            } else {
                Write-Log "Hash file does not contain a standard SHA256 hash pattern" -Level "WARNING"
            }
            
            Write-Log "Using expected hash from file: $ExpectedHash" -Level "INFO" -ShowInConsole $false
        }
        catch {
            Write-Log "Error reading hash file: $_" -Level "WARNING"
            
            if ($StrictMode) {
                Write-Log "Strict mode enabled - failing verification due to hash file read error" -Level "ERROR"
                return @{
                    Success = $false
                    Status = "ERROR"
                    ActualHash = ""
                    ExpectedHash = ""
                    ErrorMessage = "Error reading hash file: $_"
                }
            } else {
                Write-Log "Continuing without hash verification (non-strict mode)" -Level "WARNING"
            }
        }
    } else {
        if (-not [string]::IsNullOrEmpty($HashFilePath)) {
            Write-Log "Hash file not found: $HashFilePath" -Level "WARNING"
        } else {
            Write-Log "No hash file path provided" -Level "INFO"
        }
        
        if (-not [string]::IsNullOrEmpty($ExpectedHash)) {
            Write-Log "Using provided expected hash: $ExpectedHash" -Level "INFO"
        }
    }
    
    # If no expected hash is provided or could be read from file
    if ([string]::IsNullOrEmpty($ExpectedHash)) {
        if ($StrictMode) {
            Write-Log "No hash available for verification and strict mode is enabled" -Level "ERROR"
            Show-SecurityVerification -Status "ERROR"
            return @{
                Success = $false
                Status = "ERROR"
                ActualHash = ""
                ExpectedHash = ""
                ErrorMessage = "No hash available for verification and strict mode is enabled"
            }
        }
        else {
            Write-Log "No hash available for verification, skipping (non-strict mode)" -Level "WARNING"
            Show-SecurityVerification -Status "SKIPPED"
            return @{
                Success = $true
                Status = "SKIPPED"
                ActualHash = ""
                ExpectedHash = ""
            }
        }
    }
    
    # Calculate the actual hash of the file
    try {
        Write-Log "Calculating SHA256 hash for file: $FilePath" -Level "INFO" -ShowInConsole $false
        $ActualHash = Get-FileHash -Path $FilePath -Algorithm SHA256
        Write-Log "Hash calculation completed" -Level "INFO"
        
        # Compare the hashes (case-insensitive)
        if ($ActualHash.Hash -eq $ExpectedHash -or $ActualHash.Hash.ToLower() -eq $ExpectedHash.ToLower()) {
            Write-Log "Expected SHA256 hash: $($ExpectedHash.ToUpper())" -Level "INFO" -ShowInConsole $false
            Write-Log "Actual SHA256 hash: $($ActualHash.Hash)" -Level "INFO" -ShowInConsole $false
            Write-Log "Hash verification successful - hashes match" -Level "INFO" -ShowInConsole $false
            
            # Display security verification
            Show-SecurityVerification -Hash $ExpectedHash.ToUpper() -Status "OK"
            
            return @{
                Success = $true
                Status = "OK"
                ActualHash = $ActualHash.Hash
                ExpectedHash = $ExpectedHash.ToUpper()
            }
        }
        else {
            Write-Log "Hash verification failed! The file may be corrupted or tampered with." -Level "ERROR"
            Write-Log "Expected: $($ExpectedHash.ToUpper())" -Level "ERROR"
            Write-Log "Actual: $($ActualHash.Hash)" -Level "ERROR"
            
            # Display security verification failure
            Show-SecurityVerification -Hash $ExpectedHash.ToUpper() -Status "ERROR"
            
            return @{
                Success = $false
                Status = "ERROR"
                ActualHash = $ActualHash.Hash
                ExpectedHash = $ExpectedHash.ToUpper()
                ErrorMessage = "Hash verification failed"
            }
        }
    }
    catch {
        Write-Log "Error calculating file hash: $_" -Level "ERROR"
        return @{
            Success = $false
            Status = "ERROR"
            ActualHash = ""
            ExpectedHash = $ExpectedHash
            ErrorMessage = "Error calculating file hash: $_"
        }
    }
}
#endregion

#region File Extraction Functions
function Expand-ZipFile {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ZipPath,
        
        [Parameter(Mandatory = $true)]
        [string]$DestinationPath,
        
        [Parameter(Mandatory = $false)]
        [switch]$CleanDestination
    )
    
    Write-Log "Starting extraction process for zip file: $ZipPath" -Level "INFO" -ShowInConsole $false
    
    try {
        # Check if the zip file exists
        if (-not (Test-Path -Path $ZipPath)) {
            Write-Log "Zip file not found: $ZipPath" -Level "ERROR"
            return $false
        } else {
            $ZipFileInfo = Get-Item -Path $ZipPath
            $ZipFileSize = [Math]::Round($ZipFileInfo.Length / 1MB, 2)
            Write-Log "Zip file found: $ZipPath (Size: $ZipFileSize MB, Last Modified: $($ZipFileInfo.LastWriteTime))" -Level "INFO" -ShowInConsole $false
            Write-Log "Zip file attributes: $($ZipFileInfo.Attributes), Created: $($ZipFileInfo.CreationTime)" -Level "INFO"
        }
        
        # Create destination directory if it doesn't exist
        if (-not (Test-Path -Path $DestinationPath)) {
            Write-Log "Creating destination directory: $DestinationPath" -Level "INFO" -ShowInConsole $false
            $NewDir = New-Item -Path $DestinationPath -ItemType Directory -Force
            Write-Log "Destination directory created successfully: $($NewDir.FullName)" -Level "INFO" -ShowInConsole $false
        } else {
            Write-Log "Destination directory already exists: $DestinationPath" -Level "INFO" -ShowInConsole $false
            $DirInfo = Get-Item -Path $DestinationPath
            Write-Log "Directory info: Created: $($DirInfo.CreationTime), Last Modified: $($DirInfo.LastWriteTime)" -Level "INFO"
            
            # Clean destination if requested
            if ($CleanDestination) {
                Write-Log "Cleaning destination directory before extraction" -Level "INFO" -ShowInConsole $false
                $ExistingItems = Get-ChildItem -Path $DestinationPath -Force
                Write-Log "Found $($ExistingItems.Count) items to remove from destination" -Level "INFO" -ShowInConsole $false
                
                if ($ExistingItems.Count -gt 0) {
                    # Log details about items to be removed
                    $ExistingDirs = ($ExistingItems | Where-Object { $_.PSIsContainer }).Count
                    $ExistingFiles = ($ExistingItems | Where-Object { -not $_.PSIsContainer }).Count
                    Write-Log "Items to remove: $ExistingDirs directories and $ExistingFiles files" -Level "INFO" -ShowInConsole $false
                    
            try {
                # Exclude the log file and presets.json when cleaning the destination directory
                $LogFileName = Split-Path -Leaf $Config.LogPath
                $PresetsFileName = "presets.json"
                Get-ChildItem -Path $DestinationPath -Force | 
                    Where-Object { $_.Name -ne $LogFileName -and $_.Name -ne $PresetsFileName } | 
                    Remove-Item -Recurse -Force
                Write-Log "Successfully cleaned destination directory (preserved log file and presets)" -Level "INFO" -ShowInConsole $false
                } catch {
                    Write-Log "Error cleaning destination directory: $_" -Level "WARNING"
                    # Continue despite cleaning failure
                }
                } else {
                    Write-Log "Destination directory is already empty" -Level "INFO"
                }
            } else {
                Write-Log "Destination directory will not be cleaned (CleanDestination=$CleanDestination)" -Level "INFO"
            }
        }
        
            # Expand the zip file
            Write-Log "Extracting zip file to $DestinationPath" -Level "INFO" -ShowInConsole $false
        $StartTime = Get-Date
        
        try {
            Expand-Archive -Path $ZipPath -DestinationPath $DestinationPath -Force
            $EndTime = Get-Date
            $ExtractionTime = ($EndTime - $StartTime).TotalSeconds
            Write-Log "Expand-Archive command completed in $([Math]::Round($ExtractionTime, 2)) seconds" -Level "INFO" -ShowInConsole $false
        } catch {
            Write-Log "Error during Expand-Archive operation: $_" -Level "ERROR"
            return $false
        }
        
        # Verify extraction was successful
        if (Test-Path -Path $DestinationPath) {
            $ExtractedItems = Get-ChildItem -Path $DestinationPath
            
            if ($ExtractedItems.Count -gt 0) {
                Write-Log "Successfully extracted zip file with $($ExtractedItems.Count) items" -Level "INFO" -ShowInConsole $false
                
                # Log details about extracted items
                $Directories = ($ExtractedItems | Where-Object { $_.PSIsContainer }).Count
                $Files = ($ExtractedItems | Where-Object { -not $_.PSIsContainer }).Count
                Write-Log "Extracted $Directories directories and $Files files" -Level "INFO" -ShowInConsole $false
                
                # Calculate total size of extracted files
                $TotalSizeKB = 0
                $AllExtractedFiles = Get-ChildItem -Path $DestinationPath -Recurse -File
                foreach ($File in $AllExtractedFiles) {
                    $TotalSizeKB += [Math]::Round($File.Length / 1KB, 2)
                }
                Write-Log "Total size of extracted files: $TotalSizeKB KB ($([Math]::Round($TotalSizeKB / 1024, 2)) MB)" -Level "INFO" -ShowInConsole $false
                
                # Log the first few items for verification
                $TopItems = $ExtractedItems | Select-Object -First 5
                foreach ($Item in $TopItems) {
                    if ($Item.PSIsContainer) {
                        $SubItemCount = (Get-ChildItem -Path $Item.FullName -Recurse).Count
                        Write-Log "Extracted directory: $($Item.Name) (Contains $SubItemCount items)" -Level "INFO" -ShowInConsole $false
                    } else {
                        $ItemSize = [Math]::Round($Item.Length / 1KB, 2)
                        Write-Log "Extracted file: $($Item.Name) (Size: $ItemSize KB)" -Level "INFO" -ShowInConsole $false
                    }
                }
                
                if ($ExtractedItems.Count -gt 5) {
                    Write-Log "... and $($ExtractedItems.Count - 5) more items" -Level "INFO"
                }
                
                # Check for specific files
                $MainScriptFile = Get-ChildItem -Path $DestinationPath -Filter $Config.ScriptName -Recurse | Select-Object -First 1
                if ($MainScriptFile) {
                    Write-Log "Found main script file: $($MainScriptFile.FullName)" -Level "INFO" -ShowInConsole $false
                } else {
                    Write-Log "Main script file not found in extracted content" -Level "WARNING"
                }
                
                return $true
            } else {
                Write-Log "Extraction completed but no items found in destination" -Level "WARNING"
                return $false
            }
        } else {
            Write-Log "Extraction failed: Destination path not found after extraction" -Level "ERROR"
            return $false
        }
    }
    catch {
        Write-Log "Error extracting zip file: $_" -Level "ERROR"
        return $false
    }
}

function Move-ExtractedFiles {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ExtractedPath,
        
        [Parameter(Mandatory = $false)]
        [string]$MainScriptName = $Config.ScriptName
    )
    
    Write-Log "Starting reorganization of extracted files in: $ExtractedPath" -Level "INFO" -ShowInConsole $false
    Write-Log "Looking for main script: $MainScriptName" -Level "INFO" -ShowInConsole $false
    
    try {
        # Find subdirectories in the installation directory
        $SubDirs = Get-ChildItem -Path $ExtractedPath -Directory
        Write-Log "Found $($SubDirs.Count) subdirectories in extraction path" -Level "INFO" -ShowInConsole $false
        
        # Log subdirectory details
        if ($SubDirs.Count -gt 0) {
            foreach ($Dir in $SubDirs) {
                $DirItemCount = (Get-ChildItem -Path $Dir.FullName -Recurse).Count
                Write-Log "Subdirectory: $($Dir.Name) (Contains $DirItemCount items)" -Level "INFO"
            }
        }
        
        if ($SubDirs.Count -gt 0) {
            # Assume the first subdirectory contains our files
            $ExtractedDir = $SubDirs[0].FullName
            Write-Log "Using first subdirectory for reorganization: $ExtractedDir" -Level "INFO" -ShowInConsole $false
            
            # List all files in the subdirectory for logging
            $AllFiles = Get-ChildItem -Path $ExtractedDir -Recurse
            $AllDirs = $AllFiles | Where-Object { $_.PSIsContainer }
            $AllActualFiles = $AllFiles | Where-Object { -not $_.PSIsContainer }
            
            Write-Log "Subdirectory contains $($AllFiles.Count) items: $($AllDirs.Count) directories and $($AllActualFiles.Count) files" -Level "INFO" -ShowInConsole $false
            
            # Log file extensions for better understanding of content
            $Extensions = $AllActualFiles | ForEach-Object { $_.Extension } | Sort-Object -Unique
            Write-Log "File extensions found: $($Extensions -join ', ')" -Level "INFO" -ShowInConsole $false
            
            $FilesReorganized = $false
            $FilesToMove = @()
            
            # Get a list of all files that match our file management configuration
            $AllowedFiles = @()
            
            # Process required files first
            foreach ($FileConfig in $Config.FileManagement.RequiredFiles) {
                $Pattern = $FileConfig.Pattern
                $FileType = $FileConfig.Type
                $MatchingFiles = Get-ChildItem -Path $ExtractedDir -Filter $Pattern -Recurse
                
                foreach ($File in $MatchingFiles) {
                    if ($AllowedFiles -notcontains $File) {
                        $AllowedFiles += [PSCustomObject]@{
                            File = $File
                            Type = $FileType
                        }
                        Write-Log "Required file matched: $($File.Name) (Type: $FileType)" -Level "INFO"
                    }
                }
            }
            
            # Process optional files
            foreach ($FileConfig in $Config.FileManagement.OptionalFiles) {
                $Pattern = $FileConfig.Pattern
                $FileType = $FileConfig.Type
                $MatchingFiles = Get-ChildItem -Path $ExtractedDir -Filter $Pattern -Recurse
                
                foreach ($File in $MatchingFiles) {
                    if ($AllowedFiles -notcontains $File) {
                        $AllowedFiles += [PSCustomObject]@{
                            File = $File
                            Type = $FileType
                        }
                        Write-Log "Optional file matched: $($File.Name) (Type: $FileType)" -Level "INFO"
                    }
                }
            }
            
            # Process excluded files (remove any that were matched)
            foreach ($FilePattern in $Config.FileManagement.ExcludeFiles) {
                $MatchingFiles = $AllowedFiles | Where-Object { $_.File.Name -like $FilePattern }
                
                foreach ($FileObj in $MatchingFiles) {
                    $AllowedFiles = $AllowedFiles | Where-Object { $_ -ne $FileObj }
                    Write-Log "Excluded file: $($FileObj.File.Name)" -Level "INFO"
                }
            }
            
            # Now move only the allowed files
            foreach ($FileObj in $AllowedFiles) {
                $File = $FileObj.File
                $FileType = $FileObj.Type
                $TargetPath = Join-Path -Path $ExtractedPath -ChildPath $File.Name
                
                Write-Log "Moving file ($FileType) to main directory: $TargetPath" -Level "INFO"
                $FilesToMove += @{Source = $File.FullName; Destination = $TargetPath; Type = $FileType }
                
                try {
                    Copy-Item -Path $File.FullName -Destination $TargetPath -Force
                    Write-Log "Successfully copied $FileType file to main directory" -Level "INFO"
                    $FilesReorganized = $true
                    
                    # Unblock the moved file if it's a script
                    if ($FileType -eq "Script" -or $File.Extension -eq ".ps1" -or $File.Extension -eq ".psm1") {
                        Unblock-InstalledFiles -Stage "AfterMove" -InstallDirectory $ExtractedPath -SpecificFile $File.Name
                    }
                } catch {
                    Write-Log "Error copying $FileType file: $_" -Level "ERROR"
                }
            }
            
            # Summary of moved files
            if ($FilesToMove.Count -gt 0) {
                Write-Log "Summary of moved files:" -Level "INFO"
                foreach ($File in $FilesToMove) {
                    Write-Log "- $($File.Type): $($File.Source) -> $($File.Destination)" -Level "INFO"
                }
            }
            
            # Remove the subdirectory and its contents if files were reorganized
            if ($FilesReorganized) {
                Write-Log "Files were reorganized, removing original subdirectory: $ExtractedDir" -Level "INFO"
                
                try {
                    Remove-Item -Path $ExtractedDir -Recurse -Force
                    Write-Log "Successfully removed subdirectory after moving necessary files" -Level "INFO"
                } catch {
                    Write-Log "Error removing subdirectory: $_" -Level "WARNING"
                    # Continue despite removal failure
                }
            } else {
                Write-Log "No files were reorganized, keeping subdirectory structure" -Level "INFO"
            }
            
            # Verify the reorganization
            $FinalFiles = Get-ChildItem -Path $ExtractedPath -File
            Write-Log "After reorganization, main directory contains $($FinalFiles.Count) files" -Level "INFO"
            
            if ($FinalFiles.Count -gt 0) {
                Write-Log "Files in main directory after reorganization:" -Level "INFO"
                foreach ($File in $FinalFiles) {
                    $FileSize = [Math]::Round($File.Length / 1KB, 2)
                    Write-Log "- $($File.Name) (Size: $FileSize KB)" -Level "INFO"
                }
            }
            
            return $FilesReorganized
        }
        else {
            Write-Log "No subdirectories found in the extracted directory, no reorganization needed" -Level "INFO"
            
            # List files in the main directory for verification
            $MainDirFiles = Get-ChildItem -Path $ExtractedPath -File
            Write-Log "Main directory already contains $($MainDirFiles.Count) files" -Level "INFO"
            
            if ($MainDirFiles.Count -gt 0) {
                Write-Log "Files in main directory:" -Level "INFO"
                foreach ($File in $MainDirFiles) {
                    $FileSize = [Math]::Round($File.Length / 1KB, 2)
                    Write-Log "- $($File.Name) (Size: $FileSize KB)" -Level "INFO"
                }
                
                # Check if main script is already in the main directory
                $MainScriptInRoot = $MainDirFiles | Where-Object { $_.Name -eq $MainScriptName } | Select-Object -First 1
                if ($MainScriptInRoot) {
                    Write-Log "Main script $MainScriptName already exists in the root directory" -Level "INFO"
                } else {
                    Write-Log "Main script $MainScriptName not found in the root directory" -Level "WARNING"
                    
                    # Look for any PS1 files that might be the main script
                    $PsFilesInRoot = $MainDirFiles | Where-Object { $_.Extension -eq ".ps1" }
                    if ($PsFilesInRoot.Count -gt 0) {
                        Write-Log "Found PS1 files in root: $($PsFilesInRoot.Name -join ', ')" -Level "INFO"
                    }
                }
            }
            
            return $true  # No reorganization needed is still a success
        }
    }
    catch {
        Write-Log "Error reorganizing extracted files: $_" -Level "ERROR"
        return $false
    }
}
#endregion

#region GitHub API Functions
function Get-GitHubReleaseInfo {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string]$Username = $Config.GitHubUsername,
        
        [Parameter(Mandatory = $false)]
        [string]$Repository = $Config.RepositoryName,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("Latest", "PreRelease")]
        [string]$ReleaseType = "Latest",
        
        [Parameter(Mandatory = $false)]
        [string]$Token = $Config.GitHubToken
    )
    
    try {
        Write-Log "Fetching GitHub release information for $Username/$Repository" -Level "INFO" -ShowInConsole $false
        Write-Log "Release type requested: $ReleaseType" -Level "INFO" -ShowInConsole $false
        
        # Initialize headers
        $Headers = @{
            "Accept" = "application/vnd.github+json"
            "X-GitHub-Api-Version" = "2022-11-28"
        }
        
        # Add authorization header if token is provided
        if (-not [string]::IsNullOrWhiteSpace($Token)) {
            Write-Log "Using GitHub token for authentication" -Level "INFO" -ShowInConsole $false
            $Headers["Authorization"] = "token $Token"
        } else {
            Write-Log "No GitHub token provided, accessing as public repository" -Level "INFO" -ShowInConsole $false
        }
        
        # Determine which API endpoint to use based on ReleaseType
        if ($ReleaseType -eq "Latest") {
            $ReleaseUrl = "https://api.github.com/repos/$Username/$Repository/releases/latest"
            Write-Log "Using latest release API URL: $ReleaseUrl" -Level "INFO" -ShowInConsole $false
            
            Write-Log "Sending request to GitHub API..." -Level "INFO" -ShowInConsole $false
            $ReleaseInfo = Invoke-RestMethod -Uri $ReleaseUrl -Headers $Headers -Method Get
            Write-Log "Successfully retrieved latest release information" -Level "INFO" -ShowInConsole $false
            Write-Log "Release tag: $($ReleaseInfo.tag_name), Name: $($ReleaseInfo.name), Created: $($ReleaseInfo.created_at)" -Level "INFO" -ShowInConsole $false
            Write-Log "Release URL: $($ReleaseInfo.html_url)" -Level "INFO"
            
            if ($ReleaseInfo.assets.Count -gt 0) {
                Write-Log "Release contains $($ReleaseInfo.assets.Count) assets" -Level "INFO" -ShowInConsole $false
            } else {
                Write-Log "Release contains no assets" -Level "WARNING" -ShowInConsole $false
            }
            
            return @{
                Success = $true
                ReleaseInfo = $ReleaseInfo
                ActualReleaseType = "Latest"
            }
        }
        else {
            # For pre-releases, we need to get all releases and filter
            $ReleasesUrl = "https://api.github.com/repos/$Username/$Repository/releases"
            Write-Log "Using all releases API URL to find pre-releases: $ReleasesUrl" -Level "INFO" -ShowInConsole $false
            
            Write-Log "Sending request to GitHub API for all releases..." -Level "INFO" -ShowInConsole $false
            $AllReleases = Invoke-RestMethod -Uri $ReleasesUrl -Headers $Headers -Method Get
            Write-Log "Retrieved $($AllReleases.Count) total releases" -Level "INFO" -ShowInConsole $false
            
            # Filter for pre-releases
            $PreReleases = $AllReleases | Where-Object { $_.prerelease -eq $true }
            Write-Log "Found $($PreReleases.Count) pre-releases" -Level "INFO" -ShowInConsole $false
            
            if ($PreReleases.Count -eq 0) {
                Write-Log "No pre-releases found, falling back to latest release" -Level "WARNING" -ShowInConsole $true
                $ReleaseInfo = $AllReleases | Where-Object { $_.prerelease -eq $false } | Select-Object -First 1
                
                if ($null -eq $ReleaseInfo) {
                    Write-Log "No releases found at all" -Level "ERROR" -ShowInConsole $false
                    return @{
                        Success = $false
                        ErrorMessage = "No releases found"
                    }
                } else {
                    Write-Log "Using latest stable release as fallback: $($ReleaseInfo.tag_name)" -Level "INFO" -ShowInConsole $false
                }
                
                # Return the actual release type as "Latest" since we're falling back
                return @{
                    Success = $true
                    ReleaseInfo = $ReleaseInfo
                    ActualReleaseType = "Latest"
                    FalledBack = $true
                }
            }
            else {
                # Get the most recent pre-release
                $ReleaseInfo = $PreReleases | Select-Object -First 1
                Write-Log "Found pre-release: $($ReleaseInfo.tag_name), Name: $($ReleaseInfo.name), Created: $($ReleaseInfo.created_at)" -Level "INFO" -ShowInConsole $false
                Write-Log "Pre-release URL: $($ReleaseInfo.html_url)" -Level "INFO"
                
                if ($ReleaseInfo.assets.Count -gt 0) {
                    Write-Log "Pre-release contains $($ReleaseInfo.assets.Count) assets" -Level "INFO" -ShowInConsole $false
                } else {
                    Write-Log "Pre-release contains no assets" -Level "WARNING" -ShowInConsole $false
                }
                
                return @{
                    Success = $true
                    ReleaseInfo = $ReleaseInfo
                    ActualReleaseType = "PreRelease"
                }
            }
        }
    }
    catch {
        Write-Log "Error retrieving GitHub release information: $_" -Level "ERROR"
        return @{
            Success = $false
            ErrorMessage = "Error retrieving GitHub release information: $_"
        }
    }
}

function Get-GitHubAsset {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [PSObject]$ReleaseInfo,
        
        [Parameter(Mandatory = $false)]
        [string]$AssetPattern = $Config.AssetPattern,
        
        [Parameter(Mandatory = $false)]
        [string]$HashFilePattern = $Config.HashFilePattern
    )
    
    try {
        Write-Log "Searching for assets in release: $($ReleaseInfo.tag_name)" -Level "INFO" -ShowInConsole $false
        Write-Log "Using asset pattern: '$AssetPattern'" -Level "INFO" -ShowInConsole $false
        Write-Log "Using hash file pattern: '$HashFilePattern'" -Level "INFO"
        
        # Log all available assets in the release
        Write-Log "Available assets in release:" -Level "INFO"
        if ($ReleaseInfo.assets.Count -eq 0) {
            Write-Log "No assets found in release" -Level "WARNING" -ShowInConsole $false
        } else {
            Write-Log "Found $($ReleaseInfo.assets.Count) assets in release" -Level "INFO" -ShowInConsole $false
            foreach ($asset in $ReleaseInfo.assets) {
                Write-Log "Asset: $($asset.name) (ID: $($asset.id), Size: $([Math]::Round($asset.size / 1KB, 2)) KB)" -Level "INFO"
            }
        }
        
        # Find the asset matching our pattern
        $Assets = $ReleaseInfo.assets | Where-Object { $_.name -like $AssetPattern }
        Write-Log "Found $($Assets.Count) assets matching pattern '$AssetPattern'" -Level "INFO" -ShowInConsole $false
        
        if ($Assets.Count -eq 0) {
            Write-Log "No assets matching pattern '$AssetPattern' found in the release" -Level "ERROR"
            return @{
                Success = $false
                ErrorMessage = "No assets matching pattern '$AssetPattern' found"
            }
        }
        
        # Use the latest version if multiple assets match
        $Asset = $Assets | Sort-Object -Property name -Descending | Select-Object -First 1
        $AssetName = $Asset.name
        $AssetSize = [Math]::Round($Asset.size / 1KB, 2)
        Write-Log "Selected asset: $AssetName (ID: $($Asset.id), Size: $AssetSize KB)" -Level "INFO" -ShowInConsole $false
        
        # Find the hash file matching our asset
        # First try to find hash file with exact asset name
        $ExactHashPattern = "$AssetName.sha256"
        Write-Log "Looking for hash file with exact name: $ExactHashPattern" -Level "INFO" -ShowInConsole $false
        
        # Check each asset individually to see if it matches
        $HashMatchFound = $false
        foreach ($asset in $ReleaseInfo.assets) {
            Write-Log "Comparing '$($asset.name)' with '$ExactHashPattern'" -Level "INFO"
            if ($asset.name -eq $ExactHashPattern) {
                Write-Log "MATCH FOUND! Hash file: $($asset.name) (ID: $($asset.id))" -Level "INFO" -ShowInConsole $false
                $HashMatchFound = $true
            }
        }
        
        if (-not $HashMatchFound) {
            Write-Log "No exact hash file match found" -Level "INFO" -ShowInConsole $false
        }
        
        $HashAssets = $ReleaseInfo.assets | Where-Object { $_.name -eq $ExactHashPattern }
        Write-Log "HashAssets count after exact match: $($HashAssets.Count)" -Level "INFO"
        
        # If not found, try using the pattern
        if ($HashAssets.Count -eq 0) {
            Write-Log "No exact match found, trying pattern: $HashFilePattern" -Level "INFO" -ShowInConsole $false
            $HashAssets = $ReleaseInfo.assets | Where-Object { $_.name -like $HashFilePattern }
            Write-Log "HashAssets count after pattern match: $($HashAssets.Count)" -Level "INFO"
            
            if ($HashAssets.Count -gt 0) {
                Write-Log "Found hash file using pattern: $($HashAssets[0].name) (ID: $($HashAssets[0].id))" -Level "INFO" -ShowInConsole $false
            }
            else {
                Write-Log "No hash file found using pattern either" -Level "INFO" -ShowInConsole $false
                if ($Config.StrictHashMode) {
                    Write-Log "WARNING: Strict hash mode is enabled but no hash file was found" -Level "WARNING" -ShowInConsole $true
                }
            }
        }
        else {
            Write-Log "Found hash file with exact name: $ExactHashPattern" -Level "INFO" -ShowInConsole $false
        }
        
        # Make sure we're selecting the correct assets
        # The main asset should be the zip file
        $Asset = $ReleaseInfo.assets | Where-Object { $_.name -like $AssetPattern -and $_.name -notlike "*.sha256" } | 
                 Sort-Object -Property name -Descending | 
                 Select-Object -First 1
        
        if ($Asset) {
            $AssetName = $Asset.name
            $AssetSize = [Math]::Round($Asset.size / 1KB, 2)
            Write-Log "Selected main asset: $AssetName (ID: $($Asset.id), Size: $AssetSize KB)" -Level "INFO" -ShowInConsole $false
            Write-Log "Asset download URL: $($Asset.browser_download_url)" -Level "INFO" -ShowInConsole $false
        }
        else {
            Write-Log "No main asset found matching pattern '$AssetPattern'" -Level "ERROR" -ShowInConsole $true
            return @{
                Success = $false
                ErrorMessage = "No main asset found matching pattern '$AssetPattern'"
            }
        }
        
        # Force the hash asset to be found by direct lookup
        $HashAsset = $ReleaseInfo.assets | Where-Object { $_.name -eq "$AssetName.sha256" } | Select-Object -First 1
        
        if ($HashAsset) {
            Write-Log "Selected hash asset: $($HashAsset.name) (ID: $($HashAsset.id))" -Level "INFO" -ShowInConsole $false
        }
        else {
            Write-Log "No hash asset found with exact name" -Level "INFO" -ShowInConsole $false
            
            # Try one more time with pattern matching
            $HashAsset = $ReleaseInfo.assets | Where-Object { $_.name -like $HashFilePattern } | Select-Object -First 1
            
            if ($HashAsset) {
                Write-Log "Found hash asset with pattern matching: $($HashAsset.name) (ID: $($HashAsset.id))" -Level "INFO" -ShowInConsole $false
            }
            else {
                Write-Log "No hash file found with pattern matching either" -Level "INFO" -ShowInConsole $false
                if ($Config.StrictHashMode) {
                    Write-Log "WARNING: Strict hash mode is enabled but no hash file was found" -Level "WARNING" -ShowInConsole $true
                }
            }
        }
        
        Write-Log "Asset selection completed successfully" -Level "INFO" -ShowInConsole $false
        return @{
            Success = $true
            Asset = $Asset
            HashAsset = $HashAsset
        }
    }
    catch {
        Write-Log "Error finding GitHub asset: $_" -Level "ERROR" -ShowInConsole $true
        return @{
            Success = $false
            ErrorMessage = "Error finding GitHub asset: $_"
        }
    }
}

function Invoke-GitHubDownload {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$OutputPath,
        
        [Parameter(Mandatory = $false)]
        [string]$ScriptName = $Config.ScriptName,
        
        [Parameter(Mandatory = $false)]
        [string]$Username = $Config.GitHubUsername,
        
        [Parameter(Mandatory = $false)]
        [string]$Repository = $Config.RepositoryName,
        
        [Parameter(Mandatory = $false)]
        [string]$Token = $Config.GitHubToken,
        
        [Parameter(Mandatory = $false)]
        [string]$Branch = "main"
    )
    
    try {
        Write-Log "Starting direct API download for $ScriptName from $Username/$Repository (branch: $Branch)" -Level "INFO"
        
        # Use GitHub API to get file content
        $apiUrl = "https://api.github.com/repos/$Username/$Repository/contents/$ScriptName"
        
        # Add branch as a query parameter if specified
        if (-not [string]::IsNullOrEmpty($Branch)) {
            $apiUrl = "$apiUrl`?ref=$Branch"
        }
        
        Write-Log "Using GitHub API URL: $apiUrl" -Level "INFO"
        
        # Configure TLS to use 1.2 (needed for GitHub)
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
        Write-Log "TLS 1.2 enabled for secure connection" -Level "INFO"
        
        # Prepare headers
        $headers = @{
            "User-Agent" = "PowerShell Installer"
        }
        
        # Add authorization token if available
        if (-not [string]::IsNullOrWhiteSpace($Token)) {
            Write-Log "Using GitHub token for authentication" -Level "INFO"
            $headers["Authorization"] = "token $Token"
        } else {
            Write-Log "No GitHub token provided, accessing as public repository" -Level "INFO"
        }
        
        Write-Log "Sending request to GitHub API" -Level "INFO"
        $response = Invoke-RestMethod -Uri $apiUrl -Headers $headers
        Write-Log "API call successful" -Level "INFO"
        
        # Determine if this is a binary file based on extension
        $isBinaryFile = $false
        $binaryExtensions = @('.ico', '.exe', '.dll', '.bin', '.jpg', '.jpeg', '.png', '.gif', '.bmp')
        $fileExtension = [System.IO.Path]::GetExtension($ScriptName).ToLower()
        
        if ($binaryExtensions -contains $fileExtension) {
            $isBinaryFile = $true
            Write-Log "Detected binary file type: $fileExtension" -Level "INFO"
        }
        
        # Handle file based on type (binary or text)
        if ($isBinaryFile) {
            # For binary files, decode Base64 directly to bytes and write to file
            Write-Log "Processing as binary file" -Level "INFO"
            $bytes = [System.Convert]::FromBase64String($response.content)
            [System.IO.File]::WriteAllBytes($OutputPath, $bytes)
            Write-Log "Binary file written directly without text processing" -Level "INFO"
        } else {
            # For text files, use the original text processing approach
            Write-Log "Processing as text file" -Level "INFO"
            $content = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($response.content))
            
            # Clean content - remove BOM and invisible characters
            Write-Log "Cleaning content to remove any invisible characters" -Level "INFO"
            
            # Remove BOM and other invisible characters that might cause issues
            $content = $content -replace '^\xEF\xBB\xBF', '' # Remove UTF-8 BOM
            $content = $content -replace '^\xFE\xFF', ''     # Remove UTF-16 BE BOM
            $content = $content -replace '^\xFF\xFE', ''     # Remove UTF-16 LE BOM
            
            # Write content to file
            Write-Log "Writing content to $OutputPath" -Level "INFO"
            
            # Use UTF8 encoding without BOM
            $utf8NoBomEncoding = New-Object System.Text.UTF8Encoding $false
            [System.IO.File]::WriteAllText($OutputPath, $content, $utf8NoBomEncoding)
        }
        
        # Verify the download
        if (Test-Path $OutputPath) {
            $fileSize = (Get-Item $OutputPath).Length
            Write-Log "Download completed successfully. File size: $fileSize bytes" -Level "INFO"
            return $true
        } else {
            Write-Log "Download appeared to succeed but file not found" -Level "ERROR"
            return $false
        }
    }
    catch {
        $errorMessage = $_.ToString()
        
        # Try to extract HTTP status code from the error message
        $statusCode = $null
        if ($errorMessage -match "\((\d{3})\)") {
            $statusCode = $matches[1]
        } elseif ($errorMessage -match "(\d{3})") {
            $statusCode = $matches[1]
        }
        
        # Check if we're using a token
        $usingToken = -not [string]::IsNullOrWhiteSpace($Token)
        
        # Handle specific error codes with improved messages
        switch ($statusCode) {
            "404" {
                # Special handling for prerelease branch - don't show as error
                if ($Branch -eq "prerelease") {
                    Write-Log "Branch 'prerelease' not found in repository '$Username/$Repository'" -Level "INFO"
                    # Return a special value to indicate branch not found
                    return $false
                } else {
                    Write-Log "Repository or file not found (404): The repository '$Username/$Repository' or script '$ScriptName' does not exist" -Level "ERROR"
                }
            }
            "401" {
                if ($usingToken) {
                    Write-Log "Authentication failed (401): The GitHub token provided is invalid or expired" -Level "ERROR"
                } else {
                    Write-Log "Authentication required (401): This repository requires authentication. Set the GitHubToken variable." -Level "ERROR"
                }
            }
            "403" {
                if ($usingToken) {
                    Write-Log "Access forbidden (403): Your GitHub token doesn't have permission to access this repository, or you may be rate limited" -Level "ERROR"
                } else {
                    Write-Log "Access forbidden (403): You may be rate limited or the repository requires authentication. Set the GitHubToken variable." -Level "ERROR"
                }
            }
            default {
                Write-Log "Download error: $errorMessage" -Level "ERROR"
            }
        }
        
        return $false
    }
}

function Save-GitHubAsset {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [PSObject]$Asset,
        
        [Parameter(Mandatory = $true)]
        [string]$DestinationPath,
        
        [Parameter(Mandatory = $false)]
        [string]$Token = $Config.GitHubToken,
        
        [Parameter(Mandatory = $false)]
        [int]$RetryCount = 3
    )
    
    try {
        $AssetId = $Asset.id
        $AssetName = $Asset.name
        $DownloadUrl = "https://api.github.com/repos/$($Config.GitHubUsername)/$($Config.RepositoryName)/releases/assets/$AssetId"
        
        Write-Log "Downloading asset from: $DownloadUrl" -Level "INFO" -ShowInConsole $false
        Write-Log "Asset name: $AssetName, Asset ID: $AssetId" -Level "INFO"
        Write-Log "Destination path: $DestinationPath" -Level "INFO" -ShowInConsole $false
        
        # Set security protocol to TLS 1.2
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
        Write-Log "Set security protocol to TLS 1.2" -Level "INFO"
        
        # Try multiple download methods with retry
        $DownloadSuccess = $false
        $AttemptCount = 0
        $ErrorMessages = @()
        
        while (-not $DownloadSuccess -and $AttemptCount -lt $RetryCount) {
            $AttemptCount++
            Write-Log "Download attempt $AttemptCount of $RetryCount" -Level "INFO" -ShowInConsole $false
            
            try {
                # Primary download method using WebClient
                Write-Log "Using primary download method (WebClient)" -Level "INFO" -ShowInConsole $false
                $WebClient = New-Object System.Net.WebClient
                $WebClient.Headers.Add("Accept", "application/octet-stream")
                $WebClient.Headers.Add("User-Agent", "PowerShell Script")
                
                # Add authorization header if token is provided
                if (-not [string]::IsNullOrWhiteSpace($Token)) {
                    Write-Log "Adding authorization token to request headers" -Level "INFO"
                    $WebClient.Headers.Add("Authorization", "token $Token")
                }
                
                # Download the file
                Write-Log "Starting WebClient download to: $DestinationPath" -Level "INFO" -ShowInConsole $false
                $StartTime = Get-Date
                $WebClient.DownloadFile($DownloadUrl, $DestinationPath)
                $EndTime = Get-Date
                $DownloadTime = ($EndTime - $StartTime).TotalSeconds
                
                if (Test-Path -Path $DestinationPath) {
                    $FileSize = (Get-Item -Path $DestinationPath).Length
                    $FileSizeMB = [Math]::Round($FileSize / 1MB, 2)
                    $DownloadSpeedMBps = [Math]::Round($FileSizeMB / $DownloadTime, 2)
                    Write-Log "Successfully downloaded asset to $DestinationPath (Size: $FileSizeMB MB, Time: $([Math]::Round($DownloadTime, 2)) seconds, Speed: $DownloadSpeedMBps MB/s)" -Level "INFO" -ShowInConsole $false
                    $DownloadSuccess = $true
                    break
                } else {
                    Write-Log "Download completed but file not found at destination path" -Level "WARNING" -ShowInConsole $true
                }
            }
            catch {
                $ErrorMessages += "Primary download method failed: $_"
                Write-Log "Primary download method failed: $_" -Level "WARNING" -ShowInConsole $true
                
                try {
                    # Alternative download method using Invoke-WebRequest
                    Write-Log "Attempting alternative download method with Invoke-WebRequest..." -Level "INFO" -ShowInConsole $false
                    
                    # Initialize download headers
                    $DownloadHeaders = @{
                        "Accept" = "application/octet-stream"
                        "User-Agent" = "PowerShell Script"
                    }
                    
                    # Add authorization header if token is provided
                    if (-not [string]::IsNullOrWhiteSpace($Token)) {
                        Write-Log "Adding authorization token to request headers" -Level "INFO"
                        $DownloadHeaders["Authorization"] = "token $Token"
                    }
                    
                    # Use Invoke-WebRequest as fallback
                    Write-Log "Starting Invoke-WebRequest download to: $DestinationPath" -Level "INFO" -ShowInConsole $false
                    $StartTime = Get-Date
                    Invoke-WebRequest -Uri $DownloadUrl -Headers $DownloadHeaders -OutFile $DestinationPath -Method Get
                    $EndTime = Get-Date
                    $DownloadTime = ($EndTime - $StartTime).TotalSeconds
                    
                    if (Test-Path -Path $DestinationPath) {
                        $FileSize = (Get-Item -Path $DestinationPath).Length
                        $FileSizeMB = [Math]::Round($FileSize / 1MB, 2)
                        $DownloadSpeedMBps = [Math]::Round($FileSizeMB / $DownloadTime, 2)
                        Write-Log "Alternative download succeeded to $DestinationPath (Size: $FileSizeMB MB, Time: $([Math]::Round($DownloadTime, 2)) seconds, Speed: $DownloadSpeedMBps MB/s)" -Level "INFO" -ShowInConsole $false
                        $DownloadSuccess = $true
                        break
                    } else {
                        Write-Log "Alternative download completed but file not found at destination path" -Level "WARNING" -ShowInConsole $true
                    }
                }
                catch {
                    $ErrorMessages += "Alternative download method failed: $_"
                    Write-Log "Alternative download method failed: $_" -Level "WARNING" -ShowInConsole $true
                    
                    try {
                        # Try one more approach with direct browser download URL if available
                        Write-Log "Attempting direct browser download URL approach..." -Level "INFO" -ShowInConsole $false
                        
                        # Get the browser download URL from the asset
                        $BrowserDownloadUrl = $Asset.browser_download_url
                        if ($BrowserDownloadUrl) {
                            Write-Log "Using browser_download_url: $BrowserDownloadUrl" -Level "INFO" -ShowInConsole $false
                            
                            $DirectWebClient = New-Object System.Net.WebClient
                            $DirectWebClient.Headers.Add("User-Agent", "PowerShell Script")
                            
                            # Add authorization header if token is provided
                            if (-not [string]::IsNullOrWhiteSpace($Token)) {
                                Write-Log "Adding authorization token to direct request headers" -Level "INFO"
                                $DirectWebClient.Headers.Add("Authorization", "token $Token")
                            }
                            
                            Write-Log "Starting direct URL download to: $DestinationPath" -Level "INFO" -ShowInConsole $false
                            $StartTime = Get-Date
                            $DirectWebClient.DownloadFile($BrowserDownloadUrl, $DestinationPath)
                            $EndTime = Get-Date
                            $DownloadTime = ($EndTime - $StartTime).TotalSeconds
                            
                            if (Test-Path -Path $DestinationPath) {
                                $FileSize = (Get-Item -Path $DestinationPath).Length
                                $FileSizeMB = [Math]::Round($FileSize / 1MB, 2)
                                $DownloadSpeedMBps = [Math]::Round($FileSizeMB / $DownloadTime, 2)
                                Write-Log "Direct URL download succeeded (Size: $FileSizeMB MB, Time: $([Math]::Round($DownloadTime, 2)) seconds, Speed: $DownloadSpeedMBps MB/s)" -Level "INFO" -ShowInConsole $false
                                $DownloadSuccess = $true
                                break
                            } else {
                                Write-Log "Direct URL download completed but file not found at destination path" -Level "WARNING" -ShowInConsole $true
                            }
                        } else {
                            Write-Log "No browser_download_url available for direct download" -Level "WARNING" -ShowInConsole $true
                        }
                    }
                    catch {
                        $ErrorMessages += "Direct URL download failed: $_"
                        Write-Log "Direct URL download failed: $_" -Level "WARNING" -ShowInConsole $true
                    }
                }
            }
            finally {
                # Clean up WebClient objects
                if ($WebClient) { 
                    Write-Log "Disposing WebClient object" -Level "INFO"
                    $WebClient.Dispose() 
                }
                if ($DirectWebClient) { 
                    Write-Log "Disposing DirectWebClient object" -Level "INFO"
                    $DirectWebClient.Dispose() 
                }
            }
            
            # Wait before retrying
            if (-not $DownloadSuccess -and $AttemptCount -lt $RetryCount) {
                $WaitTime = [Math]::Pow(2, $AttemptCount) # Exponential backoff
            Write-Log "Waiting $WaitTime seconds before retry..." -Level "INFO" -ShowInConsole $false
                Start-Sleep -Seconds $WaitTime
            }
        }
        
        if ($DownloadSuccess) {
            Write-Log "Download completed successfully" -Level "INFO" -ShowInConsole $false
            # Verify file integrity
            $FileInfo = Get-Item -Path $DestinationPath
            Write-Log "Downloaded file details: Size=$([Math]::Round($FileInfo.Length / 1KB, 2)) KB, LastWriteTime=$($FileInfo.LastWriteTime)" -Level "INFO" -ShowInConsole $false
            
            return @{
                Success = $true
                FilePath = $DestinationPath
                FileSize = $FileInfo.Length
            }
        }
        else {
            Write-Log "All download methods failed after $RetryCount attempts" -Level "ERROR" -ShowInConsole $true
            return @{
                Success = $false
                ErrorMessage = "All download methods failed after $RetryCount attempts: $($ErrorMessages -join ' | ')"
            }
        }
    }
    catch {
        Write-Log "Error downloading GitHub asset: $_" -Level "ERROR" -ShowInConsole $true
        return @{
            Success = $false
            ErrorMessage = "Error downloading GitHub asset: $_"
        }
    }
}
#endregion

#region Installation Functions
function Test-Prerequisites {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string]$InstallType = $InstallType
    )
    
    try {
        Write-Log "Checking prerequisites for $InstallType installation" -Level "INFO" -ShowInConsole $false
        
        # Check PowerShell version
        $PSVersion = $PSVersionTable.PSVersion
        Write-Log "PowerShell version: $($PSVersion.Major).$($PSVersion.Minor)" -Level "INFO" -ShowInConsole $false
        
        if ($PSVersion.Major -lt 5) {
            Write-Log "PowerShell 5.1 or later is required" -Level "WARNING" -ShowInConsole $true
            return @{
                Success = $false
                ErrorMessage = "PowerShell 5.1 or later is required. Current version: $($PSVersion.Major).$($PSVersion.Minor)"
            }
        }
        
        # Check if running as administrator
        $IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        Write-Log "Running as administrator: $IsAdmin" -Level "INFO"
        
        # Log hash verification mode (skip for DirectAPI)
        if ($InstallType -ne "DirectAPI") {
            if ($Config.EnableHashVerification) {
                if ($Config.StrictHashMode) {
                    Write-Log "Hash verification enabled with strict mode (will fail if hash file is missing)" -Level "INFO" -ShowInConsole $false
                } else {
                    Write-Log "Hash verification enabled (will continue with warning if hash file is missing)" -Level "INFO" -ShowInConsole $false
                }
            } else {
                Write-Log "Hash verification disabled" -Level "INFO" -ShowInConsole $false
            }
        }
        
        # For online installation, check internet connectivity
        if ($InstallType -eq "Online") {
            Write-Host "> Checking internet connectivity..." -ForegroundColor Cyan
            $InternetAvailable = Test-InternetConnectivity
            
            if (-not $InternetAvailable) {
                Write-Host ""
                Write-Host "[ERROR] GitHub API is unavailable" -ForegroundColor Red
                Write-Host "Unable to connect to GitHub. Please check your internet connection." -ForegroundColor Red
                Write-Host ""
                
                return @{
                    Success = $false
                    ErrorMessage = "Unable to connect to GitHub. Please check your internet connection."
                    SuggestOffline = $true
                }
            }
            
            Write-Host "      Internet connectivity verified" -ForegroundColor Gray
            Write-Host "      [OK]" -ForegroundColor Green
            Write-Host ""
        }
        
        # Check if installation directory can be created/accessed
        try {
            if (-not (Test-Path -Path $Config.InstallDir)) {
                New-Item -Path $Config.InstallDir -ItemType Directory -Force | Out-Null
                Write-Log "Created installation directory: $($Config.InstallDir)" -Level "INFO" -ShowInConsole $false
            } else {
                Write-Log "Installation directory already exists: $($Config.InstallDir)" -Level "INFO"
            }
            
            # Test write access by creating a temporary file
            $TestFile = Join-Path -Path $Config.InstallDir -ChildPath "write_test.tmp"
            "Test" | Out-File -FilePath $TestFile -Force
            Remove-Item -Path $TestFile -Force
            Write-Log "Installation directory is writable" -Level "INFO"
        }
        catch {
            Write-Log "Cannot write to installation directory: $_" -Level "ERROR" -ShowInConsole $true
            return @{
                Success = $false
                ErrorMessage = "Cannot write to installation directory: $_"
            }
        }
        
        return @{
            Success = $true
        }
    }
    catch {
        Write-Log "Error checking prerequisites: $_" -Level "ERROR" -ShowInConsole $true
        return @{
            Success = $false
            ErrorMessage = "Error checking prerequisites: $_"
        }
    }
}

function Install-OfflineApp {
    [CmdletBinding()]
    param()
    
    Write-Log "Starting offline installation" -Level "INFO"
    
    # Define total steps for offline installation
    $TotalSteps = 3
    
    try {
        # Step 1: Validate installation files
        Show-ProgressStep -StepNumber 1 -TotalSteps $TotalSteps -StepDescription "Validating installation files..."
        
        # Look for the script file directly
        Write-Log "Looking for script file in the current directory..." -Level "INFO"
        $SourceScript = Join-Path -Path $PSScriptRoot -ChildPath $Config.ScriptName
        
        if (-not (Test-Path -Path $SourceScript)) {
            Write-Log "Script file $($Config.ScriptName) not found in the current directory" -Level "INFO"
            
            # Try looking two directories up (for repository clones where installer is in tools/installer)
            $ParentDir = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
            $ParentScript = Join-Path -Path $ParentDir -ChildPath $Config.ScriptName
            
            if (Test-Path -Path $ParentScript) {
                Write-Log "Found script file in parent directory: $ParentScript" -Level "INFO"
                $SourceScript = $ParentScript
            } else {
                Write-Log "Script file $($Config.ScriptName) not found in parent directory" -Level "INFO"
            }
        }
        
        # If script found, use it directly (no hash verification needed)
        if (Test-Path -Path $SourceScript) {
            Write-Log "Found script file: $SourceScript" -Level "INFO"
            $SourceType = "ps1"
            $SourcePath = $SourceScript
            
            Show-StepResult -Message "Found script file: $($Config.ScriptName)" -Status "OK"
            
            # Unblock the script file before copying
            Unblock-InstalledFiles -Stage "BeforeCopy" -InstallDirectory $Config.InstallDir -SourcePath $SourcePath
        }
        # If no script file found, look for a zip file
        else {
            Write-Log "No script file found, looking for zip file..." -Level "INFO"
            $ZipFiles = Get-ChildItem -Path $PSScriptRoot -Filter $Config.AssetPattern
            $SourceZip = if ($ZipFiles.Count -gt 0) { $ZipFiles[0].FullName } else { $null }
            
            if (Test-Path -Path $SourceZip) {
                Write-Log "Found zip file: $SourceZip" -Level "INFO"
                $SourceType = "zip"
                $SourcePath = $SourceZip
                
                # Get file size for display
                $FileSize = [Math]::Round((Get-Item -Path $SourceZip).Length / 1MB, 2)
                Show-StepResult -Message "Found: $([System.IO.Path]::GetFileName($SourceZip))" -Status "OK"
                Show-StepResult -Message "Size: $FileSize MB" -Status "OK"
            }
            else {
                Show-StepResult -Message "No installation files (script or zip) found" -Status "ERROR"
                Write-Log "Error: No installation files (script or zip) found in the current directory" -Level "ERROR"
                return $false
            }
        }
        
        # Create destination directory if it doesn't exist
        if (!(Test-Path -Path $Config.InstallDir)) {
            Write-Log "Creating directory: $($Config.InstallDir)" -Level "INFO"
            New-Item -Path $Config.InstallDir -ItemType Directory -Force | Out-Null
        }
        
        # Process based on source type
        if ($SourceType -eq "ps1") {
            # Step 2: Copy files
            Show-ProgressStep -StepNumber 2 -TotalSteps $TotalSteps -StepDescription "Copying files..."
            
            # Copy files based on file management configuration
            Write-Log "Copying files based on file management configuration" -Level "INFO"
            $SourceDir = Split-Path -Parent $SourcePath
            $FilesProcessed = 0
            
            # Process required files
            foreach ($FileConfig in $Config.FileManagement.RequiredFiles) {
                $FilePattern = $FileConfig.Pattern
                
                # First check parent directory
                $MatchingFiles = Get-ChildItem -Path $SourceDir -Filter $FilePattern -ErrorAction SilentlyContinue
                
                # If not found in parent directory, check current directory (where installer is)
                if (-not $MatchingFiles -or $MatchingFiles.Count -eq 0) {
                    $MatchingFiles = Get-ChildItem -Path $PSScriptRoot -Filter $FilePattern -ErrorAction SilentlyContinue
                    if ($MatchingFiles) {
                        Write-Log "Found $FilePattern in installer directory" -Level "INFO"
                    }
                }
                
                if ($MatchingFiles) {
                    foreach ($File in $MatchingFiles) {
                        Write-Log "Copying required file: $($File.Name)" -Level "INFO"
                        Copy-Item -Path $File.FullName -Destination $Config.InstallDir -Force
                        Write-Log "Copied $($File.Name)" -Level "INFO"
                        $FilesProcessed++
                    }
                } else {
                    if ($FilePattern -eq $Config.ScriptName) {
                        # Special case for main script which we already found
                        Write-Log "Copying main script: $($Config.ScriptName)" -Level "INFO"
                        Copy-Item -Path $SourcePath -Destination $Config.InstallDir -Force
                        Write-Log "Copied $($Config.ScriptName)" -Level "INFO"
                        $FilesProcessed++
                    } else {
                        Write-Log "No files matching required pattern '$FilePattern' found" -Level "WARNING"
                    }
                }
            }
            
            # Process optional files
            foreach ($FileConfig in $Config.FileManagement.OptionalFiles) {
                $FilePattern = $FileConfig.Pattern
                
                $MatchingFiles = Get-ChildItem -Path $SourceDir -Filter $FilePattern -ErrorAction SilentlyContinue
                
                if ($MatchingFiles) {
                    foreach ($File in $MatchingFiles) {
                        Write-Log "Copying optional file: $($File.Name)" -Level "INFO"
                        Copy-Item -Path $File.FullName -Destination $Config.InstallDir -Force
                        Write-Log "Copied $($File.Name)" -Level "INFO"
                        $FilesProcessed++
                    }
                } else {
                    Write-Log "No files matching optional pattern '$FilePattern' found" -Level "INFO"
                }
            }
            
            Write-Log "Processed $FilesProcessed files during installation" -Level "INFO"
            
            # Unblock the copied script and module files
            Unblock-InstalledFiles -Stage "AfterCopy" -InstallDirectory $Config.InstallDir
            
            # Set the icon path to the already copied icon file
            $IconPath = Join-Path -Path $Config.InstallDir -ChildPath $Config.IconName
            if (-not (Test-Path -Path $IconPath)) {
                Write-Log "Icon file not found at: $IconPath" -Level "INFO"
                $IconPath = $null
            } else {
                Write-Log "Using already copied icon file: $IconPath" -Level "INFO"
            }
            
            Show-StepResult -Message "Destination: $($Config.InstallDir)" -Status "OK"
            
            # Step 3: Create desktop shortcut
            Show-ProgressStep -StepNumber 3 -TotalSteps $TotalSteps -StepDescription "Creating desktop shortcut..."
            
            # Create shortcut
            $ScriptPath = Join-Path -Path $Config.InstallDir -ChildPath $Config.ScriptName
            $ShortcutPath = Join-Path -Path $Config.DesktopPath -ChildPath "$($Config.ShortcutName).lnk"
            
            # Create PowerShell shortcut to execute the script
            $PowerShellExe = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
            $Arguments = "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ScriptPath`""
            
            # Only pass IconLocation if an icon file was found
            if ($IconPath -and (Test-Path -Path $IconPath)) {
                $ShortcutCreated = New-Shortcut -TargetPath $PowerShellExe -ShortcutPath $ShortcutPath -Description $Config.ShortcutDescription -Arguments $Arguments -WorkingDirectory $Config.InstallDir -IconLocation $IconPath
            } else {
                $ShortcutCreated = New-Shortcut -TargetPath $PowerShellExe -ShortcutPath $ShortcutPath -Description $Config.ShortcutDescription -Arguments $Arguments -WorkingDirectory $Config.InstallDir
            }
            
            Show-StepResult -Message "Found script: $($Config.ScriptName)" -Status "OK"
            Show-StepResult -Message "Created shortcut at: $ShortcutPath" -Status "OK"
            
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
            # Verify hash if enabled and hash file exists (only applies to zip files)
            if ($Config.EnableHashVerification) {
                # For offline installation, we check if a hash file exists for the zip file
                $HashFilePath = "$SourcePath.sha256"
                
                # Also look for hash files using the pattern in case the hash file doesn't exactly match the zip name
                if (-not (Test-Path -Path $HashFilePath)) {
                    $HashFiles = Get-ChildItem -Path $PSScriptRoot -Filter $Config.HashFilePattern
                    if ($HashFiles.Count -gt 0) {
                        $HashFilePath = $HashFiles[0].FullName
                        Write-Log "Found hash file using pattern: $HashFilePath" -Level "INFO"
                    }
                }
                
                # Verify the hash
                $HashResult = Test-FileHash -FilePath $SourcePath -HashFilePath $HashFilePath -StrictMode $Config.StrictHashMode
                
                if (-not $HashResult.Success -and $HashResult.Status -eq "ERROR") {
                    return $false
                }
            }
            else {
                Write-Log "Hash verification disabled, skipping" -Level "INFO"
                Show-SecurityVerification -Status "DISABLED"
            }
            
            # Step 2: Extract files
            Show-ProgressStep -StepNumber 2 -TotalSteps $TotalSteps -StepDescription "Extracting files..."
            
            # Extract zip file to destination
            $ExtractResult = Expand-ZipFile -ZipPath $SourcePath -DestinationPath $Config.InstallDir -CleanDestination
            
            if (-not $ExtractResult) {
                Show-StepResult -Message "Failed to extract zip file" -Status "ERROR"
                return $false
            }
            
            Show-StepResult -Message "Destination: $($Config.InstallDir)" -Status "OK"
            
            # Unblock all script files in the installation directory
            Unblock-InstalledFiles -Stage "AfterExtract" -InstallDirectory $Config.InstallDir
            
            # Reorganize files - move script and LICENSE from subdirectory to main directory
            Move-ExtractedFiles -ExtractedPath $Config.InstallDir -MainScriptName $Config.ScriptName
            
            # Step 3: Create desktop shortcut
            Show-ProgressStep -StepNumber 3 -TotalSteps $TotalSteps -StepDescription "Creating desktop shortcut..."
            
            # Set the icon path to the already copied icon file
            $IconPath = Join-Path -Path $Config.InstallDir -ChildPath $Config.IconName
            if (-not (Test-Path -Path $IconPath)) {
                Write-Log "Icon file not found at: $IconPath" -Level "INFO"
                $IconPath = $null
            } else {
                Write-Log "Using already copied icon file: $IconPath" -Level "INFO"
            }
            
            # Create shortcut
            $ScriptFiles = Get-ChildItem -Path $Config.InstallDir -Filter "*.ps1" -Recurse
            
            if ($ScriptFiles.Count -eq 0) {
                Write-Log "No PowerShell script found in the extracted files" -Level "ERROR"
                Show-StepResult -Message "No PowerShell script found in the extracted files" -Status "ERROR"
                return $false
            }
            
            # Look for the script file by name
            $MainScript = $ScriptFiles | Where-Object { $_.Name -eq $Config.ScriptName } | Select-Object -First 1
            
            # If not found, log an error
            if ($null -eq $MainScript) {
                Write-Log "ERROR: Could not find $($Config.ScriptName) in the installation directory" -Level "ERROR"
                Write-Log "Available scripts: $($ScriptFiles.Name -join ', ')" -Level "INFO"
                Show-StepResult -Message "Could not find $($Config.ScriptName) in the installation directory" -Status "ERROR"
                return $false
            }
            else {
                Write-Log "Found main script: $($MainScript.FullName)" -Level "INFO"
                Show-StepResult -Message "Found script: $($MainScript.Name)" -Status "OK"
            }
            
            $ShortcutPath = Join-Path -Path $Config.DesktopPath -ChildPath "$($Config.ShortcutName).lnk"
            
            # Create PowerShell shortcut to execute the script
            $PowerShellExe = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
            $Arguments = "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$($MainScript.FullName)`""
            $WorkingDirectory = Split-Path -Parent $MainScript.FullName
            
            # Only pass IconLocation if an icon file was found
            if ($IconPath -and (Test-Path -Path $IconPath)) {
                $ShortcutCreated = New-Shortcut -TargetPath $PowerShellExe -ShortcutPath $ShortcutPath -Description $Config.ShortcutDescription -Arguments $Arguments -WorkingDirectory $WorkingDirectory -IconLocation $IconPath
            } else {
                $ShortcutCreated = New-Shortcut -TargetPath $PowerShellExe -ShortcutPath $ShortcutPath -Description $Config.ShortcutDescription -Arguments $Arguments -WorkingDirectory $WorkingDirectory
            }
            
            if ($ShortcutCreated) {
                Write-Log "Offline installation completed successfully" -Level "INFO"
                Show-StepResult -Message "Created shortcut at: $ShortcutPath" -Status "OK"
                return $true
            }
            else {
                Write-Log "Failed to create shortcut" -Level "ERROR"
                Show-StepResult -Message "Failed to create shortcut" -Status "ERROR"
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

function Install-DirectAPIApp {
    [CmdletBinding()]
    param()
    
    Write-Log "Starting SecurePassGenerator direct API installation" -Level "INFO"
    
    # Define total steps for direct API installation
    $TotalSteps = 4
    
    try {
        # Determine which branch to use based on ReleaseType
        $Branch = if ($ReleaseType -eq "PreRelease") { "prerelease" } else { "main" }
        Write-Log "Using branch: $Branch for $ReleaseType release" -Level "INFO"
        
        # Step 1: Check internet connectivity
        Show-ProgressStep -StepNumber 1 -TotalSteps $TotalSteps -StepDescription "Checking connectivity..."
        
        Write-Host "> Checking internet connectivity..." -ForegroundColor Cyan
        $InternetAvailable = Test-InternetConnectivity
        
        if (-not $InternetAvailable) {
            Write-Host ""
            Write-Host "[WARNING] GitHub API connectivity check failed" -ForegroundColor Yellow
            Write-Host "This may be due to proxy restrictions or firewall settings." -ForegroundColor Yellow
            Write-Host "The Direct API method may still work in some proxy environments." -ForegroundColor Yellow
            Write-Host ""
            
            $continueChoice = Read-Host "Would you like to continue anyway? (C)ontinue or (A)bort [Default: Continue]"
            
            if ($continueChoice -eq "A" -or $continueChoice -eq "a") {
                Write-Log "User chose to abort after connectivity check failed" -Level "INFO"
                Show-StepResult -Message "Installation aborted by user" -Status "ERROR"
                return $false
            } else {
                Write-Log "User chose to continue despite connectivity check failure" -Level "INFO"
                Show-StepResult -Message "Continuing despite connectivity check failure" -Status "WARNING"
            }
        } else {
            Show-StepResult -Message "Internet connectivity verified" -Status "OK"
        }
        
        Write-Host ""
        
        # Step 2: Fetch release information
        Show-ProgressStep -StepNumber 2 -TotalSteps $TotalSteps -StepDescription "Fetching release information..."
        
        Show-StepResult -Message "Repository: $($Config.GitHubUsername)/$($Config.RepositoryName)" -Status "OK"
        Show-StepResult -Message "Branch: $Branch ($ReleaseType)" -Status "OK"
        
        # Step 3: Download script directly using GitHub API
        Show-ProgressStep -StepNumber 3 -TotalSteps $TotalSteps -StepDescription "Downloading script using direct API..."
        
        Write-Host ""
        
        # Create destination directory if it doesn't exist
        if (!(Test-Path -Path $Config.InstallDir)) {
            Write-Log "Creating directory: $($Config.InstallDir)" -Level "INFO"
            New-Item -Path $Config.InstallDir -ItemType Directory -Force | Out-Null
        }
        
        # Download script from GitHub
        $ScriptPath = Join-Path -Path $Config.InstallDir -ChildPath $Config.ScriptName
        Write-Log "Downloading $($Config.ScriptName) from GitHub to $ScriptPath (branch: $Branch)" -Level "INFO"
        
        $DownloadSuccess = Invoke-GitHubDownload -OutputPath $ScriptPath -Branch $Branch
        
        # If download fails and we're using the prerelease branch, try falling back to main branch
        if (-not $DownloadSuccess -and $Branch -eq "prerelease") {
            Write-Log "Branch 'prerelease' not found, falling back to 'main' branch" -Level "WARNING" -ShowInConsole $true
            
            # Try downloading from main branch instead
            $FallbackBranch = "main"
            Write-Log "Attempting download from fallback branch: $FallbackBranch" -Level "INFO"
            $DownloadSuccess = Invoke-GitHubDownload -OutputPath $ScriptPath -Branch $FallbackBranch
            
            # Show repository info with OK status instead of warning
            Show-StepResult -Message "Repository: $($Config.GitHubUsername)/$($Config.RepositoryName)" -Status "OK"
        }
        
        if (-not $DownloadSuccess) {
            Write-Log "Failed to download script from GitHub" -Level "ERROR"
            Show-StepResult -Message "Failed to download script from GitHub" -Status "ERROR"
            
            # Check if we have a local copy to use as fallback
            $LocalScript = Join-Path -Path $PSScriptRoot -ChildPath $Config.ScriptName
            if (Test-Path -Path $LocalScript) {
                Write-Log "Using local copy as fallback" -Level "WARNING"
                Show-StepResult -Message "Using local copy as fallback" -Status "WARNING"
                Copy-Item -Path $LocalScript -Destination $ScriptPath -Force
            } else {
                return $false
            }
        } else {
            # Get file size for display
            $FileSize = [Math]::Round((Get-Item -Path $ScriptPath).Length / 1KB, 2)
            Show-StepResult -Message "Downloaded: $($Config.ScriptName)" -Status "OK"
            Show-StepResult -Message "Size: $FileSize KB" -Status "OK"
            Show-StepResult -Message "Destination: $($Config.InstallDir)" -Status "OK"
        }
        
        # Unblock the downloaded script
        Unblock-FileIfNeeded -Path $ScriptPath
        
        # Step 4: Create desktop shortcut
        Show-ProgressStep -StepNumber 4 -TotalSteps $TotalSteps -StepDescription "Creating desktop shortcut..."
        
        # Create shortcut
        $ShortcutPath = Join-Path -Path $Config.DesktopPath -ChildPath "$($Config.ShortcutName).lnk"
        
        # Create PowerShell shortcut to execute the script
        $PowerShellExe = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
        $Arguments = "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ScriptPath`""
        
        # Try to find icon file
        $IconPath = Join-Path -Path $Config.InstallDir -ChildPath $Config.IconName
        
        # If icon doesn't exist, try to download it
        if (-not (Test-Path -Path $IconPath)) {
            Write-Log "Icon file not found, attempting to download it" -Level "INFO"
            $IconDownloadSuccess = Invoke-GitHubDownload -OutputPath $IconPath -ScriptName $Config.IconName -Branch $Branch
            
            # If download fails and we're using the prerelease branch, try falling back to main branch
            if (-not $IconDownloadSuccess -and $Branch -eq "prerelease") {
                Write-Log "Failed to download icon from 'prerelease' branch, trying 'main' branch" -Level "INFO" -ShowInConsole $false
                $FallbackBranch = "main"
                $IconDownloadSuccess = Invoke-GitHubDownload -OutputPath $IconPath -ScriptName $Config.IconName -Branch $FallbackBranch
            }
            
            if (-not $IconDownloadSuccess) {
                Write-Log "Failed to download icon file, will use default PowerShell icon" -Level "WARNING"
                $IconPath = $null
            } else {
                Write-Log "Successfully downloaded icon file" -Level "INFO"
            }
        }
        
        # Only pass IconLocation if an icon file was found
        if ($IconPath -and (Test-Path -Path $IconPath)) {
            $ShortcutCreated = New-Shortcut -TargetPath $PowerShellExe -ShortcutPath $ShortcutPath -Description $Config.ShortcutDescription -Arguments $Arguments -WorkingDirectory $Config.InstallDir -IconLocation $IconPath
        } else {
            $ShortcutCreated = New-Shortcut -TargetPath $PowerShellExe -ShortcutPath $ShortcutPath -Description $Config.ShortcutDescription -Arguments $Arguments -WorkingDirectory $Config.InstallDir
        }
        
        if ($ShortcutCreated) {
            Write-Log "Direct API installation completed successfully" -Level "INFO"
            Show-StepResult -Message "Created shortcut at: $ShortcutPath" -Status "OK"
            return $true
        }
        else {
            Write-Log "Failed to create shortcut" -Level "ERROR"
            Show-StepResult -Message "Failed to create shortcut" -Status "ERROR"
            return $false
        }
    }
    catch {
        Write-Log "Error during direct API installation: $_" -Level "ERROR"
        return $false
    }
}

function Install-OnlineApp {
    [CmdletBinding()]
    param()
    
    Write-Log "Starting SecurePassGenerator online installation" -Level "INFO"
    
    # Define total steps for online installation
    $TotalSteps = 4
    
    try {
        # Step 1: Get the latest release information
        Show-ProgressStep -StepNumber 1 -TotalSteps $TotalSteps -StepDescription "Fetching release information..."
        Write-Log "Fetching latest release information from GitHub API" -Level "INFO"
        
        $ReleaseResult = Get-GitHubReleaseInfo -ReleaseType $ReleaseType
        
        if (-not $ReleaseResult.Success) {
            Write-Log "Failed to get release information: $($ReleaseResult.ErrorMessage)" -Level "ERROR"
            return $false
        }
        
        $ReleaseInfo = $ReleaseResult.ReleaseInfo
        
        # Find the asset matching our pattern
        $AssetResult = Get-GitHubAsset -ReleaseInfo $ReleaseInfo
        
        if (-not $AssetResult.Success) {
            Write-Log "Failed to find asset: $($AssetResult.ErrorMessage)" -Level "ERROR"
            return $false
        }
        
        $Asset = $AssetResult.Asset
        $HashAsset = $AssetResult.HashAsset
        $AssetName = $Asset.name
        
        # Update the temp zip path with the actual asset name
        $TempZipPath = Join-Path -Path $env:TEMP -ChildPath $AssetName
        
        # Display repository and release information
        Show-StepResult -Message "Repository: $($Config.GitHubUsername)/$($Config.RepositoryName)" -Status "OK"
        Show-StepResult -Message "Release: $($ReleaseInfo.tag_name) ($($ReleaseResult.ActualReleaseType))" -Status "OK"
        
        # Step 2: Download the asset
        Show-ProgressStep -StepNumber 2 -TotalSteps $TotalSteps -StepDescription "Downloading installation files..."
        
        $DownloadResult = Save-GitHubAsset -Asset $Asset -DestinationPath $TempZipPath
        
        if (-not $DownloadResult.Success) {
            Write-Log "Failed to download asset: $($DownloadResult.ErrorMessage)" -Level "ERROR"
            return $false
        }
        
        $FileSize = [Math]::Round($DownloadResult.FileSize / 1MB, 2)
        Show-StepResult -Message "Downloaded: $AssetName" -Status "OK"
        Show-StepResult -Message "Size: $FileSize MB" -Status "OK"
        
        # Step 2.5: Verify file hash if enabled
        if ($Config.EnableHashVerification) {
            Write-Log "Hash verification enabled, looking for hash file" -Level "INFO"
            
            if ($HashAsset) {
                # Download the hash file using direct WebClient approach (matching original script)
                $HashAssetId = $HashAsset.id
                $HashDownloadUrl = "https://api.github.com/repos/$($Config.GitHubUsername)/$($Config.RepositoryName)/releases/assets/$HashAssetId"
                $TempHashPath = Join-Path -Path $env:TEMP -ChildPath "$AssetName.sha256"
                
                Write-Log "Downloading hash file from: $HashDownloadUrl" -Level "INFO"
                
                try {
                    # Download the hash file
                    $WebClient = New-Object System.Net.WebClient
                    $WebClient.Headers.Add("Accept", "application/octet-stream")
                    $WebClient.Headers.Add("User-Agent", "PowerShell Script")
                    
                    # Add authorization header if token is provided
                    if (-not [string]::IsNullOrWhiteSpace($Config.GitHubToken)) {
                        $WebClient.Headers.Add("Authorization", "token $($Config.GitHubToken)")
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
                            
                            # Display security verification success
                            Show-SecurityVerification -Hash $ExpectedHashUpper -Status "OK"
                        }
                        else {
                            Write-Log "Hash verification failed! The downloaded file may be corrupted or tampered with." -Level "ERROR"
                            Write-Log "Expected: $ExpectedHashUpper" -Level "ERROR"
                            Write-Log "Actual: $($ActualHash.Hash)" -Level "ERROR"
                            
                            # Display security verification failure
                            Show-SecurityVerification -Hash $ExpectedHashUpper -Status "ERROR"
                            return $false
                        }
                    }
                    else {
                        Write-Log "Failed to download hash file, skipping verification" -Level "WARNING"
                        
                        if ($Config.StrictHashMode) {
                            Write-Log "Strict hash verification is enabled, cannot proceed without hash file" -Level "ERROR"
                            Show-SecurityVerification -Status "ERROR"
                            return $false
                        }
                        else {
                            # Display security verification skipped
                            Show-SecurityVerification -Status "SKIPPED"
                        }
                    }
                }
                catch {
                    Write-Log "Error during hash verification: $_" -Level "WARNING"
                    
                    if ($Config.StrictHashMode) {
                        Write-Log "Strict hash verification is enabled, cannot proceed with failed hash verification" -Level "ERROR"
                        Show-SecurityVerification -Status "ERROR"
                        return $false
                    }
                    else {
                        Write-Log "Continuing installation without hash verification" -Level "WARNING"
                        # Display security verification skipped
                        Show-SecurityVerification -Status "SKIPPED"
                    }
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
            else {
                if ($Config.StrictHashMode) {
                    Write-Log "No hash file found for asset '$AssetName' and strict hash verification is enabled" -Level "ERROR"
                    Write-Log "Installation cannot proceed without hash verification in strict mode" -Level "ERROR"
                    
                    # Display security verification failure
                    Show-SecurityVerification -Status "ERROR"
                    return $false
                } else {
                    Write-Log "No hash file found for asset '$AssetName', skipping verification" -Level "WARNING"
                    
                    # Display security verification skipped
                    Show-SecurityVerification -Status "SKIPPED"
                }
            }
        }
        else {
            Write-Log "Hash verification disabled, skipping" -Level "INFO"
            Show-SecurityVerification -Status "DISABLED"
        }
        
        # Step 3: Extract the zip file
        Show-ProgressStep -StepNumber 3 -TotalSteps $TotalSteps -StepDescription "Extracting files..."
        
            # Extract zip file to destination
            $ExtractResult = Expand-ZipFile -ZipPath $TempZipPath -DestinationPath $Config.InstallDir -CleanDestination
        
        if (-not $ExtractResult) {
            Show-StepResult -Message "Failed to extract zip file" -Status "ERROR"
            return $false
        }
        
        Show-StepResult -Message "Destination: $($Config.InstallDir)" -Status "OK"
        
        # Unblock all script files in the installation directory
        Unblock-InstalledFiles -Stage "AfterExtract" -InstallDirectory $Config.InstallDir
        
            # Reorganize files - move script and LICENSE from subdirectory to main directory
            Move-ExtractedFiles -ExtractedPath $Config.InstallDir -MainScriptName $Config.ScriptName
        
        # Step 4: Create desktop shortcut
        Show-ProgressStep -StepNumber 4 -TotalSteps $TotalSteps -StepDescription "Creating desktop shortcut..."
        
        # Set the icon path to the already copied icon file
        $IconPath = Join-Path -Path $Config.InstallDir -ChildPath $Config.IconName
        if (-not (Test-Path -Path $IconPath)) {
            Write-Log "Icon file not found at: $IconPath" -Level "INFO"
            $IconPath = $null
        } else {
            Write-Log "Using already copied icon file: $IconPath" -Level "INFO"
        }
        
        # Create shortcut
        $ScriptFiles = Get-ChildItem -Path $Config.InstallDir -Filter "*.ps1" -Recurse
        
        if ($ScriptFiles.Count -eq 0) {
            Write-Log "No PowerShell script found in the extracted files" -Level "ERROR"
            Show-StepResult -Message "No PowerShell script found in the extracted files" -Status "ERROR"
            return $false
        }
        
        # Look for the script file by name
        $MainScript = $ScriptFiles | Where-Object { $_.Name -eq $Config.ScriptName } | Select-Object -First 1
        
        # If not found, log an error
        if ($null -eq $MainScript) {
            Write-Log "ERROR: Could not find $($Config.ScriptName) in the installation directory" -Level "ERROR"
            Write-Log "Available scripts: $($ScriptFiles.Name -join ', ')" -Level "INFO"
            Show-StepResult -Message "Could not find $($Config.ScriptName) in the installation directory" -Status "ERROR"
            return $false
        }
        else {
            Write-Log "Found main script: $($MainScript.FullName)" -Level "INFO"
            Show-StepResult -Message "Found script: $($MainScript.Name)" -Status "OK"
        }
        
        $ShortcutPath = Join-Path -Path $Config.DesktopPath -ChildPath "$($Config.ShortcutName).lnk"
        
        # Create PowerShell shortcut to execute the script
        $PowerShellExe = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
        $Arguments = "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$($MainScript.FullName)`""
        $WorkingDirectory = Split-Path -Parent $MainScript.FullName
        
        # Only pass IconLocation if an icon file was found
        if ($IconPath -and (Test-Path -Path $IconPath)) {
            $ShortcutCreated = New-Shortcut -TargetPath $PowerShellExe -ShortcutPath $ShortcutPath -Description $Config.ShortcutDescription -Arguments $Arguments -WorkingDirectory $WorkingDirectory -IconLocation $IconPath
        } else {
            $ShortcutCreated = New-Shortcut -TargetPath $PowerShellExe -ShortcutPath $ShortcutPath -Description $Config.ShortcutDescription -Arguments $Arguments -WorkingDirectory $WorkingDirectory
        }
        
        if (-not $ShortcutCreated) {
            Write-Log "Failed to create desktop shortcut" -Level "ERROR"
            Show-StepResult -Message "Failed to create desktop shortcut" -Status "ERROR"
            return $false
        }
        
        Show-StepResult -Message "Created shortcut at: $ShortcutPath" -Status "OK"
        
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

#region Main Script Execution
# Check if script is being run without parameters by checking if InstallType is null or empty
if ([string]::IsNullOrEmpty($InstallType)) {
    Show-InstallerHelp
}

# Initialize log file
Write-Log "======================================================" -Level "INFO"
Write-Log "Installer started with InstallType: $InstallType, ReleaseType: $ReleaseType" -Level "INFO"
Write-Log "Script location: $PSScriptRoot" -Level "INFO"
Write-Log "User: $env:USERNAME" -Level "INFO"
Write-Log "Computer: $env:COMPUTERNAME" -Level "INFO"
Write-Log "======================================================" -Level "INFO"

# Display banner and installation details
Show-InstallerBanner
Show-InstallationDetails -InstallType $InstallType -ReleaseType $ReleaseType

# Check prerequisites
$PrereqResult = Test-Prerequisites -InstallType $InstallType
if (-not $PrereqResult.Success) {
    if ($PrereqResult.SuggestOffline) {
        # Ask if user wants to run offline installation instead
        $offlineChoice = Read-Host "Would you like to run the Offline installation instead? (Y/n) [Default: Y]"
        
        if ([string]::IsNullOrEmpty($offlineChoice) -or $offlineChoice.ToLower() -eq "y") {
            Write-Log "Switching to offline installation after connectivity check failed" -Level "INFO"
            $InstallType = "Offline"
        }
        else {
            Write-Log "User declined offline installation after connectivity check failed" -Level "INFO"
            Show-CompletionMessage -Success $false -ErrorMessage $PrereqResult.ErrorMessage
            exit 1
        }
    }
    else {
        Show-CompletionMessage -Success $false -ErrorMessage $PrereqResult.ErrorMessage
        exit 1
    }
}

# Execute installation based on type
$Success = $false

switch ($InstallType) {
    "Offline" {
        $Success = Install-OfflineApp
    }
    "Online" {
        $Success = Install-OnlineApp
    }
    "DirectAPI" {
        $Success = Install-DirectAPIApp
    }
}

# Display final message
if ($Success) {
    # Find the script path for the completion message
    $ScriptPath = ""
    $ScriptFiles = Get-ChildItem -Path $Config.InstallDir -Filter "*.ps1" -Recurse
    $MainScript = $ScriptFiles | Where-Object { $_.Name -eq $Config.ScriptName } | Select-Object -First 1
    if ($MainScript) {
        $ScriptPath = $MainScript.FullName
    }
    
    Show-CompletionMessage -Success $true -ScriptPath $ScriptPath
    
    # Add a brief pause before exiting to allow user to see the completion message
    Write-Host ""
    Write-Host "Installation completed successfully. Exiting in 3 seconds..." -ForegroundColor Green
    Start-Sleep -Seconds 3
    
    exit 0
}
else {
    Show-CompletionMessage -Success $false
    
    # Add a brief pause before exiting to allow user to see the error message
    Write-Host ""
    Write-Host "Installation failed. Exiting in 3 seconds..." -ForegroundColor Red
    Start-Sleep -Seconds 3
    
    exit 1
}
#endregion
