#==============================================================
# Installer.ps1
# Interactive installer for SecurePassGenerator
#
# Options:
# -InstallType Offline: Installs SecurePassGenerator.ps1 from local files
# -InstallType Online: Downloads and installs SecurePassGenerator.ps1 from GitHub
#==============================================================

param (
    [Parameter(Mandatory=$true)]
    [ValidateSet("Offline", "Online")]
    [string]$InstallType
)

# Script configuration variables
$scriptName = "SecurePassGenerator.ps1"
$appFolder = "securepassgenerator"
$shortcutName = "SecurePassGenerator"
$shortcutDescription = "Generate secure passwords with a modern GUI"

# GitHub configuration for online installation
$gitHubUsername = "onlyalex1984"
$repositoryName = "securepassgenerator-powershell"
$gitHubToken = "" # Leave empty for public repositories

# Initialize logging
$logFile = Join-Path -Path $PSScriptRoot -ChildPath "Installer.log"

function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO",
        [switch]$Silent = $false
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    # Write to log file
    Add-Content -Path $logFile -Value $logEntry
    
    # Also write to console if not silent
    if (-not $Silent) {
        Write-Host $logEntry
    }
}

function Create-Shortcut {
    param (
        [string]$TargetPath,
        [string]$ShortcutName,
        [string]$Arguments = "",
        [string]$Description = "",
        [string]$WorkingDirectory = ""
    )
    
    try {
        Write-Log "Creating shortcut: $ShortcutName pointing to $TargetPath"
        
        # Create a temporary VBScript file to create the shortcut
        $vbsFile = Join-Path -Path $env:TEMP -ChildPath "CreateShortcut.vbs"
        
        # Escape any quotes in the Arguments string to prevent VBScript syntax errors
        $escapedArguments = $Arguments.Replace('"', '""')
        
        # Create the VBScript content line by line to avoid formatting issues
        $vbsLines = @()
        $vbsLines += 'Set WshShell = CreateObject("WScript.Shell")'
        $vbsLines += 'Set lnk = WshShell.CreateShortcut(WshShell.SpecialFolders("Desktop") & "\' + $ShortcutName + '.lnk")'
        $vbsLines += 'lnk.TargetPath = "' + $TargetPath + '"'
        $vbsLines += 'lnk.Arguments = "' + $escapedArguments + '"'
        $vbsLines += 'lnk.Description = "' + $Description + '"'
        
        # Add working directory if specified
        if ($WorkingDirectory -ne "") {
            $vbsLines += 'lnk.WorkingDirectory = "' + $WorkingDirectory + '"'
        }
        
        # Add icon and save
        $vbsLines += 'lnk.IconLocation = "powershell.exe,0"'
        $vbsLines += 'lnk.Save'
        $vbsLines += 'WScript.Echo "SUCCESS"'
        
        # Write the VBScript file
        $vbsLines | Out-File -FilePath $vbsFile -Encoding ASCII
        
        # Log the shortcut details but not the entire VBScript
        Write-Log "Shortcut details:" -Silent
        Write-Log "  Target: $TargetPath" -Silent
        Write-Log "  Arguments: $Arguments" -Silent
        Write-Log "  Description: $Description" -Silent
        if ($WorkingDirectory -ne "") {
            Write-Log "  Working directory: $WorkingDirectory" -Silent
        }
        
        # Execute the VBScript
        Write-Log "Creating desktop shortcut..." 
        $result = (cscript //nologo $vbsFile) | Out-String
        
        # Check result
        if ($result.Trim() -eq "SUCCESS") {
            Write-Log "Shortcut created successfully: $env:USERPROFILE\Desktop\$ShortcutName.lnk" "SUCCESS"
            
            # Clean up the VBScript file
            if (Test-Path -Path $vbsFile) {
                Remove-Item -Path $vbsFile -Force
            }
            
            return $true
        } else {
            Write-Log "Failed to create shortcut: $result" "ERROR"
            return $false
        }
    }
    catch {
        Write-Log "Error creating shortcut: $_" "ERROR"
        return $false
    }
}

function Install-OfflineScript {
    Write-Log "Starting offline installation" "INFO"
    
    # Set paths
    $destDir = Join-Path -Path $env:APPDATA -ChildPath $appFolder
    $sourceScript = Join-Path -Path $PSScriptRoot -ChildPath $scriptName
    
    try {
        # Check if source script exists in current directory
        if (!(Test-Path -Path $sourceScript)) {
            Write-Log "Warning: $scriptName not found in the current directory, checking parent directories..." "WARNING"
            
            # Try to find the script two directories up (relative path)
            $parentDir = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
            $parentSourceScript = Join-Path -Path $parentDir -ChildPath $scriptName
            
            if (Test-Path -Path $parentSourceScript) {
                Write-Log "Found $scriptName in parent directory: $parentDir" "INFO"
                $sourceScript = $parentSourceScript
            } else {
                Write-Log "Error: $scriptName not found in current directory or parent directories" "ERROR"
                return $false
            }
        }
        
        # Create destination directory if it doesn't exist
        if (!(Test-Path -Path $destDir)) {
            Write-Log "Creating directory: $destDir"
            New-Item -Path $destDir -ItemType Directory -Force | Out-Null
        }
        
        # Copy script to destination
        Write-Log "Copying $scriptName to $destDir"
        Copy-Item -Path $sourceScript -Destination $destDir -Force
        
        # Create shortcut
        $powershellExe = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
        $scriptPath = Join-Path -Path $destDir -ChildPath $scriptName
        $shortcutArgs = "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`""
        
        $result = Create-Shortcut -TargetPath $powershellExe -ShortcutName $shortcutName -Arguments $shortcutArgs -Description $shortcutDescription -WorkingDirectory $destDir
        
        if ($result) {
            Write-Log "Offline installation completed successfully" "SUCCESS"
            return $true
        }
        else {
            Write-Log "Failed to create shortcut for offline installation" "ERROR"
            return $false
        }
    }
    catch {
        Write-Log "Error during offline installation: $_" "ERROR"
        return $false
    }
}

function Invoke-GitHubDownload {
    param (
        [string]$OutputPath
    )
    
    try {
        Write-Log "Starting download process for $scriptName from $gitHubUsername/$repositoryName"
        
        # Use GitHub API to get file content
        $apiUrl = "https://api.github.com/repos/$gitHubUsername/$repositoryName/contents/$scriptName"
        Write-Log "Using GitHub API URL: $apiUrl"
        
        # Configure TLS to use 1.2 (needed for GitHub)
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
        Write-Log "TLS 1.2 enabled for secure connection"
        
        # Prepare headers
        $headers = @{
            "User-Agent" = "PowerShell Installer"
        }
        
        # Add authorization token if available
        if (-not [string]::IsNullOrWhiteSpace($gitHubToken)) {
            Write-Log "Using GitHub token for authentication"
            $headers["Authorization"] = "token $gitHubToken"
        } else {
            Write-Log "No GitHub token provided, accessing as public repository"
        }
        
        Write-Log "Sending request to GitHub API"
        $response = Invoke-RestMethod -Uri $apiUrl -Headers $headers
        Write-Log "API call successful"
        
        # GitHub API returns content as Base64 encoded string
        Write-Log "Decoding Base64 content"
        $content = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($response.content))
        
        # Clean content - remove BOM and invisible characters
        Write-Log "Cleaning content to remove any invisible characters"
        
        # Remove BOM and other invisible characters that might cause issues
        $content = $content -replace '^\xEF\xBB\xBF', '' # Remove UTF-8 BOM
        $content = $content -replace '^\xFE\xFF', ''     # Remove UTF-16 BE BOM
        $content = $content -replace '^\xFF\xFE', ''     # Remove UTF-16 LE BOM
        
        # Write content to file
        Write-Log "Writing content to $OutputPath"
        
        # Use UTF8 encoding without BOM
        $utf8NoBomEncoding = New-Object System.Text.UTF8Encoding $false
        [System.IO.File]::WriteAllText($OutputPath, $content, $utf8NoBomEncoding)
        
        # Verify the download
        if (Test-Path $OutputPath) {
            $fileSize = (Get-Item $OutputPath).Length
            Write-Log "Download completed successfully. File size: $fileSize bytes"
            return $true
        } else {
            Write-Log "Download appeared to succeed but file not found" "ERROR"
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
        $usingToken = -not [string]::IsNullOrWhiteSpace($gitHubToken)
        
        # Handle specific error codes with improved messages
        switch ($statusCode) {
            "404" {
                Write-Log "Repository or file not found (404): The repository '$gitHubUsername/$repositoryName' or script '$scriptName' does not exist" "ERROR"
            }
            "401" {
                if ($usingToken) {
                    Write-Log "Authentication failed (401): The GitHub token provided is invalid or expired" "ERROR"
                } else {
                    Write-Log "Authentication required (401): This repository requires authentication. Set the GitHubToken variable." "ERROR"
                }
            }
            "403" {
                if ($usingToken) {
                    Write-Log "Access forbidden (403): Your GitHub token doesn't have permission to access this repository, or you may be rate limited" "ERROR"
                } else {
                    Write-Log "Access forbidden (403): You may be rate limited or the repository requires authentication. Set the GitHubToken variable." "ERROR"
                }
            }
            default {
                Write-Log "Download error: $errorMessage" "ERROR"
            }
        }
        
        return $false
    }
}

function Install-OnlineScript {
    Write-Log "Starting online installation" "INFO"
    
    # Set paths
    $destDir = Join-Path -Path $env:APPDATA -ChildPath $appFolder
    
    try {
        # Create destination directory if it doesn't exist
        if (!(Test-Path -Path $destDir)) {
            Write-Log "Creating directory: $destDir"
            New-Item -Path $destDir -ItemType Directory -Force | Out-Null
        }
        
        # Download script from GitHub
        $scriptPath = Join-Path -Path $destDir -ChildPath $scriptName
        Write-Log "Downloading $scriptName from GitHub to $scriptPath"
        
        $downloadSuccess = Invoke-GitHubDownload -OutputPath $scriptPath
        
        if (-not $downloadSuccess) {
            Write-Log "Failed to download script from GitHub" "ERROR"
            
            # Check if we have a local copy to use as fallback
            $localScript = Join-Path -Path $PSScriptRoot -ChildPath $scriptName
            if (Test-Path -Path $localScript) {
                Write-Log "Using local copy as fallback" "WARNING"
                Copy-Item -Path $localScript -Destination $scriptPath -Force
            } else {
                return $false
            }
        }
        
        # Create shortcut
        $powershellExe = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
        $shortcutArgs = "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`""
        
        $result = Create-Shortcut -TargetPath $powershellExe -ShortcutName $shortcutName -Arguments $shortcutArgs -Description $shortcutDescription -WorkingDirectory $destDir
        
        if ($result) {
            Write-Log "Online installation completed successfully" "SUCCESS"
            return $true
        }
        else {
            Write-Log "Failed to create shortcut for online installation" "ERROR"
            return $false
        }
    }
    catch {
        Write-Log "Error during online installation: $_" "ERROR"
        return $false
    }
}

# Initialize log file
Write-Log "======================================================" "INFO"
Write-Log "Installer started with InstallType: $InstallType" "INFO"
Write-Log "Script location: $PSScriptRoot" "INFO"
Write-Log "User: $env:USERNAME" "INFO"
Write-Log "Computer: $env:COMPUTERNAME" "INFO"
Write-Log "======================================================" "INFO"

# Execute installation based on type
$success = $false

switch ($InstallType) {
    "Offline" {
        $success = Install-OfflineScript
    }
    "Online" {
        $success = Install-OnlineScript
    }
}

# Return exit code based on success
if ($success) {
    Write-Log "Installation completed successfully" "SUCCESS"
    exit 0
} else {
    Write-Log "Installation failed" "ERROR"
    exit 1
}
