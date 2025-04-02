# SecurePassGenerator PowerShell Installer

This directory contains the installer for SecurePassGenerator PowerShell Edition. The installer provides multiple methods to set up the application on your computer with flexible options for different environments.

## Installation Methods

The installer supports several installation options:

1. **Offline Installation**

   - Installs from a local copy of SecurePassGenerator.ps1
   - Falls back to using a zip file if script is not found
   - Creates a desktop shortcut with execution policy bypass
   - Includes hash verification for zip files
   - No internet connection required
   - Searches parent directories to find the script (useful for repository clones)

2. **Online Installation**

   - **Latest**: Downloads the latest stable version from GitHub
   - **PreRelease**: Downloads the latest pre-release version if available
   - Creates a desktop shortcut with execution policy bypass
   - Includes hash verification for security
   - Falls back to local copy if download fails
   - Supports both public and private GitHub repositories

3. **Direct API Installation**

   - Uses GitHub API to download files directly (proxy-friendly)
   - Bypasses the need for downloading zip files
   - Works in environments where standard GitHub downloads might be blocked
   - Available for both Latest and PreRelease versions
   - Creates a desktop shortcut with execution policy bypass

4. **Advanced Options**
   - Offline Installation without hash verification
   - Online Installation (Latest) without hash verification
   - Online Installation (PreRelease) without hash verification
   - Direct API Installation for both Latest and PreRelease versions
   - Useful for environments with specific security or network restrictions

## How to Use

1. **Unblock the installer files** (if downloaded):

   - Right-click on `RUN_ME_FIRST.bat` and select "Properties"
   - Check the "Unblock" box and click OK
   - The installer will automatically unblock the PowerShell script during installation

2. **Run the installer**:

   - Double-click `RUN_ME_FIRST.bat`
   - No administrator privileges required

3. **Choose installation method** from the menu:

   - Select option 1 for Offline Installation (uses local files)
   - Select option 2 for Online Installation - Latest (downloads latest stable version)
   - Select option 3 for Online Installation - PreRelease (downloads latest pre-release version)
   - Select option 4 for Advanced Installation Options
   - Select option 5 to exit

4. **Start using SecurePassGenerator**:
   - Once installation is complete, you'll find a shortcut on your desktop
   - The application is installed to `%AppData%\securepassgenerator-ps\`

## Technical Features

- **Interactive Menu**: Simple menu-driven interface for easy installation
- **Multiple Download Methods**: Uses several approaches with fallbacks for reliable downloads
- **Repository Support**:
  - Works with both public and private GitHub repositories
  - For private repositories: Add your GitHub token to the $GitHubToken variable
  - For public repositories: Leave the $GitHubToken variable empty
- **Security**:
  - Hash verification for downloaded files
  - Automatic script unblocking
  - Multiple verification modes (strict, normal, disabled)
- **Intelligent Installation**:
  - Fallback mechanisms ensure successful installation
  - Relative paths support for repository clones
  - Detailed logging for troubleshooting
- **Clean Installation**:
  - Installs to %AppData% without system modifications
  - No administrator privileges required
  - Creates properly configured shortcuts

## Technical Details

### Offline Installation Process

- First looks for SecurePassGenerator.ps1 in the current directory
- If not found, checks two directories up (useful when cloning the repository)
- If script file not found, looks for a zip file matching the asset pattern
- For zip files, performs hash verification if a hash file is available
- Extracts zip contents to the installation directory
- Unblocks the PowerShell script before copying/after extracting
- Creates a desktop shortcut with execution policy bypass

### Online Installation Process

- Downloads the latest release or pre-release from GitHub based on your selection
- Verifies the downloaded file's hash against the published hash file
- Extracts the contents to the installation directory
- Reorganizes files if needed (moving from subdirectories to main directory)
- Creates a desktop shortcut with execution policy bypass
- Cleans up temporary files

## Requirements

- Windows operating system
- PowerShell 5.1 or higher
- Internet connection (for online installation only)

## Troubleshooting

If you encounter any issues:

1. Check the install_log.txt file in the installation directory (%AppData%\securepassgenerator-ps\)
2. Make sure your computer has internet access for online installation
3. Ensure Windows Defender or your antivirus isn't blocking the scripts
4. For offline installation, verify that SecurePassGenerator.ps1 exists in the expected locations
