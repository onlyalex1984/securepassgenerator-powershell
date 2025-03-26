# SecurePassGenerator PowerShell Interactive Installer

An interactive installer that provides flexible options for installing and using the PowerShell version of SecurePassGenerator.

## What This Installer Does

This installer offers multiple ways to set up SecurePassGenerator PowerShell Edition on your computer:

1. **Offline Installation**:

   - Installs from a local copy of SecurePassGenerator.ps1
   - Creates a desktop shortcut with execution policy bypass
   - Includes a fallback mechanism to find the script in parent directories

2. **Online Installation - Latest**:

   - Downloads the latest stable version of the script directly from GitHub
   - Creates a desktop shortcut with execution policy bypass
   - Falls back to local copy if download fails

3. **Online Installation - PreRelease**:

   - Downloads the latest pre-release version from GitHub (if available)
   - Creates a desktop shortcut with execution policy bypass
   - Falls back to local copy if download fails

4. **Advanced Installation Options**:
   - Provides installation options with additional customization
   - Useful for specific deployment scenarios

## How to Use

1. **Unblock the installer files**:

   - Right-click on `RUN_ME_FIRST.bat` and select "Properties"
   - At the bottom of the Properties window, check the "Unblock" box and click OK
   - This step is necessary because Windows often blocks downloaded files for security reasons
   - Note: The installer will automatically unblock the PowerShell script during installation

2. **Run the installer**:

   - Double-click `RUN_ME_FIRST.bat`
   - No administrator privileges required

3. **Choose installation method**:

   - Select option 1 for Offline Installation (uses local files)
   - Select option 2 for Online Installation - Latest (downloads latest stable version)
   - Select option 3 for Online Installation - PreRelease (downloads latest pre-release version if available)
   - Select option 4 for Advanced Installation Options (with additional customization)
   - Select option 5 to exit

4. **Start using SecurePassGenerator PowerShell Edition**:
   - Once installation is complete, you'll find a shortcut on your desktop
   - Click the shortcut anytime you need to generate or share secure passwords

## Features

- **Interactive**: Simple menu-driven interface
- **Flexible**: Choose between offline and online installation
- **Release Options**: Install stable releases or pre-release versions
- **Intelligent**: Fallback mechanisms ensure successful installation
- **Automatic Unblocking**: Unblocks PowerShell script before and after installation
- **Repository Support**: Works with both public and private GitHub repositories
- **Multiple Download Methods**: Uses several approaches with fallbacks for reliable downloads
- **Detailed Logging**: All installation steps are logged for troubleshooting
- **Relative Paths**: Works from any location, ideal for repository clones
- **Clean**: Installs to %AppData% without system modifications

## Technical Details

### Offline Installation

The offline installer:

- First looks for SecurePassGenerator.ps1 in the current directory
- If not found, checks two directories up (useful when cloning the repository)
- Unblocks the PowerShell script before copying it to the installation directory
- Copies the script to %AppData%\securepassgenerator\
- Creates a desktop shortcut with execution policy bypass

### Online Installation

The online installer:

- Supports both public and private GitHub repositories
- For private repositories: Add your GitHub token to the $GitHubToken variable
- For public repositories: Leave the $GitHubToken variable empty
- Downloads the latest release or pre-release script from GitHub based on your selection
- Uses multiple download methods with fallbacks for reliability
- Stores the script in %AppData%\securepassgenerator\
- Falls back to local copy if download fails
- Creates a desktop shortcut with execution policy bypass

### Advanced Installation

The advanced installation options:

- Allow customization of installation parameters
- Provide options for specific deployment scenarios
- Enable integration with existing automation workflows

## Troubleshooting

If you encounter any issues:

1. Check the Installer.log file created in the same directory
2. Make sure your computer has internet access for online installation
3. Ensure Windows Defender or your antivirus isn't blocking the scripts
4. For offline installation, verify that SecurePassGenerator.ps1 exists in the expected locations
5. For private repositories, ensure your GitHub token has the correct permissions

## Requirements

- Windows operating system
- PowerShell 5.1 or higher
- Internet connection (for online installation only)

## C# Version

A C# WPF version of this tool is also available with the same functionality and an enhanced graphical interface:
[SecurePassGenerator C#](https://github.com/onlyalex1984/securepassgenerator-csharp)
