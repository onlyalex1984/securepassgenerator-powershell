# SecurePassGenerator Interactive Installer

An interactive installer that provides flexible options for installing and using SecurePassGenerator.

## What This Installer Does

This installer offers multiple ways to set up SecurePassGenerator on your computer:

1. **Offline Installation**:

   - Installs from a local copy of SecurePassGenerator.ps1
   - Creates a desktop shortcut with execution policy bypass
   - Includes a fallback mechanism to find the script in parent directories

2. **Online Installation**:
   - Downloads the latest version of the script directly from GitHub
   - Creates a desktop shortcut with execution policy bypass
   - Falls back to local copy if download fails

## How to Use

1. **Unblock the installer files**:

   - Right-click on `RUN_ME_FIRST.bat` and select "Properties"
   - At the bottom of the Properties window, check the "Unblock" box and click OK
   - This step is necessary because Windows often blocks downloaded files for security reasons

2. **Run the installer**:

   - Double-click `RUN_ME_FIRST.bat`
   - No administrator privileges required

3. **Choose installation method**:

   - Select option 1 for Offline Installation (uses local files)
   - Select option 2 for Online Installation (downloads from GitHub)
   - Select option 3 to exit

4. **Start using SecurePassGenerator**:
   - Once installation is complete, you'll find a shortcut on your desktop
   - Click the shortcut anytime you need to generate or share secure passwords

## Features

- **Interactive**: Simple menu-driven interface
- **Flexible**: Choose between offline and online installation
- **Intelligent**: Fallback mechanisms ensure successful installation
- **Detailed Logging**: All installation steps are logged for troubleshooting
- **Relative Paths**: Works from any location, ideal for repository clones
- **Clean**: Installs to %AppData% without system modifications

## Technical Details

### Offline Installation

The offline installer:

- First looks for SecurePassGenerator.ps1 in the current directory
- If not found, checks two directories up (useful when cloning the repository)
- Copies the script to %AppData%\securepassgenerator\
- Creates a desktop shortcut with execution policy bypass

### Online Installation

The online installer:

- Downloads SecurePassGenerator.ps1 from GitHub using the API
- Stores the script in %AppData%\securepassgenerator\
- Falls back to local copy if download fails
- Creates a desktop shortcut with execution policy bypass

## Troubleshooting

If you encounter any issues:

1. Check the Installer.log file created in the same directory
2. Make sure your computer has internet access for online installation
3. Ensure Windows Defender or your antivirus isn't blocking the scripts
4. For offline installation, verify that SecurePassGenerator.ps1 exists in the expected locations

## Requirements

- Windows operating system
- PowerShell 5.1 or higher
- Internet connection (for online installation only)
