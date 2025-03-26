@echo off
:: ===============================================
:: SecurePassGenerator PowerShell Interactive Installer
:: ===============================================
:: 
:: This script provides an interactive menu to install
:: SecurePassGenerator PowerShell Edition on your computer.
:: 
:: Author: onlyalex1984
:: Version: 1.1
:: GitHub: https://github.com/onlyalex1984
:: 
:: ===============================================

title SecurePassGenerator PowerShell Installer

:: Check for administrative privileges
net session >nul 2>&1
if %errorLevel% == 0 (
    echo Running with administrative privileges.
) else (
    echo This script doesn't require administrative privileges.
)

:menu
cls
echo ===============================================
echo     SecurePassGenerator PowerShell Installer
echo ===============================================
echo.
echo  Please select an installation method:
echo.
echo  1. Offline Installation
echo     - Installs from local PowerShell script or zip file
echo     - Hash verification only applies if installing from zip
echo     - Creates desktop shortcut
echo.
echo  2. Online Installation - Latest
echo     - Downloads latest stable version from GitHub
echo     - Verifies downloaded file hash (strict mode)
echo     - Creates desktop shortcut
echo.
echo  3. Online Installation - PreRelease
echo     - Downloads latest pre-release version from GitHub
echo     - Verifies downloaded file hash (strict mode)
echo     - Creates desktop shortcut
echo.
echo  4. Advanced Installation Options
echo     - Installation options without hash verification
echo.
echo  5. Exit
echo.
echo ===============================================
echo.

set /p choice=Enter your choice (1-5): 

if "%choice%"=="1" goto offline
if "%choice%"=="2" goto online_latest
if "%choice%"=="3" goto online_prerelease
if "%choice%"=="4" goto advanced_options
if "%choice%"=="5" goto exit
goto invalid

:invalid
echo.
echo Invalid selection. Please try again.
timeout /t 2 >nul
goto menu

:offline
cls
echo ===============================================
echo         Offline Installation
echo ===============================================
echo.
echo Installing SecurePassGenerator PowerShell Edition from local files...
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Installer.ps1" -InstallType Offline -ReleaseType Latest
if %errorlevel% neq 0 (
    echo.
    echo Installation failed. Please check the error message above.
    echo.
    pause
    goto menu
) else (
    echo.
    echo Installation completed successfully!
    echo.
    pause
    goto exit
)

:online_latest
cls
echo ===============================================
echo         Online Installation - Latest
echo ===============================================
echo.
echo Installing SecurePassGenerator PowerShell Edition (Latest) from GitHub with strict hash verification...
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Installer.ps1" -InstallType Online -ReleaseType Latest -StrictHashVerification
if %errorlevel% neq 0 (
    echo.
    echo Installation failed. Please check the error message above.
    echo.
    pause
    goto menu
) else (
    echo.
    echo Installation completed successfully!
    echo.
    pause
    goto exit
)

:online_prerelease
cls
echo ===============================================
echo         Online Installation - PreRelease
echo ===============================================
echo.
echo Installing SecurePassGenerator PowerShell Edition (PreRelease) from GitHub with strict hash verification...
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Installer.ps1" -InstallType Online -ReleaseType PreRelease -StrictHashVerification
if %errorlevel% neq 0 (
    echo.
    echo Installation failed. Please check the error message above.
    echo.
    pause
    goto menu
) else (
    echo.
    echo Installation completed successfully!
    echo.
    pause
    goto exit
)

:advanced_options
cls
echo ===============================================
echo     Advanced Installation Options
echo ===============================================
echo.
echo  Please select an installation method:
echo.
echo  1. Offline Installation (No Hash Verification)
echo     - Installs from local PowerShell script or zip file
echo     - Skips hash verification if installing from zip
echo     - Creates desktop shortcut
echo.
echo  2. Online Installation - Latest (No Hash Verification)
echo     - Downloads latest stable version from GitHub without hash verification
echo     - Creates desktop shortcut
echo.
echo  3. Online Installation - PreRelease (No Hash Verification)
echo     - Downloads latest pre-release version from GitHub without hash verification
echo     - Creates desktop shortcut
echo.
echo  4. Return to Main Menu
echo.
echo ===============================================
echo.

set /p advanced_choice=Enter your choice (1-4): 

if "%advanced_choice%"=="1" goto offline_no_hash
if "%advanced_choice%"=="2" goto online_latest_no_hash
if "%advanced_choice%"=="3" goto online_prerelease_no_hash
if "%advanced_choice%"=="4" goto menu
goto advanced_invalid

:advanced_invalid
echo.
echo Invalid selection. Please try again.
timeout /t 2 >nul
goto advanced_options

:offline_no_hash
cls
echo ===============================================
echo     Offline Installation (No Hash Verification)
echo ===============================================
echo.
echo Installing SecurePassGenerator PowerShell Edition from local files without hash verification...
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Installer.ps1" -InstallType Offline -ReleaseType Latest -DisableHashVerification
if %errorlevel% neq 0 (
    echo.
    echo Installation failed. Please check the error message above.
    echo.
    pause
    goto advanced_options
) else (
    echo.
    echo Installation completed successfully!
    echo.
    pause
    goto exit
)

:online_latest_no_hash
cls
echo ===============================================
echo     Online Installation - Latest (No Hash Verification)
echo ===============================================
echo.
echo Installing SecurePassGenerator PowerShell Edition (Latest) from GitHub without hash verification...
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Installer.ps1" -InstallType Online -ReleaseType Latest -DisableHashVerification
if %errorlevel% neq 0 (
    echo.
    echo Installation failed. Please check the error message above.
    echo.
    pause
    goto advanced_options
) else (
    echo.
    echo Installation completed successfully!
    echo.
    pause
    goto exit
)

:online_prerelease_no_hash
cls
echo ===============================================
echo     Online Installation - PreRelease (No Hash Verification)
echo ===============================================
echo.
echo Installing SecurePassGenerator PowerShell Edition (PreRelease) from GitHub without hash verification...
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Installer.ps1" -InstallType Online -ReleaseType PreRelease -DisableHashVerification
if %errorlevel% neq 0 (
    echo.
    echo Installation failed. Please check the error message above.
    echo.
    pause
    goto advanced_options
) else (
    echo.
    echo Installation completed successfully!
    echo.
    pause
    goto exit
)

:exit
exit
