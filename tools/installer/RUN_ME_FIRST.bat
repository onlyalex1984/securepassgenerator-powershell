@echo off
title SecurePassGenerator Installer

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
echo         SecurePassGenerator Installer
echo ===============================================
echo.
echo  Please select an installation method:
echo.
echo  1. Offline Installation
echo     - Installs from local files
echo     - Creates desktop shortcut
echo.
echo  2. Online Installation
echo     - Downloads latest version from GitHub
echo     - Creates desktop shortcut
echo.
echo  3. Exit
echo.
echo ===============================================
echo.

set /p choice=Enter your choice (1-3): 

if "%choice%"=="1" goto offline
if "%choice%"=="2" goto online
if "%choice%"=="3" goto exit
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
echo Installing SecurePassGenerator from local files...
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "Installer.ps1" -InstallType Offline
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

:online
cls
echo ===============================================
echo         Online Installation
echo ===============================================
echo.
echo Installing SecurePassGenerator from GitHub...
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "Installer.ps1" -InstallType Online
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

:exit
exit
