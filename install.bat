@echo off
:: =============================================================================
:: install.bat -- Protocol DEE Quick Installer
:: Protocol DEE v2 -- Project Controls Add-in
:: =============================================================================
:: Simple batch wrapper for build_installer.ps1
:: Run as Administrator if UAC issues occur.
:: =============================================================================

title Protocol DEE -- Installer

echo ============================================
echo   Protocol DEE v2 -- Quick Install
echo ============================================
echo.

:: Check PowerShell availability
where powershell >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo ERROR: PowerShell not found. Please install PowerShell.
    pause
    exit /b 1
)

:: Check if already installed
set ADDIN_PATH=%APPDATA%\Microsoft\AddIns\ProtocolDEE.xlam
if exist "%ADDIN_PATH%" (
    echo Existing installation found: %ADDIN_PATH%
    echo.
    set /p OVERWRITE="Overwrite? [Y/N]: "
    if /i "%OVERWRITE%" neq "Y" (
        echo Installation cancelled.
        pause
        exit /b 0
    )
)

echo.
echo Running build script...
echo.

:: Run PowerShell build script
powershell.exe -ExecutionPolicy Bypass -File "%~dp0build_installer.ps1" -InstallAfterBuild

if %ERRORLEVEL% neq 0 (
    echo.
    echo BUILD FAILED. Check the error messages above.
    echo.
    echo Common issues:
    echo   1. Excel is not installed
    echo   2. "Trust access to VBA project" not enabled in Trust Center
    echo   3. Excel is currently open (close Excel first)
    echo.
) else (
    echo.
    echo ============================================
    echo   INSTALLATION COMPLETE
    echo ============================================
    echo.
    echo Next steps:
    echo   1. Open Microsoft Excel
    echo   2. Go to: File ^> Options ^> Add-ins
    echo   3. In "Manage" dropdown, select "Excel Add-ins" and click "Go..."
    echo   4. Check "Protocol DEE" from the list and click OK
    echo   5. The "Protocol DEE" tab will appear in the Ribbon
    echo.
    echo Trust Center (required):
    echo   File ^> Options ^> Trust Center ^> Trust Center Settings ^> Macro Settings
    echo   Enable: "Trust access to the VBA project object model"
    echo.
)

pause
