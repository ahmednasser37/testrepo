@echo off
title Protocol DEE -- Installer
echo ============================================
echo   Protocol DEE v2 -- Install
echo ============================================
echo.
echo Close Excel before continuing!
echo.
pause
powershell.exe -ExecutionPolicy Bypass -File "%~dp0build_installer.ps1" -InstallAfterBuild
if %ERRORLEVEL% neq 0 (
    echo.
    echo FAILED -- see errors above.
    echo Common fix: Enable "Trust access to VBA project object model" in Excel Trust Center
) else (
    echo.
    echo SUCCESS -- open Excel, then File > Options > Add-ins > Go > Browse to enable it.
)
pause
