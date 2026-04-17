# =============================================================================
# build_installer.ps1 -- Protocol DEE Build Script
# Protocol DEE v2 -- Project Controls Add-in
# =============================================================================
# Builds ProtocolDEE.xlam via Excel COM automation.
# The Ribbon tab is served by IRibbonExtensibility (ThisWorkbook.GetCustomUI)
# so NO ZIP/customUI injection is needed.
#
# PREREQUISITES:
#   - Microsoft Excel installed
#   - Excel CLOSED before running this script
#   - "Trust access to VBA project object model" enabled:
#       Excel > File > Options > Trust Center > Trust Center Settings
#       > Macro Settings > check "Trust access to the VBA project object model"
# =============================================================================

[CmdletBinding()]
param(
    [string]$OutputPath    = "$env:APPDATA\Microsoft\AddIns\ProtocolDEE.xlam",
    [switch]$InstallAfterBuild,
    [switch]$SkipInstall
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  Protocol DEE v2 -- Build & Install"
Write-Host "============================================"
Write-Host ""

# ---------------------------------------------------------------------------
# Validate
# ---------------------------------------------------------------------------
$modulesPath = Join-Path $ScriptDir "modules"
if (-not (Test-Path $modulesPath)) {
    Write-Error "modules/ folder not found: $modulesPath"
    exit 1
}

$modules = Get-ChildItem "$modulesPath\*.bas" | Sort-Object Name
Write-Host "Modules found: $($modules.Count)"
foreach ($m in $modules) { Write-Host "  $($m.Name)" }
Write-Host ""

# Block if Excel is running
if (Get-Process EXCEL -EA SilentlyContinue) {
    Write-Error "Excel is open. Close Excel completely, then run this script again."
    exit 1
}

# ---------------------------------------------------------------------------
# Clean old output
# ---------------------------------------------------------------------------
$tempPath = "$env:TEMP\ProtocolDEE_build.xlam"
if (Test-Path $tempPath) { Remove-Item $tempPath -Force }
if ((Test-Path $OutputPath) -and -not $SkipInstall) {
    Remove-Item $OutputPath -Force -EA SilentlyContinue
    Write-Host "Removed old add-in"
}

# ---------------------------------------------------------------------------
# Step 1: Create workbook and import modules via COM
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "[1/3] Building add-in via Excel COM..." -ForegroundColor Cyan

$xl  = $null
$wb  = $null

try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible        = $false
    $xl.DisplayAlerts  = $false

    $wb  = $xl.Workbooks.Add()
    $vbp = $wb.VBProject

    # Import all .bas modules
    foreach ($m in $modules) {
        try {
            $vbp.VBComponents.Import($m.FullName)
            Write-Host "  + $($m.Name)"
        } catch {
            Write-Warning "  SKIP $($m.Name): $_"
        }
    }

    # Add IRibbonExtensibility to ThisWorkbook
    # This makes Excel call DEE_RibbonXML.GetRibbonXML() at load -- no file injection needed
    Write-Host "  Wiring IRibbonExtensibility to ThisWorkbook..."
    $thisWb = $vbp.VBComponents.Item("ThisWorkbook")
    $code = "Implements IRibbonExtensibility" & [Environment]::NewLine & _
            "" & [Environment]::NewLine & _
            "Public Function GetCustomUI(ByVal ribbonID As String) As String" & [Environment]::NewLine & _
            "    GetCustomUI = DEE_RibbonXML.GetRibbonXML()" & [Environment]::NewLine & _
            "End Function"

    $thisWb.CodeModule.AddFromString($code)

    # References
    try {
        # Microsoft Scripting Runtime
        $hasScripting = ($vbp.References | Where-Object { $_.Name -eq "Scripting" }) -ne $null
        if (-not $hasScripting) {
            $vbp.References.AddFromGuid("{420B2830-E718-11CF-893D-00A0C9054228}", 1, 0)
            Write-Host "  + Microsoft Scripting Runtime"
        }
        # Microsoft Office Object Library (needed for IRibbonExtensibility)
        $hasOffice = ($vbp.References | Where-Object { $_.Name -eq "Office" }) -ne $null
        if (-not $hasOffice) {
            $vbp.References.AddFromGuid("{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}", 2, 8)
            Write-Host "  + Microsoft Office Object Library"
        }
    } catch {
        Write-Warning "  Reference add failed (may already be present): $_"
    }

    # ---------------------------------------------------------------------------
    # Step 2: Save as .xlam
    # ---------------------------------------------------------------------------
    Write-Host ""
    Write-Host "[2/3] Saving as .xlam..." -ForegroundColor Cyan
    $wb.SaveAs($tempPath, 55)   # 55 = xlOpenXMLAddIn
    $wb.Close($false)
    $wb = $null
    Write-Host "  Saved: $tempPath"

} finally {
    if ($wb  -ne $null) { try { $wb.Close($false)  } catch {} }
    if ($xl  -ne $null) { try { $xl.Quit()          } catch {} }
    if ($xl  -ne $null) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    Start-Sleep -Milliseconds 800
}

# ---------------------------------------------------------------------------
# Step 3: Copy to AddIns folder
# ---------------------------------------------------------------------------
if (-not $SkipInstall) {
    Write-Host ""
    Write-Host "[3/3] Installing..." -ForegroundColor Cyan
    $dir = Split-Path $OutputPath
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    Copy-Item $tempPath $OutputPath -Force
    Write-Host "  Installed: $OutputPath" -ForegroundColor Green
} else {
    $local = Join-Path $ScriptDir "ProtocolDEE.xlam"
    Copy-Item $tempPath $local -Force
    Write-Host "  Output: $local" -ForegroundColor Green
}

Remove-Item $tempPath -Force -EA SilentlyContinue

Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host "  BUILD COMPLETE"
Write-Host "============================================"
Write-Host ""
Write-Host "To activate in Excel:"
Write-Host "  File > Options > Add-ins > Manage: Excel Add-ins > Go"
Write-Host "  Click Browse > select ProtocolDEE.xlam > OK > check it > OK"
Write-Host ""
Write-Host "Trust Center (required once):"
Write-Host "  File > Options > Trust Center > Trust Center Settings > Macro Settings"
Write-Host "  Enable: 'Trust access to the VBA project object model'"
Write-Host ""
