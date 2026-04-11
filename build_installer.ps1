# =============================================================================
# build_installer.ps1 -- Protocol DEE Build Script
# Protocol DEE v2 -- Project Controls Add-in
# =============================================================================
# Creates the ProtocolDEE.xlam add-in file by:
# 1. Importing all .bas modules into a new Excel workbook via COM
# 2. Saving as .xlam
# 3. Injecting customUI14.xml into the ZIP package
# 4. Updating [Content_Types].xml and _rels/.rels
# 5. Copying to %APPDATA%\Microsoft\AddIns\
#
# PREREQUISITES:
# - Microsoft Excel must be installed
# - "Trust access to the VBA project object model" must be enabled in Excel:
#   File > Options > Trust Center > Trust Center Settings > Macro Settings
# =============================================================================

[CmdletBinding()]
param(
    [string]$OutputPath = "$env:APPDATA\Microsoft\AddIns\ProtocolDEE.xlam",
    [switch]$InstallAfterBuild = $false,
    [switch]$SkipInstall = $false
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host "============================================"
Write-Host "  Protocol DEE -- Build Script v2"
Write-Host "============================================"
Write-Host ""

# ---------------------------------------------------------------------------
# Validate prerequisites
# ---------------------------------------------------------------------------
$modulesPath = Join-Path $ScriptDir "modules"
$customUIPath = Join-Path $ScriptDir "customUI"

if (-not (Test-Path $modulesPath)) {
    Write-Error "modules/ directory not found at: $modulesPath"
    exit 1
}

if (-not (Test-Path $customUIPath)) {
    Write-Error "customUI/ directory not found at: $customUIPath"
    exit 1
}

$modules = Get-ChildItem "$modulesPath\*.bas" | Sort-Object Name
Write-Host "Found $($modules.Count) VBA module(s):"
foreach ($m in $modules) {
    Write-Host "  - $($m.Name)"
}
Write-Host ""

# ---------------------------------------------------------------------------
# Step 1: Create Excel add-in via COM automation
# ---------------------------------------------------------------------------
Write-Host "[1/5] Creating Excel workbook via COM..."

$tempXlamPath = Join-Path $env:TEMP "ProtocolDEE_temp.xlam"
if (Test-Path $tempXlamPath) { Remove-Item $tempXlamPath -Force }

$excel = $null
$wb = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    # Check Trust Center setting
    try {
        $vbaTrust = $excel.Application.VBE.ActiveVBProject
    } catch {
        Write-Warning "Cannot access VBA project. Ensure 'Trust access to VBA project object model' is enabled."
        Write-Warning "Excel > File > Options > Trust Center > Trust Center Settings > Macro Settings"
    }

    # Create new workbook
    $wb = $excel.Workbooks.Add()

    # ---------------------------------------------------------------------------
    # Step 2: Import all .bas modules
    # ---------------------------------------------------------------------------
    Write-Host "[2/5] Importing VBA modules..."

    $vbProject = $wb.VBProject

    foreach ($module in $modules) {
        Write-Host "  Importing: $($module.Name)"
        try {
            $vbProject.VBComponents.Import($module.FullName)
        } catch {
            Write-Warning "  Failed to import $($module.Name): $_"
        }
    }

    # Remove default Sheet1 Module1 etc. that come with new workbook
    $defaultModules = @("Module1", "Sheet1", "ThisWorkbook")
    foreach ($dm in $defaultModules) {
        try {
            $comp = $vbProject.VBComponents.Item($dm)
            if ($comp -and $comp.Type -ne 100) {  # 100 = vbext_ct_Document
                $vbProject.VBComponents.Remove($comp)
            }
        } catch {
            # Ignore - component may not exist or be removable
        }
    }

    # Add Microsoft Scripting Runtime reference
    try {
        $refs = $vbProject.References
        # Check if already referenced
        $hasScripting = $false
        foreach ($ref in $refs) {
            if ($ref.Name -eq "Scripting") { $hasScripting = $true; break }
        }
        if (-not $hasScripting) {
            $refs.AddFromGuid("{420B2830-E718-11CF-893D-00A0C9054228}", 1, 0)
            Write-Host "  Added Microsoft Scripting Runtime reference"
        }
    } catch {
        Write-Warning "  Could not add Scripting Runtime reference: $_"
        Write-Warning "  You may need to add it manually: Tools > References > Microsoft Scripting Runtime"
    }

    # ---------------------------------------------------------------------------
    # Step 3: Save as .xlam
    # ---------------------------------------------------------------------------
    Write-Host "[3/5] Saving as .xlam..."
    $wb.SaveAs($tempXlamPath, 55)  # 55 = xlOpenXMLAddIn
    $wb.Close($false)
    $wb = $null

} finally {
    if ($wb -ne $null) { try { $wb.Close($false) } catch {} }
    if ($excel -ne $null) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        $excel = $null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Start-Sleep -Milliseconds 500

# ---------------------------------------------------------------------------
# Step 4: Inject customUI14.xml into the ZIP package
# ---------------------------------------------------------------------------
Write-Host "[4/5] Injecting Ribbon XML..."

Add-Type -AssemblyName System.IO.Compression.FileSystem

$customUIXml = Join-Path $customUIPath "customUI14.xml"
if (-not (Test-Path $customUIXml)) {
    Write-Error "customUI14.xml not found at: $customUIXml"
    exit 1
}

# Work with a copy
$workXlamPath = Join-Path $env:TEMP "ProtocolDEE_work.xlam"
Copy-Item $tempXlamPath $workXlamPath -Force

try {
    $zip = [System.IO.Compression.ZipFile]::Open($workXlamPath, 'Update')

    # Add customUI/customUI14.xml
    $customUIEntry = $zip.GetEntry("customUI/customUI14.xml")
    if ($customUIEntry -ne $null) { $customUIEntry.Delete() }

    $xmlContent = Get-Content $customUIXml -Raw -Encoding UTF8
    $newEntry = $zip.CreateEntry("customUI/customUI14.xml")
    $writer = New-Object System.IO.StreamWriter($newEntry.Open())
    $writer.Write($xmlContent)
    $writer.Close()
    Write-Host "  Wrote customUI/customUI14.xml"

    # Update [Content_Types].xml
    $contentTypesEntry = $zip.GetEntry("[Content_Types].xml")
    if ($contentTypesEntry -ne $null) {
        $reader = New-Object System.IO.StreamReader($contentTypesEntry.Open())
        $ctContent = $reader.ReadToEnd()
        $reader.Close()

        # Check if already has customUI override
        if ($ctContent -notmatch 'customUI14\.xml') {
            # Add Override entry before closing tag
            $newOverride = '<Override PartName="/customUI/customUI14.xml" ContentType="application/vnd.ms-office.activeX+xml"/>'
            $ctContent = $ctContent -replace '</Types>', "$newOverride`n</Types>"

            $contentTypesEntry.Delete()
            $newCtEntry = $zip.CreateEntry("[Content_Types].xml")
            $ctWriter = New-Object System.IO.StreamWriter($newCtEntry.Open())
            $ctWriter.Write($ctContent)
            $ctWriter.Close()
            Write-Host "  Updated [Content_Types].xml"
        } else {
            Write-Host "  [Content_Types].xml already contains customUI reference"
        }
    }

    # Update _rels/.rels to reference customUI
    $relsEntry = $zip.GetEntry("_rels/.rels")
    if ($relsEntry -ne $null) {
        $reader = New-Object System.IO.StreamReader($relsEntry.Open())
        $relsContent = $reader.ReadToEnd()
        $reader.Close()

        if ($relsContent -notmatch 'customUI14\.xml') {
            $customUIRel = '<Relationship Id="rIdCustomUI" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>'
            $relsContent = $relsContent -replace '</Relationships>', "$customUIRel`n</Relationships>"

            $relsEntry.Delete()
            $newRelsEntry = $zip.CreateEntry("_rels/.rels")
            $relsWriter = New-Object System.IO.StreamWriter($newRelsEntry.Open())
            $relsWriter.Write($relsContent)
            $relsWriter.Close()
            Write-Host "  Updated _rels/.rels"
        } else {
            Write-Host "  _rels/.rels already contains customUI relationship"
        }
    }

    $zip.Dispose()
    Write-Host "  Ribbon XML injection complete"

} catch {
    Write-Warning "Failed to inject Ribbon XML: $_"
    Write-Warning "The add-in will work but without the custom Ribbon tab."
}

# ---------------------------------------------------------------------------
# Step 5: Copy to AddIns folder
# ---------------------------------------------------------------------------
if (-not $SkipInstall) {
    Write-Host "[5/5] Installing to: $OutputPath"

    $outputDir = Split-Path -Parent $OutputPath
    if (-not (Test-Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }

    # Close Excel if open and has this add-in loaded (best effort)
    $excelProcs = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue
    if ($excelProcs) {
        Write-Warning "Excel is running. Please close Excel before installing the add-in."
        $cont = Read-Host "Continue anyway? (y/N)"
        if ($cont -ne 'y' -and $cont -ne 'Y') {
            Write-Host "Installation cancelled."
            exit 0
        }
    }

    Copy-Item $workXlamPath $OutputPath -Force
    Write-Host "  Installed: $OutputPath"
} else {
    Write-Host "[5/5] Skipping install (--SkipInstall flag set)"
    $localOutput = Join-Path $ScriptDir "ProtocolDEE.xlam"
    Copy-Item $workXlamPath $localOutput -Force
    Write-Host "  Output: $localOutput"
}

# Cleanup temp files
Remove-Item $tempXlamPath -Force -ErrorAction SilentlyContinue
Remove-Item $workXlamPath -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "============================================"
Write-Host "  Build COMPLETE"
Write-Host "============================================"
Write-Host ""
Write-Host "To activate the add-in in Excel:"
Write-Host "  File > Options > Add-ins > Manage: Excel Add-ins > Go..."
Write-Host "  Check 'Protocol DEE' and click OK"
Write-Host ""
Write-Host "IMPORTANT -- Trust Center:"
Write-Host "  File > Options > Trust Center > Trust Center Settings > Macro Settings"
Write-Host "  Enable: 'Trust access to the VBA project object model'"
