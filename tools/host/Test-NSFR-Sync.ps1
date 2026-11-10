<#
.SYNOPSIS
    Test the template and source synchronization round trip on the supported host.

.DESCRIPTION
    Builds a macro-enabled workbook from templates/NSFR_Template.xlsx outside
    the repository and runs NSFR_Host_Sync against this clean clone:

    1. checks that the worktree is clean and takes the commit SHA;
    2. opens a copy of the template in a new Excel instance, saves it as
       Before.xlsm, imports src/host/NSFR_Host_Sync.bas (the bootstrap a
       person does by hand) and saves again;
    3. saves the workbook as After.xlsm and runs SyncImport twice: every
       component of src/manifest.txt but the sync module is imported, and a
       backup copy is written each time;
    4. compiles the project and runs Test_NSFR_HarnessSelfTest.RunSelfTest;
    5. runs SyncExport into this clone, saves, closes and quits Excel;
    6. checks that the export left the worktree clean (the round trip changed
       no byte Git stores), that the sheets, styles, shared strings and
       workbook names of Before.xlsm and After.xlsm are identical, and that
       the backups exist;
    7. prints a summary, restores any file the export changed, and removes
       the staging folder unless -Keep is given.

    Importing modules needs Excel's "Trust access to the VBA project object
    model". The script never changes that setting or any other Trust Center
    or macro setting: when the access is off, it stops.

    This Source Code Form is subject to the terms of the Mozilla Public
    License, v. 2.0. If a copy of the MPL was not distributed with this
    file, You can obtain one at https://mozilla.org/MPL/2.0/.

.PARAMETER Hidden
    Hide the Excel window during the run.

.PARAMETER Keep
    Keep the staging folder with Before.xlsm, After.xlsm and the backups.

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File tools\host\Test-NSFR-Sync.ps1
#>
[CmdletBinding()]
param(
    [switch]$Hidden,
    [switch]$Keep
)

Set-StrictMode -Version 3
$ErrorActionPreference = 'Stop'

$CompileControlId = 578     # Debug > Compile VBAProject in the Visual Basic Editor
$MacroEnabledFormat = 52    # xlOpenXMLWorkbookMacroEnabled
$SelfTest = 'Test_NSFR_HarnessSelfTest.RunSelfTest'

function Invoke-Git {
    # Windows PowerShell 5.1 turns a native command's stderr into an error under
    # 'Stop', so the preference is relaxed here and the exit code decides.
    param([Parameter(Mandatory)][string[]]$Arguments)
    $ErrorActionPreference = 'Continue'
    $output = & git -C $script:Root @Arguments 2>$null
    if ($LASTEXITCODE -ne 0) { throw "git $($Arguments -join ' ') failed." }
    return $output
}

function Clear-ComObject {
    param($Object)
    if ($null -ne $Object) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function ConvertTo-CrLfCopy {
    # The Visual Basic Editor imports a module reliably only with CR LF.
    param([string]$Source, [string]$Target)
    $text = [IO.File]::ReadAllText($Source) -replace "`r`n", "`n" -replace "`n", "`r`n"
    [IO.File]::WriteAllText($Target, $text, (New-Object Text.ASCIIEncoding))
}

function Get-LayoutParts {
    # The parts of a saved workbook that hold its sheets, styles, strings and
    # names, as SHA-256 per part; the VBA project and properties are left out.
    param([string]$Path)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $parts = [ordered]@{}
    $zip = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        foreach ($entry in ($zip.Entries | Sort-Object FullName)) {
            $name = $entry.FullName
            if ($name -notlike 'xl/worksheets/*' -and $name -ne 'xl/styles.xml' -and
                $name -ne 'xl/sharedStrings.xml' -and $name -ne 'xl/workbook.xml') { continue }
            $reader = New-Object IO.StreamReader($entry.Open())
            try { $text = $reader.ReadToEnd() } finally { $reader.Dispose() }
            if ($name -eq 'xl/workbook.xml') {
                # Only the defined names: the rest records the VBA project and the save.
                $match = [regex]::Match($text, '<definedNames>.*</definedNames>')
                $text = $match.Value
                $name = 'xl/workbook.xml#definedNames'
            }
            $bytes = [Text.Encoding]::UTF8.GetBytes($text)
            $hash = [Security.Cryptography.SHA256]::Create()
            try { $parts[$name] = [BitConverter]::ToString($hash.ComputeHash($bytes)) }
            finally { $hash.Dispose() }
        }
    } finally {
        $zip.Dispose()
    }
    return $parts
}

function Get-RunResult {
    # The last line of a report: RESULT PASS or RESULT FAIL ...
    param([string]$Report)
    $lines = @($Report -split "`r`n|`n" | Where-Object { $_ -ne '' })
    if ($lines.Count -eq 0) { return '' }
    return $lines[-1]
}

$script:Root = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$checks = [ordered]@{}
$log = New-Object System.Collections.Generic.List[string]

$sha = ([string](Invoke-Git @('rev-parse', 'HEAD'))).Trim()
if (Invoke-Git @('status', '--porcelain')) {
    throw 'The worktree is not clean: commit, stash or discard the changes, or use a fresh clone.'
}
$staging = Join-Path ([IO.Path]::GetTempPath()) "k-nsfr-sync-$($sha.Substring(0, 12))"
if (Test-Path $staging) { Remove-Item -Path $staging -Recurse -Force }
[void](New-Item -ItemType Directory -Path $staging)
$templateCopy = Join-Path $staging 'NSFR_Template.xlsx'
Copy-Item -Path (Join-Path $script:Root 'templates\NSFR_Template.xlsx') -Destination $templateCopy
$bootstrap = Join-Path $staging 'NSFR_Host_Sync.bas'
ConvertTo-CrLfCopy (Join-Path $script:Root 'src\host\NSFR_Host_Sync.bas') $bootstrap
$before = Join-Path $staging 'Before.xlsm'
$after = Join-Path $staging 'After.xlsm'

$excel = $null
$workbook = $null
$project = $null
$compile = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = -not $Hidden
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $workbook = $excel.Workbooks.Open($templateCopy)
    $workbook.SaveAs($before, $MacroEnabledFormat)
    try {
        $project = $workbook.VBProject
        $null = $project.VBComponents.Count
    } catch {
        throw ("Excel does not allow access to the VBA project. Turn on 'Trust access to the " +
               "VBA project object model' in File > Options > Trust Center > Trust Center " +
               'Settings > Macro Settings, then run the script again.')
    }
    [void]$project.VBComponents.Import($bootstrap)
    $workbook.Save()
    $workbook.SaveAs($after, $MacroEnabledFormat)
    $prefix = "'$($workbook.Name)'!NSFR_Host_Sync"

    foreach ($round in 1, 2) {
        $report = [string]$excel.Run("$prefix.SyncImport", $script:Root)
        $log.Add("SyncImport $round"); $log.AddRange([string[]]@($report -split "`r`n"))
        $checks["import $round"] = if ((Get-RunResult $report) -eq 'RESULT PASS') { 'PASS' } else { 'FAIL' }
    }

    $excel.VBE.ActiveVBProject = $project
    $compile = $excel.VBE.CommandBars.FindControl([Reflection.Missing]::Value, $CompileControlId)
    if ($compile.Enabled) { [void]$compile.Execute() }
    $checks['compile'] = if ($compile.Enabled) { 'FAIL' } else { 'PASS' }

    if ($checks['compile'] -eq 'PASS') {
        $selfTest = [string]$excel.Run("'$($workbook.Name)'!$SelfTest")
        $log.Add('Self-test'); $log.AddRange([string[]]@($selfTest -split "`r`n"))
        $checks['self-test'] = if ((Get-RunResult $selfTest) -eq 'RESULT PASS') { 'PASS' } else { 'FAIL' }
    } else {
        $checks['self-test'] = 'NOT_RUN'
    }

    $report = [string]$excel.Run("$prefix.SyncExport", $script:Root)
    $log.Add('SyncExport'); $log.AddRange([string[]]@($report -split "`r`n"))
    $checks['export'] = if ((Get-RunResult $report) -eq 'RESULT PASS') { 'PASS' } else { 'FAIL' }
    $workbook.Save()
} finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch { $log.Add("Close failed: $_") } }
    if ($null -ne $excel) { try { $excel.Quit() } catch { $log.Add("Quit failed: $_") } }
    Clear-ComObject $compile
    Clear-ComObject $project
    Clear-ComObject $workbook
    Clear-ComObject $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# The round trip: the export must leave every byte Git stores unchanged.
$changed = @(Invoke-Git @('status', '--porcelain'))
$checks['round trip'] = if ($changed.Count -eq 0) { 'PASS' } else { 'FAIL' }
if ($changed.Count -gt 0) {
    $log.Add('Changed by the export:'); $log.AddRange([string[]]$changed)
    $log.AddRange([string[]]@(Invoke-Git @('diff', '--stat')))
    [void](Invoke-Git @('checkout', '--', '.'))
}

# The layout: sheets, styles, strings and names are those the template gave.
$partsBefore = Get-LayoutParts $before
$partsAfter = Get-LayoutParts $after
$differ = @($partsBefore.Keys | Where-Object { $partsAfter[$_] -ne $partsBefore[$_] }) +
          @($partsAfter.Keys | Where-Object { -not $partsBefore.Contains($_) })
$checks['layout preserved'] = if ($differ.Count -eq 0) { 'PASS' } else { 'FAIL' }
if ($differ.Count -gt 0) { $log.Add("Layout parts that differ: $($differ -join ', ')") }

$backups = @(Get-ChildItem -Path $staging -Filter 'After.backup-*.xlsm')
$checks['backup'] = if ($backups.Count -ge 1) { 'PASS' } else { 'FAIL' }

$result = if (@($checks.Values | Where-Object { $_ -ne 'PASS' }).Count -eq 0) { 'PASS' } else { 'FAIL' }
$rule = '-' * 70
Write-Host $rule
Write-Host 'K-NSFR template and synchronization round trip'
Write-Host "Commit: $sha"
foreach ($key in $checks.Keys) { Write-Host ('{0,-18} {1}' -f $key, $checks[$key]) }
Write-Host "Result: $result"
Write-Host ''
$log | ForEach-Object { Write-Host $_ }
Write-Host $rule
if ($Keep) {
    Write-Host "Staging folder kept: $staging"
} else {
    Remove-Item -Path $staging -Recurse -Force
}
if ($result -ne 'PASS') { exit 1 }
