<#
.SYNOPSIS
    Test the template and source synchronization round trip on the supported
    host and write its host evidence record.

.DESCRIPTION
    Builds a macro-enabled workbook from templates/NSFR_Template.xlsx outside
    the repository and runs NSFR_Host_Sync against this clean clone:

    1. checks that the worktree is clean and takes the commit SHA;
    2. takes the SHA-256 of the template, src/manifest.txt, every file it
       lists and .gitattributes, over the bytes stored in Git, and reads the
       Windows and Excel versions, builds and bitness;
    3. opens a copy of the template in a new Excel instance, saves it as
       Before.xlsm, imports src/host/NSFR_Host_Sync.bas (the bootstrap a
       person does by hand) and saves again;
    4. saves the workbook as After.xlsm and runs SyncImport twice: every
       component of src/manifest.txt but the sync module is imported, and a
       backup copy is written each time;
    5. compiles the project and runs Test_NSFR_HarnessSelfTest.RunSelfTest;
    6. runs SyncExport into this clone, saves, closes and quits Excel;
    7. checks that the export left the worktree clean (the round trip changed
       no byte Git stores); that the sheets, styles, shared strings and
       workbook names of Before.xlsm and After.xlsm are identical, and the
       sheets (names, order and visibility) and workbook names of After.xlsm
       are the template's; and that the backups exist;
    8. restores any file the export changed, writes
       tests/evidence/sync/<SHA>.json and, when Python is available,
       validates it with tools/check_host_evidence.py; prints a summary and
       removes the staging folder unless -Keep is given.

    Importing modules needs Excel's "Trust access to the VBA project object
    model". The script never changes that setting or any other Trust Center
    or macro setting: when the access is off, the import is recorded as not
    run, with the reason.

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

$Schema = 'k-nsfr-sync-evidence/1'
$CompileControlId = 578     # Debug > Compile VBAProject in the Visual Basic Editor
$MacroEnabledFormat = 52    # xlOpenXMLWorkbookMacroEnabled
$SelfTest = 'Test_NSFR_HarnessSelfTest.RunSelfTest'

function Invoke-Git {
    # Windows PowerShell 5.1 turns a native command's stderr into an error under
    # 'Stop', so the preference is relaxed here and the exit code decides.
    param([Parameter(Mandatory)][string[]]$Arguments, [switch]$AllowFailure)
    $ErrorActionPreference = 'Continue'
    $output = & git -C $script:Root @Arguments 2>$null
    if ($LASTEXITCODE -ne 0 -and -not $AllowFailure) { throw "git $($Arguments -join ' ') failed." }
    return $output
}

function Get-RegistryValue {
    param([string]$Path, [string]$Name)
    try { return (Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop).$Name }
    catch { return $null }
}

function Get-LfBytes {
    # A text file as Git stores it: every CR LF of the checkout becomes LF.
    param([string]$Path)
    $bytes = [IO.File]::ReadAllBytes($Path)
    $result = New-Object System.Collections.Generic.List[byte] ($bytes.Length)
    for ($i = 0; $i -lt $bytes.Length; $i++) {
        if ($bytes[$i] -eq 13 -and $i + 1 -lt $bytes.Length -and $bytes[$i + 1] -eq 10) { continue }
        $result.Add($bytes[$i])
    }
    return ,$result.ToArray()
}

function Get-Sha256 {
    param([byte[]]$Bytes)
    $algorithm = [Security.Cryptography.SHA256]::Create()
    try { return (($algorithm.ComputeHash($Bytes) | ForEach-Object { $_.ToString('x2') }) -join '') }
    finally { $algorithm.Dispose() }
}

function Get-PublicText {
    # Text without the local paths, which must not be published.
    param([string]$Text)
    $text = $Text
    foreach ($path in @($script:Staging, [IO.Path]::GetTempPath().TrimEnd('\'), $env:TEMP,
                        $script:Root, $env:USERPROFILE)) {
        if ($path) { $text = $text.Replace($path, '<local path>') }
    }
    return $text
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

function Get-PartText {
    # One part of a workbook, as text, or $null when it is absent.
    param([string]$Path, [string]$Part)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entry = $zip.GetEntry($Part)
        if ($null -eq $entry) { return $null }
        $reader = New-Object IO.StreamReader($entry.Open())
        try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
    } finally {
        $zip.Dispose()
    }
}

function Get-LayoutParts {
    # The parts of a saved workbook that hold its sheets, styles, strings and
    # names, as SHA-256 per part; the VBA project and properties are left out.
    param([string]$Path)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $parts = [ordered]@{}
    $zip = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $names = @($zip.Entries | ForEach-Object { $_.FullName } | Sort-Object)
    } finally {
        $zip.Dispose()
    }
    foreach ($name in $names) {
        if ($name -notlike 'xl/worksheets/*' -and $name -ne 'xl/styles.xml' -and
            $name -ne 'xl/sharedStrings.xml' -and $name -ne 'xl/workbook.xml') { continue }
        $text = Get-PartText $Path $name
        $texts = [ordered]@{}
        if ($name -eq 'xl/workbook.xml') {
            # Only the sheets and the defined names: the rest records the VBA
            # project and the save.
            $texts['xl/workbook.xml#sheets'] = [regex]::Match($text, '<sheets>.*</sheets>').Value
            $texts['xl/workbook.xml#definedNames'] = [regex]::Match($text, '<definedNames>.*</definedNames>').Value
        } else {
            $texts[$name] = $text
        }
        foreach ($key in $texts.Keys) {
            $parts[$key] = Get-Sha256 ([Text.Encoding]::UTF8.GetBytes($texts[$key]))
        }
    }
    return $parts
}

function Get-WorkbookShape {
    # The sheets in order with their visibility, and the workbook names with
    # their references, as lines that do not depend on which program saved
    # the file. Excel's own _xlnm names are left out.
    param([string]$Path)
    $text = Get-PartText $Path 'xl/workbook.xml'
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($sheet in [regex]::Matches($text, '<sheet\s[^>]*?/>')) {
        $name = [regex]::Match($sheet.Value, '\sname="([^"]*)"').Groups[1].Value
        $state = [regex]::Match($sheet.Value, '\sstate="([^"]*)"').Groups[1].Value
        if (-not $state) { $state = 'visible' }
        $lines.Add("sheet $name $state")
    }
    $names = New-Object System.Collections.Generic.List[string]
    foreach ($defined in [regex]::Matches($text, '<definedName\s([^>]*)>([^<]*)</definedName>')) {
        $name = [regex]::Match($defined.Groups[1].Value, 'name="([^"]*)"').Groups[1].Value
        if ($name -like '_xlnm.*') { continue }
        $scope = [regex]::Match($defined.Groups[1].Value, 'localSheetId="([^"]*)"').Groups[1].Value
        $names.Add("name $name $scope $($defined.Groups[2].Value.Replace("'", ''))")
    }
    $lines.AddRange([string[]]@($names | Sort-Object))
    return ,$lines.ToArray()
}

function Get-RunResult {
    # The last line of a report: RESULT PASS or RESULT FAIL ...
    param([string]$Report)
    $lines = @($Report -split "`r`n|`n" | Where-Object { $_ -ne '' })
    if ($lines.Count -eq 0) { return '' }
    return $lines[-1]
}

function Add-Report {
    param([string]$Title, [string]$Report)
    $script:Log.Add($Title)
    $script:Log.AddRange([string[]]@($Report -split "`r`n|`n" | Where-Object { $_ -ne '' }))
}

$script:Root = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$script:Staging = $null
$script:Log = New-Object System.Collections.Generic.List[string]
$reasons = New-Object System.Collections.Generic.List[string]
$checks = [ordered]@{ import_1 = 'NOT_RUN'; import_2 = 'NOT_RUN'; compile = 'NOT_RUN';
                      self_test = 'NOT_RUN'; export = 'NOT_RUN'; round_trip = 'NOT_RUN';
                      layout = 'NOT_RUN'; backup = 'NOT_RUN' }

# 1. Candidate and worktree
$sha = ([string](Invoke-Git @('rev-parse', 'HEAD'))).Trim()
if (Invoke-Git @('status', '--porcelain')) {
    throw 'The worktree is not clean: commit, stash or discard the changes, or use a fresh clone.'
}
$published = $false
$refs = @(Invoke-Git @('for-each-ref', '--format=%(refname:short)', 'refs/remotes/origin/release/')) + @('origin/main')
foreach ($ref in $refs) {
    if (-not $ref) { continue }
    [void](Invoke-Git @('merge-base', '--is-ancestor', $sha, $ref) -AllowFailure)
    if ($LASTEXITCODE -eq 0) { $published = $true }
}
if (-not $published) {
    Write-Warning ("$sha is not on origin/main or an origin/release/* branch: the record can be " +
                   'bound only to a commit of those branches (git fetch, then check one out).')
}

# 2. Fingerprints over the bytes stored in Git, and the host
$fingerprints = [ordered]@{}
$template = Join-Path $script:Root 'templates\NSFR_Template.xlsx'
$fingerprints['templates/NSFR_Template.xlsx'] = Get-Sha256 ([IO.File]::ReadAllBytes($template))
$manifest = Join-Path $script:Root 'src\manifest.txt'
$fingerprints['src/manifest.txt'] = Get-Sha256 (Get-LfBytes $manifest)
foreach ($line in [IO.File]::ReadAllLines($manifest)) {
    $fields = @($line.Trim() -split '\s+')
    if ($line.StartsWith('#') -or $fields.Count -ne 2) { continue }
    $fingerprints[$fields[1]] = Get-Sha256 (Get-LfBytes (Join-Path $script:Root $fields[1]))
}
$fingerprints['.gitattributes'] = Get-Sha256 (Get-LfBytes (Join-Path $script:Root '.gitattributes'))

$currentVersion = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
$caption = ((Get-CimInstance -ClassName Win32_OperatingSystem).Caption -replace '^Microsoft\s+', '').Trim()
$os = '{0} {1}, OS build {2}.{3}' -f $caption,
    (Get-RegistryValue $currentVersion 'DisplayVersion'),
    (Get-RegistryValue $currentVersion 'CurrentBuild'),
    (Get-RegistryValue $currentVersion 'UBR')
$clickToRun = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
$officeBuild = Get-RegistryValue $clickToRun 'VersionToReport'
if (-not $officeBuild) {
    throw 'No Click-to-Run Office found: the supported host is Microsoft 365 desktop Excel.'
}
$products = [string](Get-RegistryValue $clickToRun 'ProductReleaseIds')
$product = if ($products -match 'O365|M365') { 'Microsoft 365 Apps' } else { 'Office' }
$bitness = if ((Get-RegistryValue $clickToRun 'Platform') -eq 'x64') { '64-bit' } else { '32-bit' }
$build = ($officeBuild -split '\.', 3)[2]     # 16.0.18129.20158 -> 18129.20158

# 3. to 6. The workbook, the imports, the self-test and the export
$script:Staging = Join-Path ([IO.Path]::GetTempPath()) "k-nsfr-sync-$($sha.Substring(0, 12))"
if (Test-Path $script:Staging) { Remove-Item -Path $script:Staging -Recurse -Force }
[void](New-Item -ItemType Directory -Path $script:Staging)
$templateCopy = Join-Path $script:Staging 'NSFR_Template.xlsx'
Copy-Item -Path $template -Destination $templateCopy
$bootstrap = Join-Path $script:Staging 'NSFR_Host_Sync.bas'
ConvertTo-CrLfCopy (Join-Path $script:Root 'src\host\NSFR_Host_Sync.bas') $bootstrap
$before = Join-Path $script:Staging 'Before.xlsm'
$after = Join-Path $script:Staging 'After.xlsm'

$excelText = $null
$excel = $null
$workbook = $null
$project = $null
$compile = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excelText = '{0} ({1}), Version {2} (Build {3} Click-to-Run)' -f $product, $products, $excel.Version, $build
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
        Add-Report "SyncImport $round" $report
        $checks["import_$round"] = if ((Get-RunResult $report) -eq 'RESULT PASS') { 'PASS' } else { 'FAIL' }
    }

    $excel.VBE.ActiveVBProject = $project
    $compile = $excel.VBE.CommandBars.FindControl([Reflection.Missing]::Value, $CompileControlId)
    if ($compile.Enabled) { [void]$compile.Execute() }
    $checks['compile'] = if ($compile.Enabled) { 'FAIL' } else { 'PASS' }

    if ($checks['compile'] -eq 'PASS') {
        $selfTestOutput = [string]$excel.Run("'$($workbook.Name)'!$SelfTest")
        Add-Report 'Self-test' $selfTestOutput
        $checks['self_test'] = if ((Get-RunResult $selfTestOutput) -eq 'RESULT PASS') { 'PASS' } else { 'FAIL' }
    } else {
        $reasons.Add('The project does not compile.')
    }

    $report = [string]$excel.Run("$prefix.SyncExport", $script:Root)
    Add-Report 'SyncExport' $report
    $checks['export'] = if ((Get-RunResult $report) -eq 'RESULT PASS') { 'PASS' } else { 'FAIL' }
    $workbook.Save()
} catch {
    $reasons.Add((Get-PublicText ($_.Exception.Message -replace '\s+', ' ')).Trim())
    # A self-test that a compiled project did not reach failed.
    if ($checks['compile'] -eq 'PASS' -and $checks['self_test'] -eq 'NOT_RUN') { $checks['self_test'] = 'FAIL' }
} finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch { $reasons.Add('Closing the workbook failed.') } }
    if ($null -ne $excel) { try { $excel.Quit() } catch { $reasons.Add('Quitting Excel failed.') } }
    Clear-ComObject $compile
    Clear-ComObject $project
    Clear-ComObject $workbook
    Clear-ComObject $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
if ($null -eq $excelText) { throw 'Excel could not be started through COM.' }

# 7. The round trip: the export must leave every byte Git stores unchanged.
$changed = @(Invoke-Git @('status', '--porcelain'))
if ($checks['export'] -ne 'NOT_RUN') {
    $checks['round_trip'] = if ($changed.Count -eq 0) { 'PASS' } else { 'FAIL' }
}
if ($changed.Count -gt 0) {
    $script:Log.Add('Changed by the export:'); $script:Log.AddRange([string[]]$changed)
    $script:Log.AddRange([string[]]@(Invoke-Git @('diff', '--stat')))
    [void](Invoke-Git @('checkout', '--', '.'))
    $reasons.Add('The export changed the worktree.')
}

# The layout: sheets, styles, strings and names are those the template gave.
if ($checks['import_1'] -ne 'NOT_RUN' -and (Test-Path $after)) {
    $partsBefore = Get-LayoutParts $before
    $partsAfter = Get-LayoutParts $after
    $differ = @($partsBefore.Keys | Where-Object { $partsAfter[$_] -ne $partsBefore[$_] }) +
              @($partsAfter.Keys | Where-Object { -not $partsBefore.Contains($_) })
    $shapeTemplate = Get-WorkbookShape $template
    $shapeAfter = Get-WorkbookShape $after
    $shapeDiffer = @(Compare-Object -ReferenceObject $shapeTemplate -DifferenceObject $shapeAfter -SyncWindow 0)
    $checks['layout'] = if ($differ.Count -eq 0 -and $shapeDiffer.Count -eq 0) { 'PASS' } else { 'FAIL' }
    if ($differ.Count -gt 0) { $script:Log.Add("Layout parts that differ from Before.xlsm: $($differ -join ', ')") }
    foreach ($item in $shapeDiffer) {
        $side = if ($item.SideIndicator -eq '<=') { 'template only' } else { 'workbook only' }
        $script:Log.Add("Shape ($side): $($item.InputObject)")
    }
    if ($checks['layout'] -ne 'PASS') { $reasons.Add('The layout differs from the template.') }
}

if ($checks['import_1'] -ne 'NOT_RUN') {
    $backups = @(Get-ChildItem -Path $script:Staging -Filter 'After.backup-*.xlsm')
    $checks['backup'] = if ($backups.Count -ge 2) { 'PASS' } else { 'FAIL' }
    if ($checks['backup'] -ne 'PASS') { $reasons.Add('Each import must leave its own backup.') }
}

# 8. The record
$values = @($checks.Values)
$result = if (@($values | Where-Object { $_ -ne 'PASS' }).Count -eq 0) { 'PASS' }
          elseif ($values -contains 'FAIL') { 'FAIL' } else { 'NOT_RUN' }
$record = [ordered]@{
    schema         = $Schema
    candidate_sha  = $sha
    recorded_on    = (Get-Date -Format 'yyyy-MM-dd')
    worktree_clean = $true
    host           = [ordered]@{ os = $os; excel = $excelText; bitness = $bitness }
    fingerprints   = $fingerprints
    checks         = $checks
    result         = $result
    log            = @($script:Log | ForEach-Object { Get-PublicText $_ })
}
if ($result -ne 'PASS') {
    $record.reason = if ($reasons.Count -gt 0) { $reasons -join ' ' } else { 'A check did not pass.' }
}
$json = (ConvertTo-Json -InputObject $record -Depth 5) -replace "`r`n", "`n"
$folder = Join-Path $script:Root 'tests\evidence\sync'
if (-not (Test-Path $folder)) { [void](New-Item -ItemType Directory -Path $folder) }
$path = Join-Path $script:Root "tests\evidence\sync\$sha.json"
[IO.File]::WriteAllText($path, $json + "`n", (New-Object Text.UTF8Encoding $false))

$rule = '-' * 70
Write-Host $rule
Write-Host 'K-NSFR template and synchronization round trip'
Write-Host "Commit: $sha"
foreach ($key in $checks.Keys) { Write-Host ('{0,-18} {1}' -f $key, $checks[$key]) }
Write-Host "Result: $result"
Write-Host ''
$script:Log | ForEach-Object { Write-Host (Get-PublicText $_) }
Write-Host $rule
Write-Host "Record written to tests\evidence\sync\$sha.json"

$python = Get-Command python, py -ErrorAction SilentlyContinue | Select-Object -First 1
if ($null -ne $python) {
    $arguments = @('tools/check_host_evidence.py')
    if ($python.Name -like 'py*.exe' -and $python.Name -notlike 'python*') { $arguments = @('-3') + $arguments }
    $ErrorActionPreference = 'Continue'
    & $python.Source @arguments
    $ErrorActionPreference = 'Stop'
} else {
    Write-Host 'Python not found: the record is validated by the host-evidence gate of its pull request.'
}
Write-Host 'Propose the record in a pull request that adds only the record.'

if ($Keep) {
    Write-Host "Staging folder kept: $script:Staging"
} else {
    Remove-Item -Path $script:Staging -Recurse -Force
}
if ($result -ne 'PASS') { exit 1 }
