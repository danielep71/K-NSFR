<#
.SYNOPSIS
    Run the neutral harness self-test in Excel and write its host evidence record.

.DESCRIPTION
    Records one run on the supported host (docs/validation/HOST_EVIDENCE.md)
    for the commit checked out in this repository:

    1. checks that the worktree is clean and takes the candidate SHA;
    2. takes the SHA-256 of .gitattributes and of the modules in tests/core,
       over the bytes stored in Git (LF), as the validator does;
    3. reads the Windows and Excel versions, builds and bitness;
    4. starts a new Excel instance, adds a blank workbook, imports CRLF copies
       of the modules, compiles the project and runs
       Test_NSFR_HarnessSelfTest.RunSelfTest;
    5. closes the workbook without saving, quits Excel and checks that the
       worktree is still clean;
    6. writes tests/evidence/<SHA>.json and, when Python is available,
       validates it with tools/check_host_evidence.py.

    Importing modules needs Excel's "Trust access to the VBA project object
    model" (File > Options > Trust Center > Trust Center Settings > Macro
    Settings). The script never changes that setting or any other Trust
    Center setting: when the access is off, it stops and writes no record.

    This Source Code Form is subject to the terms of the Mozilla Public
    License, v. 2.0. If a copy of the MPL was not distributed with this
    file, You can obtain one at https://mozilla.org/MPL/2.0/.

.PARAMETER Hidden
    Hide the Excel window during the run. By default it is visible, so that a
    compile error dialog can be seen and closed.

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File tools\host\Record-HostEvidence.ps1
#>
[CmdletBinding()]
param(
    [switch]$Hidden
)

Set-StrictMode -Version 3
$ErrorActionPreference = 'Stop'

$Schema = 'k-nsfr-host-evidence/1'
$SelfTest = 'Test_NSFR_HarnessSelfTest.RunSelfTest'
$CompileControlId = 578     # Debug > Compile VBAProject in the Visual Basic Editor

function Invoke-Git {
    # Windows PowerShell 5.1 turns a native command's stderr into an error under
    # 'Stop', so the preference is relaxed here and the exit code decides.
    param([Parameter(Mandatory)][string[]]$Arguments, [switch]$AllowFailure)
    $ErrorActionPreference = 'Continue'
    $output = & git @Arguments 2>$null
    if ($LASTEXITCODE -ne 0 -and -not $AllowFailure) { throw "git $($Arguments -join ' ') failed." }
    return $output
}

function Test-GitAncestor {
    param([string]$Commit, [string]$Ref)
    [void](Invoke-Git @('merge-base', '--is-ancestor', $Commit, $Ref) -AllowFailure)
    return $LASTEXITCODE -eq 0
}

function Test-CleanWorktree {
    return -not (Invoke-Git @('status', '--porcelain'))
}

function Get-RegistryValue {
    param([string]$Path, [string]$Name)
    try { return (Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop).$Name }
    catch { return $null }
}

function Get-LfBytes {
    # The file as Git stores it: every CR LF of the checkout becomes LF.
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

function ConvertTo-CrLfBytes {
    # The Visual Basic Editor imports a module reliably only with CR LF.
    param([byte[]]$LfBytes)
    $result = New-Object System.Collections.Generic.List[byte] ($LfBytes.Length + 1024)
    foreach ($byte in $LfBytes) {
        if ($byte -eq 10) { $result.Add(13) }
        $result.Add($byte)
    }
    return ,$result.ToArray()
}

function Get-PublicMessage {
    # Error text without the local paths, which must not be published.
    param([string]$Message)
    $text = $Message -replace '\s+', ' '
    foreach ($path in @($script:Root, $script:Staging, $env:USERPROFILE, $env:TEMP)) {
        if ($path) { $text = $text.Replace($path, '<local path>') }
    }
    return $text.Trim()
}

function Clear-ComObject {
    param($Object)
    if ($null -ne $Object) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

$script:Root = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$script:Staging = $null
Push-Location $script:Root
try {
    # 1. Candidate and worktree
    $sha = ([string](Invoke-Git @('rev-parse', 'HEAD'))).Trim()
    if (-not (Test-CleanWorktree)) {
        throw 'The worktree is not clean: commit, stash or discard the changes, or use a fresh clone.'
    }
    $published = $false
    foreach ($ref in @('origin/release/0.1.0', 'origin/main')) {
        if (Test-GitAncestor $sha $ref) { $published = $true }
    }
    if (-not $published) {
        Write-Warning ("$sha is not on origin/release/0.1.0 or origin/main: the record can be " +
                       'bound only to a commit of those branches (git fetch, then check one out).')
    }

    # 2. Fingerprints over the bytes stored in Git
    $modules = @(Get-ChildItem -Path 'tests\core' -File |
                 Where-Object { $_.Extension -in '.bas', '.cls', '.frm' } |
                 Sort-Object Name)
    if ($modules.Count -eq 0) { throw 'No module found in tests\core.' }
    $fingerprints = [ordered]@{}
    $fingerprints['.gitattributes'] = Get-Sha256 (Get-LfBytes (Join-Path $script:Root '.gitattributes'))
    $script:Staging = Join-Path ([IO.Path]::GetTempPath()) "k-nsfr-host-$($sha.Substring(0, 12))"
    if (Test-Path $script:Staging) { Remove-Item -Path $script:Staging -Recurse -Force }
    [void](New-Item -ItemType Directory -Path $script:Staging)
    $imports = @()
    foreach ($module in $modules) {
        $lf = Get-LfBytes $module.FullName
        $fingerprints["tests/core/$($module.Name)"] = Get-Sha256 $lf
        $copy = Join-Path $script:Staging $module.Name
        [IO.File]::WriteAllBytes($copy, (ConvertTo-CrLfBytes $lf))
        $imports += $copy
    }

    # 3. Host
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

    # 4. Import, compile and run in a new Excel instance
    $steps = [ordered]@{ import = 'NOT_RUN'; compile = 'NOT_RUN'; run = 'NOT_RUN'; cleanup = 'NOT_RUN' }
    $reasons = New-Object System.Collections.Generic.List[string]
    $output = $null
    $excelText = $null
    $excel = $null
    $workbook = $null
    $project = $null
    $compile = $null
    $closed = $true
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = -not $Hidden
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false
        $excelText = '{0} ({1}), Version {2} (Build {3} Click-to-Run)' -f $product, $products, $excel.Version, $build
        $workbook = $excel.Workbooks.Add()
        try {
            $project = $workbook.VBProject
            $null = $project.VBComponents.Count
        } catch {
            throw ("Excel does not allow access to the VBA project. Turn on 'Trust access to the " +
                   "VBA project object model' in File > Options > Trust Center > Trust Center " +
                   'Settings > Macro Settings, then run the script again. No record was written.')
        }

        try {
            foreach ($copy in $imports) { [void]$project.VBComponents.Import($copy) }
            $steps.import = 'PASS'
        } catch {
            $steps.import = 'FAIL'
            $reasons.Add("Import failed: $(Get-PublicMessage $_.Exception.Message)")
        }

        # Compilation: the command is disabled once the project is compiled
        if ($steps.import -eq 'PASS') {
            try {
                $excel.VBE.ActiveVBProject = $project
                $compile = $excel.VBE.CommandBars.FindControl([Reflection.Missing]::Value, $CompileControlId)
                if ($compile.Enabled) { [void]$compile.Execute() }
                if ($compile.Enabled) {
                    $steps.compile = 'FAIL'
                    $reasons.Add('Compile VBAProject did not complete: the project has a compile error.')
                } else {
                    $steps.compile = 'PASS'
                }
            } catch {
                $steps.compile = 'FAIL'
                $reasons.Add("Compilation failed: $(Get-PublicMessage $_.Exception.Message)")
            }
        }

        # The self-test, as the Immediate window would print it
        if ($steps.compile -eq 'PASS') {
            try {
                $text = [string]$excel.Run("'$($workbook.Name)'!$SelfTest")
                $lines = @($text -split "`r`n|`n")
                while ($lines.Count -gt 0 -and $lines[-1] -eq '') {
                    $lines = @($lines | Select-Object -First ($lines.Count - 1))
                }
                $output = $lines
                if ($lines.Count -gt 0 -and $lines[-1] -eq 'RESULT PASS') {
                    $steps.run = 'PASS'
                } else {
                    $steps.run = 'FAIL'
                    $reasons.Add('The self-test did not end with RESULT PASS.')
                }
            } catch {
                $steps.run = 'FAIL'
                $output = @()
                $reasons.Add("Run-time error: $(Get-PublicMessage $_.Exception.Message)")
            }
        }
    } finally {
        # 5. Cleanup: close without saving, quit, and check that nothing was left behind
        # Each call is tried on its own, so that a failed close still quits Excel.
        if ($null -ne $workbook) {
            try { $workbook.Close($false) } catch {
                $closed = $false
                $reasons.Add("Closing the workbook failed: $(Get-PublicMessage $_.Exception.Message)")
            }
        }
        if ($null -ne $excel) {
            try { $excel.Quit() } catch {
                $closed = $false
                $reasons.Add("Quitting Excel failed: $(Get-PublicMessage $_.Exception.Message)")
            }
        }
        Clear-ComObject $compile
        Clear-ComObject $project
        Clear-ComObject $workbook
        Clear-ComObject $excel
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
    if ($steps.import -ne 'NOT_RUN') {
        if ($closed -and (Test-CleanWorktree)) {
            $steps.cleanup = 'PASS'
        } else {
            $steps.cleanup = 'FAIL'
            if ($closed) { $reasons.Add('The run changed the worktree.') }
        }
    }
    if ($null -eq $excelText) { throw 'Excel could not be started through COM.' }

    # 6. The record
    $values = @($steps.Values)
    $result = if (@($values | Where-Object { $_ -ne 'PASS' }).Count -eq 0) { 'PASS' }
              elseif ($values -contains 'FAIL') { 'FAIL' } else { 'NOT_RUN' }
    $record = [ordered]@{
        schema         = $Schema
        candidate_sha  = $sha
        recorded_on    = (Get-Date -Format 'yyyy-MM-dd')
        worktree_clean = $true
        host           = [ordered]@{ os = $os; excel = $excelText; bitness = $bitness }
        fingerprints   = $fingerprints
        steps          = $steps
        result         = $result
    }
    if ($steps.run -ne 'NOT_RUN') { $record.output = @($output) }
    if ($result -ne 'PASS') { $record.reason = ($reasons -join ' ') }

    $json = (ConvertTo-Json -InputObject $record -Depth 5) -replace "`r`n", "`n"
    $path = Join-Path $script:Root "tests\evidence\$sha.json"
    [IO.File]::WriteAllText($path, $json + "`n", (New-Object Text.UTF8Encoding $false))
    Write-Host "Result: $result. Record written to tests\evidence\$sha.json"

    $python = Get-Command python, py -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($null -ne $python) {
        $arguments = @('tools/check_host_evidence.py', '--candidate', $sha)
        if ($python.Name -like 'py*.exe' -and $python.Name -notlike 'python*') { $arguments = @('-3') + $arguments }
        $ErrorActionPreference = 'Continue'
        & $python.Source @arguments
    } else {
        Write-Host 'Python not found: the record is validated by the host-evidence gate of its pull request.'
    }
    Write-Host 'Propose the record in a pull request, or send the file for review.'
} finally {
    if ($script:Staging -and (Test-Path $script:Staging)) {
        Remove-Item -Path $script:Staging -Recurse -Force
    }
    Pop-Location
}
