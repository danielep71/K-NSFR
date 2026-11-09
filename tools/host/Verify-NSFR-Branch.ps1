<#
.SYNOPSIS
    Verify one branch of K-NSFR on the supported host in one command.

.DESCRIPTION
    Fetches the branch from origin, checks its head out in a separate Git
    worktree beside this clone, runs that commit's own
    tools/host/Record-HostEvidence.ps1 there, and prints a summary to report:
    the branch, the commit, the result and the complete host evidence record
    (docs/validation/HOST_EVIDENCE.md).

    Everything is written to a bundle folder beside this clone,
    ..\nsfr-verify-<first 7 characters of the SHA>:

    - <SHA>.json         the host evidence record, ready for a pull request;
    - run.log            the output of Record-HostEvidence.ps1;
    - verify-summary.txt the summary printed at the end.

    The worktree is removed at the end; this clone and its branch are not
    changed. Like Record-HostEvidence.ps1, the script never changes a Trust
    Center or macro setting. It needs Git, Windows 11 and Microsoft 365
    desktop Excel, 64-bit; Python is optional.

    This Source Code Form is subject to the terms of the Mozilla Public
    License, v. 2.0. If a copy of the MPL was not distributed with this
    file, You can obtain one at https://mozilla.org/MPL/2.0/.

.PARAMETER Branch
    The branch to verify, for example release/0.1.0 or main.

.PARAMETER Force
    Replace the bundle folder of an earlier run of the same commit.

.PARAMETER Hidden
    Hide the Excel window during the run.

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\host\Verify-NSFR-Branch.ps1 -Branch release/0.1.0
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Branch,
    [switch]$Force,
    [switch]$Hidden
)

Set-StrictMode -Version 3
$ErrorActionPreference = 'Stop'

function Invoke-Git {
    # Windows PowerShell 5.1 turns a native command's stderr into an error under
    # 'Stop', so the preference is relaxed here and the exit code decides.
    param([Parameter(Mandatory)][string[]]$Arguments)
    $ErrorActionPreference = 'Continue'
    $output = & git -C $script:Repo @Arguments 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "git $($Arguments -join ' ') failed: $(($output | Out-String).Trim())"
    }
    return $output
}

function Write-Step {
    param([string]$Text)
    $line = '{0}  {1}' -f (Get-Date -Format 'HH:mm:ss'), $Text
    Write-Host $line
    $script:Summary.Add($line)
}

if ($Branch -notmatch '^[A-Za-z0-9._/-]+$' -or $Branch -match '\.\.') {
    throw "Not a branch name: $Branch"
}
$script:Repo = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$script:Summary = New-Object System.Collections.Generic.List[string]

[void](Invoke-Git @('fetch', '--quiet', 'origin', "+refs/heads/$($Branch):refs/remotes/origin/$Branch"))
$sha = ([string](Invoke-Git @('rev-parse', "refs/remotes/origin/$Branch^{commit}"))).Trim()
$bundle = Join-Path (Split-Path $script:Repo -Parent) "nsfr-verify-$($sha.Substring(0, 7))"
$checkout = Join-Path $bundle 'checkout'
if (Test-Path $bundle) {
    if (-not $Force) {
        throw "$bundle already exists. Remove it, or run again with -Force to replace it."
    }
    if (Test-Path $checkout) {
        $ErrorActionPreference = 'Continue'
        & git -C $script:Repo worktree remove --force $checkout 2>&1 | Out-Null
        $ErrorActionPreference = 'Stop'
    }
    Remove-Item -Path $bundle -Recurse -Force
}
[void](New-Item -ItemType Directory -Path $bundle)

Write-Step "K-NSFR host verification of $Branch at $sha"
Write-Step "Bundle folder: $bundle"
$result = 'NOT_RUN'
$record = $null
try {
    [void](Invoke-Git @('worktree', 'add', '--detach', $checkout, $sha))
    Write-Step 'Worktree checked out; running that commit''s Record-HostEvidence.ps1'
    $recorder = Join-Path $checkout 'tools\host\Record-HostEvidence.ps1'
    if (-not (Test-Path $recorder)) {
        throw "$sha has no tools\host\Record-HostEvidence.ps1: it predates the host evidence script."
    }
    $arguments = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $recorder)
    if ($Hidden) { $arguments += '-Hidden' }
    $ErrorActionPreference = 'Continue'
    & powershell.exe @arguments 2>&1 | Tee-Object -FilePath (Join-Path $bundle 'run.log') |
        ForEach-Object { Write-Host "  record: $_" }
    $exitCode = $LASTEXITCODE
    $ErrorActionPreference = 'Stop'
    Write-Step "Record-HostEvidence.ps1 exited with code $exitCode"

    $written = Join-Path $checkout "tests\evidence\$sha.json"
    if (Test-Path $written) {
        $record = Join-Path $bundle "$sha.json"
        Copy-Item -Path $written -Destination $record
        $result = [string]((Get-Content -Raw -Path $record | ConvertFrom-Json).result)
    } else {
        Write-Step 'No record was written; run.log gives the reason'
    }
} finally {
    if (Test-Path $checkout) {
        $ErrorActionPreference = 'Continue'
        & git -C $script:Repo worktree remove --force $checkout 2>&1 | Out-Null
        & git -C $script:Repo worktree prune 2>&1 | Out-Null
        $ErrorActionPreference = 'Stop'
    }
}
Write-Step 'Worktree removed'

$rule = '-' * 70
$report = New-Object System.Collections.Generic.List[string]
$report.Add($rule)
$report.Add('K-NSFR host verification: Record-HostEvidence.ps1 on one exact commit')
$report.Add("Branch: $Branch")
$report.Add("Commit: $sha")
$report.Add("Result: $result")
if ($null -ne $record) {
    $report.Add('')
    $report.Add("Host evidence record ($sha.json, also in the bundle folder):")
    foreach ($line in (Get-Content -Path $record)) { $report.Add($line) }
}
$report.Add($rule)
$report.Add('Report the text above. Propose the record in a pull request as')
$report.Add("tests/evidence/$sha.json, or send it for review.")
$report | ForEach-Object { Write-Host $_ }
$script:Summary.AddRange($report)
[IO.File]::WriteAllLines((Join-Path $bundle 'verify-summary.txt'), $script:Summary,
                         (New-Object Text.UTF8Encoding $false))
if ($result -ne 'PASS') { exit 1 }
