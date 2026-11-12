<a id="host-evidence"></a>

# 🖥️ Host evidence

**K-NSFR** · [Overview](../../README.md) · [Documentation](../../README.md#documentation) · [Governance](../GOVERNANCE.md)

[![Stage](https://img.shields.io/badge/Stage-Repository_setup-BF8700?style=flat-square)](../../README.md#status)

Host evidence records version **1**. Tracking: #14.

A host evidence record states what one run on the
[supported host](../IMPLEMENTATION.md#supported-and-tested-environments) did
with one exact commit: the fresh import, the compilation, the run of the
[neutral harness](../../tests/README.md) self-test and the cleanup. The
portable checks never run Excel, so a green Repository integrity check is
never host evidence. Only a valid record is.

`tools/check_host_evidence.py` (gate `host-evidence`) validates every record
in [`tests/evidence/`](../../tests/evidence) against the Git history. A
record there covers the neutral harness. A
[synchronization record](#synchronization-record), in `tests/evidence/sync/`,
covers the workbook template and the source synchronization (#13, #195).

<a id="record"></a>

## 🧾 Record

One JSON file per candidate, named `<candidate SHA>.json`, with exactly these
fields:

| Field | Content |
| --- | --- |
| `schema` | `k-nsfr-host-evidence/1` |
| `candidate_sha` | The full 40-character SHA of the tested commit. It must be a commit of a published release branch (`origin/release/*`; the active one is `release/0.2.0`) or of `main`: a task-branch commit disappears with its squash merge |
| `recorded_on` | The ISO date of the run, from the commit date to today |
| `worktree_clean` | `true`: `git status --porcelain` printed nothing before the import |
| `host.os` | The Windows edition, version and OS build, starting with `Windows 11`, for example `Windows 11 Pro 24H2, OS build 26100.2314` (**Settings → System → About**) |
| `host.excel` | The Microsoft 365 product and its product release IDs, then `Version … (Build … Click-to-Run)`: the scripts record the version Excel reports through COM, such as `16.0`, and the build of the installation, which **File → Account → About Excel** also shows |
| `host.bitness` | `64-bit` |
| `fingerprints` | The SHA-256 of every module imported from `tests/core/` and of `.gitattributes`, over the bytes stored in Git (LF). A Windows checkout holds the VBA source with CRLF, so `Get-FileHash` on it gives another value: `python tools/check_host_evidence.py --candidate <SHA> --inventory` prints the expected ones |
| `steps` | `import`, `compile`, `run` and `cleanup`, each `PASS`, `FAIL` or `NOT_RUN` |
| `result` | `PASS`, `FAIL` or `NOT_RUN`, as the steps give it |
| `output` | The complete Immediate-window output of the run, one string per line, without the trailing empty line; an empty list when a runtime error stopped the run before it printed; absent when the run did not happen |
| `reason` | Required unless the result is `PASS`: what failed or why a step did not run |

Example of a passing record. The SHA and fingerprints are placeholders:

```json
{
  "schema": "k-nsfr-host-evidence/1",
  "candidate_sha": "0123456789abcdef0123456789abcdef01234567",
  "recorded_on": "2026-11-08",
  "worktree_clean": true,
  "host": {
    "os": "Windows 11 Pro 24H2, OS build 26100.2314",
    "excel": "Microsoft 365 Apps (O365ProPlusRetail), Version 16.0 (Build 18129.20158 Click-to-Run)",
    "bitness": "64-bit"
  },
  "fingerprints": {
    ".gitattributes": "<SHA-256>",
    "tests/core/Test_NSFR_Harness.bas": "<SHA-256>",
    "tests/core/Test_NSFR_HarnessSelfTest.bas": "<SHA-256>"
  },
  "steps": {"import": "PASS", "compile": "PASS", "run": "PASS", "cleanup": "PASS"},
  "result": "PASS",
  "output": ["SUITE HARNESS-SELF-TEST", "H-001 PASS", "...", "RESULT PASS"]
}
```

<a id="recording-a-run"></a>

## ▶️ Recording a run

`tools/host/Verify-NSFR-Branch.ps1` verifies a branch in one command, from
any clone of the repository on the supported host:

```text
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\host\Verify-NSFR-Branch.ps1 -Branch release/0.2.0
```

It fetches the branch, checks its head out in a separate Git worktree beside
the clone, runs that commit's own `Record-HostEvidence.ps1` there, removes the
worktree, and prints a summary with the commit, the result and the complete
record. The bundle folder `..\nsfr-verify-<first 7 characters of the SHA>`
keeps the record (`<SHA>.json`), the run log and the summary. Your clone and
its branch are not changed. `-Force` replaces the bundle of an earlier run of
the same commit, and `-Hidden` hides Excel.

`tools/host/Record-HostEvidence.ps1` records a run of the commit already
checked out. You need the supported host and Git. Python is optional and
validates the record at the end.

1. Clone the repository fresh, or check that `git status --porcelain` prints
   nothing, and check out a commit of a published release branch (the active
   one is `release/0.2.0`) or of `main`.
2. Close Excel, then run from the repository folder:

   ```text
   powershell.exe -NoProfile -ExecutionPolicy Bypass -File tools\host\Record-HostEvidence.ps1
   ```

3. Propose the record that it writes, `tests/evidence/<SHA>.json`, in a pull
   request that adds only the record.

The script:

1. checks the clean worktree and takes the candidate SHA;
2. takes the fingerprints over the bytes stored in Git: it reads each file and
   turns CR LF into LF, which the clean worktree makes equal to the stored
   bytes;
3. reads the Windows version and build from the registry, and the Excel
   build and bitness from the Click-to-Run configuration;
4. drives a new Excel instance through COM: a new blank workbook, the import
   of CRLF copies of every module of `tests/core/`, made in a temporary
   folder because the Visual Basic Editor imports reliably only CR LF,
   *Compile VBAProject* and `Test_NSFR_HarnessSelfTest.RunSelfTest`;
5. closes the workbook without saving, quits Excel and checks that the
   worktree is still clean;
6. writes the record, and validates it with `check_host_evidence.py
   --candidate` when Python is available.

It records a failed step as `FAIL`, with the error message stripped of local
paths. It records the Excel line as `Microsoft 365 Apps (<products>),
Version 16.0 (Build <build> Click-to-Run)`: the build identifies the release
exactly.

Importing modules through COM needs Excel's *Trust access to the VBA project
object model* (**File → Options → Trust Center → Trust Center Settings →
Macro Settings**). The script never changes it, nor any other Trust Center or
macro setting: when the access is off, it stops before the import and writes
no record. Turn the access on yourself if you accept it, or use the manual
route. Excel stays visible during the run: if a compile error opens a
dialog, close it, and the step is recorded as `FAIL`. Pass `-Hidden` to hide
it.

Without the script, follow the [manual route](../../tests/README.md#running-in-excel):

1. Print the fingerprints with
   `python tools/check_host_evidence.py --candidate <SHA> --inventory`, on
   any machine with a clone.
2. Import, compile, run and close as the manual route says.
3. Write the record by hand, and run
   `python tools/check_host_evidence.py --candidate <SHA>` if you can.

A failed or interrupted run is recorded too, with `FAIL` or `NOT_RUN` and the
reason. If no supported host is available, the issue that needs the evidence
stays open with that blocker.

<a id="what-the-gate-rejects"></a>

## 🚦 What the gate rejects

- **Unbound evidence:**
  - an unknown candidate, one that is not an ancestor of the checked commit,
    or one absent from published `origin/main` and `origin/release/*` history;
    local task branches and synthetic PR merge ancestry are insufficient.
    Fetch those remote branches before validation; missing published refs fail
    validation when evidence records exist;
  - a file name that differs from the SHA;
  - a dirty worktree;
  - a fingerprint that differs from the files Git stores at the candidate, or a missing or extra one.
- **Another host:** anything but Windows 11 with its version and OS build, Microsoft 365 with its version and build, and 64-bit.
- **Inconsistent results:**
  - a result that the steps do not give, or a step that ran after a failed one;
  - a missing reason;
  - output for a run that did not happen;
  - a `PASS` run whose output is not exactly the self-test output documented in `tests/README.md` at the candidate;
  - a `FAIL` run whose output is that passing output.
- **Inconsistent or non-finite counts:**
  - a count that is not a non-negative integer;
  - counts that contradict the case, cleanup and problem lines;
  - a `RESULT` line that contradicts the counts.
- **Private details:**
  - user folder paths, network share paths, e-mail addresses and IP addresses;
  - any field outside the record. Do not add machine names, user names or license details.

With `--candidate <SHA>`, the checker also fails when that commit has no
harness record (*missing*), when only other commits have one (*stale*), or
when its record is not `PASS`. A record proves one commit only: a later commit
needs its own run.

<a id="synchronization-record"></a>

## 🔁 Synchronization record

`tools/host/Test-NSFR-Sync.ps1` runs the
[template and synchronization](../INSTALLATION.md#template-and-synchronization)
round trip and writes `tests/evidence/sync/<candidate SHA>.json`, schema
`k-nsfr-sync-evidence/1`. Its `candidate_sha`, `recorded_on`,
`worktree_clean` and `host` follow the harness record, and the gate binds and
rejects them the same way. The other fields:

| Field | Content |
| --- | --- |
| `fingerprints` | The SHA-256 of `templates/NSFR_Template.xlsx`, `src/manifest.txt`, every file the manifest lists and `.gitattributes`, over the bytes stored in Git. `python tools/check_host_evidence.py --candidate <SHA> --inventory --sync` prints the expected ones |
| `checks` | `import_1`, `import_2`, `compile`, `self_test`, `export`, `round_trip`, `layout` and `backup`, each `PASS`, `FAIL` or `NOT_RUN`. The self-test runs exactly when the project compiles |
| `result` | `PASS` only when every check passed; `FAIL` when one failed; otherwise `NOT_RUN` |
| `log` | The reports of the two imports, the self-test and the export, and any difference found, one string per line, with local paths replaced by `<local path>`. A passing self-test must print the output documented in `tests/README.md` at the candidate |
| `reason` | Required unless `PASS` |

The `layout` check passes when the sheets, styles, shared strings and workbook
names of the workbook are those of its first save, and its sheets (names,
order and visibility) and workbook names are the template's. The `backup`
check needs one backup per import. Commit the record as for the harness: in
a PR that adds only the record.
