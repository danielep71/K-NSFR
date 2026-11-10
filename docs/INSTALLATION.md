<a id="installation-and-host-setup"></a>

# 💻 Installation and host setup

**K-NSFR** · [Overview](../README.md) · [Documentation](../README.md#documentation) · [Governance](GOVERNANCE.md)

[![Stage](https://img.shields.io/badge/Stage-Repository_setup-BF8700?style=flat-square)](../README.md#status)

This guide sets up the two environments of the repository: the portable
checks and the supported Excel host. Tracking: #18.

> [!IMPORTANT]
> There is no NSFR engine to install yet. On the host you can install and run
> the [neutral test harness](../tests/README.md), and build a workbook from the
> macro-free template with its VBA synchronized from `src/`
> ([template and synchronization](#template-and-synchronization)).

<a id="environments"></a>

## 🧭 Environments

| Environment | What runs there | Status |
| --- | --- | --- |
| Windows 11, Microsoft 365 desktop Excel, 64-bit | The neutral harness: import, compilation, self-test, cleanup | Supported and tested: three PASS [host evidence records](../tests/evidence/README.md) |
| Linux x86_64 with Python 3.13, Node 22 and Git | The portable checks, `python tools/check.py --ci` | Reference environment of the required Repository integrity check |
| Windows or macOS with Python and Git | The portable checks with compatible tools | Not the reference: the wheel lock targets Linux x86_64. NOT RUN here |
| 32-bit Office, older or perpetual Excel, Excel for Mac, Excel on the web, LibreOffice | Nothing | Unsupported; NOT RUN |

The [support matrix](IMPLEMENTATION.md#supported-and-tested-environments)
states what each environment is declared for. Passing portable checks never
shows Excel compatibility or regulatory correctness.

<a id="get-the-source"></a>

## 📥 Get the source

Clone the repository with Git, not as a source archive: the checks and the
host scripts read the Git history. Work on the active release branch:

```sh
git clone https://github.com/danielep71/K-NSFR.git
cd K-NSFR
git checkout release/0.2.0
```

Keep the clone clean. A host run refuses a worktree with changes, because its
record must describe an exact commit.

<a id="portable-checks"></a>

## 🧰 Portable checks

You need Python 3.13, Node 22 and Git. [Running the checks](../tools/README.md#running-the-checks)
gives the exact steps:

1. Create a virtual environment outside the checkout and install the
   hash-locked tools from `tools/requirements-quality-ci.txt`.
2. Install actionlint 1.7.12 and verify its checksum.
3. Commit your change, then run `python tools/check.py --ci`.

The run prints one line per gate and ends with `repository-integrity: PASS`
when every gate passes. Without a tool, `python tools/check.py --allow-missing`
reports **INCOMPLETE**, which is never merge evidence.

<a id="host-prerequisites"></a>

## 🖥️ Host prerequisites

- Windows 11 and Microsoft 365 desktop Excel, 64-bit, installed as
  Click-to-Run. **File → Account → About Excel** shows the version, build and
  bitness.
- Git for Windows, so that the clone and the scripts can read the history.
- Windows PowerShell 5.1, which Windows 11 includes.
- Python 3 on `PATH`: optional. Without it, the record is validated later by
  the `host-evidence` gate of its pull request.
- For the scripted routes only: Excel's *Trust access to the VBA project
  object model* (**File → Options → Trust Center → Trust Center Settings →
  Macro Settings**). The scripts import modules through COM, which needs it.
  They never change it, nor any other Trust Center or macro setting. Turn it
  on yourself if you accept it, and off again after the run if you prefer.
  The manual route of the harness does not need it; the source
  synchronization below does, while it runs.

No add-in, COM reference or library is installed. The harness routes use a
new blank workbook and close it without saving; only the
[template route](#template-and-synchronization) saves a workbook, outside the
clone.

<a id="install-and-run-the-harness"></a>

## ▶️ Install and run the harness

Three routes build the same thing: the two modules of `tests/core/` in a new
blank workbook, compiled, with the self-test run and the workbook closed.

| Route | When to use it | Command or steps |
| --- | --- | --- |
| Branch verification | To verify the head of a branch in one command, without touching your clone | `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\host\Verify-NSFR-Branch.ps1 -Branch release/0.2.0` |
| Record the checked-out commit | To record a run of the commit already checked out in a clean clone | `powershell -NoProfile -ExecutionPolicy Bypass -File tools\host\Record-HostEvidence.ps1` |
| Manual | When you do not accept the trust setting, or a script fails | [Running in Excel](../tests/README.md#running-in-excel) |

Close Excel first and run the commands from the clone's folder.
[Recording a run](validation/HOST_EVIDENCE.md#recording-a-run) describes what
each script does. The branch verification writes its bundle beside the
clone, in `..\nsfr-verify-<first 7 characters of the SHA>`: put the clone in
the folder where you want the bundle.

A passing run ends with the self-test output `RESULT PASS` and, for the
scripts, `Result: PASS`. Propose the record `tests/evidence/<SHA>.json` in a
pull request: a run proves one commit only.

<a id="template-and-synchronization"></a>

## 🔁 Template and synchronization

`templates/NSFR_Template.xlsx` is the macro-free layout of the workbook (#13).
The VBA module `NSFR_Host_Sync` (`src/host/`) imports and exports exactly the
components that `src/manifest.txt` lists, and nothing else: no sheet, cell,
format, workbook name or unlisted component changes. Build a workbook from the
template outside the repository:

1. Copy `templates/NSFR_Template.xlsx` to a folder outside the clone, open it
   and save it as an Excel Macro-Enabled Workbook (`.xlsm`).
2. Open the Visual Basic Editor (Alt+F11) and **File → Import File…**
   `src/host/NSFR_Host_Sync.bas`: the sync cannot import itself.
3. With the trust setting on, run in the Immediate window (Ctrl+G), with your
   clone's folder:

   ```text
   ?NSFR_Host_Sync.SyncImport("C:\path\to\K-NSFR")
   ```

   It saves a backup copy beside the workbook, imports every other listed
   component, and prints one line per component and `RESULT PASS`.
4. **Debug → Compile VBAProject**, and save.

After editing a listed module in the Visual Basic Editor, write the changes
back with `?NSFR_Host_Sync.SyncExport("C:\path\to\K-NSFR")`, then review
them with `git diff`. A new module is synchronized only once it is listed in
`src/manifest.txt`. Never commit the `.xlsm` or its backups: `.gitignore`
excludes them.

`tools/host/Test-NSFR-Sync.ps1` runs this round trip in one command, from a
clean clone with Excel closed:

```text
powershell -NoProfile -ExecutionPolicy Bypass -File tools\host\Test-NSFR-Sync.ps1
```

It builds the workbook from the template in a temporary folder, imports twice,
compiles, runs the harness self-test, exports into the clone, and checks that
the export changed no byte Git stores, that the sheets, styles, strings and
workbook names are the template's, and that the backups exist. It restores any
file the export changed and prints a summary.

**Recovery:**

| Situation | What to do |
| --- | --- |
| `RESULT FAIL at manifest` | Pass the clone's folder; the manifest names each file as `name path`, and the file must exist |
| `RESULT FAIL at trust access` | Turn on the trust setting for the run, then off again if you prefer |
| `RESULT FAIL at backup` | Save the workbook as `.xlsm` first |
| `RESULT FAIL at import <name>` | The workbook is left half imported: close it without saving and reopen its backup, `<workbook>.backup-<time>.xlsm` |
| A sheet or workbook module has a listed name | The sync refuses it: rename the module or the manifest entry |
| An export wrote something unexpected | `git diff` shows it; `git checkout -- <file>` restores the committed source |

<a id="troubleshooting"></a>

## 🩺 Troubleshooting

| Message or symptom | Cause | What to do |
| --- | --- | --- |
| `Excel does not allow access to the VBA project` | The trust setting is off | Turn it on as above, or use the manual route. No record was written |
| `The worktree is not clean` | Changed or untracked files in the clone | Commit, stash or discard them, or clone fresh |
| `… is not on origin/main or an origin/release/* branch` (a warning) | The checked-out commit is not published | The record is still written, but the `host-evidence` gate rejects it: run `git fetch`, then record a commit of those branches |
| `No Click-to-Run Office found` | Excel is not a Microsoft 365 Click-to-Run installation | Use the supported host |
| `Excel could not be started through COM` | Excel is missing, blocked or still closing | Close every Excel window, wait, and run again |
| A compile error dialog stays open | The modules do not compile | Close the dialog: the step is recorded as `FAIL` with the reason |
| `… already exists. Remove it, or run again with -Force` | The bundle of the same commit exists | Delete the folder or add `-Force` |
| `… has no tools\host\Record-HostEvidence.ps1` | The branch predates the host scripts | Verify a later commit |
| `Result: FAIL` or `NOT_RUN` | A step failed or did not run | Read `reason` in the record and `run.log` in the bundle |
| `Python not found` | Python is not on `PATH` | Nothing: the pull request validates the record |
| The script does not start | The execution policy blocks scripts | Run it as shown, with `-ExecutionPolicy Bypass` on the command line; this changes no machine setting |
| `Get-FileHash` differs from a fingerprint | The Windows checkout holds CR LF, Git stores LF | Expected: compare with `python tools/check_host_evidence.py --candidate <SHA> --inventory` |

If a script stops in a way this table does not cover, run the manual route.
Record the failure as a `FAIL` or `NOT_RUN` record with its reason, as
[host evidence](validation/HOST_EVIDENCE.md#recording-a-run) describes.

<a id="executed-outcomes"></a>

## ✅ Executed outcomes

These instructions were followed from fresh clones, with these outcomes:

| Date | Route | Clean checkout | Outcome |
| --- | --- | --- | --- |
| 2026-11-09 | Record the checked-out commit, on the supported host | Fresh clone of `release/0.1.0` at `5231985` | PASS ([record](../tests/evidence/523198552eee8296b1dfa4e72344c2a8e5d55cfd.json), #175) |
| 2026-11-09 | Branch verification `-Branch release/0.1.0`, on the supported host | Fresh clone, verified at `dec4ad5` in its own worktree | PASS, recorder exit code 0 ([record](../tests/evidence/dec4ad555db1240f8c2f0361c5de47e36f059528.json), #178) |
| 2026-11-10 | Branch verification `-Branch release/0.1.0`, on the supported host | Fresh clone, verified at `87d4618`, the v0.1.0 candidate, in its own worktree | PASS ([record](../tests/evidence/87d461818831636be270c42db8d45c835c9d6e6b.json), #186) |
| 2026-11-10 | Portable checks, Linux x86_64 with Python 3.13.17, Node 22.22.0 and Git 2.43.0 | Fresh clone of `release/0.1.0` at `cd2860f` | All 17 gates PASS. The locked tools were installed with `uv pip install --require-hashes` into a virtual environment, from the same lock file. Not the reference environment: CI pins Python 3.13.18 and Node 22.23.4 and installs with `pip --only-binary=:all:` |
| 2026-11-10 | Template and synchronization, `Test-NSFR-Sync.ps1`, on the supported host | Clean clone of `feat/13-host-sync` at `cadf916` | PASS on all eight checks: two imports with a backup each, compilation, the harness self-test (`RESULT PASS`), the export, an unchanged worktree after the export, the template's sheets, styles, strings and names in the workbook, and the backups (#13) |

Not executed:

- the manual routes of the harness and of the synchronization, step by step;
- the scripts' `-Hidden`, `-Force` and `-Keep` options;
- the portable checks on Windows or macOS;
- any host other than the one above.
