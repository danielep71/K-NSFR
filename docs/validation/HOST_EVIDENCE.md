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
in [`tests/evidence/`](../../tests/evidence) against the Git history. The
source synchronization round trip and the host template belong to #13 in
v0.2.0. Until then, the [manual route](../../tests/README.md#running-in-excel)
is the only way to build the host project.

<a id="record"></a>

## 🧾 Record

One JSON file per candidate, named `<candidate SHA>.json`, with exactly these
fields:

| Field | Content |
| --- | --- |
| `schema` | `k-nsfr-host-evidence/1` |
| `candidate_sha` | The full 40-character SHA of the tested commit. It must be a commit of `release/0.1.0` or `main`: a task-branch commit disappears with its squash merge |
| `recorded_on` | The ISO date of the run, from the commit date to today |
| `worktree_clean` | `true`: `git status --porcelain` printed nothing before the import |
| `host.os` | The Windows edition, version and OS build, starting with `Windows 11`, for example `Windows 11 Pro 24H2, OS build 26100.2314` (**Settings → System → About**) |
| `host.excel` | The Microsoft 365 product line with `Version … (Build …)`, as **File → Account → About Excel** shows it |
| `host.bitness` | `64-bit` |
| `fingerprints` | The SHA-256 of every module imported from `tests/core/` and of `.gitattributes`, as checked out (VBA source with CRLF) |
| `steps` | `import`, `compile`, `run` and `cleanup`, each `PASS`, `FAIL` or `NOT_RUN` |
| `result` | `PASS`, `FAIL` or `NOT_RUN`, as the steps give it |
| `output` | The complete Immediate-window output of the run, one string per line, without the trailing empty line; absent when the run did not happen |
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
    "excel": "Microsoft 365 Apps for enterprise, Version 2410 (Build 18129.20158 Click-to-Run)",
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

1. Clone the repository fresh, or check that `git status --porcelain` prints
   nothing, and check out the commit to test. Note `git rev-parse HEAD`.
2. In PowerShell, from the repository folder, take the fingerprints:
   `Get-FileHash -Algorithm SHA256 .gitattributes, tests\core\*.bas`. The
   values are case-insensitive.
3. Follow the [manual route](../../tests/README.md#running-in-excel): new
   blank workbook, import both modules, compile, run
   `?Test_NSFR_HarnessSelfTest.RunSelfTest()` and close the workbook without
   saving. Each of these is a step of the record.
4. Copy the output and the host details into the record, save it as
   `tests/evidence/<SHA>.json` and propose it in a pull request. Run
   `python tools/check_host_evidence.py --candidate <SHA>` first if you can.

A failed or interrupted run is recorded too, with `FAIL` or `NOT_RUN` and the
reason. If no supported host is available, #14 stays open with that blocker.

<a id="what-the-gate-rejects"></a>

## 🚦 What the gate rejects

- **Unbound evidence:**
  - an unknown candidate, or one that is not an ancestor of the checked commit;
  - a file name that differs from the SHA;
  - a dirty worktree;
  - a fingerprint that differs from the files Git checks out at the candidate, or a missing or extra one.
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
  - any field outside the record. Do not add machine names, user names or licence details.

With `--candidate <SHA>`, the checker also fails when that commit has no
record (*missing*), when only other commits have one (*stale*), or when its
record is not `PASS`. A record proves one commit only: a later commit needs
its own run.
