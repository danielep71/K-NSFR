<a id="neutral-test-harness"></a>

# 🧪 Neutral test harness

**K-NSFR** · [Overview](../README.md) · [Documentation](../README.md#documentation) · [Governance](../docs/GOVERNANCE.md)

[![Stage](https://img.shields.io/badge/Stage-Prototype_setup-BF8700?style=flat-square)](../README.md#status)

The harness records the outcome of each test case and prints a deterministic
summary. It is neutral: it implements no NSFR calculation and its self-test
asserts nothing about one. Tracking: #12.

> [!NOTE]
> The harness was imported, compiled and run in Excel, with the self-test
> result `RESULT PASS`, first on 2026-11-09 at `5231985` and last on
> 2026-11-10 at `87d4618`, the v0.1.0 candidate
> ([host evidence records](evidence/README.md); #12, #14). The portable checks
> read its source only (`vba-static`). A run on one commit is no evidence for
> another, so a change to the harness or to `.gitattributes` needs a new run.

<a id="layout"></a>

## 🗂️ Layout

| Path | Content |
| --- | --- |
| `core/Test_NSFR_Harness.bas` | The harness: run state, outcomes, summary |
| `core/Test_NSFR_HarnessSelfTest.bas` | Its self-test, with a suite that passes and suites that fail on purpose |
| `fixtures/` | The [numerical fixtures](../docs/validation/VALIDATION.md); the harness will execute them once an engine exists |
| `evidence/` | The [host evidence records](../docs/validation/HOST_EVIDENCE.md), one per tested commit |

Both modules are standard modules with `Option Private Module`. They use VBA
only: no Excel object model, workbook, sheet, file, clock or random number.

<a id="writing-a-suite"></a>

## ✍️ Writing a suite

A suite declares every case first, records one outcome per case, and reads
the summary:

```vb
Dim Suite As NsfrTestRun
BeginRun Suite, "EXAMPLE", "FX-0001,FX-0002"
CheckExact Suite, "FX-0001", "6.766915", ActualRatio
CheckDiagnostic Suite, "FX-0002", "NSFR-W001", ActualCodes
Debug.Print Summary(Suite)
```

The snippet shows the call form only; `ActualRatio` and `ActualCodes` stand
for what a subject under test would return.

| Procedure | Outcome |
| --- | --- |
| `BeginRun` | Resets the run and declares its case IDs: 1 to 32 upper-case letters, digits and hyphens, comma-separated |
| `CheckExact` | `PASS` when the two texts are equal character for character; numbers are compared as text at their reported scale |
| `CheckText` | The same for multi-line text, naming the first line that differs |
| `CheckTolerance` | `PASS` within an approved tolerance: `absolute`, or `relative` below 1 with a floor near zero; values are decimal text with a point, parsed without the system locale. Values whose difference or relative allowance `Decimal` cannot hold exactly fail, so the comparison is never made on a rounded value |
| `CheckDiagnostic` | Expected-error outcome: `PASS` when the subject reported exactly the expected diagnostics; reporting none fails |
| `SkipCase` | `SKIPPED`, with a reason |
| `RecordCleanup` | A cleanup failure for a case, kept apart from its outcome; empty text means cleanup succeeded |
| `Summary`, `RunPassed` | The summary text, and whether the run passed |

Tolerances follow the [comparison rules](../docs/validation/VALIDATION.md#comparison):
exact by default, a tolerance only where a fixture records an approved one.

<a id="what-fails-a-run"></a>

## 🚦 What fails a run

A run passes only when every declared case passed. It fails when:

- a case fails, is skipped or is missing (declared but never recorded);
- a cleanup fails, even if every case passed;
- a run-level problem arises: no case declared, an invalid or repeated
  declaration, an outcome recorded twice, or an outcome for an undeclared
  case.

The summary has one line per declared case, in ID order, then the cleanup
failures and the problems, each in text order. Then come the counts
(`declared`, `executed`, `passed`, `failed`, `skipped`, `missing`,
`undeclared`, `cleanup_failed`, `problems`) and `RESULT PASS` or
`RESULT FAIL`. Lines end with CR LF. The same outcomes give the same text,
whatever the order they were recorded in.

<a id="self-test"></a>

## ✅ Self-test

`RunSelfTest` builds three suites and checks each against its exact
expected summary:

- `DEMO-PASS`: an exact, an absolute-tolerance, a relative-tolerance near
  zero and an expected-diagnostic case, all passing, with a successful
  cleanup;
- `DEMO-FAIL`: fails on purpose in every reported way. It has a wrong value,
  a value outside its tolerance, a missing diagnostic and a different one, a
  skipped and a missing case, a cleanup failure, a case recorded twice, an
  undeclared case, a value that is not a decimal, three invalid
  tolerances and a relative allowance that `Decimal` cannot compute exactly. It is built twice, in opposite orders, to show that the
  summary does not depend on order;
- `DEMO-DECLARE`: a repeated, a lower-case and an empty declared ID.

When the harness behaves as documented, the self-test prints:

```text
SUITE HARNESS-SELF-TEST
H-001 PASS
H-002 PASS
H-003 PASS
H-004 PASS
H-005 PASS
H-006 PASS
H-007 PASS
COUNTS declared=7 executed=7 passed=7 failed=0 skipped=0 missing=0 undeclared=0 cleanup_failed=0 problems=0
RESULT PASS
```

Any other output is a failure: a `FAIL` line names the case and the first
line of its summary that differs.

<a id="running-in-excel"></a>

## ▶️ Running in Excel

You need the [intended host](../docs/IMPLEMENTATION.md#supported-and-tested-environments):
Windows 11, Microsoft 365 desktop Excel, 64-bit.
`tools/host/Record-HostEvidence.ps1` runs the steps below through COM and
writes the [host evidence record](../docs/validation/HOST_EVIDENCE.md#recording-a-run);
`tools/host/Verify-NSFR-Branch.ps1 -Branch <branch>` does the same for the
head of a branch, in a separate worktree, in one command.
The manual route, with a new blank workbook, is below; the
[template and synchronization](../docs/INSTALLATION.md#template-and-synchronization)
route builds the workbook from the template instead:

1. Check out the commit to test and note its full SHA.
2. Open a new blank workbook and the Visual Basic Editor (Alt+F11).
3. **File → Import File…** `tests/core/Test_NSFR_Harness.bas`, then
   `tests/core/Test_NSFR_HarnessSelfTest.bas`.
4. **Debug → Compile VBAProject**; it must finish without a message.
5. In the Immediate window (Ctrl+G) run
   `?Test_NSFR_HarnessSelfTest.RunSelfTest()`.
6. Close the workbook without saving it; the run leaves nothing behind.

Record the run as a [host evidence record](../docs/validation/HOST_EVIDENCE.md#recording-a-run)
in [`evidence/`](evidence/README.md): the commit SHA and the file fingerprints,
the Windows version, the Excel version and build with its bitness
(**File → Account → About Excel**), the outcome of the import, the
compilation, the run and the cleanup, and the complete output of step 5. The
`host-evidence` gate rejects a record that is not bound to its commit. A run
on one commit is no evidence for another.

<a id="reproducibility"></a>

## 🔁 Reproducibility

The summary depends only on the recorded outcomes. It uses no time, random
number, workbook state or system setting: numbers are parsed without the
system locale, booleans are written as `yes` and `no`, and text is compared
in binary order. The same commit therefore gives the same output on every
run of the supported host.
