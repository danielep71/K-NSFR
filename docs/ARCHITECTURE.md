# 🏗️ Architecture and implementation profile

**K-NSFR** · [Overview](../README.md) · [Documentation](../README.md#documentation) · [Governance](GOVERNANCE.md)

[![Stage](https://img.shields.io/badge/Stage-Repository_setup-BF8700?style=flat-square)](../README.md#status)

## 🔎 Evidence and selection

At baseline commit `0708e5b3bdd6564c2c99fd63d1f0500594864583`,
K-NSFR contains a short README, governance documentation, GitHub issue forms,
label/workflow configuration, JavaScript label tooling, and Python issue-triage
tooling and tests. It contains no calculation source, numerical fixtures,
workbook, or host runner. No AGENTS.md or implementation manifest is present.
The setup backlog explicitly provides conditional VBA work in #5, #11, and #13.

Select **Excel/VBA for the reference engine and host**, retaining Python for
portable repository validation and JavaScript for existing label tooling. Excel
provides inspectable tabular inputs and reconciliations for the users in
[SCOPE.md](SCOPE.md); exported text modules make calculations reviewable in Git.
A deterministic core and a narrow host adapter prevent workbook state from
becoming an implicit calculation input.

A Python engine would simplify portable numerical testing but would require
another deployment/interface path for workbook users. A web/service stack would
add hosting, authentication, and operational concerns outside setup scope.
Neither alternative is selected for the initial engine. Python is not a second
regulatory implementation or an automatic test oracle.

## 🧰 Host and tooling profile

| Surface | Declared target | Execution evidence |
| --- | --- | --- |
| Reference engine and workbook adapter | Windows 11; Microsoft 365 desktop Excel; VBA7; 64-bit Office | NOT_RUN; #14 must record exact OS/Excel build and bitness |
| Portable repository checks | Ubuntu 24.04 runner; Python 3.13; Node supplied by the current runner | Existing workflow configuration; results must be attached to each candidate |
| 32-bit Office, perpetual Excel versions, Excel for Mac | Not supported in the initial profile | No compatibility claim |
| Excel web/mobile, LibreOffice, server-side Excel automation | Not supported | Cannot substitute for the selected host |

The host target is a support intention, not a tested compatibility certification.
[Implementation conventions](IMPLEMENTATION.md) specify source encoding,
numeric/error behavior and the detailed supported-versus-tested matrix; #6 owns
runtime consistency for portable tooling.
#13 supplies reproducible packaging and source synchronization, and #14 supplies
genuine host evidence. No worksheet API, Windows API, COM reference, or external
library is needed by the pure calculation core. Any future host dependency must
be justified and isolated.

## 🧱 Dependency and responsibility boundaries

| Layer | Responsibility | Permitted dependencies |
| --- | --- | --- |
| Core | Deterministic classification and calculation from explicit validated inputs and versioned parameters | Plain VBA data and pure helper procedures only |
| API | Validate request shape/profile, call core, return structured results and diagnostics | Core; no workbook or filesystem access |
| Host | Read workbook inputs, invoke API, render results and export evidence | API and approved Excel/host interfaces |
| Tests | Exercise core/API cases and separately exercise host behavior | Tested layer; synthetic fixtures |
| Tools | Repository checks and source import/export orchestration | Files and explicitly approved host automation; never business rules |

Dependency direction is host to API to core. Core and API must not reference
worksheets, ranges, active workbooks, selection, application settings, clocks,
randomness, network, or filesystem. No layer may infer a regulatory version
from the machine date. Configuration is explicit and unchanged during a run.

The API will accept validated records, an as-of date, a reporting perimeter,
and an approved profile/parameter identity; it will return results, diagnostics,
and input/configuration provenance. Exact fields, types, precision, tolerances,
and error codes are deliverables of #5, #9, and #10, not finalized by this issue.
Invalid inputs must produce explicit diagnostics, never fabricated zero results.
The host owns cleanup and restoration of any Excel state it changes, including
on failure; host failures must remain distinguishable from calculation failures.

## 🗂️ Layout and stable names

The following is the intended layout. Create a directory only when it contains a
real deliverable; this issue adds documentation only, with no empty scaffolding.

| Path | Content and ownership |
| --- | --- |
| `src/core/` | Exported pure VBA calculation modules; later production work |
| `src/api/` | Public validation and result boundary; #5 and #9 define contracts |
| `src/host/` | Workbook adapters, isolated from core; #13 |
| `tests/core/`, `tests/api/` | Deterministic neutral harness first (#12), then reviewed numerical cases |
| `tests/host/` | Import, cleanup, and workbook round-trip tests; #13/#14 |
| `tests/fixtures/` | Synthetic input/expected-result records and provenance; #10 |
| `docs/` | Scope, architecture, decisions, governance, and future methodology/evidence guides |
| `tools/` | Existing Python governance tools; future checks and sync utilities |
| `.github/scripts/` | Existing JavaScript label tools |
| `.github/workflows/` | CI and repository automation |
| `templates/` | Only the exact approved macro-free workbook template when delivered by #13 |

Use `NSFR_Core_*`, `NSFR_API_*`, and `NSFR_Host_*` for VBA module names,
with filenames matching the component name and host limits respected.
Test modules use `Test_NSFR_*`; procedure/type names use PascalCase.
Use lowercase snake_case for Python modules and contract field identifiers,
lowercase kebab-case for profile and fixture IDs, and the existing uppercase
convention for top-level documentation files. Stable fixture IDs identify cases
independently of row order. #5 will define the public API naming details before
code is introduced.

Repository text is the future canonical source; generated macro-enabled
workbooks are not source of truth. [Source-control conventions](SOURCE_CONTROL.md)
define encoding, line endings, generated files and exact binary exceptions,
and #13 implements the manifest and synchronization controls. Preserve existing
governance tooling in place; do not move tests merely to match the future layout.

## ✅ Validation boundaries

Portable CI checks label policy, issue triage and committed source conventions. It does not compile
VBA, run Excel, or validate regulatory results. Future static checks (#11),
neutral harness tests (#12), and actual host tests (#14) provide different
evidence and must report their own status. Independent numerical validation is
later production work, gated by #8-#10 and the [backlog](BACKLOG.md).
