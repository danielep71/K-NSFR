<a id="architecture-and-implementation-profile"></a>

# 🏗️ Architecture and implementation profile

**K-NSFR** · [Overview](../README.md) · [Documentation](../README.md#documentation) · [Governance](GOVERNANCE.md)

[![Stage](https://img.shields.io/badge/Stage-Repository_setup-BF8700?style=flat-square)](../README.md#status)

<a id="evidence-and-selection"></a>

## 🔎 Evidence and selection

At baseline commit `0708e5b3bdd6564c2c99fd63d1f0500594864583`, when this
selection was made, K-NSFR contained a short README, governance documentation,
GitHub issue forms, label/workflow configuration, JavaScript label tooling, and
Python issue-triage tooling and tests. It contained no calculation source,
numerical fixtures, workbook, host runner, AGENTS.md or implementation
manifest. The setup backlog provided conditional VBA work in #5, #11, and #13.

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

<a id="host-and-tooling-profile"></a>

## 🧰 Host and tooling profile

| Surface | Declared target | Execution evidence |
| --- | --- | --- |
| Reference engine and workbook adapter | Windows 11; Microsoft 365 desktop Excel; VBA7; 64-bit Office | Not yet built, so NOT RUN. The neutral harness has run on this host ([host evidence](../tests/evidence/523198552eee8296b1dfa4e72344c2a8e5d55cfd.json)) |
| Portable repository checks | Ubuntu 24.04 runner; Python 3.13; Node 22 explicitly selected | Existing workflow configuration; results must be attached to each candidate |
| 32-bit Office, perpetual Excel versions, Excel for Mac | Not supported in the initial profile | No compatibility claim |
| Excel web/mobile, LibreOffice, server-side Excel automation | Not supported | Cannot substitute for the selected host |

The host target is a support intention, not a tested compatibility certification.
[Implementation conventions](IMPLEMENTATION.md) specify source encoding,
numeric/error behavior and the detailed supported-versus-tested matrix; the
portable tooling's runtime is pinned as the
[tooling guide](../tools/README.md#tool-versions) describes.
#14 delivered candidate-bound host evidence: the neutral harness self-test
passed on the supported host at `5231985` and `dec4ad5`
([records](../tests/evidence/README.md)). #13, in v0.2.0, supplies the
workbook template, packaging and the source synchronization round trip. No worksheet API, Windows API, COM reference, or external
library is needed by the pure calculation core. Any future host dependency must
be justified and isolated.

<a id="dependency-and-responsibility-boundaries"></a>

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
and input/configuration provenance. #5 recorded the
[implementation conventions](IMPLEMENTATION.md); the exact fields, types,
precision, tolerances and error codes are fixed by the
[contracts](contracts/CONTRACTS.md) (#9) and the
[validation rules](validation/VALIDATION.md) (#10).
Invalid inputs must produce explicit diagnostics, never fabricated zero results.
The host owns cleanup and restoration of any Excel state it changes, including
on failure; host failures must remain distinguishable from calculation failures.

<a id="layout-and-stable-names"></a>

## 🗂️ Layout and stable names

The following is the intended layout. Create a directory only when it contains a
real deliverable; the repository holds no empty scaffolding.

| Path | Content and ownership |
| --- | --- |
| `src/core/` | Exported pure VBA calculation modules; later production work |
| `src/api/` | Public validation and result boundary; #9 defines the contracts |
| `src/host/` | Workbook adapters, isolated from core; `NSFR_Host_Sync` synchronizes the components that `src/manifest.txt` lists (#13) |
| `tests/core/`, `tests/api/` | The deterministic [neutral harness](../tests/README.md) and its self-test (#12), then reviewed numerical cases |
| `tests/host/` | Import, cleanup, and workbook round-trip tests; #13 (v0.2.0) |
| `tests/fixtures/` | Registered numerical fixtures: synthetic inputs, parameters, independently derived expected results and provenance; #10 |
| `tests/evidence/` | Host evidence records, one per tested commit; #14 |
| `examples/modules/` | Example modules that call the API with synthetic data; only once the API exists |
| `docs/` | Scope, architecture, decisions, governance, installation, release, and the source register and validation guides |
| `docs/contracts/` | Input, parameter and result contracts with their machine-readable schema and synthetic examples; #9 |
| `docs/validation/` | Numerical fixture rules and their machine-readable schema (#10); the host evidence record format (#14) |
| `tools/` | Python repository checks and their tests; `tools/host/` PowerShell scripts that record host evidence; future sync utilities (#13) |
| `prototype/` | My benchmark prototype, stored as delivered and not admitted ([ADR-011](DECISIONS.md#recorded-decisions)); v0.2.0 moves it into this layout piece by piece |
| `.github/scripts/` | Existing JavaScript label tools |
| `.github/workflows/` | CI and repository automation |
| `templates/` | The macro-free workbook template, `NSFR_Template.xlsx`, generated by `tools/build_template.py` and checked byte for byte; #13 |

Standard VBA modules use `NSFR_Core_*`, `NSFR_API_*` or `NSFR_Host_*` by
layer; classes, forms, sheet code names, test and example modules follow the
[VBA house style](VBA_HOUSE_STYLE.md#naming), which holds the complete naming
table ([ADR-008](DECISIONS.md#recorded-decisions)). Filenames match the
component name and respect host limits; procedure/type names use PascalCase.
Use lowercase snake_case for Python modules and contract field identifiers,
lowercase kebab-case for profile and fixture IDs, and the existing uppercase
convention for top-level documentation files. Stable fixture IDs identify cases
independently of row order. The contracts (#9) fix the field names; the VBA
request and result types follow them when code is introduced.

Repository text is the future canonical source; generated macro-enabled
workbooks are not source of truth. [Source-control conventions](SOURCE_CONTROL.md)
define encoding, line endings, generated files and exact binary exceptions,
and #13 implements the manifest and synchronization controls. Preserve existing
governance tooling in place; do not move tests merely to match the future layout.

<a id="validation-boundaries"></a>

## ✅ Validation boundaries

Portable CI checks the labels and issue forms, committed bytes, Markdown and
YAML, the contracts and their examples, the numerical fixtures, the VBA static
rules, the host evidence records and the workflows
([gates](../tools/README.md#gates-and-reports)). It does not compile
VBA, run Excel, or validate regulatory results. The VBA static checks (#11),
neutral harness tests (#12), and actual host tests (#14) provide different
evidence and must report their own status. Independent numerical validation is
later production work, gated by #8-#10 and the [backlog](BACKLOG.md).
