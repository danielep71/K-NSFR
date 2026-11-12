<a id="implementation-host-and-api-conventions"></a>

# ✍️ Implementation, host and API conventions

**K-NSFR** · [Overview](../README.md) · [Documentation](../README.md#documentation) · [Governance](GOVERNANCE.md)

[![Stage](https://img.shields.io/badge/Stage-Prototype_setup-BF8700?style=flat-square)](../README.md#status)

This is the implementation contract for the planned K-NSFR Excel/VBA reference
engine. The repository is in setup: there is no numerical engine, engine
workbook or validated regulatory profile; only a macro-free layout template
exists (#13). [Architecture](ARCHITECTURE.md) selected VBA;
Python and JavaScript remain repository tooling, not parallel NSFR engines.
These conventions constrain later code without asserting a regulatory formula.

<a id="public-boundary-and-deterministic-core"></a>

## 🧩 Public boundary and deterministic core

The only planned calculation entry point is an explicitly named
`NSFR_API_Run` procedure. It accepts a typed request containing validated
records, an as-of date, reporting perimeter, and explicit methodology/profile
and parameter identities. It returns a status plus a structured result with
diagnostics and provenance. The [contracts](contracts/CONTRACTS.md) (#9) define
its field names, units, requiredness and serialization. Standard modules start
`NSFR_API_`, `NSFR_Core_` or `NSFR_Host_` by layer; classes, forms, sheet code
names, tests and examples follow the
[VBA house style](VBA_HOUSE_STYLE.md#naming). Exported filenames match
component names; procedures and types use PascalCase. An API change requires
a documented contract version and compatibility decision, not an implicit
worksheet-column change.

The host alone reads/writes workbook ranges, files and Excel state. The API
validates the request, copies configuration and mutable arrays into a run-local
snapshot, calls the core, and returns a complete response. The core consumes
only explicit inputs and the snapshot; it performs no worksheet/COM, file,
clock, locale, randomness, network or global-state access. It never alters the
caller request or profile/parameter data. Results must be reproducible for
identical validated inputs and versioned configuration. No default profile,
effective date, reporting perimeter, currency basis or missing value is
inferred. Unknown or unapproved configuration is rejected. Regulatory
sources belong to the [register](methodology/SOURCE_REGISTER.md) (#8), the
parameter-set format to the [contracts](contracts/CONTRACTS.md) (#9), and
accepted parameter values to PROD-01 of the [backlog](BACKLOG.md).

<a id="numbers-validation-and-diagnostics"></a>

## 🧮 Numbers, validation and diagnostics

Represent quantities and money with an explicit unit and decimal scale in
the #9 contract. Amounts, weighted amounts, totals and the ratio are held in
a `Variant` containing a `Decimal` subtype, built with explicit `CDec`
conversion; `Decimal` is not a directly declared VBA variable type.
`Currency` cannot hold the contract's largest amount, so it is not used for
them (ADR-010, superseding the preference of ADR-007). Do not silently round,
accumulate past the `Decimal` range, or use binary `Double` as an
authoritative money or ratio representation. The
[validation rules](validation/VALIDATION.md#arithmetic-scale-and-rounding)
fix the ratio's scale and rounding, the zero-denominator behavior, the
aggregation order and the comparison; no regulatory percentages or results
are set here. Reject NaN/infinite or
non-numeric host values before constructing a valid typed request.

Return a documented status and diagnostics for invalid input, unknown
profile/version, arithmetic overflow, and host/import failures. Core and API
detect overflow, division by zero and other invalid arithmetic with range
checks before the operation and return that status; they hold no `On Error`
handler, and an unexpected runtime error propagates to the host boundary
([ADR-008](DECISIONS.md#recorded-decisions)). Never turn
an error, blank, or missing value into zero, a plausible NSFR result, or a
partially successful response. The [contracts](contracts/CONTRACTS.md) (#9)
define the diagnostic codes and field paths; use stable codes for tests and descriptive text for people. An error
result contains no authoritative totals. Host errors remain distinguishable
from calculation validation. The host restores any Excel state it changed
even after failure and does not suppress the original diagnostic.

<a id="supported-and-tested-environments"></a>

## 💻 Supported and tested environments

| Environment | Declared support | Execution evidence |
| --- | --- | --- |
| Windows 11, Microsoft 365 desktop Excel, VBA7, 64-bit Office | Initial intended host only | Neutral harness self-test PASS in three records, on 2026-11-09 and 2026-11-10: Windows 11 Enterprise 25H2 (OS build 26200.9711), Microsoft 365 Apps build 20430.20180, 64-bit ([host evidence records](../tests/evidence/README.md)). No engine has run. [Installation](INSTALLATION.md) gives the routes |
| Ubuntu 24.04, Python 3.13 and Node 22 at the exact patch versions of the [tool versions](../tools/README.md#tool-versions) | Portable repository checks | Repository integrity runs on candidate PRs; not an Excel or numerical test |
| 32-bit Office, older/perpetual Excel, Excel for Mac | Outside initial host profile | NOT RUN; no compatibility claim |
| Excel web/mobile, LibreOffice, server Excel automation | Unsupported | NOT RUN |

A declared target does not become tested merely because portable CI passes.
A [host evidence record](validation/HOST_EVIDENCE.md) (#14) binds the exact
candidate commit, the Windows and Excel builds, the bitness, the fingerprints
of the imported files and the observed outcome; follow-up host variants need
their own record. #13 delivered the template and the source-to-workbook synchronization; #195
records their host evidence.

<a id="vba-source-and-environment"></a>

## 📦 VBA source and environment

Every exported module starts with `Option Explicit`; keep its module name and
source filename in sync. Use the repository's [source-control policy](SOURCE_CONTROL.md):
committed text uses LF, Git checks `.bas`, `.cls` and `.frm` out as CRLF.
The VBE exports according to the host system code page, which differs between
hosts, so VBA source is ASCII only and a non-ASCII character is built at run
time with `ChrW` ([ADR-008](DECISIONS.md#recorded-decisions)). No charset or
Git working-tree encoding is configured for VBA. The export/import
round trip passed once on the intended host on 2026-11-10 at `cadf916`, a
task-branch commit of #13: an export changed no byte Git stores. That run has
no host evidence record; #195 records the round trip at a commit of the
release branch. Never normalize an exported module
with a text editor merely to make a test pass. Keep form binaries (`.frx`)
paired with their source and admit each through the exact binary allowlist.

The declared target is VBA7/Win64. Conditional compilation is allowed only
at an isolated host or Windows API boundary, using explicit `VBA7` and
`Win64` guards where declarations differ. Do not imply support for an
untested branch. No Windows API or COM dependency belongs in core/API. Excel
object-model calls belong in host modules, with concrete interfaces and
explicit ownership/cleanup. A new external COM reference, library or macro
trust requirement needs an approved dependency/host decision and evidence;
do not require blanket macro enablement.

Prohibit `On Error`, unstructured jumps, `CallByName`, `Application.Run`,
`Evaluate` and dynamic procedure names in core and API. A narrow
`On Error GoTo` handler is permitted at a host resource/Excel-state boundary
to restore state, emit a diagnostic and exit. Test and example modules may
use `On Error GoTo` to record a failure, and `On Error Resume Next` is
forbidden everywhere ([ADR-008](DECISIONS.md#recorded-decisions)). Do not use
`GoTo` for ordinary branching. Use explicit argument types,
qualified calls, and no implicit coercion from Excel cells. Avoid active
workbook, selection or active sheet as an input; pass the workbook and ranges
explicitly. The `vba-static` gate enforces the feasible source rules
([VBA static checks](../tools/README.md#vba-static-checks)), and the
[neutral harness](../tests/README.md) (#12) runs test suites in the host.

<a id="repository-tooling"></a>

## 🛠️ Repository tooling

Python governance utilities use explicit arguments and typed, deterministic
checks; JavaScript label tools retain their manifest-driven behavior. They
remain separate from regulatory calculations. Python and JavaScript tooling
files use UTF-8/LF; the PowerShell host scripts check out as CRLF
([source control](SOURCE_CONTROL.md)). Tooling follows
the existing naming rules in [ARCHITECTURE.md](ARCHITECTURE.md). VBA-specific
module and conditional-compilation rules are inapplicable to Python/JavaScript,
because neither is selected as a calculation engine. No new runtime or
dependency is introduced by this document.

<a id="numeric-design-rationale-and-admission-tests"></a>

## 📏 Numeric design rationale and admission tests

Decimal arithmetic keeps scale and rounding explicit and avoids accidental
binary rounding. The initial preference for `Currency`, a fixed-point type
with four fractional digits, did not survive the range check: the contract
admits amounts up to 999,999,999,999,999.9999, beyond the `Currency` maximum,
and weighted amounts carry up to eight decimals. ADR-010 therefore selects
`Decimal`, whose Variant subtype needs careful coercion control. This choice
is not admitted to production merely by this document.

The [validation rules](validation/VALIDATION.md) record the input, weighted,
aggregate and ratio magnitudes, the scales, the rounding mode and where it is
applied, the zero-denominator behavior and the comparison. Their
[case plan](validation/VALIDATION.md#case-plan) assigns the tests of range
boundaries, fractional weights, cancellation, aggregation order, zero
denominators and serialization round trips to the engine work. Never cast
silently between `Currency`, `Decimal` and `Double`.

The detailed [VBA house style](VBA_HOUSE_STYLE.md) defines procedure banners,
visibility, argument ownership, array bounds, module layout and cleanup sketches.
Its examples are not executed host evidence.
