# Implementation, host and API conventions

This is the implementation contract for the planned K-NSFR Excel/VBA reference
engine. The repository is in setup: there is no numerical engine, workbook or
validated regulatory profile. [Architecture](ARCHITECTURE.md) selected VBA;
Python and JavaScript remain repository tooling, not parallel NSFR engines.
These conventions constrain later code without asserting a regulatory formula.

## Public boundary and deterministic core

The only planned calculation entry point is an explicitly named
`NSFR_API_Run` procedure. It accepts a typed request containing validated
records, an as-of date, reporting perimeter, and explicit methodology/profile
and parameter identities. It returns a status plus a structured result with
diagnostics and provenance. #9 will finalize field names, units, requiredness,
and serialization before code is added. Public names start `NSFR_API_`;
core and host modules use `NSFR_Core_` and `NSFR_Host_`. Exported filenames
match module names; procedures and types use PascalCase. An API change requires
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
methodology and accepted parameter versions belong to #8 and #9.

## Numbers, validation and diagnostics

Represent quantities and money with an explicit unit and decimal scale in
the #9 contract. For amounts whose accepted scale is at most four fractional
digits and whose full aggregation range fits, use VBA `Currency`; it is a
fixed-point type with four decimal places. For a calculation that requires
another scale or larger range, define a reviewed alternative before coding:
a `Variant` containing a `Decimal` subtype may be used with explicit `CDec`
conversion, but `Decimal` is not a directly declared VBA variable type.
Do not silently round into `Currency`, accumulate past its range, or use
binary `Double` as an authoritative money or ratio representation. The
ratio's exact scale, rounding mode, zero-denominator behavior, aggregation
order and tolerance are contract/validation decisions in #9/#10; no
regulatory percentages or results are set here. Reject NaN/infinite or
non-numeric host values before constructing a valid typed request.

Return a documented status and diagnostics for invalid input, unknown
profile/version, arithmetic overflow, and host/import failures. Never turn
an error, blank, or missing value into zero, a plausible NSFR result, or a
partially successful response. Define diagnostic codes and field paths in
#9; use stable codes for tests and descriptive text for people. An error
result contains no authoritative totals. Host errors remain distinguishable
from calculation validation. The host restores any Excel state it changed
even after failure and does not suppress the original diagnostic.

## Supported and tested environments

| Environment | Declared support | Execution evidence at v0.1.0 setup |
| --- | --- | --- |
| Windows 11, Microsoft 365 desktop Excel, VBA7, 64-bit Office | Initial intended host only | NOT_RUN; #14 must record OS/Excel build, bitness, import, execution and cleanup |
| Ubuntu 24.04, Python 3.13, Node supplied by Actions runner | Portable repository checks | Repository integrity runs on candidate PRs; not an Excel or numerical test |
| 32-bit Office, older/perpetual Excel, Excel for Mac | Outside initial host profile | NOT_RUN; no compatibility claim |
| Excel web/mobile, LibreOffice, server Excel automation | Unsupported | NOT_RUN |

A declared target does not become tested merely because portable CI passes.
Host evidence in #14 binds the exact candidate commit, builds, bitness, source
import/export and observed outcome; follow-up host variants need their own
record. #13 owns packaging and a reproducible source-to-workbook process.

## VBA source and environment

Every exported module starts with `Option Explicit`; keep its module name and
source filename in sync. Use the repository's [source-control policy](SOURCE_CONTROL.md):
committed text uses LF, Git checks `.bas`, `.cls` and `.frm` out as CRLF.
The VBE exports according to the host system code page. No charset or Git
working-tree encoding is imposed on VBA. Until #13/#14 verifies an exact host
code page and export/import round trip, keep source identifiers and literals
within the portable ASCII repertoire; review any non-ASCII source separately
with its byte-level round-trip evidence. Never normalize an exported module
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

Prohibit unstructured jumps, `On Error Resume Next` across logic,
`CallByName`, `Application.Run`, `Evaluate` and dynamic procedure names
in core and API. A narrow `On Error GoTo` handler is permitted at a host
resource/Excel-state boundary to restore state, emit a diagnostic and exit;
do not use `GoTo` for ordinary branching. Use explicit argument types,
qualified calls, and no implicit coercion from Excel cells. Avoid active
workbook, selection or active sheet as an input; pass the workbook and ranges
explicitly. Static checks in #11 will enforce feasible source rules once
modules exist, and a neutral harness in #12 will test the interfaces.

## Repository tooling

Python governance utilities use explicit arguments and typed, deterministic
checks; JavaScript label tools retain their manifest-driven behavior. They
remain separate from regulatory calculations. Tooling files use UTF-8/LF and
the existing naming rules in [ARCHITECTURE.md](ARCHITECTURE.md). VBA-specific
module and conditional-compilation rules are inapplicable to Python/JavaScript,
because neither is selected as a calculation engine. No new runtime or
dependency is introduced by this document.
