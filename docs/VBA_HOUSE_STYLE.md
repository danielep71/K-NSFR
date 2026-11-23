<a id="vba-house-style"></a>

# ✍️ VBA house style

**K-NSFR** · [Overview](../README.md) · [Documentation](../README.md#documentation) · [Governance](GOVERNANCE.md)

[![Stage](https://img.shields.io/badge/Stage-Prototype_setup-BF8700?style=flat-square)](../README.md#status)

These rules govern every exported VBA component: standard modules (`.bas`),
class modules (`.cls`) and the code part of forms (`.frm`). I adopted them on
2026-11-23 ([ADR-012](DECISIONS.md#recorded-decisions), #222); where they
differ from an earlier rule, they prevail. The admitted modules are the
[neutral test harness](../tests/README.md) and its self-test (#12) and the
source synchronization `NSFR_Host_Sync` (#13). The seven modules of the
non-admitted `prototype/` (ADR-011) follow the same layout (#223) but are not
checked by the gate until they are admitted.

The `vba-static` gate enforces the mechanical rules
([VBA static checks](../tools/README.md#vba-static-checks)). Portable CI reads
VBA source text only; it never compiles or runs it. The public boundary and
numeric decisions remain in [IMPLEMENTATION.md](IMPLEMENTATION.md).

<a id="naming"></a>

## 🏷️ Naming

The VBE lists every component in one flat project, so a component's name
carries its role in a prefix. This table is the naming rule for every VBA
component; the other documents refer to it.

| Element | Convention | Example |
| --- | --- | --- |
| Core module (`src/core/`) | `NSFR_Core_` + subject | `NSFR_Core_Validation` |
| Public facade module (`src/api/`) | `NSFR_API_` + subject | `NSFR_API_Functions` |
| Class (`src/api/`) | `C` + noun | `CFundingRecord` |
| Workbook module (`src/host/`) | `ThisWorkbook`; sheets by code name `sh` + noun; standard modules `NSFR_Host_` + subject | `shInputs`, `NSFR_Host_Main` |
| Form (`src/host/`) | `frm` + noun | `frmRunOptions` |
| Test module (`tests/core/`) | `Test_NSFR_` + subject | `Test_NSFR_Harness` |
| Example module (`examples/modules/`) | `Example` + subject | `ExampleSingleRecord` |
| Procedure in a standard module | `<ModulePrefix>_<Area>_<Verb><Object>`, the module prefix declared in the module header's `NOTES` | `Val_Record_Check` |
| Class member, event handler | PascalCase; event handlers keep the names VBA requires | `Amount`, `UserForm_Initialize` |
| Parameter, local variable | PascalCase, spelled out, no type prefix and no single letter | `Amount`, `RecordIndex` |
| Module-level variable | `m` + PascalCase | `mRunState` |
| Host-only public variable (exception requiring review) | `gNSFR_` + PascalCase | `gNSFR_HostStatus` |
| Constant | `NSFR_` + UPPER_SNAKE_CASE; Win32 constants keep their SDK names | `NSFR_MAX_RECORDS` |
| Error number constant | `NSFR_ERR_` + UPPER_SNAKE_CASE, raised as `vbObjectError +` the constant | `NSFR_ERR_INVALID_INPUT` |
| Enum, type | `NSFR_` + PascalCase; Win32 structures keep their SDK names | `NSFR_ProductType` |

Apart from `Test_NSFR_Harness` and the non-admitted prototype's
`NSFR_API_Functions` and `NSFR_Host_Main`, the examples show the form of a
name only; no product, quantity or parameter is decided by them.

- Formatting never renames a public procedure, a constant, a type or a
  module-level variable. Names given before ADR-012 stay until a reviewed
  change renames them; the calculation entry point keeps the name
  `NSFR_API_Run` that the [contracts](contracts/CONTRACTS.md) (#9) give it.
- Names are English and spelled out. Abbreviations are limited to established
  terms of the documented method, written in full in the procedure's header
  the first time they appear.
- API arguments use the terms of the [source register](methodology/SOURCE_REGISTER.md)
  (#8) and the [contracts](contracts/CONTRACTS.md) (#9).
- VBA names are case-insensitive and the VBE shows one spelling per name
  across the project, so the source files, not the editor, carry the
  convention. The component name (`Attribute VB_Name`) must equal the file
  name without its extension.

<a id="module-rules"></a>

## ⚙️ Module rules

- Every module declares `Option Explicit`.
- Every standard module in `src/core/` and all test folders declares
  `Option Private Module`, so its `Public` procedures are visible inside the
  project only.
- No `Option Base 1`: arrays are declared with explicit bounds
  (`Dim Values(1 To ValueCount)`).
- No `Variant` where a specific type works. Every parameter states `ByVal` or
  `ByRef` and its type, and every function its return type. VBA arrays and
  user-defined types that require `ByRef` are documented in `INPUTS`; copy
  mutable inputs into run-local storage before using them as a snapshot.
- `src/core/` never touches the Excel object model (`Application`,
  workbooks, sheets, ranges) and holds no mutable module-level state: a core
  procedure depends only on its arguments
  ([dependency direction](ARCHITECTURE.md)).
- A `Declare` statement, if ever needed, uses `PtrSafe` and `LongPtr`, because
  64-bit Excel is the
  [intended target](IMPLEMENTATION.md#supported-and-tested-environments).

<a id="module-layout"></a>

## 🧱 Module layout

1. The `Attribute` lines, one blank line, `Option Explicit` (and
   `Option Private Module` or `Option Compare` where used), one blank line.
2. The component header: a lone `'` line, a banner, `' MODULE: <Name>` (or
   `' CLASS: <Name>` for a class; forms and document modules use `MODULE:`),
   a banner, the [nine fields](#documentation-fields), a closing banner.
3. The declarations, each kind under its section banner and in this order,
   empty sections left out: `PRIVATE TYPES`, `WINDOWS API DECLARATIONS`,
   `PRIVATE CONSTANTS`, `PUBLIC CONSTANTS`, `PUBLIC ENUMS`, `PUBLIC TYPES`,
   `PUBLIC STATE`, `PRIVATE STATE` (a class uses `PRIVATE VARIABLES`). A
   module may add a clearly named section of its own.
4. Two blank lines, then the procedures, families separated by a group
   banner: an empty `'` line, a banner, an empty `'` line, the centered
   title, an empty `'` line, a banner, an empty `'` line.

Every banner line is exactly 79 characters: an apostrophe and 78 dashes. A
test-harness module may use `=` instead of `-` for its component header only.
A section banner holds an UPPERCASE title after one space between two banner
lines; a procedure or group title is UPPERCASE and centered in the 78 columns
after the apostrophe. Source decorations are ASCII only.

Declarations under a section banner are indented four spaces, with `As` in
one column and a trailing comment (no space after the apostrophe) in another,
both on multiples of four. Every declaration carries a comment; blank lines
separate logical groups, and a block of error numbers has a comment line
above it.

A generated module is exempt from the layout rules; its header says it is
generated, names the generator and says it must not be edited by hand.

<a id="documentation-fields"></a>

<a id="procedure-contracts"></a>

## 📜 Documentation fields

The component header and every procedure header carry the same nine fields,
in this order. Each is a single-space-indented UPPERCASE label, its text
indented three spaces, with a lone `'` line between fields. No field is
omitted: write `None` instead.

| Field | Explain |
| --- | --- |
| `PURPOSE` | What it does, in one or two sentences |
| `WHY THIS EXISTS` | The problem it solves, or why it is shaped this way |
| `INPUTS` | `None`, or each parameter on its own line and its description below it: meaning, **unit**, **valid domain** and ownership. Module level: `None at module level` |
| `RETURNS` | `Nothing` for a Sub, otherwise the value and its edge cases |
| `BEHAVIOR` | What happens, in order |
| `ERROR POLICY` | What raises, what is contained, what is best-effort |
| `DEPENDENCIES` | Object models, other components and members it calls |
| `NOTES` | Design decisions, caveats, call examples; the methodology entry, by source ID and locator, that a calculation implements; the Excel or module state it reads, changes or restores |
| `UPDATED` | The date of the last substantive change, `YYYY-MM-DD` |

A `Private` procedure whose body has at most 15 code lines may use a compact
header: `PURPOSE` and `UPDATED`, plus any other field that says something
useful, in the same order. `Public`, `Friend` and event procedures always
carry all nine. A header may add its own UPPERCASE sections after the
standard fields and before `UPDATED`, which always comes last.

Headers state facts the code shows. A change keeps every header, `INPUTS`
entry and intent comment true, and sets `UPDATED` of each procedure and
module whose behavior it changes; a formatting-only change keeps the dates.

<a id="units-and-domains"></a>

### Units and domains

- Amount representation, scale, range and rounding follow the numeric decision
  in [IMPLEMENTATION.md](IMPLEMENTATION.md#numbers-validation-and-diagnostics).
  State currency and unit explicitly; never silently scale by thousands or
  millions.
- Currencies are three-letter ISO 4217 codes held as `String`.
- Approved weights and other proportions are decimals: `0.05`, never `5`
  meaning 5%.
- Dates are VBA `Date`. Period representations and units are the explicit
  decisions of the [contracts](contracts/CONTRACTS.md) (#9); never assume a
  year fraction or day-count convention.
- Every numeric input states its valid domain, for example `Amount >= 0` or
  `0 <= Proportion < 1`.
- Maturity, encumbrance, classification and every other regulatory
  convention belong to the methodology approved under #8 and
  [PROD-01](BACKLOG.md#later-calculation-work). A procedure cites the
  convention it implements in `NOTES`; it never chooses one silently.

<a id="procedure-layout"></a>

## 🧩 Procedure layout

The signature has explicit visibility, `ByVal` or `ByRef` and a type on every
parameter, and a return type on every function. A signature that does not fit
on one line puts each parameter on its own line, indented four spaces and
ending in `, _`. A one-line signature is followed by one blank line; then come
a lone `'` line, a banner, the centered title (a short verb phrase, not the
identifier), a banner, the header fields and a closing banner. One blank line
separates the header from the body and the body from `End Sub` or
`End Function`, and one blank line separates a procedure from the next.

The body is a sequence of section banners, in this order:

| Section | Contains |
| --- | --- |
| `DECLARE` | `Const PROC_NAME As String = "<Component>.<Procedure>"`, a blank line, then every `Dim`, aligned and commented |
| `INITIALIZE` | `On Error GoTo ErrorHandler` and initial state |
| Task sections | One per logical step, a short UPPERCASE verb phrase: `VALIDATE INPUTS`, `LOAD SETTINGS`, `RETURN VALUE` |
| `EXIT PROCEDURE` | Cleanup on the success path, then `Exit Sub` or `Exit Function` |
| `ERROR HANDLER` | The `ErrorHandler:` label at column 0, then capture, cleanup and re-raise |

A trivial procedure keeps only the task sections it needs; `DECLARE`,
`INITIALIZE`, `EXIT PROCEDURE` and `ERROR HANDLER` appear only when there is
something to put in them.

Inside a section, every statement or small group of statements follows a
one-line intent comment at four spaces (no space after the apostrophe,
sentence case, no final period), and the code sits at eight spaces; nested
blocks add four spaces per level. Labels are at column 0.

```vb
'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject a negative amount
        If Amount < 0 Then
            Err.Raise vbObjectError + NSFR_ERR_INVALID_INPUT, PROC_NAME, _
                "Amount cannot be negative"
        End If
```

<a id="statements"></a>

### Statements

- Write the VBA library functions with their qualifier: `VBA.Len`,
  `VBA.Trim$`, `VBA.Left$`, `VBA.Mid$`, `VBA.InStr`, `VBA.StrComp`,
  `VBA.CStr`, `VBA.CLng` and the like, preferring the `$` string variants.
  Test emptiness with `VBA.Len(Text) = 0` and write `vbNullString` rather
  than `""`.
- `InStr`, `InStrRev` and `StrComp` state their compare mode
  (`vbBinaryCompare` or `vbTextCompare`).
- `Next` names its counter. One statement per line, no `:` separator; a
  one-line `If ... Then` is allowed for a guard or an `Exit`.
- A continued line breaks after an operator or comma and is indented four
  more than its statement; `Err.Raise` keeps the number and source on the
  first line and the message on the next.
- A `Select Case` has a `Case Else` whenever an unexpected value matters.
- Object variables are released with `Set x = Nothing` on the exit path.
- Aim for 100 characters a line; 120 is the limit. String contents and the
  trailing comment of an aligned declaration do not count.

<a id="errors"></a>

<a id="return-values-and-errors"></a>

## 🚨 Errors

- The API returns a documented status, diagnostics and no authoritative totals
  on failure. Stable codes, field paths and error domains are defined in the
  [contracts](contracts/CONTRACTS.md) (#9).
- Expected invalid inputs are explicit validation outcomes, checked in a
  `VALIDATE INPUTS` section before the operation. Never substitute zeros or
  partially valid results.
- A procedure that can fail unexpectedly holds the error-handling pattern
  ([ADR-012](DECISIONS.md#recorded-decisions)), in every layer:
  `On Error GoTo ErrorHandler` in `INITIALIZE`, the clean exit in
  `EXIT PROCEDURE`, and in `ERROR HANDLER`:

  ```vb
  ErrorHandler:
      'Capture the error number
          ErrorNumber = Err.Number
      'Capture the error description
          ErrorDescription = Err.Description
      'Suppress cleanup errors
          On Error Resume Next
      '...cleanup...
      'Restore normal error handling
          On Error GoTo 0
      'Raise a descriptive error
          Err.Raise ErrorNumber, PROC_NAME, "<What failed>: " & ErrorDescription
  ```

  Capture `Err` before any other call. Raise a custom error as
  `vbObjectError +` its constant, with `PROC_NAME` as the source.
  `On Error Resume Next` appears only inside `ERROR HANDLER`, to make cleanup
  best-effort. A procedure whose contract is to contain failures, such as a
  report that ends in `RESULT FAIL`, returns that result from its handler
  instead of re-raising, and says so in `ERROR POLICY`.
- Adding a handler to an existing procedure changes what its caller sees, so
  it is a behavior change with its own review, never part of a formatting
  change.
- No worksheet-function surface or `CVErr` mapping is approved at this stage.

<a id="public-and-internal-procedures"></a>

### Public and internal procedures

- Procedures are `Private` unless another component calls them.
- A `Public` procedure in `src/core/` is internal: `Option Private Module`
  keeps it off the workbook's external surface.
- The only planned calculation entry point is `NSFR_API_Run` in `src/api/`.
  Its request/result types follow the [contracts](contracts/CONTRACTS.md) (#9)
  and are declared with the first API module.
- Changing arguments, defaults, units, errors or return values requires an
  issue, updated contracts and regression cases, and a changelog entry.

<a id="operation-and-cleanup"></a>

## 🧹 Operation and cleanup

A procedure that changes Excel application state owns restoring it, and only
it.

1. **Capture** the values it will change before changing them
   (`Calculation`, `ScreenUpdating`, `EnableEvents`, `DisplayAlerts`,
   `Cursor`, `StatusBar`).
2. **Change** only what the operation needs.
3. **Restore** on both paths: in `EXIT PROCEDURE` before the clean exit, and
   in `ERROR HANDLER` under `On Error Resume Next` after capturing the error.
   Restore only the settings this procedure changed, to the values it
   captured. A setting the caller changed is the caller's to restore.
4. **Preserve the primary error.** The captured number and description are
   re-raised after cleanup; a failure inside cleanup never replaces them.

The host evidence of the first host module that changes application settings
must exercise success, a primary error with cleanup, and a failed restore.
`NSFR_Host_Sync` (#13) changes no application setting: it changes only the
VBA project, after a backup copy. Core procedures never change Excel state,
so they need none of this.

<a id="export-compatibility"></a>

## 💾 Export compatibility

- **Encoding:** the VBE reads and writes source in the Windows ANSI code
  page, which differs between hosts. Source is therefore **ASCII only**
  ([ADR-008](DECISIONS.md#recorded-decisions)); a non-ASCII character is built
  at run time with `ChrW`. No byte-order mark. No working-tree encoding is
  configured in `.gitattributes`.
- **Line endings and indentation:** four spaces, no tabs, no trailing
  whitespace, CRLF line endings ending the file. The files are stored with LF
  in Git and checked out with CRLF by `.gitattributes`, whatever
  `core.autocrlf` says; the VBE expects CRLF. Never change these rules or the
  export header lines when reformatting.
- **Document modules:** `ThisWorkbook` and the sheet modules are exported as
  `.cls` files in `src/host/`. They are pasted into the workbook, never
  imported, because importing a `.cls` creates a new class module instead of
  replacing the document module. The source synchronization in #13 follows
  this rule.
- **Forms:** a `.frm` and its `.frx` are one component, kept side by side in
  `src/host/`. The `.frx` is binary and never edited as text, and neither is
  the designer block at the top of the `.frm`.
- **Workbooks and macro-enabled files** are generated outputs. The one
  approved workbook is the macro-free template `templates/NSFR_Template.xlsx`
  (#13): its exact path is allowlisted, the `template` gate rebuilds it, and
  `src/manifest.txt` lists the components synchronized into a workbook built
  from it. Do not add broad exceptions.

<a id="review-and-validation"></a>

## 🔎 Review and validation

- A formatting-only change leaves names, signatures, visibility, constants,
  string contents, calculations, tests and error behavior unchanged; a defect
  found on the way goes in a separate change.
- The `vba-static` gate checks the banner width, the component and procedure
  headers, one statement per line, `Next` counters and the `VBA.` qualifier;
  alignment, intent comments and the wording of headers are reviewed by hand.
- `python tools/check.py` is a static check and never compiles or runs VBA: a
  static PASS proves neither that the project compiles nor that a result is
  right ([tools/README.md](../tools/README.md#validation-boundaries)).
- Host execution is evidenced only by a [host evidence record](validation/HOST_EVIDENCE.md)
  of the exact commit: OS/Excel build, bitness, source fingerprints, import,
  compile, execution and cleanup. A reformatted module needs a new run. Never
  promote a sketch or a portable PASS into host evidence.

---

**Style principle:** state the contract in nine fields, say the intent of
each step, restore only what you change, and never lose the original error.
