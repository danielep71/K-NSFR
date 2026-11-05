<a id="decision-log"></a>

# 🧭 Decision log

**K-NSFR** · [Overview](../README.md) · [Documentation](../README.md#documentation) · [Governance](GOVERNANCE.md)

[![Stage](https://img.shields.io/badge/Stage-Repository_setup-BF8700?style=flat-square)](../README.md#status)

<a id="baseline-decision"></a>

## 📌 Baseline decision

Issue: [#2](https://github.com/danielep71/K-NSFR/issues/2).
Decision date: 2026-10-28.
Status: selected implementation baseline for the development branch after
integration of issue #2. These decisions do not certify host execution or
regulatory correctness.

<a id="recorded-decisions"></a>

## 🗂️ Recorded decisions

| ID | Decision and rationale | Consequence / revisit trigger |
| --- | --- | --- |
| ADR-001 | Keep v0.1.0 setup-only; the repository has governance tools but no numerical engine | No regulatory formulas or parameter tables; production capability requires separately reviewed work |
| ADR-002 | Use Excel/VBA for an inspectable workbook reference implementation; retain Python/JavaScript for repository tooling | Windows 11 / Microsoft 365 desktop Excel / VBA7 / 64-bit is the initial intended host; revisit if deployment or scale requirements exceed this profile |
| ADR-003 | Select EU CRR standard NSFR as the first planned profile; Basel is reference-only | No executable profile is currently supported; #8 must establish exact applicability and version before implementation |
| ADR-004 | Separate core, API, and host with one-way dependencies | Determinism and independent testing take precedence over worksheet convenience; no hidden host inputs in calculations |
| ADR-005 | Keep exported source canonical and data synthetic | #4 defines text conventions, #13 packaging; real data and operational deployment require separate approval and controls |
| ADR-006 | Distinguish planned support, portable checks, host execution, and numerical validation | A green repository workflow cannot be cited as proof of Excel or NSFR correctness |
| ADR-007 | Prefer explicit decimal scale for monetary inputs; Currency/Decimal remain provisional pending #9/#10 range, rounding and coercion tests | Do not implement numerical logic before those tests; supersede the preference if domain limits require another representation |
| ADR-008 | VBA source is ASCII only; standard modules carry the `NSFR_Core_`, `NSFR_API_` or `NSFR_Host_` prefix, and classes, forms, sheets, tests and examples follow the [house-style naming table](VBA_HOUSE_STYLE.md#naming); core and API validate before an operation instead of trapping errors, and `On Error` appears only at the host boundary. Decided on 2026-11-05 in #118, after the baseline | One rule each in the house style, implementation conventions, architecture and source-control documents; revisit if host round-trip evidence (#13/#14) or the static checks (#11) require otherwise |
| ADR-009 | Version 1 of `eu-crr-standard` is bound to the CRR consolidated version `02013R0575-20260626` and to a single institution on an individual basis; consolidation, group and scheme treatments and the simplified calculation are outside it. Decided on 2026-11-05 in #8, after the baseline | The [source register](methodology/SOURCE_REGISTER.md) records the sources and their verification; a later legal version or a wider entity scope needs a new register version and a new decision |
| ADR-010 | Results are exact decimal arithmetic: weighted amounts are never rounded, the ASF and RSF totals are reported with eight decimals, and the ratio is reported as a decimal with six decimals, rounded half-even once from the exact quotient; fixtures compare results exactly, and a tolerance is an exception I approve per fixture. The engine holds amounts, weighted amounts, totals and the ratio in VBA `Decimal`, because the contract's largest amount exceeds the `Currency` range; this supersedes the `Currency` preference of ADR-007. Decided on 2026-11-05 in #10, after the baseline | The [validation rules](validation/VALIDATION.md#arithmetic-scale-and-rounding) record the scales, the [numeric range](validation/VALIDATION.md#numeric-range) and the comparison; an input beyond the `Decimal` range is rejected with a diagnostic added by PROD-03 |

<a id="changing-a-decision"></a>

## 🔄 Changing a decision

Alternatives and the selected layout are documented in
[ARCHITECTURE.md](ARCHITECTURE.md). Regulatory versions, interpretation choices,
numeric/error conventions, workbook schema, and production-readiness approval
remain assigned follow-on decisions in [BACKLOG.md](BACKLOG.md). Record a new
dated decision with its rationale and superseded ID when changing this baseline;
do not silently rewrite historical decisions.
