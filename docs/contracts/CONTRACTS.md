<a id="nsfr-input-parameter-and-result-contracts"></a>

# 📐 NSFR input, parameter and result contracts

**K-NSFR** · [Overview](../../README.md) · [Documentation](../../README.md#documentation) · [Governance](../GOVERNANCE.md)

[![Stage](https://img.shields.io/badge/Stage-Prototype_setup-BF8700?style=flat-square)](../../README.md#status)

Contract version **1.0.0** for profile `eu-crr-standard` version 1 and
[source register](../methodology/SOURCE_REGISTER.md) version 1. Tracking: #9.

This document defines what the planned engine accepts and returns. The
machine-readable form is [nsfr-contract-v1.json](nsfr-contract-v1.json);
`tools/check_contracts.py` (gate `contracts`) checks that every field and
diagnostic below matches it and that the [examples](#synthetic-examples)
give their expected results. Contract validation checks the form of an
input. It never proves that a classification or a result is regulatorily
correct; that belongs to the parameter set ([PROD-01](../BACKLOG.md#later-calculation-work))
and to independent validation (#10).

<a id="scope-of-version-1"></a>

## 📌 Scope of version 1

- One institution on an individual basis (ADR-009), one reporting currency,
  one as-of date per run.
- One record per position: an own-funds item, a liability, an asset or an
  off-balance-sheet item. A position is never split or netted by the
  contract.
- Not supported, and rejected with `NSFR-E012`: the consolidated basis,
  derivative contracts (S-01 Article 428d) and netted transactions
  (S-01 Article 428e). Each needs its own decision and issue before admission.
- No value is inferred: a missing value is an error, never a zero or a
  default.

<a id="input-tables"></a>

## 📥 Input tables

An input is two UTF-8 CSV tables with a header row, the columns in the order
listed: `run.csv`, with exactly one row, and `positions.csv`, with one row per
position. An empty cell is a missing value.

<a id="run-table"></a>

### Run

| Field | Type | Required | Meaning |
| --- | --- | --- | --- |
| `contract_version` | enum `1.0.0` | Yes | The contract the input follows |
| `profile_id` | enum `eu-crr-standard` | Yes | The regulatory profile |
| `profile_version` | enum `1` | Yes | The profile version (ADR-009) |
| `register_version` | enum `1` | Yes | The source register version the classifications cite |
| `parameter_set_id` | identifier | Yes | The parameter set holding the factors |
| `parameter_set_version` | positive integer, 1 to 9 ASCII digits, no leading zero | Yes | Its version |
| `as_of_date` | date | Yes | The reporting date; every residual maturity is measured from it |
| `reporting_entity_id` | identifier | Yes | The institution reported on |
| `basis` | enum `individual` | Yes | `consolidated` is rejected in version 1 |
| `reporting_currency` | currency | Yes | The only currency of the run |
| `amount_unit` | enum `units` | Yes | Amounts are in units of the reporting currency, never thousands or millions |
| `source_system` | text | Yes | The system the input was extracted from |
| `extracted_at` | timestamp | Yes | When it was extracted, in UTC |

<a id="positions-table"></a>

### Positions

| Field | Type | Required | Meaning |
| --- | --- | --- | --- |
| `record_id` | identifier | Yes | Stable, unique key of the position; never its row number |
| `side` | enum `asf`, `rsf` | Yes | Available or required stable funding |
| `item_type` | enum `own_funds`, `liability`, `asset`, `off_balance_sheet` | Yes | `own_funds` and `liability` belong to `asf`, the others to `rsf`; `derivative` is rejected in version 1 |
| `description` | text | Yes | A human-readable label; never parsed |
| `amount` | amount | Yes | The carrying amount, a non-negative decimal with at most four decimal places |
| `currency` | currency | Yes | Must equal `reporting_currency` |
| `maturity_kind` | enum `dated`, `no_stated_maturity` | Yes | Whether the position has a contractual maturity date |
| `maturity_date` | date | When `maturity_kind` is `dated`, otherwise empty | Contractual maturity; not before `as_of_date` |
| `classification` | classification | Yes | The register article that governs the position (see below) |
| `encumbered` | enum `true`, `false` | Yes | Whether the position is encumbered at the as-of date |
| `encumbrance_end_date` | date | When `encumbered` is `true`, otherwise empty | End of the encumbrance; not before `as_of_date` |
| `counterparty_type` | enum `financial_customer`, `other`, `not_applicable` | Yes | Needed for the S-01 Article 510 treatment of short-term transactions with financial customers |
| `transaction_type` | enum `secured_lending`, `capital_market_driven`, `unsecured`, `none` | Yes | `derivative` and `netted` are rejected in version 1 |
| `source_record` | text | Yes | Lineage: the record in `source_system` the position comes from |

<a id="value-types"></a>

### Value types

| Type | Form |
| --- | --- |
| identifier | 1 to 64 characters: a letter or digit, then letters, digits, `.`, `_` or `-` |
| date | ISO 8601 calendar date `YYYY-MM-DD` |
| timestamp | ISO 8601 UTC `YYYY-MM-DDThh:mm:ssZ` |
| currency | three upper-case letters (ISO 4217); the code list itself is not checked |
| amount | digits, optionally a point and one to four decimals; at most 15 integer digits; no sign, exponent, grouping or spaces. Positive and negative positions are told apart by `side` and `item_type`, never by sign |
| classification | `428` plus the article letters, with optional paragraph, point and sub-point, for example `428l`, `428l(2)` or `428l(2)(a)(i)`. The article must be one of the register's ASF factor articles (428k to 428o) for `asf`, or RSF factor articles (428r to 428ah) for `rsf` |
| text | any non-empty text |

The classification key names the provision; the factor it carries comes
only from the parameter set, with its citation. The keys in the examples
show the syntax only; they do not state which article governs any item.

The contract carries dates, not residual maturities or bands: the engine
derives each residual maturity from `as_of_date` and `maturity_date` (or
`encumbrance_end_date`), and the band boundaries come, with their citations,
from the parameter set. A position with `no_stated_maturity` has no maturity
date; how it is treated is the cited article's, not the contract's.

<a id="parameter-set"></a>

## ⚙️ Parameter set

A parameter set is versioned configuration, never input data. It carries
`parameter_set_id`, `parameter_set_version`, `profile_id`, `register_version`,
`effective_from`, `review_status` and a list of factors, each with its
`classification`, `factor` (a decimal from 0 to 1), the register `source_id`
and the `locator` (article, paragraph and point) it is taken from. No
regulatory parameter set exists in this repository yet; the only one is the
synthetic set of the neutral fixture FX-0001 (the non-admitted prototype's
transcribed tables are not one; ADR-011). Regulatory values are reviewed under
[PROD-01](../BACKLOG.md#later-calculation-work) and cited through the
register, as [AGENTS.md](../../AGENTS.md#what-the-documentation-describes)
requires. The engine rejects a run whose parameter set is unknown, does not
match the profile and register, or lacks a factor for a classification used.

<a id="result"></a>

## 📤 Result

A result has `status` `ok` or `rejected`, echoes the run identity
(`contract_version`, profile, register, parameter set, `as_of_date`,
`reporting_entity_id`) and carries:

- `input_fingerprint` and `configuration_fingerprint` (see
  [repeatability](#repeatability));
- when `ok`: `available_stable_funding`, `required_stable_funding` and
  `ratio`, plus the reconciliation below;
- `diagnostics`, possibly empty.

A contract-stage rejection carries no fingerprint; an engine-stage rejection
carries the input fingerprint. A `rejected` result carries no total and no
ratio. When required stable
funding is zero, the ratio is not reported and `NSFR-W001` explains why. The
[validation rules](../validation/VALIDATION.md#arithmetic-scale-and-rounding)
fix the reported scales: both totals with eight decimals, and the ratio as a
decimal with six decimals, rounded half-even once from the exact quotient.

<a id="reconciliation"></a>

### Reconciliation

Every result of a valid input reports, before any factor: the record count;
the count and amount total of each side; and the count and amount total of
each pair of side and classification. Totals are exact decimal sums, shown
with four decimals, computed with enough precision that no input within the
amount limits can overflow them. They let a reviewer tie the input to the
source system independently of the calculation. The engine's numeric type
must carry the same limits; the [validation rules](../validation/VALIDATION.md#numeric-range)
record its range.

<a id="diagnostics"></a>

## 🩺 Diagnostics

Each diagnostic names its code, the `record_id` (or `run`, `positions` for a
table-level problem) and the field. Contract-stage diagnostics come from
validating the input alone; engine-stage ones need the parameter set. Any
error rejects the whole run: no partial result is produced.

| Code | Stage | Meaning |
| --- | --- | --- |
| `NSFR-E001` | Contract | A required value is missing |
| `NSFR-E002` | Contract | A value does not match its type, pattern or enumeration |
| `NSFR-E003` | Contract | A `record_id` appears more than once (reported once per repeated ID) |
| `NSFR-E004` | Contract | An amount is not a finite decimal number, such as `NaN` or `Infinity` |
| `NSFR-E005` | Contract | An amount is negative |
| `NSFR-E006` | Contract | An amount has more than four decimal places |
| `NSFR-E007` | Contract | A currency differs from the reporting currency |
| `NSFR-E008` | Contract | A maturity or encumbrance end date is before the as-of date |
| `NSFR-E009` | Contract | `item_type` does not belong to the `side` |
| `NSFR-E010` | Contract | The classification is outside the articles of its side |
| `NSFR-E011` | Contract | A conditional value is missing, or present when it must be empty |
| `NSFR-E012` | Contract | The input uses a feature that version 1 does not support |
| `NSFR-E013` | Contract | The run table does not have exactly one row, or a table's columns or a row's cells differ from the contract |
| `NSFR-E016` | Contract | An amount has more than fifteen integer digits |
| `NSFR-E014` | Engine | The parameter set is unknown, or does not match the profile and register |
| `NSFR-E015` | Engine | No factor in the parameter set covers a classification |
| `NSFR-W001` | Engine | Required stable funding is zero, so the ratio is not reported |

<a id="repeatability"></a>

## 🔁 Immutability and repeatability

- The engine never changes the input or the parameter set; the API copies
  both into a run-local snapshot before calling the core.
- The input fingerprint is the SHA-256 of the canonical input: a `[run]`
  line, the run header and row, a `[positions]` line, the positions header
  and the positions sorted by `record_id`, columns in contract order, amounts
  with exactly four decimals, CSV with LF line ends. Row order and trailing
  zeros therefore do not change it.
- The configuration fingerprint is the SHA-256 of the parameter set in its
  canonical form, defined with the parameter-set format under PROD-01; the
  result also names the source commit of the engine.
- The same input fingerprint, configuration fingerprint and engine commit
  give a byte-identical result.

<a id="synthetic-examples"></a>

## 🧪 Synthetic examples

[`examples/`](examples) holds one directory per case, each with `run.csv`,
`positions.csv` and the `expected.json` the checker must reproduce: the
diagnostics, or for a valid input its reconciliation and input fingerprint.
All data is synthetic ([data policy](../DATA_POLICY.md)).

| Example | Shows |
| --- | --- |
| `valid-basic` | A valid input on both sides, with dated, undated, encumbered and S-01 Article 510 positions, and its reconciliation totals |
| `invalid-positions` | At least one row per contract error: missing, malformed, non-finite, negative, over-precise and over-limit amounts, a malformed date, another currency, a past maturity date, side mismatches, a missing and an unexpected conditional date, a derivative, an extra cell and a duplicate `record_id` |
| `invalid-run-basis` | The consolidated basis, rejected in version 1 |
| `invalid-run-rows` | A run table with two rows |

<a id="changing-the-contract"></a>

## 🔀 Changing the contract

A change to a field, a type, an enumeration or a diagnostic changes the
contract version, the JSON file, this document and the examples together,
under an issue I approve. Adding an optional field or a new diagnostic is a
minor version; removing or redefining one is a major version.
