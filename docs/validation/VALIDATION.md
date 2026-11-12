<a id="numerical-fixtures-and-validation-rules"></a>

# 🧪 Numerical fixtures and validation rules

**K-NSFR** · [Overview](../../README.md) · [Documentation](../../README.md#documentation) · [Governance](../GOVERNANCE.md)

[![Stage](https://img.shields.io/badge/Stage-Prototype_setup-BF8700?style=flat-square)](../../README.md#status)

Fixture rules version **1**, for [contract](../contracts/CONTRACTS.md) version
1.0.0. Tracking: #10.

This document defines how a numerical fixture is recorded, how its expected
result is derived and how an engine result is compared with it. The
machine-readable form is [fixture-v1.json](fixture-v1.json);
`tools/check_fixtures.py` (gate `fixtures`) checks every fixture under
[`tests/fixtures/`](../../tests/fixtures) against it. The repository holds
these rules and one [neutral example](#neutral-example) only. Validating
production numbers belongs to the [later calculation work](../BACKLOG.md#later-calculation-work):
PROD-01 admits the parameters, PROD-02 the validation and classification,
PROD-03 the aggregation and the ratio, and PROD-04 the end-to-end validation.

<a id="principles"></a>

## 📌 Principles

- **Independent oracle.** An expected result is derived by hand, by an
  independent tool or from a published source, never solely from the
  implementation under test. An engine output may be compared with a fixture;
  it is never copied into one. The `fixtures` gate derives every expected
  result a second time with exact rational arithmetic in Python, which shares
  no code with the VBA engine.
- **Registered is not executed.** A fixture is *registered* when it is in
  `tests/fixtures/` and passes the `fixtures` gate. It is *executed* only when
  a run of the engine at a recorded commit, on a recorded host, is listed in
  its `execution.runs` with evidence. The gate never runs the engine, so a
  green gate is not execution evidence (ADR-006).
- **Synthetic data.** Every fixture follows the
  [data policy](../DATA_POLICY.md).
- **Cited factors.** A validation fixture takes every factor from the
  [source register](../methodology/SOURCE_REGISTER.md), by entry ID and
  locator, as [AGENTS.md](../../AGENTS.md#what-the-documentation-describes)
  requires; an entry still marked "to verify" is never used for an expected
  result. Only a neutral example uses synthetic factors.

<a id="fixture-layout"></a>

## 🗂️ Fixture layout

A fixture is a directory under `tests/fixtures/` holding `fixture.json`, the
two input tables of the contract (`run.csv` and `positions.csv`) and the
parameter set it uses (`parameters.json`).

<a id="fixture-fields"></a>

### Fixture fields

| Field | Required | Meaning |
| --- | --- | --- |
| `fixture_schema_version` | Yes | `1`, the version of these rules |
| `fixture_id` | Yes | `FX-` and four digits; stable, never reused |
| `title` | Yes | What the fixture shows |
| `kind` | Yes | `neutral-example` (synthetic factors, no regulatory meaning) or `validation` (cited factors) |
| `contract_version` | Yes | The contract the inputs follow; must equal the contract's |
| `inputs` | Yes | The input files, relative to the fixture directory |
| `inputs.run` | Yes | The run table |
| `inputs.positions` | Yes | The positions table |
| `inputs.parameters` | Yes | The parameter set |
| `provenance` | Yes | Where the data comes from |
| `provenance.data` | Yes | `synthetic`; no other value is admitted |
| `provenance.author` | Yes | Who wrote the fixture |
| `provenance.created` | Yes | When, as an ISO date |
| `provenance.sources` | Yes | Register entries and locators the expected result relies on, each starting with a verified `S-` entry ID, such as `S-01 Article 428b`; empty only for a neutral example |
| `provenance.note` | No | Any further remark |
| `derivation` | Yes | How the expected result was obtained |
| `derivation.method` | Yes | `hand`, `independent-tool` or `published`; there is no value for the implementation under test |
| `derivation.independent_of_implementation` | Yes | Must be `true` |
| `derivation.description` | Yes | The steps, the tool and its version, or the publication and its locator |
| `derivation.weighted` | When `expected.status` is `ok` | One step per position: `record_id`, `amount`, `factor` and `weighted`, which is exactly `amount` × `factor` |
| `derivation.available_stable_funding` | When `expected.status` is `ok` | The derived ASF total |
| `derivation.required_stable_funding` | When `expected.status` is `ok` | The derived RSF total |
| `expected` | Yes | The result the engine must give |
| `expected.status` | Yes | `ok` or `rejected` |
| `expected.input_fingerprint` | When the input passes contract validation | The contract's input fingerprint, as `check_contracts.py` computes it; a hash is not derived by hand |
| `expected.available_stable_funding` | When `ok` | ASF with eight decimals |
| `expected.required_stable_funding` | When `ok` | RSF with eight decimals |
| `expected.ratio` | When `ok` | The ratio with six decimals, or `null` with `NSFR-W001` |
| `expected.diagnostics` | Yes | Every diagnostic, sorted by code, `record_id` and field |
| `comparison` | Yes | How results are compared |
| `comparison.default` | Yes | `exact`; no other value is admitted |
| `comparison.exceptions` | Yes | Approved tolerances, usually none (see [comparison](#comparison)) |
| `review` | Yes | Independent review of the fixture |
| `review.status` | Yes | `draft`, `reviewed` or `approved` |
| `review.reviewer` | Yes | Who reviewed it; `null` while `draft` |
| `review.date` | Yes | When, as an ISO date; `null` while `draft` |
| `execution` | Yes | Whether the engine has run it |
| `execution.status` | Yes | `not-executed` or `executed` |
| `execution.runs` | Yes | Each run: `engine_commit`, `host`, `date`, `result` and `evidence`; empty while `not-executed` |

The checker requires every required field of `fixture-v1.json` to head
exactly one row of this table.

<a id="fixture-parameter-set"></a>

### Parameter set

`parameters.json` carries the parameter-set fields of the
[contract](../contracts/CONTRACTS.md#parameter-set). In a fixture:

- each factor is a decimal from 0 to 1 with at most four decimals, and a
  classification has at most one factor;
- a factor applies to a position whose classification equals its key
  exactly; whether a paragraph or point inherits its article's factor is
  decided under PROD-01 and needs a change to these rules;
- in a validation fixture each factor names a normative register entry
  (`S-` and two digits) and its locator; the checker rejects an entry that
  is not in the register, or whose status is not verified or still has
  something to verify;
- in a neutral example `review_status` is `synthetic`, and every factor has
  `source_id` `synthetic` and `locator` `none`. Synthetic factors are chosen
  for arithmetic; any coincidence with a regulatory factor carries no
  meaning. How the engine refuses synthetic parameters outside tests is
  decided with the parameter-set format under PROD-01.

<a id="arithmetic-scale-and-rounding"></a>

## 🧮 Arithmetic, scale and rounding

These rules settle what the contract leaves to #10 (ADR-010).

- **Weighted amounts.** Each position's weighted amount is its amount times
  its factor, exactly: at most four plus four decimals, never rounded.
- **Totals.** Available and required stable funding are the exact sums of the
  weighted amounts of their side, reported with exactly eight decimals, so
  reporting never rounds them.
- **Ratio.** The ratio is available divided by required stable funding,
  reported as a decimal, not a percentage (`1.000000` is 100%), with six
  decimals. It is the exact quotient rounded once, half-even: a tie goes to
  the even last digit, so `1.2345675` gives `1.234568` and `1.2345665` gives
  `1.234566`. An implementation that rounds an intermediate approximation,
  such as a 28-digit VBA `Decimal` quotient, must still give the result of
  rounding the exact quotient; the tie and near-tie fixtures of the
  [case plan](#case-plan) detect double rounding.
- **Zero denominator.** When required stable funding is zero, the result is
  `ok`, `ratio` is `null` and `NSFR-W001` explains why. When only available
  stable funding is zero, the ratio is `0.000000`.
- **Order.** Exact sums do not depend on the order of the positions; a result
  never changes when the rows are permuted, and every list in a result is
  sorted by its keys.
- **Cancellation.** Contract amounts are never negative and a sign is never
  carried, so no cancellation occurs within a side; the case returns if a
  later contract admits signed amounts.
- **Non-finite values.** The contract rejects `NaN` and `Infinity`
  (`NSFR-E004`). Expected and actual values are finite decimal text; a
  comparison involving a non-finite value fails.

<a id="numeric-range"></a>

### Numeric range of the engine

| Quantity | Largest value | Fits VBA `Currency` (922,337,203,685,477.5807, 4 decimals) | Fits VBA `Decimal` (79,228,162,514,264,337,593,543,950,335, up to 28 decimals) |
| --- | --- | --- | --- |
| Amount | 999,999,999,999,999.9999 | No | Yes |
| Weighted amount | 999,999,999,999,999.9999 × 1, with up to eight decimals | No | Yes |
| ASF or RSF total, eight decimals | The sum over all positions of a side | No | Yes, while the total stays below about 7.9 × 10<sup>20</sup>, about 790,000 positions at the largest amount |
| Ratio | ASF divided by the smallest non-zero RSF, 0.00000001 | No | Not always |

`Currency` therefore cannot hold every contract amount, and the engine uses
`Decimal` (a `Variant` built with `CDec`) for amounts, weighted amounts,
totals and the ratio; `Double` is never authoritative (ADR-010). An input
whose totals or ratio would exceed `Decimal` must be rejected with a
diagnostic, never rounded or truncated. That diagnostic is a contract
change, made with the aggregation work (PROD-03).

<a id="comparison"></a>

## ⚖️ Comparison

The default is exact: every field of `expected` must equal the engine's
result, as text in the reported scale. A field must be present on both
sides: a missing field never equals an explicit `null`. Decimal arithmetic with one defined
rounding makes the result reproducible character for character, so no
tolerance is needed by default.

A tolerance is an exception. It applies to one numeric field
(`available_stable_funding`, `required_stable_funding` or `ratio`), names
the issue I approved it in, and states why exact comparison cannot hold:

| Key | Meaning |
| --- | --- |
| `field` | The tolerated field |
| `mode` | `absolute`: the values may differ by at most `tolerance`; `relative`: by at most `tolerance` × the expected value, and never less than `floor` |
| `tolerance` | A non-negative decimal |
| `floor` | The absolute allowance near zero; required with `relative`, where a relative tolerance alone would shrink to zero |
| `reason` | Why exact comparison cannot hold |
| `approved_in` | The issue that approved it, such as `#10` |

Status, fingerprints and diagnostics are always compared exactly.

<a id="engine-stage-diagnostics"></a>

## 🩺 Engine-stage diagnostics

Contract validation comes first; a contract-stage error rejects the input
before any factor is looked up, and the result then has no fingerprint. After
it, in a fixture:

| Diagnostic | `record_id` | Field | When |
| --- | --- | --- | --- |
| `NSFR-E014` | `run` | The run field that differs | `parameter_set_id`, `parameter_set_version`, `profile_id` or `register_version` of the run differs from the parameter set |
| `NSFR-E015` | The position | `classification` | No factor has the position's classification |
| `NSFR-W001` | `run` | `required_stable_funding` | Required stable funding is zero |

An engine-stage error rejects the run with the input fingerprint and no total.

<a id="case-plan"></a>

## 🗺️ Case plan

Planned cases are registered as fixtures with the engine work named in the
last column, and executed by the test harness (#12); the contract examples
already cover contract-stage validation.

| Area | Case | Expected | Status |
| --- | --- | --- | --- |
| Valid | Both sides, dated, undated and encumbered positions, a zero factor, rows out of order | `ok` with totals and ratio | Registered: `FX-0001`, neutral |
| Valid | The same with cited factors | `ok` | Planned: PROD-02, after PROD-01 |
| Invalid | Every contract error | The contract diagnostics | Covered by the [contract examples](../contracts/CONTRACTS.md#synthetic-examples) |
| Parameters | The parameter set differs from the run | `NSFR-E014` | Planned: PROD-02 |
| Parameters | A classification without a factor | `NSFR-E015` | Planned: PROD-02 |
| Zero denominator | No RSF position, or only zero RSF factors | `ok`, no ratio, `NSFR-W001` | Planned: PROD-03 |
| Zero numerator | No ASF position | Ratio `0.000000` | Planned: PROD-03 |
| Maturity boundary | Maturity date equal to the as-of date; the day before | Accepted; `NSFR-E008` | Accepted case planned: PROD-02; rejection of an earlier date covered by the contract examples, the day-before boundary planned: PROD-02 |
| Maturity boundary | Residual maturity on each band boundary of the parameter set, and one day either side | The band of each side | Planned: PROD-02 |
| Encumbrance boundary | End date equal to the as-of date, on each band boundary and one day either side | The band of each side | Planned: PROD-02 |
| Classification | An article of the other side; a paragraph or point key | `NSFR-E010`; the matching rule | `NSFR-E010` covered by the contract examples; matching planned: PROD-02 |
| Currency and entity | Another currency; the consolidated basis | `NSFR-E007`; `NSFR-E012` | Covered by the contract examples. A second entity cannot occur: the run has one row (`NSFR-E013`) |
| Range | Amounts at the contract maximum with factor 1; totals and a ratio beyond `Decimal` | Exact totals; a rejection | Planned: PROD-03, with its diagnostic |
| Rounding | Quotients exactly on a half at the seventh decimal, with an odd and an even sixth digit; quotients one unit either side | Half-even; no double rounding | Planned: PROD-03 |
| Order | The same positions in two orders | Identical results | Planned: PROD-03; `FX-0001` lists its rows out of order |
| Round trip | Trailing zeros and the reported scales written and read back | Unchanged values and fingerprint | Planned: PROD-03 |

<a id="neutral-example"></a>

## 🧪 Neutral example

[`tests/fixtures/neutral-synthetic/`](../../tests/fixtures/neutral-synthetic)
is `FX-0001`. It contains six synthetic positions and six synthetic factors,
including a zero factor. Its derivation was worked out by hand:

- available stable funding is 1,125,000.25;
- required stable funding is 166,250.09255;
- the ratio is 6.766915.

The `fixtures` gate re-derives these values. The fixture stays `draft`
until I review it, and `not-executed` until the engine exists. It shows the
format only and is no evidence of any NSFR treatment.

<a id="changing-the-rules"></a>

## 🔀 Changing the rules

A change to a field, the scales, the rounding or the comparison changes
`fixture_schema_version`, `fixture-v1.json`, this document, the checker and
every fixture together, under an issue I approve, and is recorded in the
[decision log](../DECISIONS.md).
