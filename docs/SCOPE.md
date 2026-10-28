# 🎯 Scope and maturity

**K-NSFR** · [Overview](../README.md) · [Documentation](../README.md#documentation) · [Governance](GOVERNANCE.md)

[![Stage](https://img.shields.io/badge/Stage-Repository_setup-BF8700?style=flat-square)](../README.md#status)

This is the scope decision for [issue #2](https://github.com/danielep71/K-NSFR/issues/2),
dated 2026-10-28. It defines planned capabilities, not implemented calculations.

## 👥 Users and use cases

| User | Intended use | Boundary |
| --- | --- | --- |
| Risk analyst or advisor | Inspect synthetic funding and asset classifications, scenario inputs, and reconciliations | No client data or operational reliance in setup |
| Instructor or learner | Explain ASF, RSF, and the effect of classification assumptions using reviewed examples | Teaching examples must identify their profile and assumptions |
| Methodology reviewer | Trace future rules and parameters to sources and independently derived expected results | A passing software check is not regulatory approval |
| Developer or maintainer | Build deterministic components and reproduce tests and host imports | Setup tooling does not constitute an NSFR engine |

Future reference calculations should expose available stable funding (ASF),
required stable funding (RSF), diagnostics, and traceable reconciliation outputs.
Input scope must explicitly identify the as-of date, reporting perimeter,
currency basis, and methodology version. No default entity consolidation,
currency conversion, or missing-data interpretation is implied.

## 📜 Regulatory profiles

| Profile identifier | Role | Status and admission condition |
| --- | --- | --- |
| `eu-crr-standard` | First planned calculation profile for the standard EU CRR NSFR | Selected target only; exact legal version, effective date, applicability, and interpretations require the source register in #8 |
| `basel-reference` | Conceptual and educational comparison | Reference-only; no executable mapping or equivalence to the EU profile is claimed |

**Currently supported executable regulatory profiles: none.** These identifiers
are repository design choices, not official regulatory labels. A future profile
must bind a jurisdiction, applicability decision, exact source version, effective
date range, parameter version, and review status. An unknown or unapproved profile
must fail explicitly; it must never silently fall back to another profile.

Simplified NSFR treatments, national discretions, other jurisdictions, and
institution-specific interpretations are outside the initial target. They require
separate applicability decisions and independently validated fixtures before
being admitted. #8 owns source verification; this document specifies no legal
thresholds, weightings, reporting obligations, or effective legal text.

## 🛡️ Reference use and production admission

The setup baseline may be used to review architecture and repository processes.
Even an educational numerical example requires reviewed provenance before it
can be presented as an NSFR result. Later reference releases may demonstrate
synthetic calculations with explicit limitations.

Production use remains out of scope until a separately approved workstream
establishes regulatory applicability, independent numerical validation, data
quality and reconciliation controls, access and operational controls, change
management, supported-host evidence, and accountable owner sign-off. Passing
portable CI or an Excel smoke test alone does not satisfy this gate.

## 📦 v0.1.0 inclusions and exclusions

The setup milestone defines governance, architecture, source provenance,
contracts, static checks, neutral test infrastructure, host packaging, and
candidate-bound evidence through the existing setup issues.

It does not introduce regulatory formulas or parameter percentages, calculate
ASF/RSF or a ratio, certify compliance, ingest production/client data, implement
bank connectors, create regulatory submissions, automate consolidation or FX,
build a web service, or publish a release. Neutral scaffold tests must not be
described as numerical NSFR validation.

See [architecture](ARCHITECTURE.md), [decisions](DECISIONS.md), and
[follow-on backlog](BACKLOG.md) for delivery boundaries.
