# 🗺️ Follow-on backlog

**K-NSFR** · [Overview](../README.md) · [Documentation](../README.md#documentation) · [Governance](GOVERNANCE.md)

[![Stage](https://img.shields.io/badge/Stage-Repository_setup-BF8700?style=flat-square)](../README.md#status)

This backlog separates existing setup work from later numerical implementation.
All links refer to K-NSFR. Production work items below are planning records, not
created GitHub issues, scheduled commitments, or implemented features.

## 🧩 Existing setup delivery

| Work | Existing issue | Deliverable required before calculation work |
| --- | --- | --- |
| Licensing and synthetic-data policy | [#3](https://github.com/danielep71/K-NSFR/issues/3) | Approved ownership/distribution and data boundaries |
| Text and generated-file conventions | [#4](https://github.com/danielep71/K-NSFR/issues/4) | Canonical source handling |
| Implementation and host conventions | [#5](https://github.com/danielep71/K-NSFR/issues/5) | API/numeric/error conventions and target matrix |
| Regulatory source register | [#8](https://github.com/danielep71/K-NSFR/issues/8) | Versioned sources, applicability, and unresolved interpretations |
| Data and reconciliation contracts | [#9](https://github.com/danielep71/K-NSFR/issues/9) | Input, parameter, output, and diagnostic schemas |
| Independent validation design | [#10](https://github.com/danielep71/K-NSFR/issues/10) | Fixture provenance, oracle, and acceptance rules |
| Static checks and neutral harness | [#11](https://github.com/danielep71/K-NSFR/issues/11), [#12](https://github.com/danielep71/K-NSFR/issues/12) | Meaningful failure/pass checks with no NSFR claims |
| Host packaging and actual evidence | [#13](https://github.com/danielep71/K-NSFR/issues/13), [#14](https://github.com/danielep71/K-NSFR/issues/14) | Reproducible import and genuine supported-host execution |
| Setup certification | [#22](https://github.com/danielep71/K-NSFR/issues/22) | Exact-candidate setup readiness, subject to all milestone dependencies |

## 🚀 Later calculation work

Accountable backlog owner: @danielep71. Before starting each item, create a scoped
issue with priority, milestone, target end date, dependencies, acceptance
criteria, and evidence requirements. No production dates are assigned here.

| ID | Scope | Prerequisites | Acceptance gate |
| --- | --- | --- | --- |
| PROD-01 | Approve executable `eu-crr-standard` methodology and parameter package | #8, #9; applicable source and interpretation review | Every rule/parameter traces to an approved version; unresolved applicability blocks admission |
| PROD-02 | Implement input validation and ASF/RSF classification | PROD-01, #5, #10, #11 | Independent valid/invalid and maturity/encumbrance boundary fixtures; no silent defaults |
| PROD-03 | Implement aggregation, ratio outputs, and reconciliations | PROD-02 | Independently derived totals; explicit zero-denominator, sign, rounding, and currency/perimeter behavior |
| PROD-04 | Validate the reference engine end to end | PROD-03, #12-#14 | Independent oracle, regression coverage, deterministic results, and exact-candidate host evidence |
| PROD-05 | Assess production admission and operational controls | PROD-04, #3, #22 | Named owner review of applicability, data controls, deployment, security, change management, monitoring, and documented limitations |
| PROD-06 | Assess additional jurisdictions or simplified profiles | PROD-01; separate owner scope decision | Separate profile/version and independent fixtures; no implicit reuse of EU parameters |

PROD-04 may support a reference release; it does not by itself authorize PROD-05
or production use. These records introduce no formulas, regulatory weightings,
or executable calculation code.
