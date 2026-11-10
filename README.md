<div align="center">

<a id="k-nsfr"></a>

# 🏦 K-NSFR

<a id="an-auditable-net-stable-funding-ratio-reference-project-in-excelvba"></a>

### An auditable Net Stable Funding Ratio reference project in Excel/VBA

**Planned EU CRR standard profile · Synthetic examples · Traceable rules and evidence · Setup stage**

<br>

[![Stage](https://img.shields.io/badge/Stage-Repository_setup-BF8700?style=for-the-badge)](#status)
[![Target](https://img.shields.io/badge/Target-Excel_VBA_64--bit-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)](#requirements)
[![Platform](https://img.shields.io/badge/Platform-Windows_11-0078D6?style=for-the-badge)](#requirements)
[![License](https://img.shields.io/badge/License-MPL_2.0-2ea44f?style=for-the-badge)](LICENSE)

<br>

[Status](#status) · [Scope](#scope) · [Requirements](#requirements) · [Quick start](#quick-start) · [Layout](#layout) · [Documentation](#documentation) · [Validation](#validation) · [Roadmap](#roadmap)

</div>

> [!IMPORTANT]
> **Repository setup (v0.1.0).** No admitted NSFR calculation engine, workbook
> or validated regulatory profile exists yet; the benchmark
> [prototype](prototype/README.md) is stored but not admitted. Do not use this baseline for
> regulatory reporting, production decisions or a claim of compliance.

<a id="what-this-project-is"></a>

## ✨ What this project is

K-NSFR is being prepared as a reviewable Excel/VBA reference implementation of
the Net Stable Funding Ratio. Its initial audience is risk analysts, reviewers,
instructors and developers working with synthetic cases. The first planned
calculation profile is the EU CRR standard NSFR; Basel material is a conceptual
reference, not an interchangeable rule set. Source and interpretation decisions
must be approved before any numerical implementation.

<a id="status"></a>

## 🎯 Status

| Area | Current state |
| --- | --- |
| Repository | Governance, licensing, architecture, source conventions and portable CI are in place |
| Calculation | Not implemented; no ASF, RSF or ratio output is available. A benchmark [prototype](prototype/README.md) is stored but not admitted |
| Tests | A neutral VBA [test harness](tests/README.md) with its self-test; it passed in Excel on 2026-11-09 at `5231985` ([host evidence](tests/evidence/523198552eee8296b1dfa4e72344c2a8e5d55cfd.json)) |
| Regulatory profiles | None executable or validated; EU CRR standard is a planned target |
| Workbook host | Windows 11 / Microsoft 365 desktop Excel / 64-bit VBA7 is the declared target; only the neutral harness has run there ([host evidence](tests/evidence/523198552eee8296b1dfa4e72344c2a8e5d55cfd.json)); no engine or workbook exists yet |
| Data | Synthetic-only policy; no production or client data |

<a id="scope"></a>

## 🧭 Scope

The setup milestone establishes the process and evidence needed to build a
deterministic calculation core, a typed API and a narrow Excel host adapter.
Future work must trace methodology and parameters to applicable sources,
separately validate numerical cases and record actual host execution. Read the
[scope](docs/SCOPE.md) and [architecture](docs/ARCHITECTURE.md) before using
the backlog as an implementation plan.

<a id="requirements"></a>

## 🧰 Requirements

The intended engine host is **Windows 11 with Microsoft 365 desktop Excel,
VBA7 and 64-bit Office**. On that host only the neutral test harness has run,
with two PASS [host evidence records](tests/evidence/README.md); no engine
exists to run. Portable repository checks run on Ubuntu 24.04 with Python 3.13
and Node 22; passing them does not prove Excel compatibility or regulatory
correctness. See [implementation conventions](docs/IMPLEMENTATION.md#supported-and-tested-environments)
for the support matrix and [source-control conventions](docs/SOURCE_CONTROL.md)
for exported modules.

<a id="quick-start"></a>

## 🚀 Quick start

Clone the repository and work on the active release branch:

```sh
git clone https://github.com/danielep71/K-NSFR.git
cd K-NSFR
git checkout release/0.1.0
```

- **Portable checks:** install the [locked tools](tools/README.md#running-the-checks),
  then run `python tools/check.py --ci` on a clean, committed checkout.
- **Supported host:** with Excel closed, run
  `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\host\Verify-NSFR-Branch.ps1 -Branch release/0.1.0`.
  It builds the neutral harness in a new workbook, runs its self-test and
  writes a host evidence record.

The [installation guide](docs/INSTALLATION.md) gives the prerequisites, the
manual route and troubleshooting.

<a id="layout"></a>

## 🗂️ Layout

| Path | Content |
| --- | --- |
| `docs/` | Scope, architecture, conventions, governance and decisions; `contracts/`, `methodology/` and `validation/` hold the contracts, the source register and the validation rules |
| `tests/` | The neutral VBA harness (`core/`), the numerical fixtures (`fixtures/`) and the host evidence records (`evidence/`) |
| `tools/` | The portable checks, their tests and the host scripts (`host/`) |
| `prototype/` | My benchmark prototype, stored as delivered and not admitted (ADR-011) |
| `.github/` | Workflows, issue forms, the PR template, the Dependabot configuration and the label catalogue with its scripts |

There is no `src/` yet: admitted VBA modules will live in `src/core/`,
`src/api/` and `src/host/` ([architecture](docs/ARCHITECTURE.md)).

<a id="documentation"></a>

## 📚 Documentation

| Read | Purpose |
| --- | --- |
| [Scope](docs/SCOPE.md) · [Architecture](docs/ARCHITECTURE.md) | Users, profiles, boundaries and planned layout |
| [Implementation conventions](docs/IMPLEMENTATION.md) · [Source control](docs/SOURCE_CONTROL.md) | API, numeric, host, encoding and generated-file rules |
| [Installation and host setup](docs/INSTALLATION.md) | Portable checks, host prerequisites, the three routes to run the harness, and troubleshooting |
| [Decision log](docs/DECISIONS.md) · [Backlog](docs/BACKLOG.md) | Selected decisions and later calculation work |
| [Regulatory source register](docs/methodology/SOURCE_REGISTER.md) | NSFR sources, their verification, concept mapping and applicability of the planned profile |
| [Input and result contracts](docs/contracts/CONTRACTS.md) | Field dictionary, parameter and result schemas, diagnostics and synthetic examples |
| [Numerical fixtures and validation rules](docs/validation/VALIDATION.md) · [Host evidence](docs/validation/HOST_EVIDENCE.md) | Fixture schema, independent derivation, scales, rounding, comparison and the case plan; the record of a run on the supported host |
| [Contributor guide](CONTRIBUTING.md) · [Code of conduct](CODE_OF_CONDUCT.md) · [Working rules](AGENTS.md) · [Tooling](tools/README.md) · [VBA house style](docs/VBA_HOUSE_STYLE.md) | Expected behavior, editorial rules, reproducible checks and implementation review |
| [Neutral test harness](tests/README.md) | The VBA test harness, its self-test and how to run it in Excel |
| [NSFR benchmark prototype](prototype/README.md) | My prototype, stored as delivered and not admitted (ADR-011) |
| [Documentation style](docs/WRITING_STYLE.md) | Consistent glyphs, badges, navigation and evidence wording |
| [Governance](docs/GOVERNANCE.md) | Issues, labels, branches and integration |
| [Release procedure](docs/RELEASE.md) · [Changelog](CHANGELOG.md) | Versions, release candidate, tag, rollback, milestone closure and what each version adds |
| [Repository settings](docs/REPOSITORY_SETTINGS.md) · [Settings evidence](docs/REPOSITORY_SETTINGS_EVIDENCE.md) | Non-default settings and their dated read-back |
| [License](LICENSE) · [Distribution](docs/LICENSING.md) · [Third-party notices](docs/licenses/THIRD_PARTY_NOTICES.txt) · [Provenance](docs/PROVENANCE.md) | MPL-2.0 selection, rights, attributions and provenance inventory |
| [Synthetic data](docs/DATA_POLICY.md) · [Security](SECURITY.md) | Data boundaries and private vulnerability reporting |
| [Traffic history](docs/TRAFFIC.md) | Daily traffic collection, its credential and its alert |

<a id="validation"></a>

## ✅ Validation

The required **Repository integrity** check validates committed text, Markdown
file and explicit-anchor links, fences, YAML and workflow syntax, strict Python
typing/lint/tests, the VBA static rules, the label catalogue and the
issue-form configuration. Live
issues are audited separately by the issue triage workflow. The check runs on
documentation-only and synchronization PRs. It never runs Excel: a run on
the supported host is proved only by a [host evidence record](docs/validation/HOST_EVIDENCE.md),
bound to its commit and validated by the `host-evidence` gate (#14). The
workbook template and the export/import round trip come with #13;
independent numerical validation comes later. Follow the commands in
[governance](docs/GOVERNANCE.md#local-validation) for portable local checks.

<a id="roadmap"></a>

## 🗺️ Roadmap

The [v0.1.0 repository setup milestone](https://github.com/danielep71/K-NSFR/milestone/1)
uses `release/0.1.0` as its active development branch. Task PRs target that
branch; I control integration into `main`. The
[follow-on backlog](docs/BACKLOG.md) separates future calculation work from
this setup milestone. No release or numerical capability is implied by the
branch name.

<a id="license-data-and-security"></a>

## ⚖️ License, data and security

K-NSFR uses [MPL-2.0](LICENSE) for contributions the rights holders can license.
The repository has been public since 2026-10-28; the post-publication review
is recorded in #19. Examples and fixtures must be synthetic. Report
suspected vulnerabilities privately as described in [SECURITY.md](SECURITY.md).
