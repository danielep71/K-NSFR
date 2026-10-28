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

[Status](#status) · [Scope](#scope) · [Requirements](#requirements) · [Documentation](#documentation) · [Validation](#validation) · [Roadmap](#roadmap)

</div>

> [!IMPORTANT]
> **Repository setup (v0.1.0).** No NSFR calculation engine, workbook or
> validated regulatory profile exists yet. Do not use this baseline for
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
| Calculation | Not implemented; no ASF, RSF or ratio output is available |
| Regulatory profiles | None executable or validated; EU CRR standard is a planned target |
| Workbook host | Windows 11 / Microsoft 365 desktop Excel / 64-bit VBA7 is the declared target; execution **NOT_RUN** |
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
VBA7 and 64-bit Office**. That combination has not been tested here. Portable
repository checks run on Ubuntu 24.04 with Python 3.13 and Node 22; passing
them does not prove Excel compatibility or regulatory correctness. See
[implementation conventions](docs/IMPLEMENTATION.md) for the support matrix
and [source-control conventions](docs/SOURCE_CONTROL.md) for exported modules.

<a id="documentation"></a>

## 📚 Documentation

| Read | Purpose |
| --- | --- |
| [Scope](docs/SCOPE.md) · [Architecture](docs/ARCHITECTURE.md) | Users, profiles, boundaries and planned layout |
| [Implementation conventions](docs/IMPLEMENTATION.md) · [Source control](docs/SOURCE_CONTROL.md) | API, numeric, host, encoding and generated-file rules |
| [Decision log](docs/DECISIONS.md) · [Backlog](docs/BACKLOG.md) | Selected decisions and later calculation work |
| [Contributor guide](CONTRIBUTING.md) · [Tooling](tools/README.md) · [VBA house style](docs/VBA_HOUSE_STYLE.md) | Reproducible checks and implementation review |
| [Documentation style](docs/WRITING_STYLE.md) | Consistent glyphs, badges, navigation and evidence wording |
| [Governance](docs/GOVERNANCE.md) | Issues, labels, branches and integration |
| [License](LICENSE) · [Distribution](docs/LICENSING.md) · [Provenance](docs/PROVENANCE.md) | MPL-2.0 selection, rights and attribution inventory |
| [Synthetic data](docs/DATA_POLICY.md) · [Security](SECURITY.md) | Data boundaries and private vulnerability reporting |

<a id="validation"></a>

## ✅ Validation

The required **Repository integrity** check validates committed text, Markdown
file and explicit-anchor links, fences, YAML and workflow syntax, strict Python
typing/lint/tests, label policy and
issue triage on PRs. It runs on documentation-only and synchronization PRs.
Actual Excel import/export and host evidence are planned under issue #14;
independent numerical validation comes later. Follow the commands in
[governance](docs/GOVERNANCE.md#local-validation) for portable local checks.

<a id="roadmap"></a>

## 🗺️ Roadmap

The [v0.1.0 repository setup milestone](https://github.com/danielep71/K-NSFR/milestone/1)
uses `release/0.1.0` as its active development branch. Task PRs target that
branch; the owner controls integration into `main`. The
[follow-on backlog](docs/BACKLOG.md) separates future calculation work from
this setup milestone. No release or numerical capability is implied by the
branch name.

<a id="license-data-and-security"></a>

## ⚖️ License, data and security

K-NSFR uses [MPL-2.0](LICENSE) for contributions the rights holders can license.
The repository is private during setup and intended to become public after the
separate publication review. Examples and fixtures must be synthetic. Report
suspected vulnerabilities privately as described in [SECURITY.md](SECURITY.md).
