<a id="security-policy"></a>

# 🔒 Security policy

**K-NSFR** · [Overview](README.md) · [Documentation](README.md#documentation) · [Data policy](docs/DATA_POLICY.md)

[![Stage](https://img.shields.io/badge/Stage-Repository_setup-BF8700?style=flat-square)](README.md#status)

<a id="supported-versions"></a>

## 📌 Supported versions

No version of K-NSFR has been released. The active release branch named in
[governance](docs/GOVERNANCE.md#branches-and-integration) and `main` are
supported on a best-effort basis. Reports must identify a full commit SHA.
Modified copies are unsupported unless the problem reproduces in this
repository.

<a id="report-a-vulnerability"></a>

## 📨 Report a vulnerability

K-NSFR is private during setup and is intended to become public. Report
suspected vulnerabilities privately to the repository owner through a private
GitHub channel available to you, or contact the owner outside the repository.
Before publication, the owner must verify that a usable private reporting route
is documented and available to external reporters. Do not open a public issue or
attach exploit details, credentials, client data, or personal data to an issue or
PR. The owner will acknowledge, triage, and coordinate remediation and any
necessary disclosure within the targets below.

Include the full commit SHA and the affected file, procedure or workflow; the
Excel/Office build, bitness and Windows version where relevant; the impact and
minimal reproduction steps, with synthetic data only; and any workaround you
tested.

<a id="response-targets"></a>

## ⏱️ Response targets

K-NSFR is maintained by one person; these targets are best effort, not a
contractual commitment or a security-support guarantee:

- acknowledgement within 5 business days;
- an initial assessment within 10 business days of receiving enough evidence;
- an update at least every 14 days while the report is investigated.

Remediation and disclosure timing are agreed with the reporter according to
severity.

<a id="scope"></a>

## 🎯 Scope

**In scope:** the repository tooling in [tools/](tools/README.md), the GitHub
Actions workflows and their configuration, and, once added, the VBA source,
workbook template and tests.

**Out of scope:**

- vulnerabilities in Excel, Office, Windows, GitHub, Python or Node.js
  themselves;
- organization-controlled macro and deployment policy;
- modified copies;
- ordinary defects without concrete security impact, such as an incorrect
  NSFR figure. Report those as issues.

<a id="handle-secrets-and-incidents"></a>

## 🔑 Handle secrets and incidents

Never commit secrets, tokens, connection strings, real customer records, or
workbooks containing them. Use the least-privilege credential store for local
work, review staged changes, and keep diagnostic logs free of secrets. If a
secret or sensitive record is exposed, stop sharing it, alert the owner privately,
revoke/rotate affected credentials, assess exposure and remove the material from
active artifacts. History cleanup requires a separate coordinated decision; a
normal deletion commit does not erase previous versions.

<a id="existing-controls"></a>

## 🛡️ Existing controls

GitHub Actions has a read-only default token and SHA-pinned actions; these
configuration controls do not constitute a security certification. Dependency
alerts and host validation have separate evidence requirements. See
[synthetic-data policy](docs/DATA_POLICY.md) and [provenance](docs/PROVENANCE.md).

<a id="safe-harbor"></a>

## 🤝 Safe harbor

The project will not initiate or recommend legal action for good-faith
research that follows this policy, stays within this repository's own source,
avoids privacy violations and destructive actions, and is reported privately.
This does not authorize testing third-party systems; no bug bounty is offered.
