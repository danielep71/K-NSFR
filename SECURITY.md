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

Report a suspected vulnerability privately through GitHub's private
vulnerability reporting: **Security → Report a vulnerability**
(<https://github.com/danielep71/K-NSFR/security/advisories/new>). Do not open a public issue or
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

<a id="automation-risk-surface"></a>

## ⚙️ Automation risk surface

No workflow runs VBA, Excel or an NSFR calculation. The workflows use only the
workflow token; no repository secret or other credential is configured in them.

| Workflow | Job token permissions | Runs on |
| --- | --- | --- |
| `Repository integrity` | `contents: read` | Pull requests into and pushes to `main` and `release/**`; manual dispatch |
| `Issue triage` | `contents: read`, `issues: read`, `pull-requests: read` | Issue events, from the default branch; manual dispatch |
| `Detect issue-label drift` | `contents: read`; `issues: read` for the live comparison | Label-file pull requests; daily schedule; manual dispatch |
| `Sync issue labels` | `contents: read`; `issues: write` only in the reconcile job | Label-file pull requests (validation only); label-file pushes to `main` and manual dispatch on the default branch (reconciliation) |
| `Keep milestones issue-only` | None at workflow level; `pull-requests: write` in its one job | `pull_request_target` (opened, reopened, edited, milestoned); daily schedule; its own pushes to `main`; manual dispatch on the default branch |

In these workflows every checkout sets `persist-credentials: false` and every
action is pinned to a full commit SHA. The workflow-level default permission is
`contents: read`, except in `Keep milestones issue-only`, which declares none.
That workflow uses `pull_request_target`, so it runs with write access on pull
requests from forks. It never checks out code and runs only its inline script,
from the base branch; the script refuses to change an issue and only clears PR
milestones. No workflow approves, merges or dismisses alerts. The
[tooling guide](tools/README.md) describes what each
workflow runs.

CodeQL code scanning runs separately as GitHub's default setup: a
GitHub-managed workflow with no file in this repository. It analyzes the
workflows and the Python tools on GitHub's default-setup triggers and weekly,
and uploads its results to code scanning. It uses GitHub's actions by version
tag (its first run used `github/codeql-action/analyze@v4`), so the full-SHA
pins above do not apply to it.

<a id="existing-controls"></a>
<a id="security-features"></a>

## 🛡️ Security features

Recorded on **2026-10-28** for the public repository and updated on
**2026-10-29** (#75). The source column says
how each state is known; a state marked *not verified* has not been read back
and must not be assumed.

| Feature | State | Source |
| --- | --- | --- |
| Rulesets on `main` and `release/**` | Active, as described in [governance](docs/GOVERNANCE.md#installed-settings-and-limits) | Owner statement in governance; enforcement scenarios tracked in #15 |
| Default workflow token | Read-only | Owner statement in governance; each workflow also declares `contents: read` at workflow level, or no permission at all (`Keep milestones issue-only`) |
| Full-SHA action pins | Required by the repository setting | Owner statement in governance; every `uses:` in the workflow files is a full SHA |
| Dependabot version updates | GitHub Actions: configured and active. Python tools in `tools/`: configured (#85); not yet observed running. Proposals arrive as pull requests for review | `.github/dependabot.yml`; Actions activation evidenced in #6 by Dependabot PR #41 |
| Private vulnerability reporting | Enabled | GitHub API read-back, 2026-10-28 |
| Secret scanning | Enabled | GitHub API read-back, 2026-10-28 |
| Secret scanning push protection | Enabled | GitHub API read-back, 2026-10-28 |
| Secret scanning non-provider patterns and validity checks | Disabled | GitHub API read-back, 2026-10-28 |
| Dependabot security updates | Enabled | GitHub API read-back, 2026-10-28 |
| Dependabot alerts | Enabled by inference: security updates require alerts | Not readable without admin access; not verified directly |
| Code scanning | CodeQL default setup, enabled 2026-10-29: Actions and Python, default query suite, remote threat model, weekly schedule. It covers the workflows and Python tools, not the VBA | GitHub API read-back, 2026-10-29. The first analysis, on `main` at `d272338`, succeeded and scanned 4 of 4 workflow files and 8 of 8 Python files ([run 40838760015](https://github.com/danielep71/K-NSFR/actions/runs/40838760015)). Alerts are not readable without security-events access; *not verified* |

Update this table, with its date and source, when a feature is enabled,
changed or read back. These controls do not constitute a security
certification; host validation has separate evidence requirements. See
[synthetic-data policy](docs/DATA_POLICY.md) and [provenance](docs/PROVENANCE.md).

<a id="safe-harbor"></a>

## 🤝 Safe harbor

The project will not initiate or recommend legal action for good-faith
research that follows this policy, stays within this repository's own source,
avoids privacy violations and destructive actions, and is reported privately.
This does not authorize testing third-party systems; no bug bounty is offered.
