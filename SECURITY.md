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
PR. I will acknowledge, triage, and coordinate remediation and any
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
secret or sensitive record is exposed, stop sharing it, alert me privately,
revoke/rotate affected credentials, assess exposure and remove the material from
active artifacts. History cleanup requires a separate coordinated decision; a
normal deletion commit does not erase previous versions.

<a id="automation-risk-surface"></a>

## ⚙️ Automation risk surface

The workflows below are on `main`, except `CodeQL` (#146) and `External
links` (#147) until I integrate them; meanwhile CodeQL runs on pull requests
and pushes from the release branch, and the link check only on manual dispatch
from a branch that has it.
Traffic history has collected since 2026-11-07 (#16).

No workflow runs VBA, Excel or an NSFR calculation. The workflows use the
workflow token, except in the collect job of `Traffic history`. There, one step
turns the private key of my GitHub App, a secret of the `analytics` environment
restricted to the default branch, into a short-lived installation token with
only *Administration: read* on this repository. Only the fetch step receives
that token, and it runs `curl` and no repository code
([traffic history](docs/TRAFFIC.md#credential-setup-and-renewal)).

| Workflow | Job token permissions | Runs on |
| --- | --- | --- |
| `Repository integrity` | `contents: read` | Pull requests into and pushes to `main` and `release/**`; manual dispatch |
| `Issue triage` | `contents: read`, `issues: read`, `pull-requests: read` | Issue and milestone events, from the default branch; manual dispatch |
| `Detect issue-label drift` | `contents: read`; `issues: read` for the live comparison | Label-file pull requests; daily schedule; manual dispatch |
| `Sync issue labels` | `contents: read`; `issues: write` and `pull-requests: read` only in the reconcile job | Label-file pull requests (validation only); label-file pushes to `main` and manual dispatch on the default branch (reconciliation) |
| `Traffic history` | `contents: read`; `contents: write` in `collect` (default branch only, `analytics` environment); `issues: write` in `alert`, which never receives the token | Daily schedule; manual dispatch (collection runs only from the default branch) |
| `CodeQL` | `contents: read` on PRs; `security-events: write` only in the trusted-event job | Pull requests into and pushes to `main` and `release/**`; weekly schedule; manual dispatch |
| `External links` | `contents: read` | Weekly schedule; manual dispatch |
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

Pull requests run a separate read-only CodeQL job with result and database
uploads disabled. Each language's SARIF is retained as an artifact for 30 days;
PRs receive no code scanning annotations. Pushes, scheduled runs and manual
runs publish results with `security-events: write`.

CodeQL code scanning runs from the `CodeQL` workflow
([.github/workflows/codeql.yml](.github/workflows/codeql.yml)), which replaces
GitHub's default setup (#146). One job per language analyzes the workflows
(`actions`), the Python tools (`python`) and the label scripts in
`.github/scripts/` (`javascript-typescript`) with the `security-extended`
queries, and uploads the results to code scanning. Its actions are pinned to a
full SHA like the others. The default setup must stay off: while it is on,
code scanning rejects the results of a workflow. The VBA is not analyzed,
because CodeQL has no extractor for it.

<a id="existing-controls"></a>
<a id="security-features"></a>

## 🛡️ Security features

Recorded on **2026-10-28** for the public repository and updated on
**2026-10-29** (#75), **2026-11-05** (#115), **2026-11-07** (#146) and **2026-11-09** (#15). The source column says
how each state is known; a state marked *not verified* has not been read back
and must not be assumed.

| Feature | State | Source |
| --- | --- | --- |
| Rulesets on `main` and `release/**` | Active, as described in [governance](docs/GOVERNANCE.md#installed-settings-and-limits) | GitHub API read-back, 2026-10-29 and 2026-11-09 ([settings evidence](docs/REPOSITORY_SETTINGS_EVIDENCE.md#rulesets)); enforcement demonstrated on #172 ([enforcement](docs/REPOSITORY_SETTINGS_EVIDENCE.md#enforcement)) |
| Default workflow token | Read-only | Verified by its effect on 2026-11-09: a job without declared permissions received read access only ([enforcement](docs/REPOSITORY_SETTINGS_EVIDENCE.md#enforcement)); each workflow also declares `contents: read` at workflow level, or no permission at all (`Keep milestones issue-only`) |
| Full-SHA action pins | Required by the repository setting | Verified by its effect on 2026-11-09: an action used by tag was refused ([enforcement](docs/REPOSITORY_SETTINGS_EVIDENCE.md#enforcement)); every `uses:` in the workflow files is a full SHA |
| Dependabot version updates | GitHub Actions: configured and active. Python tools in `tools/`: configured (#85), read from `main` since 2026-11-03 and active; the first update, #109, was integrated into `main` on 2026-11-04. Proposals arrive as pull requests for review | `.github/dependabot.yml`; Actions activation evidenced in #6 by Dependabot PR #41; Python activation evidenced by Dependabot PR #109; Python dependency graph built on 2026-11-03 ([run 41699792588](https://github.com/danielep71/K-NSFR/actions/runs/41699792588)) |
| Private vulnerability reporting | Enabled | GitHub API read-back, 2026-10-28; **Settings → Advanced Security**, read by me on 2026-11-09 |
| Secret scanning | Enabled | GitHub API read-back, 2026-10-28 |
| Secret scanning push protection | Enabled | GitHub API read-back, 2026-10-28 |
| Secret scanning non-provider patterns and validity checks | Disabled | GitHub API read-back, 2026-10-28 |
| Dependabot security updates | Enabled | GitHub API read-back, 2026-10-28 |
| Dependency graph | On | **Settings → Advanced Security**, read by me on 2026-11-09 |
| Automatic dependency submission | Disabled | **Settings → Advanced Security**, read by me on 2026-11-09 |
| Dependabot alerts | On | **Settings → Advanced Security**, read by me on 2026-11-09 (#15); the API does not report it with the access used for the other rows |
| Dependabot malware alerts | On | Off when read on 2026-11-09; I turned them on the same day (**Settings → Advanced Security**) |
| Dependabot auto-triage rules | One rule enabled | **Settings → Advanced Security**, read by me on 2026-11-09 |
| Code scanning | CodeQL from the `CodeQL` workflow (#146): Actions, Python and JavaScript/TypeScript, `security-extended` queries, on pull requests, pushes and weekly. It covers the workflows, the Python tools and the label scripts, not the VBA. The default setup, enabled on 2026-10-29 for Actions and Python with the default queries, is off | The workflow file and its first run, on #153 ([run 42356390686](https://github.com/danielep71/K-NSFR/actions/runs/42356390686)). The first attempt was rejected while the default setup was on. I turned the default setup off on 2026-11-07, and the second attempt uploaded all three analyses: 7 workflow files, 21 Python files and the 3 label scripts. The default setup's first analysis, on `main` at `d272338`, succeeded ([run 40838760015](https://github.com/danielep71/K-NSFR/actions/runs/40838760015)). Alerts are not readable without security-events access; *not verified* |

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
