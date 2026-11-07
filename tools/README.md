<a id="repository-tooling"></a>

# 🧰 Repository tooling

**K-NSFR** · [Overview](../README.md) · [Documentation](../README.md#documentation) · [Governance](../docs/GOVERNANCE.md)

[![Stage](https://img.shields.io/badge/Stage-Repository_setup-BF8700?style=flat-square)](../README.md#status)

<a id="running-the-checks"></a>

## ▶️ Running the checks

Use a Git clone, Python 3.13, Node 22 and Git on Linux x86_64 for the reference
portable environment. A source archive deliberately omits repository controls.
Create a virtual environment outside the checkout, activate it, then install:

```sh
python -m pip install --only-binary=:all: --require-hashes -r tools/requirements-quality-ci.txt
```

Install actionlint 1.7.12 on PATH. Verify its Linux amd64 archive against
SHA-256 `8aca8db96f1b94770f1b0d72b6dddcb1ebb8123cb3712530b08cc387b349a3d8`,
as the integrity workflow does. Commit the candidate before validating committed
content, then run the same entry point used in CI:

```sh
python tools/check.py --ci
```

Other platforms can run compatible tools locally, but the wheel lock targets
CPython 3.13 on Linux x86_64. CI remains the reference environment. A diagnostic
`python tools/check.py --allow-missing` run may report **INCOMPLETE**, never PASS,
if a tool is unavailable. It is not merge evidence. `--ci` rejects a dirty tree
and missing tools. No gate is skipped after an earlier gate fails.

<a id="gates-and-reports"></a>

## ✅ Gates and reports

The gates run in this order; each needs the tool or libraries shown, and is
reported **NOT RUN** without them.

| Gate | Needs | Evidence |
| --- | --- | --- |
| `tool-tests` | PyYAML, markdown-it-py, Node | Positive and negative fixtures for the runner, source control, content, contracts, numerical fixtures, issues, workflow invariants, traffic history and the VBA static checks |
| `labels-sync-self-test` | Node | Label synchronization self-tests |
| `labels-manifest` | Node | Label catalogue validation |
| `labels-drift-self-test` | Node | Label drift self-tests |
| `issue-config` | Node | Issue forms, label guide and Dependabot labels against the issue rules (`check_issues.py --check-config`) |
| `source-control` | Python | Committed bytes (encoding, LF, trailing whitespace outside Markdown and `LICENSE`, final newline, blank lines at end of file) and exact binary allowlist; isolated committed Git attributes |
| `content` | PyYAML, markdown-it-py | Parsed Markdown links/images, reference links, HTML links, explicit section anchors, fences and duplicate YAML keys |
| `contracts` | Python | `check_contracts.py`: every field and diagnostic of [the contracts](../docs/contracts/CONTRACTS.md) documented once and matching `nsfr-contract-v1.json`, and every synthetic example giving its `expected.json` |
| `fixtures` | Python | `check_fixtures.py`: every [numerical fixture](../docs/validation/VALIDATION.md) under `tests/fixtures/` against `fixture-v1.json`, its hand derivation re-derived and its expected result re-evaluated with exact rational arithmetic; it never runs the engine |
| `vba-static` | Python | `check_vba.py`: the [VBA static rules](#vba-static-checks) on every tracked `.bas`, `.cls` and `.frm` file, with reviewed exceptions from `vba-exceptions.json` |
| `ruff` | Ruff | Import/error rules and complexity ceiling 15 |
| `ruff-complexity-probe` | Ruff | Complexity 15 passes and 16 fails |
| `mypy` | mypy, PyYAML, markdown-it-py | Strict typing for every Python tool and test |
| `mypy-strict-probe` | mypy | An unannotated function is rejected |
| `actionlint` | actionlint | Workflow syntax and expressions |
| `workflow-invariants` | PyYAML | `check_workflows.py`: full-SHA action pins and `persist-credentials: false`; declared, read-only workflow permissions and a timeout on every job; write access only for the listed jobs; the required check's job, conditions and branch filters against governance and the rulesets; the required check's concurrency (only a pull request's superseded run is cancelled); the label-reconcile condition and permissions; issue triage on milestone events; the milestone workflow's events, per-pull-request concurrency, condition, permissions and single inline script; the traffic workflow's App token (only *Administration: read* on this repository, the key in the token step only, the token in the fetch step only, a preflight that sees only whether the key is set) and the alert job's checkout of the published history commit; Dependabot coverage of `/` and `/tools` with exactly `P3` and no milestone |

On GitHub Actions a gate whose Python libraries are missing runs anyway and
fails, so the required check never skips one.

Internal links to Markdown sections must point to explicit `<a id="..."></a>`
anchors. This policy avoids reliance on renderer-specific heading slugs and
keeps links stable when headings or glyphs change. Remote URLs are not fetched;
remote availability and final visual rendering still require review. Fragments
inside non-Markdown assets are not validated.

The runner writes `test-results/repository-integrity.json`, its Markdown summary
and one log per gate. Reports identify the candidate SHA, dirty state and tool
versions. CI uploads them even after failure, retains them for 30 days, and names
the artifact with candidate SHA, run ID and attempt. Download evidence needed
for longer retention before expiry. Setup failures before the runner starts
produce no gate report and fail the upload rather than inventing evidence.

For a pull request, `github.sha` identifies GitHub's synthetic merge commit.
Record both PR head and tested merge SHA in review evidence; push runs test the
integrated commit. A run on one SHA is not evidence for an unrelated later SHA.

<a id="validation-boundaries"></a>

## 🔎 Validation boundaries

These checks neither compile VBA nor run Excel or an NSFR calculation. The
VBA static checks read source text only (see their
[limitations](#vba-static-limitations)). The [neutral harness](../tests/README.md)
(#12) is VBA and runs only in Excel; packaging and actual host evidence remain
#13 and #14. Live issue triage is read-only and separate from offline CI;
passing forms tests does not prove every live issue satisfies governance.

<a id="vba-static-checks"></a>

## 🧾 VBA static checks

`check_vba.py` applies the rules of the
[implementation conventions](../docs/IMPLEMENTATION.md#vba-source-and-environment),
the [VBA house style](../docs/VBA_HOUSE_STYLE.md) and ADR-008 to every tracked
`.bas`, `.cls` and `.frm` file. Each finding names the file, the line where the
statement starts and a rule:

```text
src/core/NSFR_Core_Example.bas:12: [goto] GoTo is allowed only as On Error GoTo
```

| Rule | `src/core/`, `src/api/` | `src/host/` | `tests/`, `examples/` |
| --- | --- | --- | --- |
| `ascii`, `option-explicit`, `vb-name` | Required | Required | Required |
| `option-private-module` (standard modules) | Required in `src/core/` | Not required | Required in `tests/` |
| `goto` (any GoTo but `On Error GoTo`), `gosub` (GoSub, Return), `on-error-resume-next` | Forbidden | Forbidden | Forbidden |
| `error-handler` (`On Error GoTo`, `Resume`) | Forbidden | Allowed | Allowed |
| `declare-ptrsafe` (Declare without PtrSafe) | Forbidden | Forbidden | Forbidden |
| `declare` (Declare statements) | Forbidden | Reviewed exception | Forbidden |
| `com-object` (`CreateObject`, `GetObject`, `New` on anything but a project class or `Collection`) | Forbidden | Reviewed exception | Forbidden |
| `dynamic-call` (`CallByName`, `Application.Run`, `Evaluate`, `ExecuteExcel4Macro`) | Forbidden | Reviewed exception | Forbidden |
| `conditional-compilation` (`#If`, `#ElseIf`) | Forbidden | Only on `VBA7` and `Win64` | Forbidden |
| `compile-constant` (`#Const`) | Forbidden | Forbidden | Forbidden |

A VBA file outside these folders is reported as `vba-location`. The one
exception is `prototype/`: it is not admitted ([ADR-011](../docs/DECISIONS.md#recorded-decisions)),
so its VBA is not checked and the summary line counts it as not checked. Every
conditional-compilation branch is checked, so a rule holds in each compile
environment, including branches the target never compiles.

<a id="vba-static-exceptions"></a>

### Reviewed exceptions

`vba-exceptions.json` lists the exceptions I approved. Each entry names a file
under `src/host/`, one of the rules `declare`, `com-object` or `dynamic-call`,
a text that the statement contains, the reason and the issue that reviewed
it:

```json
{"path": "src/host/NSFR_Host_Timer.bas", "rule": "declare",
 "contains": "GetTickCount", "reason": "...", "issue": 13}
```

No other path or rule can be excepted. An entry that matches no finding is
reported as `exception-unused`, so the file never keeps a stale exception.

<a id="vba-static-limitations"></a>

### Limitations

- The checks read text; they neither compile nor resolve types. A late-bound
  call through a variable (`Obj.Run`, `Sheet.Evaluate`) is not seen as a
  dynamic call, and `New` on an unqualified class is judged only against the
  `.cls` and `.frm` files checked in the same run.
- A name declared with the name of a built-in (a function called `Evaluate`,
  for example) is taken to be that declaration, so its unqualified uses are
  not reported: a procedure in every checked file, a variable, parameter or
  constant in the module that declares it, whatever its scope there. The
  qualified forms (`Application.Evaluate`, `VBA.CreateObject`) are always
  reported.
- A finding is reported on the first line of a statement continued with
  ` _`; a statement after a colon is reported on its line.
- Bracket evaluation such as `[A1]` is indistinguishable from an escaped
  identifier and is not checked.
- Rules that need the Excel object model or a compiled project, such as
  `Option Private Module` hiding a procedure or the absence of Excel calls in
  the core, need the neutral harness (#12) and host evidence (#13, #14).

<a id="continuous-integration"></a>

## ☁️ Continuous integration

Six workflows run from `.github/workflows/`, all on `main` since the
integration of 2026-11-03 (#107). Scheduled and issue-event runs use the copy
on `main`; the Traffic history schedule also requires the setup in
[traffic history](../docs/TRAFFIC.md). The milestone workflow also runs from
the PR base branch on pull-request events. In them every checkout sets
`persist-credentials: false`, every action is pinned to a full commit SHA, and
the default token is read-only. Only label reconciliation, the milestone
workflow and the two traffic jobs can write. CodeQL
code scanning also runs, as GitHub's default setup with no workflow file here;
its [risk surface](../SECURITY.md#automation-risk-surface) is described in the
security policy.

| Workflow | Triggers | Runs | Writes |
| --- | --- | --- | --- |
| Repository integrity | Every PR into and push to `main` and `release/**`; manual dispatch. No path filter, so it can be a required check. | `python tools/check.py --ci` | Nothing; uploads `test-results/` |
| Issue triage | Issue events and milestone created, edited, closed, opened or deleted (run from the default branch); manual dispatch | `tools/check_issues.py` against the live issues | Nothing; uploads its report |
| Detect issue-label drift | PRs into `main`/`release/**` that touch label files; daily at 05:17 UTC; manual dispatch | Drift self-tests on PRs; live comparison otherwise | Nothing; uploads its evidence |
| Sync issue labels | PRs that touch label files (validation only); label-file pushes to `main` and manual dispatch on the default branch (reconciliation) | Manifest validation and self-tests; live reconciliation | Issue labels, only in the reconcile job (`issues: write`; `pull-requests: read` for the prune check) |
| Traffic history | Daily at 06:17 UTC; manual dispatch (collection only on the default branch) | `tools/traffic_history.py` merges the traffic API responses; the alert job compares snapshots | The `traffic-history` branch (`collect`, `contents: write`); one alert issue (`alert`, `issues: write`). See [traffic history](../docs/TRAFFIC.md) |
| Keep milestones issue-only | `pull_request_target` (opened, reopened, edited, milestoned); daily at 03:17 UTC; its own pushes to `main`; manual dispatch on the default branch | Its inline script only; never a pull request's contents | Removes PR milestones (`pull-requests: write`); its job summary lists each pull request and the milestone title removed |

Concurrency: a new push to a pull request cancels that pull request's
superseded Repository integrity run, but every push and manual run has a
group of its own and is never cancelled. *Keep milestones issue-only* queues
each pull request's events separately from the full scans and cancels nothing.

Repository integrity checks out the exact commit under review and verifies that
`HEAD` is that commit; for a pull request this is GitHub's test merge of the
head into the base. Its Markdown report is added to the job summary and
`test-results/` is uploaded as `repository-integrity-<sha>-<run>-<attempt>`,
retained for 30 days. It downloads only the pinned actions, the locked Python
wheels and the checksum-verified actionlint archive.

<a id="issue-triage"></a>

## 🗂️ Issue triage

`check_issues.py` applies the [issue rules](../docs/GOVERNANCE.md#issues) to
every open issue and reports findings; it never changes an issue. It checks the
assignee configured in the Bug report form, exactly one priority (P1 only with `bug`), an open milestone of the
active release, a valid ISO target end date before the milestone's due date,
catalogue labels, and the dependency section against the `blocked` label. Pull
request numbers are loaded so that a dependency section may name one as a
reference; a pull request never counts as a prerequisite, so the workflow
token also needs `pull-requests: read`. Issue events, and milestone events (a
changed due date, or a closed or deleted milestone), run the workflow from the
default branch, so a rule change takes effect after integration into `main`.

```sh
GITHUB_TOKEN=<token with issues:read and pull-requests:read> python tools/check_issues.py --repository danielep71/K-NSFR
python tools/check_issues.py --input issues.json --repo-name danielep71/K-NSFR   # exported GitHub REST JSON
python tools/check_issues.py --check-config        # configuration only, offline
```

A repository is always given as `owner/name`; `--input` requires
`--repo-name`. `--check-config` validates the label catalogue with
`.github/scripts/labels-sync.mjs`, the same validator synchronization uses, so
it needs Node. It also checks that every issue form sets its default priority,
known labels, the assignee and the `dependencies` and `target_end_date` fields;
that the label tables of [governance](../docs/GOVERNANCE.md#issues) explain
each catalogue label exactly once; that `blocked` is the only dependency
status label; that `ISSUE_TEMPLATE/config.yml` disables blank issues; and that
`.github/dependabot.yml` uses only catalogue labels and the same assignee.

Exit code 0 means no findings, 1 means findings, 2 means the audit could not
run. The rules are tested by `test_check_issues.py` in the tool-tests gate, and
`--check-config` runs in Repository integrity.

<a id="tool-versions"></a>

## 🔧 Tool versions

| Tool | Version | Pinned in |
| --- | --- | --- |
| Python | 3.13 | `actions/setup-python` in the workflows, except *Keep milestones issue-only*, whose single inline step runs the runner's `python3`; `pyproject.toml` targets |
| Node.js | 22 | `actions/setup-node` in the workflows |
| Ruff, mypy | 0.17.1, 2.4.0 | `requirements-quality-ci.txt`, hash-locked and installed on CPython 3.13 / Linux x86_64 |
| PyYAML, markdown-it-py | 6.0.3, 4.2.0 | `requirements-quality-ci.txt` (the tool-tests, content, mypy and workflow-invariants gates) |
| actionlint | 1.7.12 | `ACTIONLINT_VERSION` and its SHA-256 in `repository-integrity.yml` |

Update a version only as [dependency updates](#dependency-updates) describes,
and change every place that pins it in the same pull request.

<a id="dependency-updates"></a>

## 📦 Dependency updates

Quality/content tools are hash-locked in `requirements-quality-ci.txt`. A
package that Dependabot has updated lists the hashes of every file published
for its version, sdist included; the others list only the CPython 3.13 Linux
x86_64 wheel. Either way CI installs wheels only (`--only-binary=:all:`), and
each hash must match a file published on PyPI for that version.
Review upstream release notes and licensing, download the intended wheels,
verify their hashes, update the lock and provenance together, and rerun all
gates. Do not remove hash enforcement to admit an update. The narrow mypy
configuration applies to every Python tool and test without blanket exclusions.
Actionlint has its own pinned version/checksum in the workflow.

`.github/dependabot.yml` proposes weekly updates, each Monday at 05:00
Europe/Rome, for the pinned GitHub Actions and for the Python requirements in
`/tools`. The two are grouped separately, as `github-actions` and
`python-tools`, and target the active release branch with the `ci` and `P3`
labels and danielep71 as assignee. Python updates wait 14 days after an upstream
release. Each update is a pull request that I review before merge
(for example #41, reviewed in #42); a Python update must still pass the hash and
provenance review above. Dependabot reads the configuration from `main`, so a
change takes effect once it is integrated there. When a new release branch
opens, update both `target-branch` values. Milestones stay on issues only.
