<a id="repository-tooling"></a>

# 🧰 Repository tooling

**K-NSFR** · [Overview](../README.md) · [Documentation](../README.md#documentation)

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

| Gate | Evidence |
| --- | --- |
| Tool tests | Positive and negative fixtures for source, links, YAML, issues and runner behavior |
| Source control | Committed bytes (encoding, LF, trailing whitespace, final newline, blank lines at end of file) and exact binary allowlist; isolated committed Git attributes |
| Content | Parsed Markdown links/images, reference links, HTML links, explicit section anchors, fences and duplicate YAML keys |
| Labels and forms | Manifest validation, sync/drift self-tests and issue-form configuration |
| Ruff | Import/error rules and complexity ceiling 15; a probe proves 15 passes and 16 fails |
| mypy | Strict typing for every Python tool and test; annotated/unannotated probe |
| actionlint | Workflow syntax and expressions |

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

These checks neither compile VBA nor run Excel or an NSFR calculation.
VBA-specific static gates, neutral harness, packaging and actual host evidence
remain #11–#14. Live issue triage is read-only and separate from offline CI;
passing forms tests does not prove every live issue satisfies governance.

<a id="continuous-integration"></a>

## ☁️ Continuous integration

Five workflows run from `.github/workflows/`. In them every checkout sets
`persist-credentials: false`, every action is pinned to a full commit SHA, and
the default token is read-only. Only label reconciliation and the milestone
workflow can write. CodeQL
code scanning also runs, as GitHub's default setup with no workflow file here;
its [risk surface](../SECURITY.md#automation-risk-surface) is described in the
security policy.

| Workflow | Triggers | Runs | Writes |
| --- | --- | --- | --- |
| Repository integrity | Every PR into and push to `main` and `release/**`; manual dispatch. No path filter, so it can be a required check. | `python tools/check.py --ci` | Nothing; uploads `test-results/` |
| Issue triage | Issue events (run from the default branch); manual dispatch | `tools/check_issues.py` against the live issues | Nothing; uploads its report |
| Detect issue-label drift | PRs into `main`/`release/**` that touch label files; daily at 05:17 UTC; manual dispatch | Drift self-tests on PRs; live comparison otherwise | Nothing; uploads its evidence |
| Sync issue labels | PRs that touch label files (validation only); label-file pushes to `main` and manual dispatch on the default branch (reconciliation) | Manifest validation and self-tests; live reconciliation | Issue labels, only in the reconcile job (`issues: write`) |
| Keep milestones issue-only | `pull_request_target` (opened, reopened, edited, milestoned); daily at 03:17 UTC; its own pushes to `main`; manual dispatch on the default branch | Its inline script only; never a pull request's contents | Removes PR milestones (`pull-requests: write`) |

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
assignee, exactly one priority (P1 only with `bug`), an open milestone of the
active release, a valid ISO target end date before the milestone's due date,
catalogue labels, and the dependency section against the `blocked` label. Pull
request numbers are loaded so that a dependency section may name one as a
reference; a pull request never counts as a prerequisite, so the workflow
token also needs `pull-requests: read`. Issue events run the workflow from the
default branch, so a rule change takes effect after integration into `main`.

```sh
GITHUB_TOKEN=<token with issues:read and pull-requests:read> python tools/check_issues.py --repository danielep71/K-NSFR
python tools/check_issues.py --input issues.json   # issues and pull requests exported as GitHub REST JSON
python tools/check_issues.py --check-config        # issue forms and label catalogue only, offline
```

Exit code 0 means no findings, 1 means findings, 2 means the audit could not
run. The rules are tested by `test_check_issues.py` in the tool-tests gate, and
`--check-config` runs in Repository integrity.

<a id="tool-versions"></a>

## 🔧 Tool versions

| Tool | Version | Pinned in |
| --- | --- | --- |
| Python | 3.13 | `actions/setup-python` in the workflows; `pyproject.toml` targets |
| Node.js | 22 | `actions/setup-node` in the workflows |
| Ruff, mypy | 0.16.6, 2.3.1 | `requirements-quality-ci.txt`, hash-locked for CPython 3.13 on Linux x86_64 |
| PyYAML, markdown-it-py | 6.0.3, 4.0.0 | `requirements-quality-ci.txt` (content gate) |
| actionlint | 1.7.12 | `ACTIONLINT_VERSION` and its SHA-256 in `repository-integrity.yml` |

Update a version only as [dependency updates](#dependency-updates) describes,
and change every place that pins it in the same pull request.

<a id="dependency-updates"></a>

## 📦 Dependency updates

Quality/content tools are hash-locked in `requirements-quality-ci.txt`.
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
release. Each update is a pull request that the owner reviews before merge
(for example #41, reviewed in #42); a Python update must still pass the hash and
provenance review above. Dependabot reads the configuration from `main`, so a
change takes effect once it is integrated there. When a new release branch
opens, update both `target-branch` values. Milestones stay on issues only.
