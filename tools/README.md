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

<a id="dependency-updates"></a>

## 📦 Dependency updates

Quality/content tools are hash-locked in `requirements-quality-ci.txt`.
Review upstream release notes and licensing, download the intended wheels,
verify their hashes, update the lock and provenance together, and rerun all
gates. Do not remove hash enforcement to admit an update. The narrow mypy
configuration applies to every Python tool and test without blanket exclusions.
Actionlint has its own pinned version/checksum in the workflow. Dependabot
handles GitHub Actions pins only; #42 separately reviews the pending Node setup
action major-version update.
