<a id="tracked-file-audit"></a>

# 🔍 Tracked-file audit

**K-NSFR** · [Overview](../README.md) · [Documentation](../README.md#documentation) · [Governance](GOVERNANCE.md)

[![Stage](https://img.shields.io/badge/Stage-Repository_setup-BF8700?style=flat-square)](../README.md#status)

This audit inventories every tracked file of one candidate: its purpose, its
origin and license, and why it stays. It records what the audit looked for,
what it found and how each finding was settled. Tracking: #20.

The audit is a snapshot. A later commit needs its own review; this record is
not updated file by file.

<a id="candidate-and-method"></a>

## 🎯 Candidate and method

- **Candidate:** `release/0.1.0` at `9bf5b41df99ce3d3a5a2696333061b8f0bcb606c`,
  120 tracked files, audited on 2026-11-10. This record and the corrections
  below come after it.
- **Inventory:** `git ls-files` at the candidate, with the commit that added
  each file (`git log --diff-filter=A`).
- **Origin and license:** every copied or adapted file was matched against
  its entry in the [third-party notices](licenses/THIRD_PARTY_NOTICES.txt).
  Every other file was added by a K-NSFR commit.
- **Searches over every tracked file** (`git grep -I`):
  - credentials: GitHub, AWS and Slack tokens, private keys, assignments of a
    password, secret or API key;
  - e-mail addresses and local user paths;
  - names of other projects;
  - assistant and model names;
  - placeholders: `TODO`, `TBD`, `FIXME`, `XXX`, `lorem`, `placeholder`;
  - ownership and confidentiality statements: `copyright`, `proprietary`,
    `confidential`, `all rights reserved`;
  - real-bank names and LEI-like codes in the data.
- **Binaries and generated files:** Git's binary detection over every
  tracked file, the binary allowlist, and the ignore rules for workbooks and
  reports.

<a id="inventory"></a>

## 🗂️ Inventory

**Origin** is *authored* (written for K-NSFR in the PR named), *adapted*
(copied from another project and changed) or *copied* (unchanged apart from a
header). An authored file is under the repository's license
([licensing](LICENSING.md)); a copied or adapted file keeps the source,
revision and license that the [third-party notices](licenses/THIRD_PARTY_NOTICES.txt)
record for it.

<a id="root-and-hidden-configuration"></a>

### Root and hidden configuration

| Path | Purpose | Origin | Retain because |
| --- | --- | --- | --- |
| `.editorconfig` | Editor encoding, line ends and indentation | Adapted (#34) | Keeps edits consistent with the committed-text rules |
| `.gitattributes` | Line-end, binary and diff attributes; VBA checked out as CR LF | Adapted (#34) | The source-control conventions depend on it |
| `.gitignore` | Ignores workbooks, reports, secrets and keys | Adapted (#34) | Keeps generated and secret files out of Git |
| `AGENTS.md` | Editorial and assistant rules | Authored (#100) | The one place for those rules |
| `CHANGELOG.md` | What each version adds | Authored (#182) | Required by the release procedure |
| `CODE_OF_CONDUCT.md` | Expected behavior and how to report abuse | Authored (#154) | Contributor policy |
| `CONTRIBUTING.md` | How to make a bounded change | Authored (#44) | Contributor onboarding |
| `LICENSE` | MPL-2.0 text | Standard license text (#32) | The repository's license |
| `README.md` | Purpose, status, quick start, layout, documentation map | Authored (initial commit, rewritten since) | Entry point |
| `SECURITY.md` | Vulnerability reporting and security features | Authored (#30) | Security policy |
| `pyproject.toml` | Ruff and mypy configuration for `tools/` | Adapted (#44) | Configures the lint and typing gates |

<a id="github-configuration"></a>

### GitHub configuration

| Path | Purpose | Origin | Retain because |
| --- | --- | --- | --- |
| `.github/ISSUE_TEMPLATE/bug.yml`, `documentation.yml`, `enhancement.yml` | Issue forms with priority, dates, criteria and dependencies | Adapted (initial import) | Triage reads their fields |
| `.github/ISSUE_TEMPLATE/config.yml` | Disables blank issues; links the private security report | Adapted (initial import) | Routes vulnerability reports privately |
| `.github/PULL_REQUEST_TEMPLATE.md` | Evidence sections of every PR | Authored (#43) | Review evidence |
| `.github/dependabot.yml` | Update proposals for actions and Python tools, against the release branch | Adapted (initial import) | Keeps pinned tools current |
| `.github/labels.json` | Label catalogue, schema 2 | Adapted (#84) | Source of the live labels |
| `.github/scripts/labels-sync.mjs`, `labels-drift.mjs` | Synchronize the labels and report drift, with self-tests | Adapted (#84) | Run by the label workflows |
| `.github/scripts/markdown.mjs` | Renders API values safely in job summaries | Authored (#152) | Used by both label scripts |
| `.github/workflows/repository-integrity.yml` | The required check | Adapted (initial import) | Required by the rulesets |
| `.github/workflows/codeql.yml` | CodeQL analysis | Authored (#153) | Code scanning |
| `.github/workflows/external-links.yml` | Weekly external-link check | Authored (#154) | Finds broken external links |
| `.github/workflows/issue-triage.yml` | Audits live issues | Adapted (initial import) | Enforces the issue rules |
| `.github/workflows/labels-sync.yml`, `labels-drift.yml` | Label synchronization and drift report | Adapted (initial import) | Keep live labels on the catalogue |
| `.github/workflows/pr-milestones.yml` | Removes milestones from PRs | Copied (#92); its license is in its header and the notices | Milestones belong to issues |
| `.github/workflows/traffic-history.yml` | Daily traffic collection and alert | Adapted (#104) | Accepted traffic policy (#16) |

<a id="documentation"></a>

### Documentation

| Path | Purpose | Origin | Retain because |
| --- | --- | --- | --- |
| `docs/ARCHITECTURE.md`, `docs/SCOPE.md`, `docs/DECISIONS.md`, `docs/BACKLOG.md` | Scope, layout, decisions and later work | Authored (#27) | Baseline of the project |
| `docs/DATA_POLICY.md`, `docs/LICENSING.md`, `docs/PROVENANCE.md` | Synthetic data, license and origins | Authored (#30) | Data and rights policy |
| `docs/GOVERNANCE.md` | Issues, labels, branches, review | Adapted (initial import), rewritten since | Process rules |
| `docs/IMPLEMENTATION.md` | API, numeric and host conventions | Authored (#35) | Implementation contract |
| `docs/INSTALLATION.md` | Portable checks and host setup | Authored (#181) | Installation guide |
| `docs/RELEASE.md` | Versions, release, tag, milestone closure | Authored (#182) | Release procedure |
| `docs/FILE_AUDIT.md` | This audit | Authored (#20) | Audit record |
| `docs/REPOSITORY_SETTINGS.md` | Non-default settings | Adapted (#78) | Settings reference |
| `docs/REPOSITORY_SETTINGS_EVIDENCE.md` | Dated settings read-back and enforcement | Authored (#78) | Settings evidence |
| `docs/SOURCE_CONTROL.md` | Text, encoding and generated-file rules | Authored (#34) | Source conventions |
| `docs/TRAFFIC.md` | Traffic collection, credential and alert | Adapted (#104) | Operating guide |
| `docs/VBA_HOUSE_STYLE.md` | VBA conventions | Adapted (#44) | Implementation review |
| `docs/WRITING_STYLE.md` | Page structure, badges, glyphs | Authored (#38) | Documentation style |
| `docs/contracts/CONTRACTS.md`, `docs/contracts/nsfr-contract-v1.json` | Input, parameter and result contracts and their schema | Authored (#125) | Contract of the future engine |
| `docs/contracts/examples/valid-basic/`, `invalid-positions/`, `invalid-run-basis/`, `invalid-run-rows/`: `positions.csv`, `run.csv`, `expected.json` in each (12 files) | Synthetic contract examples | Authored (#125) | The `contracts` gate validates them |
| `docs/licenses/THIRD_PARTY_NOTICES.txt` | Attribution of every copied or adapted file | Authored (#92) | Required by the licenses |
| `docs/methodology/SOURCE_REGISTER.md` | Regulatory source register | Authored (#123) | Only route to cite a rule |
| `docs/validation/VALIDATION.md`, `docs/validation/fixture-v1.json` | Validation rules and fixture schema | Authored (#128) | Numerical validation design |
| `docs/validation/HOST_EVIDENCE.md` | Host evidence record format | Authored (#157) | Host evidence rules |

<a id="tests"></a>

### Tests and evidence

| Path | Purpose | Origin | Retain because |
| --- | --- | --- | --- |
| `tests/README.md` | Harness guide | Authored (#129) | How to write and run suites |
| `tests/core/Test_NSFR_Harness.bas`, `Test_NSFR_HarnessSelfTest.bas` | Neutral VBA harness and its self-test | Authored (#129) | Host test runner |
| `tests/evidence/README.md` | Index of host records | Authored (#157) | Evidence index |
| `tests/evidence/523198552eee8296b1dfa4e72344c2a8e5d55cfd.json` | Host record of `5231985`, PASS | Authored (#175) | Host evidence (#12, #14) |
| `tests/evidence/dec4ad555db1240f8c2f0361c5de47e36f059528.json` | Host record of `dec4ad5`, PASS | Authored (#178) | Host evidence of the branch verifier |
| `tests/fixtures/neutral-synthetic/fixture.json`, `parameters.json`, `positions.csv`, `run.csv` | Neutral synthetic fixture | Authored (#128) | The `fixtures` gate re-derives it |

<a id="tools"></a>

### Tools

| Path | Purpose | Origin | Retain because |
| --- | --- | --- | --- |
| `tools/README.md` | Gates, CI, versions and updates | Authored (#44) | Tooling guide |
| `tools/check.py`, `tools/test_check.py` | Gate runner and its tests | Adapted (#44) | The required check runs it |
| `tools/requirements-quality-ci.txt` | Hash-locked Python tools | Adapted (#44) | Reproducible gates |
| `tools/check_content.py`, `test_check_content.py` | Markdown and YAML gate | Authored (#36) | `content` gate |
| `tools/check_contracts.py`, `test_check_contracts.py` | Contract gate | Authored (#125) | `contracts` gate |
| `tools/check_fixtures.py`, `test_check_fixtures.py` | Fixture gate | Authored (#128) | `fixtures` gate |
| `tools/check_host_evidence.py`, `test_check_host_evidence.py` | Host evidence gate | Authored (#157) | `host-evidence` gate |
| `tools/check_issues.py`, `test_check_issues.py` | Issue triage rules | Adapted (initial import, #90) | Run by issue triage |
| `tools/check_links.py`, `test_check_links.py` | External-link check | Authored (#154) | Run by the external-links workflow |
| `tools/check_source_control.py`, `test_source_control.py` | Committed-blob gate | Authored (#34) | `source-control` gate |
| `tools/check_vba.py`, `test_check_vba.py` | VBA static rules | Authored (#124) | `vba-static` gate |
| `tools/check_workflows.py`, `test_check_workflows.py` | Workflow and Dependabot invariants | Adapted (#94) | `workflow-invariants` gate |
| `tools/traffic_history.py`, `test_traffic_history.py` | Traffic history merge and alert | Adapted (#104) | Run by the traffic workflow |
| `tools/binary-allowlist.json` | Exact paths of admitted binaries; empty | Authored (#34) | Read by the `source-control` gate |
| `tools/vba-exceptions.json` | Reviewed VBA rule exceptions; empty | Authored (#124) | Read by the `vba-static` gate |
| `tools/host/Record-HostEvidence.ps1` | Records a host run of the checked-out commit | Authored (#158) | Host evidence |
| `tools/host/Verify-NSFR-Branch.ps1` | Verifies a branch in one command | Authored (#174) | Host evidence |

<a id="prototype"></a>

### Prototype, not admitted

All 17 files are mine, incorporated by #131 as delivered, under ADR-011. They
stay until each piece is admitted or replaced by the v0.2.0 issues (#132 to
#138).

| Path | Purpose |
| --- | --- |
| `prototype/README.md` | What the directory holds and how the repository treats it |
| `prototype/docs/methodology/PROTOTYPE.md` | Specification |
| `prototype/src/api/NSFR_API_Functions.bas` | API module |
| `prototype/src/core/NSFR_Core_Config.bas`, `NSFR_Core_Engine.bas`, `NSFR_Core_Import.bas`, `NSFR_Core_Reconcile.bas`, `NSFR_Core_Util.bas` | Core modules |
| `prototype/src/host/NSFR_Host_Main.bas` | Host module |
| `prototype/templates/README.md` | Keeps the generator's output directory |
| `prototype/tests/derivations/prototype/regulatory-tables.json` | Transcribed tables, not verified (#132) |
| `prototype/tests/derivations/prototype/demo-positions.json`, `demo-bank-cells.json`, `demo.expected.json` | Synthetic demo, its bank figures and expected values |
| `prototype/tests/derivations/prototype/nsfr_reference.py` | Independent Python reference |
| `prototype/tools/build_template.py` | Workbook template generator; basis of #13 |
| `prototype/tools/lo_regression.py` | LibreOffice regression |

<a id="findings-and-dispositions"></a>

## 🧾 Findings and dispositions

| ID | Finding | Disposition |
| --- | --- | --- |
| A-01 | No credential, private key or secret assignment in any tracked file | None needed |
| A-02 | E-mail addresses only in tests: `example.invalid`, `example.org`, `example.com` and the `git@github.com` URL form | Retained: synthetic inputs of the checks' tests |
| A-03 | Local user paths only in `tools/test_check_host_evidence.py`, as inputs the private-details rule must reject | Retained: the test needs them |
| A-04 | Other projects are named in the third-party notices; the provenance record also linked the upstream repositories of the runtime tools | Notices retained: [AGENTS.md](../AGENTS.md#what-the-documentation-describes) allows attributions there. The upstream links moved to [licensing](LICENSING.md#runtime-tool-licenses) in #21 |
| A-05 | No assistant or model name in text; `.gitignore` ignores three assistant tools' local state directories by name | Retained: ignore patterns are configuration, not documentation |
| A-06 | Every `placeholder` is an example value in an issue form | Retained: form syntax, not a dead placeholder |
| A-07 | No unsupported ownership statement: the only copyright line is in the third-party notices | None needed |
| A-08 | No binary file; the binary allowlist is empty; workbooks, `test-results/`, `.env` files, keys and keystores are ignored | None needed |
| A-09 | Every data file is synthetic: the contract examples, the neutral fixture and the prototype's demo, whose bank figures are generated by its reference script. No real-bank name or LEI-like code | None needed; the `contracts` and `fixtures` gates validate the first two |
| A-10 | No stale or generated file: `prototype/templates/README.md` only keeps a directory, as its README says | Retained |
| A-11 | The notices said the traffic files were "adapted by #167"; the changes came from #168 and #170, and the start-date extension of `tools/check_issues.py` by #168 was not recorded | Corrected in the notices |
| A-12 | The provenance inventory did not list the surfaces added since #12: the host evidence checker and records, the host scripts, the external-link check, the code of conduct, the installation guide, the release procedure, the changelog and this audit | Corrected in [provenance](PROVENANCE.md#maintenance-inventory) |
| A-13 | Documentation statements that later work superseded, found while reading the inventory (for example settings described as not read back) | Passed to #21, the documentation audit |

No file was deleted: every file has a purpose and an owner.

<a id="limitations"></a>

## ⚠️ Limitations

- The searches are text patterns. They find known credential formats and
  names, not every possible secret or confidential fact.
- Git history, issues, pull requests and Actions logs are outside this audit;
  the post-publication review of #19 covered them on 2026-10-28.
- The audit binds one candidate. Files added after it are not covered.
