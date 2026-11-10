<a id="provenance-inventory"></a>
<a id="provenance-and-attribution-inventory"></a>

# 🔗 Provenance inventory

**K-NSFR** · [Overview](../README.md) · [Documentation](../README.md#documentation) · [Governance](GOVERNANCE.md)

[![Stage](https://img.shields.io/badge/Stage-Repository_setup-BF8700?style=flat-square)](../README.md#status)

<a id="review-baseline"></a>

## 📌 Review baseline

Baseline reviewed for issue #3: K-NSFR release branch after issue #2, at
`68263b950d5fda7ff31e9ac7780f15390fce7190`. The initial governance
import is recorded in repository history at `926e49f5de0221953a5147b562631c0081f7ae06`;
its source and revision are in the
[third-party notices](licenses/THIRD_PARTY_NOTICES.txt), which also record the
license of every copied or adapted file. Imported text was adapted
for this repository, with unrelated repository references removed. This inventory
records origins at the repository level; it does not establish transferable
copyright ownership or permission to redistribute.

<a id="tracked-surfaces"></a>

## 🗂️ Tracked surfaces

| Surface | Origin and current status | Obligation review |
| --- | --- | --- |
| Label catalogue and sync/drift scripts | Imported (see the third-party notices), then K-NSFR-specific edits; `.github/scripts/markdown.mjs`, the summary helper they share, authored locally in #145 | Source, revision and license per file in the third-party notices |
| Issue templates and triage Python code/tests | Imported (see the third-party notices), adapted for the K-NSFR issue model | Source, revision and license per file in the third-party notices |
| GitHub Actions workflows and Dependabot configuration | Imported (see the third-party notices), adapted to local checks; `.github/workflows/codeql.yml` authored locally in #146 | License per file in the third-party notices; actions use SHA-pinned upstream components, so inventory each action's license and notice before redistributing bundled code |
| README and GOVERNANCE.md | Initial baseline, then rewritten for K-NSFR | Repository history records changes; I confirm notices if external distribution is proposed |
| SCOPE.md, ARCHITECTURE.md, DECISIONS.md, BACKLOG.md | Authored for issue #2 in this repository | No external code or template was adopted in these files |
| Workbook, regulatory parameters | None admitted | Future source/version, licensing and independent validation required |
| Numerical fixtures | One neutral synthetic fixture, FX-0001 (#10) | Synthetic data and factors; no regulatory claim |
| NSFR benchmark prototype (`prototype/`) | Authored by me; incorporated as delivered in #130, with only the specification's relative links changed and `templates/README.md` added; Git normalized the `.bas` line ends; not admitted (ADR-011) | Its factors are transcribed from the CRR and not verified against the source register; its demo data are synthetic; its tools need `xlsxwriter`, `openpyxl` and LibreOffice, which are not locked |

<a id="external-tools-and-future-dependencies"></a>

## 🧩 External tools and future dependencies

GitHub-hosted actions are fetched at pinned revisions and are not vendored in
this tree. No third-party application package manifest or binary is present.
A future dependency entry must record its exact version, upstream URL, license,
notices, and use/redistribution decision. Regulatory text is source material to
cite and interpret, not an assumed software license. A retained copy of a
regulatory source follows the byte-for-byte
[convention](SOURCE_CONTROL.md#regulatory-source-copies), with its source,
version and SHA-256 recorded beside it.

<a id="confidentiality-review"></a>

## 🔍 Confidentiality review

A repository-tree review at this candidate found no workbook, client data,
credentials or proprietary template in tracked paths. The tracked text and
issue descriptions use synthetic planning examples. This is a scoped inspection,
not a claim about untracked local files, inaccessible prior private material,
or every historical blob. If questionable content is found, report privately
under SECURITY.md without copying it into issue evidence.

<a id="maintenance-inventory"></a>

## 🔄 Maintenance inventory

The baseline above is historical. The
[third-party notices](licenses/THIRD_PARTY_NOTICES.txt) record, for every
copied or adapted file, the source revision, local import evidence,
adaptations and license, including the initial governance import. Source
URLs there identify origins; they are not project dependencies or product
claims. The publication review is
recorded in #19.

| Added surface | Origin / control |
| --- | --- |
| Source-control checker and fixtures | Authored locally in #4; isolation correction and regression fixture are local maintenance work |
| Markdown/YAML checker | Authored locally in #6; parsed navigation and explicit-anchor validation added during maintenance |
| Contracts, contract checker and synthetic examples | Authored locally in #9; the examples are synthetic |
| Validation rules, fixture checker and neutral fixture | Authored locally in #10; the fixture's data and factors are synthetic and its derivation was worked out by hand |
| VBA static checker, exceptions file and fixtures | Authored locally in #11 |
| Neutral VBA test harness and its self-test | Authored locally in #12; neutral, with no NSFR calculation |
| Host evidence checker, its records and the host scripts | Authored locally in #14 (#157, #158, #174); the records come from my runs on the supported host (#175, #178) |
| External-link checker and its workflow | Authored locally in #147 |
| Code of conduct | Authored locally in #147 |
| Installation guide, release procedure and changelog | Authored locally in #18 and #19 |
| Tracked-file audit | Authored locally in #20; it lists every tracked file of its candidate with its origin |
| Workbook template and its generator | Adapted in #13 from my prototype's generator, without its factors, template rows, demo data, expected results or buttons; the template is generated, never edited by hand |
| Check runner and runner tests | Adapted from the immutable source in the third-party notices; current gates cover only delivered capabilities |
| Contributor/tooling instructions and numeric rationale | Authored for the current setup; sketches remain unexecuted |
| VBA house style | Adapted with local module, API, numeric and maturity decisions; no host capability imported by implication |
| Repository settings reference | Imported (see the third-party notices) and adapted to the local documentation style; its settings rows are unchanged |
| Repository settings evidence | Authored locally in #80 from the GitHub API read-back of 2026-10-29 |
| Social-preview image | Original image I made with a generative image tool for #17; uploaded in the repository settings only, not stored here |

<a id="runtime-tools"></a>

## 🧩 Runtime tools

Tools are downloaded, not vendored. Python wheel versions and SHA-256 hashes
are recorded in `tools/requirements-quality-ci.txt`; all transitive packages
are locked. Their licenses are listed in
[licensing](LICENSING.md#runtime-tool-licenses).

| Tool | Version identity | Purpose |
| --- | --- | --- |
| GitHub checkout, Python/Node setup, artifact upload actions | Full SHA in each workflow | CI execution |
| GitHub App token action (`actions/create-github-app-token`) | v3.3.0, commit `64ab729a6294000f5aa4a4bc0b73d1f32bc45510` | Short-lived token for the traffic collection |
| CodeQL action (`github/codeql-action`, init and analyze) | Full SHA in `codeql.yml` | Code scanning |
| actionlint | 1.7.12 plus archive SHA-256 in workflow | Workflow validation |
| Ruff | 0.17.2 plus the hashes of its published files | Python lint and complexity |
| mypy and runtime dependencies | 2.4.0 (ast-serialize 0.13.1, with the hashes of its published files); exact transitive pins/hashes in lock | Strict typing |
| PyYAML | 6.0.3 plus wheel hash | YAML parsing |
| types-PyYAML | 6.0.12.20260906 plus wheel hash | YAML type stubs |
| markdown-it-py / mdurl | 4.2.0 plus the hashes of its published files / 0.1.2 plus wheel hash | Markdown parsing and URL handling |
| XlsxWriter | 3.2.9 plus wheel hash | Workbook template generation |

[Licensing](LICENSING.md#runtime-tool-licenses) gives their licenses and
upstream sources.
Publication review must assess any future bundled distributions separately.
