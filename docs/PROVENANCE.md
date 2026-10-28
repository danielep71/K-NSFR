<a id="provenance-and-attribution-inventory"></a>

# 🔗 Provenance and attribution inventory

**K-NSFR** · [Overview](../README.md) · [Documentation](../README.md#documentation) · [Governance](GOVERNANCE.md)

[![Stage](https://img.shields.io/badge/Stage-Repository_setup-BF8700?style=flat-square)](../README.md#status)

<a id="review-baseline"></a>

## 📌 Review baseline

Baseline reviewed for issue #3: K-NSFR release branch after issue #2, at
`68263b950d5fda7ff31e9ac7780f15390fce7190`. The initial governance
import is recorded in repository history at
`926e49f5de0221953a5147b562631c0081f7ae06`. Imported text was adapted
for this repository, with unrelated repository references removed. This inventory
records origins at the repository level; it does not establish transferable
copyright ownership or permission to redistribute.

<a id="tracked-surfaces"></a>

## 🗂️ Tracked surfaces

| Surface | Origin and current status | Attribution/obligation review |
| --- | --- | --- |
| Label catalogue, policy and sync/drift scripts | Initial governance import, then K-NSFR-specific edits | Owner must confirm rights to imported source and any required original attribution before distribution |
| Issue templates and triage Python code/tests | Initial governance import, adapted for the K-NSFR issue model | Preserve commit history; confirm rights and notice obligations before distribution |
| GitHub Actions workflows and Dependabot configuration | Initial governance import, adapted to local checks | Actions use SHA-pinned upstream components; inventory each action license and notice before redistribution of bundled code |
| README and GOVERNANCE.md | Initial baseline, then rewritten for K-NSFR | Repository history records changes; owner confirms notices if external distribution is proposed |
| SCOPE.md, ARCHITECTURE.md, DECISIONS.md, BACKLOG.md | Authored for issue #2 in this repository | No external code or template was adopted in these files |
| Workbook, regulatory parameters, numerical fixtures | None in the setup baseline | Future source/version, licensing and independent validation required |

<a id="external-tools-and-future-dependencies"></a>

## 🧩 External tools and future dependencies

GitHub-hosted actions are fetched at pinned revisions and are not vendored in
this tree. No third-party application package manifest or binary is present.
A future dependency entry must record its exact version, upstream URL, license,
notices, and use/redistribution decision. Regulatory text is source material to
cite and interpret, not an assumed software license.

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

The baseline above is historical. The machine-readable
[import register](provenance.json) records immutable reference revisions,
local import evidence, adaptations and license identity for later configuration,
quality tooling and coding-style adoption. Origin URLs in that register are
attribution records, not project dependencies or product claims. Initial imported
governance retains its local import commit above; its original rights review
remains part of the explicit publication gate in #19.

| Added surface | Origin / control |
| --- | --- |
| Source-control checker and fixtures | Authored locally in #4; isolation correction and regression fixture are local maintenance work |
| Markdown/YAML checker | Authored locally in #6; parsed navigation and explicit-anchor validation added during maintenance |
| Check runner and runner tests | Adapted from the immutable source in the import register; current gates cover only delivered capabilities |
| Contributor/tooling instructions and numeric rationale | Authored for the current setup; sketches remain unexecuted |
| VBA house style | Adapted with local module, API, numeric and maturity decisions; no host capability imported by implication |

<a id="runtime-tools"></a>

## 🧩 Runtime tools and licenses

Tools are downloaded, not vendored. Python wheel versions and SHA-256 hashes
are recorded in `tools/requirements-quality-ci.txt`; all transitive packages
are locked. Their installed distributions retain license files. The root MPL
license does not relicense these dependencies.

| Tool | Version identity | License | Purpose |
| --- | --- | --- | --- |
| GitHub checkout, Python/Node setup, artifact upload actions | Full SHA in each workflow | MIT | CI execution |
| actionlint | 1.7.12 plus archive SHA-256 in workflow | MIT | Workflow validation |
| Ruff | 0.16.6 plus wheel hash | MIT | Python lint and complexity |
| mypy and runtime dependencies | 2.3.1; exact transitive pins/hashes in lock | Licenses retained in installed distributions | Strict typing |
| PyYAML | 6.0.3 plus wheel hash | MIT | YAML parsing |
| types-PyYAML | 6.0.12.20260906 plus wheel hash | Apache-2.0 | YAML type stubs |
| markdown-it-py / mdurl | 4.0.0 / 0.1.2 plus wheel hashes | MIT | Markdown parsing and URL handling |

Upstream package sources: [PyYAML](https://github.com/yaml/pyyaml),
[markdown-it-py](https://github.com/executablebooks/markdown-it-py),
[mdurl](https://github.com/executablebooks/mdurl),
[typeshed](https://github.com/python/typeshed),
[Ruff](https://github.com/astral-sh/ruff),
[mypy](https://github.com/python/mypy), and
[actionlint](https://github.com/rhysd/actionlint).
Publication review must assess any future bundled distributions separately.
