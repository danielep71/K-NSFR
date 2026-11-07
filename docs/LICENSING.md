<a id="licensing-and-distribution-decision"></a>

# ⚖️ Licensing and distribution decision

**K-NSFR** · [Overview](../README.md) · [Documentation](../README.md#documentation) · [Governance](GOVERNANCE.md)

[![Stage](https://img.shields.io/badge/Stage-Repository_setup-BF8700?style=flat-square)](../README.md#status)

<a id="selected-license"></a>

## 📜 Selected license

K-NSFR has been public since 2026-10-28. I selected the Mozilla Public License, version 2.0 (MPL-2.0), for K-NSFR. The
standard text is in [LICENSE](../LICENSE). Copyright and other rights remain
with their respective rights holders; this choice grants only rights a
contributor is entitled to license.

<a id="rights-and-notices"></a>

## 🔍 Rights and notices

Confirm rights and required notices for imported material and any third-party
content before it is published. The MPL-2.0 choice applies to original K-NSFR
source files and contributions for which the contributors can grant those
rights; it cannot relicense material owned by others. Preserve existing notices
and add the MPL Exhibit A notice to new source files or a location where a
recipient is likely to find it. Do not add an Exhibit B incompatibility notice
without a separate decision. Attributions are recorded only here and in
[THIRD_PARTY_NOTICES.txt](licenses/THIRD_PARTY_NOTICES.txt), which lists every
file copied or adapted from another project with its source, revision,
adaptations and license, and the license text and copyright of any file under
a license other than MPL-2.0. Such a file keeps that license and its notice. The licensing and policy setup is complete in issue #3. Record the remaining
ownership and attribution evidence for publication under issue #19.

<a id="publication-gate"></a>

## 🚦 Publication gate

Publication gate: review every tracked file and accessible history, issue/PR
content, Actions logs and artifacts for secrets and client, employer, proprietary
or personal material. Confirm rights and notices for the imported governance
source, and remove or resolve any questionable content before changing visibility.
The repository was made public on 2026-10-28, before this review; on my
decision the review was then run as a post-publication review and is
recorded in #19. Its follow-up items #68, #69 and #70 are closed.
After publication, review public issue reporting and repository controls for
that visibility. Record the candidate commit and review evidence in issue #19.

<a id="future-dependencies"></a>

## 🧩 Future dependencies

The inventory in [PROVENANCE.md](PROVENANCE.md) identifies the existing imported
and authored surfaces. Future dependencies require a source URL, version or
commit, license text/notice, permitted use and redistribution review before
admission. A dependency manifest or vendor notice must accompany adopted code
where its terms require one.

<a id="runtime-tool-licenses"></a>

## 🧾 Runtime tool licenses

The tools listed in [PROVENANCE.md](PROVENANCE.md#runtime-tools) are
downloaded at run time and are not vendored; their installed distributions
retain their own license files, and the root MPL-2.0 license does not
relicense them.

| Tool | License |
| --- | --- |
| GitHub checkout, Python/Node setup and artifact upload actions | MIT |
| GitHub App token action (`actions/create-github-app-token`) | MIT |
| actionlint | MIT |
| Ruff | MIT |
| mypy and its runtime dependencies | Retained in the installed distributions |
| PyYAML | MIT |
| types-PyYAML | Apache-2.0 |
| markdown-it-py and mdurl | MIT |

<a id="prototype-tool-licenses"></a>

### Prototype tool licenses

The tools of the [non-admitted prototype](../prototype/README.md) need the
following, which CI neither locks nor installs. None is vendored, and the
workbook they generate is not committed.

| Tool | License |
| --- | --- |
| XlsxWriter | BSD-2-Clause |
| openpyxl | MIT |
| LibreOffice | MPL-2.0 |
