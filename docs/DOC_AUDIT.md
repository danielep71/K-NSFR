<a id="documentation-and-comment-audit"></a>

# 📝 Documentation and comment audit

**K-NSFR** · [Overview](../README.md) · [Documentation](../README.md#documentation) · [Governance](GOVERNANCE.md)

[![Stage](https://img.shields.io/badge/Stage-Repository_setup-BF8700?style=flat-square)](../README.md#status)

This audit checked every document and every source comment against what the
repository does: statements, commands, paths, counts, dates, links, scope and
regulatory provenance. It records the coverage, the findings and how each was
settled. Tracking: #21.

The audit is a snapshot of one candidate. It does not certify later commits.

<a id="candidate-and-method"></a>

## 🎯 Candidate and method

- **Candidate:** `release/0.1.0` at `b0a3ca73a109a9fbba8b45ee8d06f6e2360d0a3f`,
  audited on 2026-11-10. The files were read at
  `70550f3dbdb1a56bf124c4da19143ad40133c381`, the head of #183, whose content
  equals the candidate apart from three lines of `docs/FILE_AUDIT.md`; I
  reread those lines at the candidate.
- **Reading:** each file in full, in three groups: governance and policy
  documents; design, contract and validation documents; workflows, tools,
  configuration and VBA comments. Every statement was checked against the
  implementing file it describes, and against the closed and open issues.
- **Automated checks:** the `content` gate verifies that every relative link
  and anchor resolves; the `contracts`, `fixtures`, `host-evidence` and
  `vba-static` gates verify that the examples, the fixture, the records and
  the VBA agree with their documents.
- **After the corrections:** `python tools/check.py --ci`, all 17 gates.

<a id="coverage"></a>

## 🗂️ Coverage

| Files | Read | Findings |
| --- | --- | --- |
| `README.md` | In full | 2 (S-14, S-15) |
| `CONTRIBUTING.md`, `AGENTS.md` | In full | None |
| `CODE_OF_CONDUCT.md` | In full | 1 (S-22) |
| `SECURITY.md` | In full | 3 (S-05 to S-07) |
| `CHANGELOG.md` | In full | 1 (S-19) |
| `docs/GOVERNANCE.md` | In full | 5 (S-01 to S-03, S-13, S-16) |
| `docs/REPOSITORY_SETTINGS.md` | In full | None |
| `docs/REPOSITORY_SETTINGS_EVIDENCE.md` | In full | 4 (S-08 to S-11) |
| `docs/TRAFFIC.md` | In full | 1 (S-04) |
| `docs/RELEASE.md` | In full | 2 (S-20, S-21) |
| `docs/INSTALLATION.md` | In full | 3 (S-17, S-18, S-21) |
| `.github/PULL_REQUEST_TEMPLATE.md` | In full | 1 (S-12) |
| `docs/ARCHITECTURE.md` | In full | 5 (D-01 to D-05) |
| `docs/SCOPE.md` | In full | 2 (D-06, D-07) |
| `docs/DECISIONS.md` | In full; the ADR rows are historical and stay | 1 (D-08) |
| `docs/BACKLOG.md` | In full | 1 (D-09) |
| `docs/IMPLEMENTATION.md` | In full | 3 (D-10 to D-12) |
| `docs/SOURCE_CONTROL.md` | In full | 2 (D-11, D-13) |
| `docs/VBA_HOUSE_STYLE.md` | In full | 6 (D-11, D-12, D-14 to D-17) |
| `docs/WRITING_STYLE.md` | In full | None |
| `docs/DATA_POLICY.md` | In full | 1 (D-18) |
| `docs/LICENSING.md` | In full | 1 (D-22) |
| `docs/PROVENANCE.md` | In full | 4 (D-19 to D-22) |
| `docs/FILE_AUDIT.md` | In full; inventory checked against Git at its candidate | 2 (D-23, D-24) |
| `docs/contracts/CONTRACTS.md` | In full, against `nsfr-contract-v1.json` and the checker | 2 (D-25, D-26) |
| `docs/methodology/SOURCE_REGISTER.md` | In full | 2 (D-27, D-28) |
| `docs/validation/VALIDATION.md` | In full, against `fixture-v1.json` and the checker | None |
| `docs/validation/HOST_EVIDENCE.md` | In full, against the checker and both scripts | 2 (D-29, D-30) |
| `tests/README.md`, `tests/evidence/README.md` | In full, against the harness and both records | None |
| `prototype/README.md`, `prototype/templates/README.md` | In full | None |
| `.github/workflows/` (8 files) | In full | 2 (C-03, C-04) |
| `.github/ISSUE_TEMPLATE/` (4 files), `.github/dependabot.yml`, `.github/labels.json`, `.github/scripts/` (3 files) | In full | 1 (C-03, in the three forms) |
| `tools/README.md` | In full | 8 (C-01 to C-08) |
| `tools/*.py` (11 tools, 11 test files) | In full: comments, docstrings, help texts, messages and test names | 6 (C-04, C-06, C-10 to C-13) |
| `tools/host/Record-HostEvidence.ps1`, `Verify-NSFR-Branch.ps1` | In full | 1 (C-09) |
| `tools/requirements-quality-ci.txt`, `pyproject.toml`, `.gitignore` | In full | None |
| `.editorconfig`, `.gitattributes` | In full | 1 (C-15) |
| `tests/core/` (2 modules) | In full | 1 (C-14) |
| `docs/licenses/THIRD_PARTY_NOTICES.txt` | Against the history of each listed file | 1 (C-16) |

Not reviewed:
- `LICENSE`, which is the standard license text.
- The data files: contract examples, the fixture, the host records and the
  JSON schemas. Their gates check them against their documents.
- `prototype/` beyond its two READMEs. It is stored as delivered (ADR-011),
  and the v0.2.0 issues review each piece before it is admitted.

<a id="findings"></a>

## 🧾 Findings and dispositions

Every finding was corrected in #21 unless its row says otherwise.

<a id="governance-and-policy-findings"></a>

### Governance and policy

| ID | Finding | Disposition |
| --- | --- | --- |
| S-01 | Governance said the full-SHA pin and read-only token settings were "not read back" | Corrected: verified by effect on 2026-11-09 (#15) |
| S-02 | Governance said enforcement scenarios were "tracked by #15" | Corrected: read back on 2026-11-09 and demonstrated on #172 |
| S-03 | Governance said no traffic collection was verified | Corrected: collection runs daily since 2026-11-07 (#16) |
| S-04 | The traffic guide said the same | Corrected |
| S-05 | SECURITY.md said CodeQL and External links were not on `main` | Corrected: on `main` since #155 |
| S-06 | SECURITY.md dated the security-updates read-back 2026-10-28 only | Corrected: also 2026-11-09 |
| S-07 | SECURITY.md put the tests in scope "once added" | Corrected: the harness and fixtures exist |
| S-08 | The settings evidence said enforcement was "tracked in #15" | Corrected |
| S-09 | Its legend kept a "My statement" level no row uses | Corrected |
| S-10 | Its general settings gave only the 2026-10-29 read | Corrected: also 2026-11-09 |
| S-11 | Its security-features row gave only 2026-10-28 and 2026-10-29 | Corrected |
| S-12 | The PR template checklist omitted the start date | Corrected |
| S-13 | Governance omitted the milestone events of issue triage | Corrected |
| S-14 | The README status cited one host record of two | Corrected |
| S-15 | The README validation list omitted the contracts and fixtures gates | Corrected |
| S-16 | Governance's list of what Repository integrity checks was incomplete | Corrected: it names the 17 gates |
| S-17 | The installation guide said a task-branch commit "cannot be recorded"; the recorder only warns, and the gate rejects the record | Corrected |
| S-18 | The recorded portable run did not say it was not the reference environment | Corrected |
| S-19 | The changelog called #75 an advanced CodeQL setup | Corrected: #75 was the default setup |
| S-20 | The release procedure's practices table uses adopted and adapted | Kept: #19 requires those dispositions, and the rows describe current behavior |
| S-21 | "Tracking: #18" and "Tracking: #19" name closed issues | Kept: every document names the issue that owns it this way, open or closed |
| S-22 | The code of conduct named discussions, which are off | Corrected: comments |

<a id="design-and-validation-findings"></a>

### Design, contracts and validation

| ID | Finding | Disposition |
| --- | --- | --- |
| D-01 | Architecture gave #14 as future host evidence and #13 the round trip together | Corrected: #14 delivered; #13 in v0.2.0 |
| D-02 | The `tests/host/` row named #13/#14 | Corrected |
| D-03 | The layout table missed `tests/evidence/`, the host scripts and the delivered guides | Corrected |
| D-04 | Architecture called the contract fields "not finalized" | Corrected: #9 and #10 fixed them |
| D-05 | Architecture's list of what CI checks was incomplete | Corrected |
| D-06 | Scope said the legal version required the register of #8 | Corrected: ADR-009 binds it |
| D-07 | Scope included host packaging in v0.1.0, ignored the prototype's factors and excluded any release | Corrected |
| D-08 | Decisions listed the regulatory version and numeric conventions as open | Corrected |
| D-09 | The backlog row joined #13 and #14 as open | Corrected |
| D-10 | Implementation gave parameter values to #8 and #9 | Corrected: values belong to PROD-01 |
| D-11 | Four documents gave the round trip to "#13/#14" | Corrected: #13 |
| D-12 | Several texts said #9 would define what it has defined | Corrected |
| D-13 | Source control said the repository held only configuration and checks | Corrected |
| D-14 | The house style said the harness was the only VBA | Corrected: the prototype's seven modules are not admitted |
| D-15 | The house style said CI validates tooling only | Corrected |
| D-16 | The house style said none of its example names exists | Corrected |
| D-17 | The house style gave a cleanup test to #14 | Corrected: to the first host module of #13 |
| D-18 | The data policy said there was no fixture | Corrected: FX-0001 |
| D-19 | Provenance said no numerical fixture was admitted | Corrected |
| D-20 | Provenance linked the upstream repositories of the runtime tools, which AGENTS.md keeps in licensing | Corrected: moved to licensing |
| D-21 | Provenance described the prototype's incorporation incompletely | Corrected |
| D-22 | Licensing and provenance omitted the CodeQL action | Corrected |
| D-23 | The file audit's A-04 overlooked those upstream links | Corrected |
| D-24 | The file audit's A-05 overlooked the assistant tools named in `.gitignore` | Corrected: retained as ignore patterns |
| D-25 | The contracts said no parameter set exists; FX-0001 has a synthetic one | Corrected |
| D-26 | The `invalid-positions` description omitted the malformed and unexpected dates | Corrected |
| D-27 | The register planned derivative and netting inputs that contract v1 rejects, and cases the case plan does not hold | Corrected |
| D-28 | The register's open items waited for #9 | Corrected: PROD-01 |
| D-29 | Host evidence said #14 stays open without a host | Corrected |
| D-30 | Host evidence accepted only `release/0.1.0`; the gate accepts any published release branch | Corrected |

<a id="comment-and-configuration-findings"></a>

### Comments and configuration

| ID | Finding | Disposition |
| --- | --- | --- |
| C-01 | The tooling guide counted six workflows; there are eight | Corrected |
| C-02 | The tooling guide gave host evidence to #13 and #14 as pending | Corrected |
| C-03 | The triage texts and the three forms named only the active milestone, and the workflow comment omitted the start date | Corrected |
| C-04 | The tooling guide, `check.py` and the integrity workflow said no NSFR calculation runs; the fixtures gate computes one | Corrected: no run of the NSFR engine |
| C-05 | The `tool-tests` row omitted the link and host-evidence tests | Corrected |
| C-06 | The host-evidence descriptions omitted the published-history rule | Corrected |
| C-07 | The invariants list omitted the full-history checkout | Corrected |
| C-08 | The CI download list omitted the runtimes | Corrected |
| C-09 | `Record-HostEvidence.ps1` warns only for `release/0.1.0` and `main` | Kept: the [release procedure](RELEASE.md#next-release-branch) changes it when the next release branch opens; the script cannot be run here |
| C-10 | The `--check-config` help text was incomplete | Corrected |
| C-11 | The link checker's docstring omitted HTML links | Corrected |
| C-12 | A contract test name contradicted its assertion | Corrected |
| C-13 | A host-evidence test name contradicted its assertion | Corrected |
| C-14 | The harness modules cite the closed #14 for compatibility | Kept: editing them changes their fingerprints and would need a new host run |
| C-15 | `.editorconfig` and `.gitattributes` left the VBA encoding open; ADR-008 settled it | Corrected |
| C-16 | The notices called `pr-milestones.yml` unchanged; #151 and #160 adapted it | Corrected |

<a id="limitations"></a>

## ⚠️ Limitations

- Reading finds what a reader notices. A statement that is wrong in a way no
  file contradicts can remain.
- The two host scripts were read, not run, in this audit; their runs are the
  [host evidence records](../tests/evidence/README.md).
- Links to external sites are checked by the weekly External links workflow,
  not by this audit.
