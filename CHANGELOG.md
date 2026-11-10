<a id="changelog"></a>

# 📜 Changelog

**K-NSFR** · [Overview](README.md) · [Documentation](README.md#documentation) · [Governance](docs/GOVERNANCE.md)

[![Stage](https://img.shields.io/badge/Stage-Repository_setup-BF8700?style=flat-square)](README.md#status)

What each version adds, newest first. The [release procedure](docs/RELEASE.md)
moves the *Unreleased* entries under a version when its release candidate is
selected. No version contains an NSFR calculation or claims regulatory
compliance.

<a id="unreleased"></a>

## 🚧 Unreleased

- The macro-free workbook template `templates/NSFR_Template.xlsx`, generated
  reproducibly by `tools/build_template.py` with a hash-locked XlsxWriter, and
  the `template` gate that rebuilds and compares it (#13).
- The source manifest `src/manifest.txt`, checked by the `vba-static` gate;
  the host module `NSFR_Host_Sync`, which imports and exports exactly the
  listed components after a backup copy; and `Test-NSFR-Sync.ps1`, which runs
  the round trip on the supported host (#13).
- `release/0.2.0` is the active release branch (#189).
- The documentation describes the delivered template and synchronization, the
  three host evidence records and the tag `v0.1.0`; the installation guide
  uses a neutral example path; the `content` gate no longer needs XlsxWriter
  (#193).
- The template's gaps against contract v1 and the prototype are recorded,
  with #133, #135 and #137 as owners; ADR-008 allows `On Error GoTo` in tests
  and examples, as the static checks already did; the cleanup evidence is
  required of the first host module that changes application settings
  (#196).
- `Test-NSFR-Sync.ps1` writes a synchronization record, validated by the
  `host-evidence` gate, and its layout check also compares the sheets and
  workbook names with the template (#195).

<a id="v010"></a>

## 🏷️ 0.1.0, repository setup — 2026-11-10

The setup milestone, `v0.1.0 - Repo setup`, on `release/0.1.0`: the candidate
for the tag `v0.1.0`. #22 certifies it, and the
[setup readiness report](docs/SETUP_READINESS.md) summarizes it.

<a id="governance-and-process"></a>

### ⚖️ Governance and process

- Release-branch workflow, squash task merges, merge-commit integration into
  `main` and synchronization back (#1, #29).
- Scope, the EU CRR standard profile as the first planned target, and the
  core/API/host architecture (#2, #27); decisions ADR-001 to ADR-011.
- Licensing under MPL-2.0, provenance, synthetic-data and security policies
  (#3, #30, #32); the repository is public since 2026-10-28, with private
  vulnerability reporting (#68, #69, #71).
- Commit messages, review evidence and issue closure (#50, #54); editorial and
  assistant rules in `AGENTS.md` (#99); a code of conduct (#147).
- Issue forms, label catalogue schema 2 with synchronization and drift checks,
  and issue triage of assignee, priority, milestone, start date and target
  end date (#7, #83, #89, #91, #167).
- Repository settings reference with dated read-back, and enforcement observed
  with deliberately failing probes (#15, #80, #97).
- Release and closeout procedure and this changelog (#19); installation guide
  (#18).
- Audit of every tracked file (#20) and of the documentation and comments
  (#21), and the setup readiness report (#22).

<a id="checks"></a>

### ✅ Checks

- One required check, Repository integrity, running `tools/check.py`: 17
  portable gates over committed text, Markdown links and anchors, YAML,
  workflows, strict Python typing, lint and tests, labels, issue forms,
  contracts, fixtures, VBA static rules and host evidence (#6, #93, #117,
  #144, #145).
- Hash-locked Python tools, SHA-pinned actions and Dependabot updates against
  the release branch (#6, #85).
- CodeQL default setup (#75), replaced by the advanced setup for Actions,
  Python and JavaScript (#146).
- Byte-for-byte convention for future copies of regulatory sources (#95).

<a id="nsfr-foundations"></a>

### 🏦 NSFR foundations, without a calculation

- Regulatory source register for the planned profile (#8).
- Input, parameter and result contracts with synthetic examples (#9).
- Numerical fixture schema and validation rules: exact decimal results, scales,
  rounding and comparison (#10).
- VBA static checks with regression fixtures (#11).
- My benchmark prototype, stored as delivered and not admitted (#130).

<a id="host"></a>

### 💻 Host

- Neutral VBA test harness and its self-test (#12).
- Candidate-bound host evidence, recorded by `Record-HostEvidence.ps1` and
  verified for a branch by `Verify-NSFR-Branch.ps1` (#14, #162).
- The self-test passed on Windows 11 with Microsoft 365 Excel, 64-bit, at
  `5231985` and `dec4ad5` (#175, #178), and at the release candidate `87d4618`
  (#186).

<a id="operations"></a>

### 📈 Operations

- Daily traffic collection into the protected `traffic-history` branch, with a
  GitHub App token and an alert issue (#16, #103, #143).

<a id="not-in-this-version"></a>

### 🚫 Not in this version

- No ASF, RSF or ratio, and no admitted engine or workbook.
- No workbook template or source synchronization: #13, in v0.2.0.
- No verified regulatory parameter set, and no executed numerical fixture.
