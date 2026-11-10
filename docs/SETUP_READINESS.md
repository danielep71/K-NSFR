<a id="setup-readiness-report"></a>

# 🏁 Setup readiness report: v0.1.0

**K-NSFR** · [Overview](../README.md) · [Documentation](../README.md#documentation) · [Governance](GOVERNANCE.md)

[![Stage](https://img.shields.io/badge/Stage-Repository_setup-BF8700?style=flat-square)](../README.md#status)

This report summarizes what the repository setup delivers, what it does not,
and what comes next. Tracking: #22. The exact candidate, its runs and my
review are recorded in #22 and in the release PR, as the
[release procedure](RELEASE.md#release-candidate) requires.

> [!IMPORTANT]
> v0.1.0 is a repository setup. It contains no NSFR calculation, no admitted
> engine or workbook and no verified regulatory parameter. It certifies the
> process, not a result.

<a id="capabilities"></a>

## ✅ Completed capabilities

| Capability | Where | Issues |
| --- | --- | --- |
| Release-branch workflow, review evidence and issue closure | [Governance](GOVERNANCE.md), [release procedure](RELEASE.md) | #1, #19, #50 |
| Scope, architecture and decisions ADR-001 to ADR-011 | [Scope](SCOPE.md), [architecture](ARCHITECTURE.md), [decisions](DECISIONS.md) | #2 |
| License, provenance, synthetic data and security policies | [Licensing](LICENSING.md), [provenance](PROVENANCE.md), [data policy](DATA_POLICY.md), [SECURITY.md](../SECURITY.md) | #3, #68, #69 |
| Source-control, implementation and VBA conventions | [Source control](SOURCE_CONTROL.md), [implementation](IMPLEMENTATION.md), [house style](VBA_HOUSE_STYLE.md) | #4, #5, #118 |
| Issue forms, label catalogue and issue triage with dates | [Governance](GOVERNANCE.md#issues) | #7, #83, #89, #167 |
| One required check of 17 portable gates, hash-locked tools and SHA-pinned actions | [Tooling](../tools/README.md) | #6, #93, #117, #145 |
| Repository settings, read back and enforced | [Settings evidence](REPOSITORY_SETTINGS_EVIDENCE.md) | #15, #80 |
| Regulatory source register for the planned profile | [Register](methodology/SOURCE_REGISTER.md) | #8 |
| Input, parameter and result contracts with synthetic examples | [Contracts](contracts/CONTRACTS.md) | #9 |
| Numerical fixture schema and validation rules | [Validation](validation/VALIDATION.md) | #10 |
| VBA static checks | [Tooling](../tools/README.md#vba-static-checks) | #11 |
| Neutral VBA harness, run on the supported host | [Harness](../tests/README.md), [records](../tests/evidence/README.md) | #12, #14 |
| Host evidence bound to a commit, recorded and verified by script | [Host evidence](validation/HOST_EVIDENCE.md), [installation](INSTALLATION.md) | #14, #18 |
| Daily traffic collection | [Traffic](TRAFFIC.md) | #16 |
| Audit of every tracked file and of the documentation | [File audit](FILE_AUDIT.md), [documentation audit](DOC_AUDIT.md) | #20, #21 |

<a id="limitations"></a>

## ⚠️ Limitations

- **No calculation:** no ASF, RSF or ratio exists. The benchmark
  [prototype](../prototype/README.md) is stored and not admitted (ADR-011).
- **Host coverage:** one host build ran the neutral harness, on Windows 11
  Enterprise 25H2 with Microsoft 365 Apps build 20430.20180, 64-bit. Other
  hosts are NOT RUN.
- **No workbook template or source synchronization:** they come with #13.
- **Review:** I am the sole maintainer. The rulesets require no approving
  review, and the automated code review is not an independent approval.
- **Issue records:**
  - The 55 issues closed before the start-date rule of 2026-11-09 (#167) have
    no start date. The rule applies to open issues.
  - #1 to #7 were closed on 2026-10-28, before the closure rule of #50. Their
    evidence links are in their bodies, not in a resolution section.
- **Historical review gaps:** the gaps recorded in #22 on 2026-10-30 and
  2026-11-01 stay as recorded. Some early PRs carry an automated review of an
  earlier head, or an accepted review of the base merge only.

<a id="practices"></a>

## 🧭 Setup practices

The [release procedure](RELEASE.md#setup-practices) records each setup
practice as adopted, adapted, authored, deferred or not adopted, with its
reason.

<a id="next-milestone"></a>

## 🗺️ Next milestone: v0.2.0 - Prototype Setup

| Issue | Work |
| --- | --- |
| #13 | Host template and source synchronization |
| #132 | Verify the prototype's regulatory tables against the source register |
| #133 | Decide the scope of the reconciliation with a bank's NSFR templates |
| #134 | Use VBA Decimal in the prototype's calculation core |
| #135 | Reconcile the prototype's input model with the NSFR contract |
| #136 | Register the prototype's demo and Python reference as numerical fixtures |
| #137 | Admit the prototype's VBA modules into `src/` |
| #138 | Run the admitted prototype in Excel and record host evidence |

The calculation work after it is in the [backlog](BACKLOG.md#later-calculation-work).
