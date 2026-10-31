<a id="repository-settings-evidence"></a>

# 🔎 Repository settings evidence

**K-NSFR** · [Overview](../README.md) · [Documentation](../README.md#documentation) · [Governance](GOVERNANCE.md)

[![Stage](https://img.shields.io/badge/Stage-Repository_setup-BF8700?style=flat-square)](../README.md#status)

This record holds the dated, repository-specific observations that the portable
[non-default settings reference](REPOSITORY_SETTINGS.md) leaves out.
[Governance](GOVERNANCE.md#installed-settings-and-limits) defines the working
rules and [SECURITY.md](../SECURITY.md#security-features) records the security
features. GitHub's live settings remain authoritative. This record is a
snapshot, not a certification. It does not enforce or authorize any change,
and it must not be copied as evidence for another repository.

Snapshot date: **2026-10-29 (Europe/Rome)**. Tracking: #80.

<a id="reading-the-record"></a>

## 🧭 Reading the record

- **Verified**: read from the GitHub API on the snapshot date.
- **Owner statement**: stated by the owner and recorded in governance; not read back.
- **Unverified**: no read-back on the snapshot date. This never means disabled.

No setting was changed to prepare this record. Configuration read-back is not
an enforcement test; enforcement scenarios are tracked in #15.

<a id="general-settings"></a>

## 🏷️ General settings

Evidence: the [repository API](https://api.github.com/repos/danielep71/K-NSFR),
read 2026-10-29. Every row is **verified**.

| Setting | Observed value | Interpretation |
| --- | --- | --- |
| Visibility / default branch | Public / `main` | Development targets `release/0.1.0`, not `main`. |
| Issues / Wiki | Enabled / enabled | The wiki is a separate repository, published by the owner. |
| Projects / Discussions | Disabled / disabled | Repository feature flags only. |
| Merge methods | Merge, squash and rebase enabled | The rulesets below restrict each protected branch further. |
| Auto-merge | Disabled | Passing checks do not authorize an integration. |
| Delete merged head branches | Enabled | `release/**` and `traffic-history` are protected from deletion by rulesets. |
| Update-branch suggestions | Enabled | Governance updates branches by merging the base in, never by rebasing. |
| Web commit sign-off | Not required | Separate from reviews and commit signing. |
| Legacy downloads flag | Disabled | API metadata; not a restriction on source downloads. |
| Pages / homepage | No Pages site; no homepage | No website is published. |
| Description / topics / license | "Net Stable Funding Ratio"; no topics; MPL-2.0 detected | About information is tracked in #17. |
| Forking | Allowed | Anyone can fork a public repository. |

<a id="rulesets"></a>

## 🛡️ Branch and tag rulesets

All five rulesets were read in full through the API on **2026-10-29**. All
have active enforcement, an empty bypass list and no excluded refs.

| Ruleset (ID) | Target | Observed rules |
| --- | --- | --- |
| [Protect main](https://github.com/danielep71/K-NSFR/rules/26065725) (26065725) | `~DEFAULT_BRANCH` | PR required, merge commits only, 0 approvals, conversations resolved; strict up-to-date `Repository integrity` from GitHub Actions (integration 15368), enforced on branch creation; deletion and force pushes blocked |
| [Protect release branches](https://github.com/danielep71/K-NSFR/rules/26065772) (26065772) | `refs/heads/release/**` | PR required, squash or merge, 0 approvals, conversations resolved; the same strict check, not enforced on branch creation; deletion and force pushes blocked |
| [Prevent release branch deletion](https://github.com/danielep71/K-NSFR/rules/26065850) (26065850) | `refs/heads/release/**` | Deletion blocked |
| [Protect release tags](https://github.com/danielep71/K-NSFR/rules/26065910) (26065910) | `refs/tags/v*` | Updates, deletion and force pushes blocked; creation is not restricted |
| [Protect traffic history](https://github.com/danielep71/K-NSFR/rules/26065944) (26065944) | `refs/heads/traffic-history` | Deletion and force pushes blocked; direct pushes allowed, with no PR or check requirement |

Both PR rules disable stale-review dismissal, code-owner review and last-push
approval, and name no required reviewers. Both also report
`require_extra_approval_for_unattributed_changes` as enabled; its effect here
was not tested. Zero approvals fits the sole-maintainer model and is not
independent approval. The rulesets require only Repository integrity. CodeQL
reports a check on pull requests into `main`, but no rule requires it.

The `traffic-history` branch has existed since 2026-10-29 and its ruleset applies
to it (API branch-rules read-back). Traffic collection is described in [traffic history](TRAFFIC.md) (#103).

<a id="actions-and-security"></a>

## ⚙️ Actions and security settings

| Setting | State | Source |
| --- | --- | --- |
| Full-SHA action pins required | Required | Owner statement in [governance](GOVERNANCE.md#installed-settings-and-limits); the Actions permissions API is not readable from the maintenance environment, so this is **unverified** on the snapshot date |
| Default workflow token | Read-only | Owner statement in governance; **unverified** on the snapshot date. Each workflow also declares `contents: read` at workflow level, or no permission at all (`Keep milestones issue-only`) |
| Security features | See [SECURITY.md](../SECURITY.md#security-features) | Read back on 2026-10-28 and 2026-10-29, with each row's source |

<a id="maintenance"></a>

## ✅ Maintenance

Update this record, with its snapshot date, after a setting changes or is read
back again. Keep the portable [reference](REPOSITORY_SETTINGS.md) free of
repository identifiers, and keep the evidence here.
