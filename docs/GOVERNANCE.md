<a id="issue-and-label-governance"></a>

# ⚖️ Issue and label governance

**K-NSFR** · [Overview](../README.md) · [Documentation](../README.md#documentation) · [Governance](GOVERNANCE.md)

[![Stage](https://img.shields.io/badge/Stage-Repository_setup-BF8700?style=flat-square)](../README.md#status)

K-NSFR uses a 20-label catalogue and the issue-triage rules defined below.
Names, colors and descriptions are versioned in [.github/labels.json](../.github/labels.json).
The application profile has no additional labels.

<a id="issues"></a>

## 🏷️ Issues

Use the Bug report, Enhancement or Documentation form. Each issue is assigned
to danielep71, has exactly one priority, a valid `YYYY-MM-DD` target end date in
Europe/Rome, and belongs to an open milestone for the active release selected
during triage. The target end date must be earlier than the milestone's due
date, so every milestone carries a due date; triage reads the calendar date of
the milestone's `due_on` and reports a milestone without one rather than
guessing it. The active milestone is `v0.1.0 - Repo setup`; update the triage
configuration when a new release branch becomes active. Never assign a closed
milestone as a fallback. Forms default bugs to P2 and enhancements/documentation to P3.

- P1: confirmed calculation bug producing wrong results for valid inputs; requires bug.
- P2: material correctness, robustness or maintainability defect.
- P3: non-blocking quality, testing, documentation or setup work.

P1 is reserved for a confirmed wrong calculation on valid inputs and requires
the bug label; P3 covers planned setup work. Keep intentional labels and valid
assignments when triaging. Forms cannot preselect a milestone; the owner assigns
it and resolves read-only triage findings.

Every issue includes a Dependencies and blocking status section. Write None
if ready, name prerequisite issues with #number, or use a line beginning
External condition: for an external blocker. Apply blocked only while a
named prerequisite is unresolved. Review dependent issues when a prerequisite
closes. Pull-request references do not count as blocking prerequisites.

The read-only Issue triage workflow audits open issues on issue events and
manual dispatch. Findings require maintainer triage; the workflow never edits
issues. Its mechanical P1 check requires bug; confirming a calculation defect
remains a maintainer decision.

<a id="label-workflows"></a>

## 🔄 Label workflows

- Sync issue labels validates and self-tests changes on pull requests into
  main and release/**. Relevant pushes to main, or manual dispatch on the
  default branch, reconcile the live catalogue and verify an exact match.
  Labels outside the selected catalogue are removed (prune is enabled).
- Detect issue-label drift runs offline fixtures on relevant pull requests
  and compares live labels daily at 05:17 UTC or on manual dispatch.
  It reports differences without changing labels and retains evidence for 30 days.
- Issue triage checks assignee, exactly one priority, open active milestone,
  ISO target end date before the milestone due date, catalogue membership and
  dependency status. It reports
  findings without changing issues. A pull request is never a prerequisite.
- Before label synchronization deletes a label outside the manifest, it checks
  all-state issues and PRs for use. A used label prevents reconciliation;
  the owner decides whether to retain or migrate it. Dry runs make no writes.

<a id="local-validation"></a>

## ✅ Local validation

Install the [portable toolchain](../tools/README.md#running-the-checks), then
run all gates from a clean, committed checkout:

```sh
python tools/check.py --ci
```

[Tooling documentation](../tools/README.md) defines the gates, report artifacts,
missing-tool behavior and limits. PR runs test GitHub's synthetic merge commit;
record that SHA alongside the PR head. Push runs test the integrated commit.

<a id="repository-configuration"></a>

## ⚙️ Repository configuration

The label manifest, synchronization and drift workflows are versioned in this
repository. Issue forms and triage tooling target K-NSFR.

Repository integrity runs the local checks above on pull requests and pushes
into main and release/**. Dependabot proposes grouped GitHub Actions updates to `release/0.1.0`
weekly on Mondays at 05:00 Europe/Rome, with ci/P3 labels and danielep71 assigned.
This configuration does not set branch protections or other repository settings.

<a id="branches-and-integration"></a>

## 🌿 Branches and integration

| Branch | Role | Changes |
| --- | --- | --- |
| `main` | Stable default branch and integration record | Receives a reviewed release PR by a merge commit. |
| `release/0.1.0` | Development and release candidate for milestone v0.1.0 | Receives task PRs by squash and the subsequent `main` synchronization by merge commit. |
| Task branches | One bounded issue or related change | Start from the current release branch and target `release/0.1.0`. |

Use the release branch as the default development target while v0.1.0 is in
progress. Link the task PR to its issue and name the exact candidate head SHA
in its evidence. Milestones belong to issues; leave PR milestones empty. Record
checks, tests and any manual validation against that candidate commit. Refresh
review and evidence whenever the head changes. A green check on an earlier head
is not evidence for a later one.

For a task PR, review the diff, resolve review threads and wait for the required
Repository integrity check on the current head. The owner makes the integration
decision and squashes the task PR into `release/0.1.0`. The squash commit
subject is the PR title in the [commit convention](../CONTRIBUTING.md#commit-messages)
ending with the PR number, for example `docs: define review evidence (#<PR>)`.
Its body says what changed and why, what was validated on the final head and
what was not, and ends with `Refs #<issue>`. Do not record self-review
as an independent approval. The installed rulesets require zero approving reviews
and resolved review threads; this policy asks for an explicit owner decision,
not a fictitious second reviewer. [Review evidence](#review) is recorded in the
PR. Close the linked issue only after its criteria and evidence are complete on
the integrated release branch, as [closing issues](#closing-issues) describes. Review dependent
issues and their blocked labels then.

For release integration, open a PR from `release/0.1.0` into `main`. Record the
candidate head and validation, resolve discussions, obtain the owner's explicit
integration decision, and wait for the required check at the current head.
Pause task merges into the release branch after selecting the candidate until
the release PR has merged and `main` has been synchronized back. If a task merge
or new commit intervenes, select a new candidate and repeat the review and
validation. Merge the release PR into `main` with a merge commit, preserving
ancestry. Then open a `main` to `release/0.1.0` PR and merge it with a merge
commit; resume task merges only when synchronization completes. Never force push
or delete protected branches. Release tags, publication and version claims
require their own release procedure (#19).

Record exceptions or deferred acceptance criteria in the linked issue and PR,
including the reason, owner decision, affected candidate SHA and follow-up
issue. A skipped check or unresolved criterion is visible evidence of a gap,
not an implicit pass. The owner decides whether to defer or rework before
integration; branch rules remain in force.

<a id="review"></a>

## 🔎 Review

The owner is the sole maintainer and reviewer. Review evidence has three
separate parts, and none stands in for another:

- the owner's review of the final diff and head commit;
- the automated code review, which runs on PRs opened by people;
- the required Repository integrity check on the current head.

The code-review bot reviews a PR when it is opened or marked ready, not after
later pushes. When the head changes after a completed review, the owner decides
whether to request a new review by posting `@codex review` as a top-level
comment. Before merging, wait until the review shows completed for the
recorded head SHA, then fix or answer every finding and resolve its thread.
Record the reviewed SHA and the result as text in the PR (for example "no
findings", or each finding and how it was settled): the bot's summary comment
is edited in place, so a link to it later describes another commit.

The bot does not review PRs authored by bots such as Dependabot. For those, the
owner's review of the final diff and head SHA is the only review; do not claim
automated coverage that did not run. A `main` to `release/0.1.0`
synchronization PR whose diff is empty changes no file: it needs the required
check only. Confirm that `git diff` between the two branches is empty and
record that and the check result in the PR. A synchronization PR that changes
any file is reviewed like any other PR. If review fails or is unavailable,
leave the PR open and record the blocker.

<a id="closing-issues"></a>

## 🏁 Closing issues

Close an issue as completed only when linked evidence meets its acceptance
criteria, and tick a criterion only when that evidence demonstrates it. When
the issue is closed, add a **Resolution** section to its body with:

- links to the change (PR and commit);
- the checks actually run, with their results;
- what was not validated, and any accepted limitation or deferred criterion
  with its follow-up issue.

Close for another reason (duplicate, not planned, superseded) only with the
matching close reason and an explanation. After closing, review dependent
issues and their blocked labels.

<a id="installed-settings-and-limits"></a>

## 🔐 Installed settings and limits

The private repository currently has active rulesets for `main` and
`release/**`: pull requests, resolved review threads, no required approving
reviews, no bypass actors, strict Repository integrity status checks, and
protection against deletion and non-fast-forward updates. The `main` ruleset
allows merge commits; the release ruleset allows squash and merge commits.
Additional rulesets protect `release/**` from deletion, `v*` tags from
updates/deletion, and `traffic-history` from deletion or history rewrite.
GitHub Actions requires full SHA pins and gives the default workflow token
read access; Dependabot targets `release/0.1.0` from the configuration
installed on `main`. Activation was verified in #6. Branch-rule enforcement scenarios are tracked by #15. These are
observed configuration and workflow baselines, not a claim that private-repo
limits or features available only to public repositories have been enabled.

Repository integrity runs on every PR into main and release/**, including
documentation-only and ancestry-only synchronization PRs, without a path filter.
It checks committed Markdown links/fences, YAML syntax and duplicate keys,
workflow syntax/expressions with a checksum-verified actionlint binary, strict Python typing, lint and tests, and committed blob whitespace (trailing
whitespace, final newline and blank lines at end of file). Python 3.13 and Node 22
are selected explicitly in the workflow. These portable checks do not run Excel
or validate regulatory results. A failing check blocks integration until a new
candidate passes; no skipped outcome is treated as evidence.
