# Issue and label governance

K-NSFR uses a 20-label catalogue and the issue-triage rules defined below.
Names, colors and descriptions are versioned in [.github/labels.json](../.github/labels.json).
The application profile has no additional labels.

## Issues

Use the Bug report, Enhancement or Documentation form. Each issue is assigned
to danielep71, has exactly one priority and belongs to a milestone selected
during triage. Forms default bugs to P2 and enhancements/documentation to P3.

- P1: confirmed calculation bug producing wrong results for valid inputs; requires bug.
- P2: material correctness, robustness or maintainability defect.
- P3: non-blocking quality, testing, documentation or setup work.

Every issue includes a Dependencies and blocking status section. Write None
if ready, name prerequisite issues with #number, or use a line beginning
External condition: for an external blocker. Apply blocked only while a
named prerequisite is unresolved. Review dependent issues when a prerequisite
closes. Pull-request references do not count as blocking prerequisites.

The read-only Issue triage workflow audits open issues on issue events and
manual dispatch. Findings require maintainer triage; the workflow never edits
issues. Its mechanical P1 check requires bug; confirming a calculation defect
remains a maintainer decision.

## Label workflows

- Sync issue labels validates and self-tests changes on pull requests into
  main and release/**. Relevant pushes to main, or manual dispatch on the
  default branch, reconcile the live catalogue and verify an exact match.
  Labels outside the selected catalogue are removed (prune is enabled).
- Detect issue-label drift runs offline fixtures on relevant pull requests
  and compares live labels daily at 05:17 UTC or on manual dispatch.
  It reports differences without changing labels and retains evidence for 30 days.
- Issue triage checks assignee, priority, milestone, catalogue membership
  and dependency status.

## Local validation

```sh
node .github/scripts/labels-sync.mjs --self-test
node .github/scripts/labels-sync.mjs --mode validate
node .github/scripts/labels-drift.mjs --self-test
python tools/check_issues.py --check-config
python -m unittest discover -s tools -p 'test_check_issues.py'
python tools/check_source_control.py --ref HEAD
python -m unittest discover -s tools -p 'test_source_control.py'
```

## Repository configuration

The label manifest, synchronization and drift workflows are versioned in this
repository. Issue forms and triage tooling target K-NSFR.

Repository integrity runs the local checks above on pull requests and pushes
into main and release/**. Dependabot proposes grouped GitHub Actions updates
weekly on Mondays at 05:00 Europe/Rome, with ci/P3 labels and danielep71 assigned.
This configuration does not set branch protections or other repository settings.

## Branches and integration

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
decision and squashes the task PR into `release/0.1.0`. Do not record self-review
as an independent approval. The installed rulesets require zero approving reviews
and resolved review threads; this policy asks for an explicit owner decision,
not a fictitious second reviewer. Close the linked issue only after its criteria
and evidence are complete on the integrated release branch. Review dependent
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

## Installed settings and limits

The private repository currently has active rulesets for `main` and
`release/**`: pull requests, resolved review threads, no required approving
reviews, no bypass actors, strict Repository integrity status checks, and
protection against deletion and non-fast-forward updates. The `main` ruleset
allows merge commits; the release ruleset allows squash and merge commits.
Additional rulesets protect `release/**` from deletion, `v*` tags from
updates/deletion, and `traffic-history` from deletion or history rewrite.
GitHub Actions requires full SHA pins and gives the default workflow token
read access; Dependabot currently targets `main` and retargeting is tracked
by #6. Branch-rule enforcement scenarios are tracked by #15. These are
observed configuration and workflow baselines, not a claim that private-repo
limits or features available only to public repositories have been enabled.
