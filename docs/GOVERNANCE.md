# Issue and label governance

K-NSFR adopts the 20-label catalogue and issue-triage rules from K-LCR.
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
```

## Source

Initial configuration copied from danielep71/K-LCR at commit
ff01720c79633d28ba342a7a1ac2c6836aafdf95 on 2026-10-28 and adapted for K-NSFR.
The 20-label manifest, synchronization and drift workflows match that source.
Repository references in issue forms and triage tooling target K-NSFR.

Repository integrity runs the local checks above on pull requests and pushes
into main and release/**. Dependabot proposes grouped GitHub Actions updates
weekly on Mondays at 05:00 Europe/Rome, with ci/P3 labels and danielep71 assigned.
This configuration does not set branch protections or other repository settings.
