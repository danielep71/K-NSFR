<a id="versioning-and-release-procedure"></a>

# 🚢 Versioning and release procedure

**K-NSFR** · [Overview](../README.md) · [Documentation](../README.md#documentation) · [Governance](GOVERNANCE.md)

[![Stage](https://img.shields.io/badge/Stage-Repository_setup-BF8700?style=flat-square)](../README.md#status)

This procedure takes a milestone from its last task merge to a version on
`main`, and closes the milestone. Tracking: #19. Every step is manual: no
workflow tags, publishes a release, changes the visibility or closes a
milestone, and nothing here authorizes one.

> [!IMPORTANT]
> A version names a stage of this repository, not a certification. v0.1.0 is
> the repository setup: governance, checks, contracts, validation rules and a
> neutral host harness, with no NSFR calculation. No version claims
> regulatory compliance or production suitability.

<a id="versions"></a>

## 🏷️ Versions

| Version | Milestone | What it may claim |
| --- | --- | --- |
| 0.1.0 | `v0.1.0 - Repo setup` | The repository setup. No ASF, RSF or ratio exists |
| 0.2.0 | `v0.2.0 - Prototype Setup` | The prototype brought into line with the contracts, the source register and Excel evidence. Numerical claims only where admitted fixtures pass on the host |
| Later 0.x | One milestone each | Engine work from the [backlog](BACKLOG.md); a numerical claim needs its fixtures and host evidence |
| 1.0.0 | Not planned | My decision after PROD-04 of the backlog; even then, not a regulatory certification |

Versions follow `MAJOR.MINOR.PATCH`. Each version has one milestone, one
release branch `release/X.Y.Z` (without a `v`) and one tag `vX.Y.Z`. A
PATCH version corrects a published version without new capability.

The [changelog](../CHANGELOG.md) lists what each version adds, under the
heading of the version, and keeps unreleased changes under *Unreleased*.

<a id="release-candidate"></a>

## 🎯 Release candidate

The release candidate is the head of the release branch I select for
integration, with its full SHA.

1. **Freeze:** when every issue of the milestone except its certification
   issue is closed or rescheduled (see [milestone closure](#milestone-closure)),
   stop merging task PRs into the release branch. Only the changelog and the
   host record below are merged after the freeze; any other task merge makes
   a new candidate.
2. **Changelog:** move the *Unreleased* entries under the version and its
   date, in a task PR. Its head is then the candidate.
3. **Portable checks:** from a fresh clone at the candidate, run
   `python tools/check.py --ci`, and confirm that the Repository integrity
   push run on the candidate passed.
4. **Host evidence:** run
   `tools/host/Verify-NSFR-Branch.ps1 -Branch release/X.Y.Z` on the
   supported host ([installation](INSTALLATION.md#install-and-run-the-harness)),
   and propose its record in a task PR. That PR adds only
   `tests/evidence/<candidate SHA>.json`, so the fingerprints of the candidate
   and of its head are identical: I accept the record of the candidate for
   the head. Any other change makes a new candidate.
5. **Review:** I review the diff from the last version to the candidate and
   record that review and the evidence in the release PR.

<a id="integration-and-tag"></a>

## 🔀 Integration and tag

1. Open the release PR from `release/X.Y.Z` into `main`, as
   [governance](GOVERNANCE.md#branches-and-integration) describes. Record the
   candidate, the tested merge SHA, the checks, the host record and my
   decision. Merge it with a merge commit when the required check passes on
   the current head.
2. Synchronize `main` back into the release branch with an empty-diff PR
   merged by a merge commit.
3. Tag the integration merge commit on `main`:

   ```sh
   git fetch origin main
   git tag -a vX.Y.Z -m "K-NSFR X.Y.Z" <integration merge SHA>
   git push origin vX.Y.Z
   ```

   The tag ruleset blocks updating or deleting a `v*` tag, so check the SHA
   before pushing. A wrong tag is never moved: correct it with the next PATCH
   version.
4. Optionally, publish a GitHub release from the tag, with the changelog
   section as its notes. Mark a 0.x release as a pre-release. Attach no
   workbook or binary.
5. Close the milestone (see [milestone closure](#milestone-closure)).

<a id="next-release-branch"></a>

## 🌿 Opening the next release branch

After the tag, create `release/X.Y.Z` for the next milestone from `main`,
then change the active branch where it is named, in one task PR into the new
branch:

- [governance](GOVERNANCE.md), [contributing](../CONTRIBUTING.md), the
  [installation guide](INSTALLATION.md), the README and the PR template;
- the Dependabot `target-branch` in `.github/dependabot.yml`;
- `ACTIVE_MILESTONES` in `tools/check_issues.py`, when the closed milestone
  leaves it.

The rulesets already cover every `release/**` branch.

<a id="rollback-and-recovery"></a>

## ↩️ Rollback and recovery

Protected branches are never force-pushed, rebased or deleted. Recover
forward:

| Situation | Recovery |
| --- | --- |
| A task merge broke the release branch | Revert its squash commit in a task PR (`git revert <sha>`), with the same checks as any change |
| The release PR fails its checks or review | Fix on the release branch through task PRs; the fixed head is a new candidate |
| A defect is found on `main` after integration | Fix it on the release branch and integrate again; for a published version, release a PATCH version |
| A wrong tag was pushed | Leave it, and record the error in the changelog; the next PATCH version carries the correction |
| A synchronization PR shows a non-empty diff | Stop: find the change made outside the procedure, and review the PR like any other |

<a id="milestone-closure"></a>

## 🏁 Milestone closure

Every issue carries a start date and a target end date before its milestone's
due date; issue triage checks them ([governance](GOVERNANCE.md#issues)).
Milestones belong to issues: a workflow removes a milestone set on a pull
request.

- **Closing a milestone:** only after the version is tagged and the
  certification issue is closed with the tag in its resolution, and only when
  every other issue in it is closed with a resolution or rescheduled. I close
  it by hand.
- **Overdue issue:** an open issue past its target end date stays open, and
  I record a new target end date and the reason in its Planning section.
- **Rescheduled or deferred work:** I move the issue to a later milestone,
  update its target end date, and add a dated section with the reason. A
  deferred acceptance criterion is recorded in the issue that defers it, with
  the issue that takes it over, as #14 did for the synchronization round
  trip.
- **Milestone due date:** I move it only by decision, recorded in the
  certification issue of the milestone, never to make a late issue fit.

<a id="setup-practices"></a>

## 🧭 Setup practices

The practices of the setup milestone, and whether they were adopted as is,
adapted or deferred. [Provenance](PROVENANCE.md) and the
[third-party notices](licenses/THIRD_PARTY_NOTICES.txt) record where imported
material came from.

| Practice | Disposition | Reason and record |
| --- | --- | --- |
| Release branch with squash task merges and merge-commit integration | Adopted | Keeps `main` as the integration record (#1) |
| Rulesets: required check, resolved threads, no bypass, protected tags and branches | Adopted, with 0 approvals | I am the sole maintainer; enforcement observed in #15 |
| Label catalogue with synchronization and drift checks | Adapted | A schema 2 catalogue; workflows synchronize it and report drift (#83) |
| Issue forms and issue triage | Adapted | The forms ask for a start date and a target end date; triage checks them against the milestone due date (#7, #55, #89, #167) |
| Milestones on issues only | Adapted | A reconciling workflow removes PR milestones (#91) |
| Portable checks as one required check | Adapted | One entry point, `tools/check.py`, for CI and local runs (#6, #93) |
| Hash-locked Python tools and SHA-pinned actions | Adopted | Reproducible CI; Dependabot proposes the updates (#6, #85) |
| CodeQL advanced setup | Adapted | Replaces the default setup to cover Actions, Python and JavaScript (#146) |
| Traffic history on a protected data branch | Adapted | Collected daily with a short-lived GitHub App token (#103, #143) |
| Byte-for-byte regulatory source copies | Adopted | Convention only; no source copy is stored yet (#95) |
| VBA static checks | Authored | The VBA rules of ADR-008 (#11) |
| Candidate-bound host evidence | Authored | Portable checks never prove Excel (#14) |
| Workbook template and source synchronization | Authored | Deferred from v0.1.0 and delivered in v0.2.0; its layout is the prototype's (#13) |
| Release tag and GitHub release | Authored | This procedure; `v0.1.0` was tagged on 2026-11-10 (#22); a GitHub release is optional, and none is published |
| Required approving reviews, required commit signing, Discussions | Not adopted | A sole maintainer cannot give an independent approval; no ruleset requires signed commits; Discussions are off ([settings evidence](REPOSITORY_SETTINGS_EVIDENCE.md#general-settings)) |
| Public visibility | Adopted on 2026-10-28 | Post-publication review in #19 |
