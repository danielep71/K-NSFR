<!--
Target release/0.1.0 for task work. Use main only for owner-requested integration.
Never include secrets, client data or vulnerability details. Mark unavailable
checks NOT RUN with a reason; a portable check is not Excel evidence.
-->

<a id="outcome-and-linked-issue"></a>

## 🎯 Outcome and linked issue

Related issue: #

Describe the observable result and the reason for this change. State whether
it changes calculations, public API, host behavior, tooling or documentation.
Milestones belong to issues only; leave the PR milestone empty.

<a id="candidate-and-validation"></a>

## 🔎 Candidate and validation

- Exact current PR head (full SHA):
- Repository integrity run and result:
- Local validation commands and results:
- Excel/Windows host validation: NOT RUN — reason and follow-up issue, or exact build/bitness, candidate SHA and evidence.
- Numerical expected values: NOT APPLICABLE, or independent/published source and tolerance.

If the head changes, update this section and rerun affected checks. No
earlier green run is evidence for a later candidate. Describe any skipped
check as NOT RUN and explain it.

<a id="scope-and-risk"></a>

## ⚖️ Scope and risk

- Regulatory source/profile and applicability, if calculations change:
- Compatibility, input/output and public API effects:
- Synthetic-data and provenance review:
- Remaining risks, exceptions or deferred criteria (link the follow-up issue):

<a id="review-checklist"></a>

## ✅ Review checklist

- [ ] The PR targets the correct branch and links an issue with owner, one priority, an open active-release milestone and an ISO target end date before the milestone due date.
- [ ] The diff is focused; no secrets, real data, generated output or unapproved binary files.
- [ ] The current head has a successful required Repository integrity check and review conversations are resolved.
- [ ] Any host or numerical claim has separate evidence bound to the candidate.
- [ ] Owner integration decision and merge method follow [governance](../docs/GOVERNANCE.md#branches-and-integration).
