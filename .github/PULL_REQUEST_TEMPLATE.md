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

<a id="final-head-review"></a>

## 🔍 Final-head review

Record the review as text; the code-review summary comment is edited in place,
so a link to it later shows another commit. For a bot-authored PR, record the
owner's review of the final head and that automated review did not run. An
empty-diff synchronization PR needs the required check only: record the empty
diff instead. See [governance](https://github.com/danielep71/K-NSFR/blob/release/0.1.0/docs/GOVERNANCE.md#review).

- Final PR head (full SHA):
- Review type (automated and owner / owner only for a bot-authored PR / not required: empty-diff synchronization):
- Reviewed commit SHA and result (no findings, or each finding and how it was settled):

<a id="scope-and-risk"></a>

## ⚖️ Scope and risk

- Regulatory source/profile and applicability, if calculations change:
- Compatibility, input/output and public API effects:
- Synthetic-data and provenance review:
- Remaining risks, exceptions or deferred criteria (link the follow-up issue):

<a id="not-validated"></a>

## ⚠️ Not validated

State plainly what was not checked and why, separating portable checks from
Excel runs and numerical evidence. Write "Nothing" only if that is true.

<a id="review-checklist"></a>

## ✅ Review checklist

- [ ] The PR targets the correct branch and links an issue with owner, one priority, an open active-release milestone and an ISO target end date before the milestone due date.
- [ ] The diff is focused; no secrets, real data, generated output or unapproved binary files.
- [ ] The current head has a successful required Repository integrity check and review conversations are resolved.
- [ ] The final-head review covers the current head; every finding is fixed or answered.
- [ ] Any host or numerical claim has separate evidence bound to the candidate.
- [ ] Owner integration decision and merge method follow [governance](https://github.com/danielep71/K-NSFR/blob/release/0.1.0/docs/GOVERNANCE.md#branches-and-integration).
