<a id="contributing"></a>

# 🤝 Contributing

**K-NSFR** · [Overview](README.md) · [Documentation](README.md#documentation)

<a id="start-here"></a>

## 🧭 Start here

Read [scope](docs/SCOPE.md), [architecture](docs/ARCHITECTURE.md) and
[governance](docs/GOVERNANCE.md). The editorial and assistant rules for every
change are in [AGENTS.md](AGENTS.md). The project is setup-only: no calculation
engine or workbook is available. Use synthetic data and retain source notices.
Report vulnerabilities privately under [SECURITY.md](SECURITY.md).

<a id="make-a-bounded-change"></a>

## 🌿 Make a bounded change

1. Use an issue form; assign the owner, one priority, an open active milestone,
   an ISO target date and an explicit dependency section.
2. Branch from `release/0.1.0`; keep one coherent issue or correction per PR.
3. Follow [source conventions](docs/SOURCE_CONTROL.md),
   [VBA house style](docs/VBA_HOUSE_STYLE.md) and
   [documentation style](docs/WRITING_STYLE.md).
4. Install the [portable toolchain](tools/README.md#running-the-checks), commit
   the candidate and run `python tools/check.py --ci` from a clean checkout.
5. Open the PR against `release/0.1.0`. Record the exact head, tested merge SHA,
   check run, results, risks and deferred criteria. Mark unavailable host tests
   **NOT RUN** with their follow-up issue, and state what was not validated.
6. Resolve findings and refresh evidence after changes; record review evidence
   as [governance](docs/GOVERNANCE.md#review) describes. The owner decides
   integration; task PRs are squashed. Main integration follows governance.

<a id="commit-messages"></a>

## ✍️ Commit messages

Write a subject in the conventional form with a type and an optional scope:
`fix:`, `docs:`, `ci:`, `test:`, `chore:` or `fix(ci):`. Use the imperative
mood and no trailing period. The body states:

- **what** changed and **why**;
- **what was validated** at that commit, with the commands run and their
  results; and
- **what was not validated**, separating portable checks from Excel runs and
  numerical evidence.

End the body with the issue reference, for example `Refs #12`. Do not put
credentials, private links or unverified claims in a commit message. Squash
commits follow the same form with the PR number in the subject, as
[governance](docs/GOVERNANCE.md#branches-and-integration) describes.

<a id="review-expectations"></a>

## ✅ Review expectations

Use clear names, explicit contracts and meaningful failure fixtures. Explain
what changes for a caller or maintainer. Review docs and comments alongside code.
Do not add generated workbooks, client data, secrets or speculative scaffolding.
A passing portable check does not establish regulatory correctness or Excel
compatibility. Supported-host installation guidance will arrive with #13/#14;
there is currently no workbook installation procedure to follow.
