<a id="contributing"></a>

# 🤝 Contributing

**K-NSFR** · [Overview](README.md) · [Documentation](README.md#documentation) · [Governance](docs/GOVERNANCE.md)

[![Stage](https://img.shields.io/badge/Stage-Prototype_setup-BF8700?style=flat-square)](README.md#status)

<a id="start-here"></a>

## 🧭 Start here

Read [scope](docs/SCOPE.md), [architecture](docs/ARCHITECTURE.md) and
[governance](docs/GOVERNANCE.md). The editorial and assistant rules for every
change are in [AGENTS.md](AGENTS.md). The project is setup-only: no calculation
engine is available, and the workbook template holds only the layout. Use synthetic data only, as the
[data policy](docs/DATA_POLICY.md) defines, and retain source notices.
Follow the [code of conduct](CODE_OF_CONDUCT.md), and report abuse as it
describes. Report vulnerabilities privately under [SECURITY.md](SECURITY.md).

<a id="make-a-bounded-change"></a>

## 🌿 Make a bounded change

1. Use an issue form: keep me as assignee and the form's priority, and give
   an ISO start date and target end date and an explicit dependency section.
   Triage confirms the priority and sets an open active milestone; the start
   date must not be after the target end date, and the target end date must
   fall before that milestone's due date.
2. Branch from `release/0.2.0`; keep one coherent issue or correction per PR.
3. Follow [source conventions](docs/SOURCE_CONTROL.md),
   [VBA house style](docs/VBA_HOUSE_STYLE.md) and
   [documentation style](docs/WRITING_STYLE.md).
4. Install the [portable toolchain](tools/README.md#running-the-checks), commit
   the candidate and run `python tools/check.py --ci` from a clean checkout.
5. Open the PR against `release/0.2.0`. Record the exact head, tested merge SHA,
   check run, results, risks and deferred criteria. Mark unavailable host tests
   **NOT RUN** with their follow-up issue, and state what was not validated.
6. Resolve findings and refresh evidence after changes; record review evidence
   as [governance](docs/GOVERNANCE.md#review) describes. I decide
   integration; task PRs are squashed. Main integration follows governance.

<a id="commit-messages"></a>

## ✍️ Commit messages

Write a subject in the conventional form with a type and an optional scope:
`feat:`, `fix:`, `docs:`, `ci:`, `test:`, `chore:` or `fix(ci):`. Use the
imperative mood and no trailing period. The body states:

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
what changes for a caller or for whoever maintains the code. Review docs and comments alongside code.
Do not add generated workbooks, client data, secrets or speculative scaffolding.
The one approved binary is the workbook template, which `tools/build_template.py`
generates and the `template` gate checks.
A passing portable check does not establish regulatory correctness or Excel
compatibility. The [installation guide](docs/INSTALLATION.md) sets up the
portable checks, the supported host, the neutral harness and the template with
its source synchronization.
