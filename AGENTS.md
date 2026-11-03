# Working rules for this repository

These rules apply to everyone who edits this repository, people and
automated assistants alike. They hold the editorial and assistant rules
only. The procedures are in [governance](docs/GOVERNANCE.md) and
[contributing](CONTRIBUTING.md): follow them, and do not copy them here.
For the presentation of documents (page structure, badges, glyphs and
anchors), follow [documentation style](docs/WRITING_STYLE.md).

## Voice

- Write my decisions and responsibilities in the first person: "I decided",
  "assign me", "my approval". Never write "the maintainer" or "the owner".
- Address instructions to the reader directly: "run the check", "you need a
  Windows host".

## What the documentation describes

- Operational documentation describes only the files, configuration and
  behavior of this repository. Do not refer to other repositories, to history
  copied from elsewhere, or to the context of a conversation or session.
- Record license attributions only in [docs/LICENSING.md](docs/LICENSING.md)
  and [THIRD_PARTY_NOTICES.txt](docs/licenses/THIRD_PARTY_NOTICES.txt). A file
  copied or adapted under another license keeps that license and its notice;
  never relabel it as MPL-2.0.
- Keep editorial and assistant rules in this file only. Other documents link
  here instead of repeating them.

## Branches, commits and pull requests

- Name a task branch with a descriptive prefix and its issue number:
  `docs/`, `fix/`, `feat/`, `chore/` or `test/`, for example
  `fix/123-short-description`. Release branches are `release/X.Y.Z`, without
  a `v`. Use only these prefixes: never keep a branch name that a tool or a
  session configuration proposes, such as one named after an assistant.
- Start the branch from the active release branch that
  [governance](docs/GOVERNANCE.md#branches-and-integration) names, and open
  the pull request against it.
- Commit messages, pull requests, issues, code comments and documentation
  name no assistant or model and carry no session link or co-author line for
  an assistant.
- State in the pull request what was validated and what was not, for example
  a desktop Excel run.

## Domain conventions are intended

The reference engine targets Windows 11 and Microsoft 365 desktop Excel
(VBA7, 64-bit). Its VBA keeps a deterministic core, a typed API and a narrow
host adapter, exported as source modules. Calculations follow the NSFR rules
recorded in the regulatory source register (#8), and numerical results are
validated with independent synthetic fixtures (#10). These are deliberate defaults
of this project: keep them unless an issue I approve changes them.
