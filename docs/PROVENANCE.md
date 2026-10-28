# Provenance and attribution inventory

Baseline reviewed for issue #3: K-NSFR release branch after issue #2, at
`68263b950d5fda7ff31e9ac7780f15390fce7190`. The initial governance
import is recorded in repository history at
`926e49f5de0221953a5147b562631c0081f7ae06`. Imported text was adapted
for this repository, with unrelated repository references removed. This inventory
records origins at the repository level; it does not establish transferable
copyright ownership or permission to redistribute.

| Surface | Origin and current status | Attribution/obligation review |
| --- | --- | --- |
| Label catalogue, policy and sync/drift scripts | Initial governance import, then K-NSFR-specific edits | Owner must confirm rights to imported source and any required original attribution before distribution |
| Issue templates and triage Python code/tests | Initial governance import, adapted for the K-NSFR issue model | Preserve commit history; confirm rights and notice obligations before distribution |
| GitHub Actions workflows and Dependabot configuration | Initial governance import, adapted to local checks | Actions use SHA-pinned upstream components; inventory each action license and notice before redistribution of bundled code |
| README and GOVERNANCE.md | Initial baseline, then rewritten for K-NSFR | Repository history records changes; owner confirms notices if external distribution is proposed |
| SCOPE.md, ARCHITECTURE.md, DECISIONS.md, BACKLOG.md | Authored for issue #2 in this repository | No external code or template was adopted in these files |
| Workbook, regulatory parameters, numerical fixtures | None in the setup baseline | Future source/version, licensing and independent validation required |

GitHub-hosted actions are fetched at pinned revisions and are not vendored in
this tree. No third-party application package manifest or binary is present.
A future dependency entry must record its exact version, upstream URL, license,
notices, and use/redistribution decision. Regulatory text is source material to
cite and interpret, not an assumed software license.

A repository-tree review at this candidate found no workbook, client data,
credentials or proprietary template in tracked paths. The tracked text and
issue descriptions use synthetic planning examples. This is a scoped inspection,
not a claim about untracked local files, inaccessible prior private material,
or every historical blob. If questionable content is found, report privately
under SECURITY.md without copying it into issue evidence.
