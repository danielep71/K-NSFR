# Decision log

Issue: [#2](https://github.com/danielep71/K-NSFR/issues/2).
Decision date: 2026-10-28.
Status: selected implementation baseline for the development branch after
integration of issue #2. These decisions do not certify host execution or
regulatory correctness.

| ID | Decision and rationale | Consequence / revisit trigger |
| --- | --- | --- |
| ADR-001 | Keep v0.1.0 setup-only; the repository has governance tools but no numerical engine | No regulatory formulas or parameter tables; production capability requires separately reviewed work |
| ADR-002 | Use Excel/VBA for an inspectable workbook reference implementation; retain Python/JavaScript for repository tooling | Windows 11 / Microsoft 365 desktop / VBA7 / 64-bit is the initial host target; revisit if deployment or scale requirements exceed this profile |
| ADR-003 | Select EU CRR standard NSFR as the first planned profile; Basel is reference-only | No executable profile is currently supported; #8 must establish exact applicability and version before implementation |
| ADR-004 | Separate core, API, and host with one-way dependencies | Determinism and independent testing take precedence over worksheet convenience; no hidden host inputs in calculations |
| ADR-005 | Keep exported source canonical and data synthetic | #4 defines text conventions, #13 packaging; real data and operational deployment require separate approval and controls |
| ADR-006 | Distinguish planned support, portable checks, host execution, and numerical validation | A green repository workflow cannot be cited as proof of Excel or NSFR correctness |

Alternatives and the selected layout are documented in
[ARCHITECTURE.md](ARCHITECTURE.md). Regulatory versions, interpretation choices,
numeric/error conventions, workbook schema, and production-readiness approval
remain assigned follow-on decisions in [BACKLOG.md](BACKLOG.md). Record a new
dated decision with its rationale and superseded ID when changing this baseline;
do not silently rewrite historical decisions.
