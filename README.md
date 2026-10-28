# K-NSFR

Net Stable Funding Ratio: a repository for a planned, auditable Excel/VBA
reference implementation.

**Status: repository setup (v0.1.0).** No NSFR calculation engine or validated
regulatory profile is implemented. Do not use this baseline for regulatory
reporting, production decisions, or a claim of compliance.

The initial audience is risk analysts, reviewers, instructors, and developers
working with synthetic examples. The selected implementation target is Microsoft
365 desktop Excel on Windows 11, 64-bit VBA7. Host execution has not yet been
validated. EU CRR standard NSFR is the first planned implementation profile;
Basel material is a conceptual reference, not an interchangeable rule set.

- [Scope, users, profiles, and exclusions](docs/SCOPE.md)
- [Architecture, host targets, and repository layout](docs/ARCHITECTURE.md)
- [Decision log](docs/DECISIONS.md)
- [Follow-on implementation backlog](docs/BACKLOG.md)
- [Existing governance and local checks](docs/GOVERNANCE.md)
- [Licensing and distribution](docs/LICENSING.md)
- [Provenance and attribution](docs/PROVENANCE.md)
- [Synthetic-data policy](docs/DATA_POLICY.md)
- [Security reporting](SECURITY.md)

Development targets `release/0.1.0` through task pull requests. Integration into
`main` follows the owner decision and synchronization process in
[release governance](docs/GOVERNANCE.md#branches-and-integration).
