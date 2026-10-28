# Security policy

K-NSFR is private during setup and is intended to become public. Report
suspected vulnerabilities privately to the repository owner through a private
GitHub channel available to you, or contact the owner outside the repository.
Before publication, the owner must verify that a usable private reporting route
is documented and available to external reporters. Do not open a public issue or
attach exploit details, credentials, client data, or personal data to an issue or
PR. The owner will acknowledge, triage, and coordinate remediation and any
necessary disclosure; no response-time or security-support guarantee is claimed.

Never commit secrets, tokens, connection strings, real customer records, or
workbooks containing them. Use the least-privilege credential store for local
work, review staged changes, and keep diagnostic logs free of secrets. If a
secret or sensitive record is exposed, stop sharing it, alert the owner privately,
revoke/rotate affected credentials, assess exposure and remove the material from
active artifacts. History cleanup requires a separate coordinated decision; a
normal deletion commit does not erase previous versions.

GitHub Actions has a read-only default token and SHA-pinned actions; these
configuration controls do not constitute a security certification. Dependency
alerts and host validation have separate evidence requirements. See
[synthetic-data policy](docs/DATA_POLICY.md) and [provenance](docs/PROVENANCE.md).
