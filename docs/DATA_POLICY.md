# 🧪 Synthetic-data policy

**K-NSFR** · [Overview](../README.md) · [Documentation](../README.md#documentation) · [Governance](GOVERNANCE.md)

[![Stage](https://img.shields.io/badge/Stage-Repository_setup-BF8700?style=flat-square)](../README.md#status)

## ✅ Allowed data

Repository issues, source, fixtures, workbooks, test outputs, logs, screenshots,
and PR attachments may contain only synthetic or suitably non-identifying data.
Construct examples independently; avoid copying or lightly masking real client,
employer, employee, customer, account, or transaction records. Fictional names,
identifiers, dates, and amounts must not be reversible to real cases. Record
fixture generation method and expected-result provenance when numerical cases
are introduced. No production feeds, credentials, or proprietary templates are
admitted to this setup repository.

## 🔍 Review before committing

Before adding a fixture or artifact, inspect its content and metadata, including
hidden workbook sheets, comments, author fields, external links, embedded
objects, and cached data. Treat an unexpected real-world match or uncertain
source as a blocker: keep it outside the repo and ask the owner for private
review. Do not reproduce questionable material in an issue or PR. The setup
baseline contains no calculation fixtures or workbook.
