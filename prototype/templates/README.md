<a id="prototype-workbook-output"></a>

# 🗂️ Prototype workbook output

**K-NSFR** · [Overview](../../README.md) · [Documentation](../../README.md#documentation) · [Governance](../../docs/GOVERNANCE.md)

[![Stage](https://img.shields.io/badge/Stage-Repository_setup-BF8700?style=flat-square)](../../README.md#status)

`python tools/build_template.py`, run from `prototype/`, writes
`NSFR_Template.xlsx` into this directory, and `tools/lo_regression.py` reads it
from here. The workbook is generated output: `.gitignore` excludes it and it is
never committed. This file only keeps the directory in a fresh checkout,
because the generator does not create it. The prototype is not admitted
([ADR-011](../../docs/DECISIONS.md#recorded-decisions)).
