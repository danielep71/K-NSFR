<a id="nsfr-benchmark-prototype"></a>

# 🧪 NSFR benchmark prototype (not admitted)

**K-NSFR** · [Overview](../README.md) · [Documentation](../README.md#documentation) · [Governance](../docs/GOVERNANCE.md)

[![Stage](https://img.shields.io/badge/Stage-Repository_setup-BF8700?style=flat-square)](../README.md#status)

> [!WARNING]
> Nothing in this directory is admitted ([ADR-011](../docs/DECISIONS.md#recorded-decisions)).
> It is my benchmark prototype, stored as I delivered it. Its factors are
> transcribed from the CRR and are not verified against the
> [source register](../docs/methodology/SOURCE_REGISTER.md). Its VBA has run in
> LibreOffice only, never in Excel. Do not use it for regulatory reporting, and
> do not cite its values as expected results.

The prototype recalculates the EU CRR NSFR from positions and reconciles it
with a bank's C 80.00, C 81.00 and C 84.00 figures. Its
[specification](docs/methodology/PROTOTYPE.md) describes the calculation, the
reconciliation, the evidence and its departures from the recorded conventions.
The milestone v0.2.0 - Prototype Setup brings it into line piece by piece,
each piece under its own issue.

<a id="contents"></a>

## 🗂️ Contents

| Path | Holds |
| --- | --- |
| `src/core/`, `src/api/`, `src/host/` | The seven VBA modules of the prototype |
| `tools/build_template.py` | Generates the macro-free workbook `templates/NSFR_Template.xlsx`, which is ignored and never committed |
| `templates/README.md` | Keeps the output directory, which the generator does not create |
| `tools/lo_regression.py` | Headless regression of the VBA in LibreOffice |
| `tests/derivations/prototype/` | Transcribed regulatory tables, the synthetic demo, demo bank figures, expected values and the independent Python reference |
| `docs/methodology/PROTOTYPE.md` | The specification |

The paths in the specification are relative to this directory. Run its
commands from here, for example `python tools/build_template.py`.

<a id="how-the-repository-treats-it"></a>

## 🧭 How the repository treats it

- The byte and Markdown checks of Repository integrity apply here as to
  every other file.
- The VBA static checks skip this directory, and only this directory
  (`NOT_ADMITTED` in `tools/check_vba.py`, with a test). Ruff and mypy check
  `tools/` only, and the contract and fixture checks read only
  `docs/contracts/` and `tests/fixtures/`, so the prototype's Python and data
  are outside them too.
- Its dependencies (`xlsxwriter`, `openpyxl`, LibreOffice) are not locked or
  installed by CI; their licenses are listed in
  [licensing](../docs/LICENSING.md#prototype-tool-licenses).
- Only the relative links of the specification were changed, so that they
  resolve from this directory, and `templates/README.md` was added so that
  the output directory exists; every other file is byte for byte as
  delivered, apart from Git's line-end normalization of the `.bas` files.
