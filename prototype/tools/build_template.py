"""Build the macro-free K-NSFR prototype workbook template.

Writes templates/NSFR_Template.xlsx from the regulatory tables, the synthetic
demo and its expected values in tests/derivations/prototype/, in the layout
that src/core/NSFR_Core_Config.bas and src/host/NSFR_Host_Main.bas describe:
every table is located through the workbook name of its first header cell
(hdr*), every setting through the workbook name equal to its code. The
template holds no VBA: save it as .xlsm and import the modules of src/ (see
docs/methodology/PROTOTYPE.md). The sheet buttons call the host macros by name.

Usage: python tools/build_template.py
"""

from __future__ import annotations

import datetime as dt
import json
from pathlib import Path
from typing import Any

import xlsxwriter  # type: ignore[import-untyped,import-not-found,unused-ignore]

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "templates/NSFR_Template.xlsx"
DER = ROOT / "tests/derivations/prototype"
TABLES = json.loads((DER / "regulatory-tables.json").read_text())
DEMO = json.loads((DER / "demo-positions.json").read_text())
BANK = json.loads((DER / "demo-bank-cells.json").read_text())
EXPECTED = json.loads((DER / "demo.expected.json").read_text())

NAVY, BLUE, BAND, GOLD, INPUT, GREY = "#1F3864", "#2E75B6", "#265A99", "#FFC000", "#FFF2CC", "#BFBFBF"
TAB_DOC, TAB_IN, TAB_OUT, TAB_TEST = "#1F3864", "#BF9000", "#2E75B6", "#548235"
FONT = "Aptos"
INPUT_ROWS = 5000

wb = xlsxwriter.Workbook(str(OUT))
wb.set_properties({"title": "K-NSFR prototype", "author": "", "company": "", "comments": ""})
wb.formats[0].set_font_name(FONT)
wb.formats[0].set_font_size(10)
_cache: dict[tuple[tuple[str, Any], ...], Any] = {}


def fmt(**kw: Any) -> Any:
    key = tuple(sorted(kw.items()))
    if key not in _cache:
        base: dict[str, Any] = {"font_name": FONT, "valign": "vcenter"}
        base.update(kw)
        _cache[key] = wb.add_format(base)
    return _cache[key]


F: dict[str, Any] = {
    "brand": fmt(font_size=14, font_color=GOLD, bg_color="#000000", align="left", indent=1),
    "band1": fmt(bg_color="#000000"),
    "title": fmt(font_size=18, bold=True, font_color="#FFFFFF", bg_color=BAND, align="left", indent=1),
    "band2": fmt(bg_color=BAND),
    "note": fmt(font_size=10, italic=True, font_color="#595959"),
    "runinfo": fmt(font_size=10, italic=True, font_color="#7F6000"),
    "section": fmt(font_size=14, bold=True, font_color=NAVY, bottom=2, bottom_color=NAVY),
    "hdr_in": fmt(font_size=9, bold=True, font_color="#FFFFFF", bg_color=NAVY, text_wrap=True, border=1,
                  border_color=GREY, align="center"),
    "hdr_out": fmt(font_size=9, bold=True, font_color="#FFFFFF", bg_color=BLUE, text_wrap=True, border=1,
                   border_color=GREY, align="center"),
    "hdr_grp": fmt(font_size=9, bold=True, font_color=NAVY, bg_color="#D9E1F2", text_wrap=True, border=1,
                   border_color=GREY, align="center"),
    "code": fmt(font_size=9, bold=True, num_format="@"),
    "desc": fmt(font_size=9),
    "desc_b": fmt(font_size=9, bold=True),
    "desc_wrap": fmt(font_size=9, text_wrap=True),
    "ref": fmt(font_size=8, font_color="#595959"),
    "in_txt": fmt(font_size=9, bg_color=INPUT, border=1, border_color=GREY, num_format="@"),
    "in_ctr": fmt(font_size=9, bg_color=INPUT, border=1, border_color=GREY, num_format="@", align="center"),
    "in_amt": fmt(font_size=9, bg_color=INPUT, border=1, border_color=GREY, num_format="#,##0.00"),
    "in_gen": fmt(font_size=9, bg_color=INPUT, border=1, border_color=GREY, num_format="General", align="center"),
    "in_pct": fmt(font_size=9, bg_color=INPUT, border=1, border_color=GREY, num_format="0.0%", align="center"),
    "in_date": fmt(font_size=9, bg_color=INPUT, border=1, border_color=GREY, num_format="yyyy-mm-dd",
                   align="center"),
    "fix_pct": fmt(font_size=9, border=1, border_color=GREY, num_format="0.0%", align="center",
                   bg_color="#F2F2F2"),
    "grey": fmt(bg_color="#D9D9D9", border=1, border_color=GREY),
    "out_amt": fmt(font_size=9, num_format="#,##0;-#,##0;-", border=1, border_color=GREY),
    "out_amt_b": fmt(font_size=9, num_format="#,##0;-#,##0;-", border=1, border_color=GREY, bold=True),
    "out_pct": fmt(font_size=9, num_format="0.0%", align="center", border=1, border_color=GREY),
    "out_pct2": fmt(font_size=10, num_format="0.00%", bold=True, align="right", border=1, border_color=GREY),
    "out_txt": fmt(font_size=9, num_format="@", border=1, border_color=GREY),
    "out_big": fmt(font_size=11, bold=True, num_format="#,##0", align="right", border=1, border_color=GREY),
    "out_bigtxt": fmt(font_size=11, bold=True, align="right", border=1, border_color=GREY),
    "out_date": fmt(font_size=10, num_format="yyyy-mm-dd", align="right", border=1, border_color=GREY),
    "lbl": fmt(font_size=10, bold=True, border=1, border_color=GREY, bg_color="#F2F2F2"),
    "rm_h": fmt(font_size=14, bold=True, font_color=NAVY, bottom=2, bottom_color=NAVY, text_wrap=True),
    "rm_k": fmt(font_size=10, bold=True, text_wrap=True, valign="vcenter", bottom=3, bottom_color=GREY),
    "rm_v": fmt(font_size=9, text_wrap=True, valign="vcenter", bottom=3, bottom_color=GREY),
    "cf_bad": wb.add_format({"bg_color": "#F8CBAD"}),
    "cf_warn": wb.add_format({"bg_color": "#FFE699"}),
    "cf_pass": wb.add_format({"bg_color": "#C6EFCE", "font_color": "#006100"}),
    "cf_fail": wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"}),
}


def banner(ws: Any, title: str, last_col: int) -> None:
    """Brand and title bands over B:last_col, as in the other K- workbooks."""
    ws.hide_gridlines(2)
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 2.5)
    ws.set_row(0, 24)
    ws.set_row(1, 26.1)
    for c in range(1, last_col + 1):
        ws.write_blank(0, c, None, F["band1"])
        ws.write_blank(1, c, None, F["band2"])
    ws.write(0, 1, "K-NSFR", F["brand"])
    ws.write(1, 1, title, F["title"])


def cell_ref(row: int, col: int) -> str:
    """Absolute A1 reference of a zero-based cell."""
    letters = ""
    col += 1
    while col:
        col, rem = divmod(col - 1, 26)
        letters = chr(65 + rem) + letters
    return f"${letters}${row + 1}"


def name(nm: str, ws: Any, row: int, col: int, row2: int | None = None, col2: int | None = None) -> None:
    ref = cell_ref(row, col)
    if row2 is not None and col2 is not None:
        ref += ":" + cell_ref(row2, col2)
    wb.define_name(nm, f"='{ws.name}'!{ref}")


def buttons(ws: Any, row: int, col: int, items: list[tuple[str, str]]) -> None:
    for i, (caption, macro) in enumerate(items):
        ws.insert_button(row, col, {"macro": macro, "caption": caption, "width": 132, "height": 24,
                                    "x_offset": 4 + i * 140, "y_offset": 4})


def iso(text: str) -> dt.datetime | None:
    return dt.datetime.fromisoformat(text) if text else None


ROWS = TABLES["rows"]
CATS = TABLES["categories"]
SUMM = TABLES["summary"]
P = DEMO["params"]
OVR = P["applied_overrides"]

# ----------------------------------------------------------------------------
# README
# ----------------------------------------------------------------------------
readme = wb.add_worksheet("README")
readme.set_tab_color(TAB_DOC)
banner(readme, "ReadMe", 4)
readme.set_column(2, 2, 34)
readme.set_column(3, 3, 150)
readme.write(2, 2, "Prototype NSFR benchmark: recalculate the net stable funding ratio of the EU CRR from "
             "positions, and reconcile a bank's C 80.00 / C 81.00 / C 84.00 figures cell by cell.", F["note"])
README = [
    ("Purpose", None),
    ("What it does", "An independent recalculation of the NSFR (CRR Part Six, Title IV, Art. 428a-428ah, as amended "
     "by Regulation (EU) 2019/876) to benchmark a bank's own calculation. Two uses, separately or together: "
     "(1) a positions run - classify each position into a row and maturity bucket of C 80.00 / C 81.00, weight "
     "it with the CRR factor and derive C 84.00 and the NSFR; (2) a reconciliation of the bank's reported "
     "templates - arithmetic re-performance, factor conformity against the CRR, and, with a positions run, "
     "amounts and stable funding cell by cell."),
    ("Attribution", "Reconciliation splits the NSFR difference into: arithmetic (bank amounts x bank factors, "
     "summed bottom-up, minus what the bank reported), factor effect (bank amounts x CRR factors minus bank "
     "amounts x bank factors) and amounts / classification (benchmark minus restated)."),
    ("How to use", None),
    ("1. Enable macros", "This file is the macro-free template. Save it as an Excel Macro-Enabled Workbook "
     "(.xlsm), import the .bas modules of src/core, src/api and src/host (Alt+F11, File > Import File), "
     "save, reopen and click Enable Content."),
    ("2. Params", "Reporting date, currency, perimeter and profile (eu-crr-standard). The factor table lists every "
     "row of C 80.00 and C 81.00 with its regulatory factors (grey, transcribed from the CRR) and the applied "
     "factors (yellow, used by the run). An applied factor that differs makes the run WHAT-IF. Rows 1.6, 1.10.1, "
     "1.10.5 and 2.4 take factors set by the competent authority: enter them as applied factors. The category "
     "table maps each classification code to template rows; the summary map builds C 84.00."),
    ("3. Positions", "One row per position (balance-sheet item or off-balance-sheet commitment), in units of the "
     "reporting currency: category, carrying amount, maturity date (blank = the category's undated bucket), "
     "date until which an asset is encumbered, an optional bucket override (<6M, 6M-1Y, >=1Y, e.g. for "
     "options under Art. 428j) and an optional factor override (WHAT-IF). Derivatives: enter positive and "
     "negative fair values per netting set and the variation margin; the engine nets them."),
    ("4. Bank figures", "BankCells holds the bank's figures as (template, row, column, value). Import bank "
     "workbook reads every visible sheet titled C 80.00, C 81.00 or C 84.00 laid out as the ITS templates "
     "(column codes 0010... in one header row, row codes in one column). Import bank CSV reads the export "
     "format RowColumn_Index,Suffix,Amount (e.g. 'C 80.00 c0010 r0040,a,123.45'). AmountScale converts bank "
     "amounts to units (1000 for thousands)."),
    ("5. Run", "Run benchmark (macro RunNSFR) rebuilds Results, C80, C81, PositionCalc, Reconciliation and Checks. "
     "A rejected position withholds the benchmark totals (INVALID); the bank checks still run."),
    ("Conventions", None),
    ("Buckets", "Residual maturity < 6 months if the maturity date is before the reporting date + 6 calendar "
     "months; 6 months to < 1 year if before + 12 months; otherwise 1 year or more. HQLA assets go to the HQLA "
     "column of C 80.00. The encumbrance end date selects the row (unencumbered or < 6 months / 6 months to "
     "< 1 year / 1 year or more)."),
    ("Derivatives", "NSFR derivative assets = max(0, positive fair values - variation margin received); "
     "liabilities = max(0, negative fair values - variation margin posted). The excess of assets takes 100% RSF "
     "(1.7.2), an excess of liabilities 0% ASF (2.7); 5% RSF on the gross negative fair value (1.7.1) "
     "[Art. 428d, 428s(2), 428ah(2), 428k(3)(a)]."),
    ("Memo rows", "'of which' rows (residential mortgages, retail bonds, operational deposits...) receive the "
     "amount of the categories that name them, weighted with the main row's factor, and stay out of totals."),
    ("Known template gap", "C 81.00 has no 6M-1Y factor cell for AT1, Tier 2 and other capital instruments; the "
     "CRR gives 50% (Art. 428l(d); EBA Q&A 2021_6017). The benchmark applies 50% and the FACTOR check names it."),
    ("Sheets", None),
    ("Params / Positions / BankCells", "Inputs (yellow cells)."),
    ("Results", "Status, run class, ASF, RSF, NSFR and C 84.00."),
    ("C80 / C81", "Benchmark replica of C 80.00 and C 81.00: amounts, regulatory factor, applied factor "
     "(effective, where amounts exist) and stable funding."),
    ("PositionCalc", "Every position: row, bucket, encumbrance, factor, stable funding, status."),
    ("Reconciliation", "Attribution and every check, cell by cell: filter Result for DIFF and MISSING; sort "
     "Impact for the factor differences that matter."),
    ("Checks / TestCatalogue", "Diagnostics of the last run; regression of the demo against the independent "
     "Python reference (tests/derivations/prototype/nsfr_reference.py)."),
    ("Status", None),
    ("Prototype", "Not for regulatory reporting. Factors transcribed from the CRR and cross-checked with the "
     "standard-factor columns of the reporting templates, not yet against the registered sources of the "
     "project. Simplified NSFR (C 82.00 / C 83.00), consolidation, currencies and national discretions are "
     "out of scope. The demo data are synthetic."),
]
r = 3
for key, value in README:
    if value is None:
        readme.set_row(r, 8)
        r += 1
        readme.set_row(r, 24)
        readme.write(r, 2, key, F["rm_h"])
        readme.write_blank(r, 3, None, F["rm_h"])
    else:
        readme.set_row(r, 15 * max(1, (len(value) // 165) + 1) + 4)
        readme.write(r, 2, key, F["rm_k"])
        readme.write(r, 3, value, F["rm_v"])
    r += 1
readme.freeze_panes(3, 0)

# ----------------------------------------------------------------------------
# PARAMS
# ----------------------------------------------------------------------------
params = wb.add_worksheet("Params")
params.set_tab_color(TAB_IN)
banner(params, "Parameters", 19)
for c, w in {2: 10, 3: 8, 4: 11, 5: 52, 6: 7, 7: 8, 8: 8}.items():
    params.set_column(c, c, w)
params.set_column(9, 16, 8.5)
params.set_column(17, 17, 40)
params.set_column(18, 18, 24)
params.write(3, 2, "Run settings", F["section"])
SETTINGS = [
    ("ReportingDate", "Reporting (as-of) date", iso(P["ReportingDate"]), "in_date"),
    ("ReportingCurrency", "Reporting currency (informative: amounts are taken as entered)", P["ReportingCurrency"],
     "in_ctr"),
    ("Perimeter", "Reporting perimeter (informative)", P["Perimeter"], "in_txt"),
    ("ProfileID", "Profile: eu-crr-standard, the only one the prototype runs", P["ProfileID"], "in_ctr"),
]
for i, (code, desc, val, f) in enumerate(SETTINGS):
    rr = 4 + i
    params.write(rr, 2, code, F["code"])
    params.merge_range(rr, 3, rr, 4, val if not isinstance(val, dt.datetime) else "", F[f])
    if isinstance(val, dt.datetime):
        params.write_datetime(rr, 3, val, F[f])
    params.write(rr, 5, desc, F["desc"])
    name(code, params, rr, 3)
params.write(8, 2, "Version", F["code"])
params.write(8, 3, TABLES["parameter_version"], F["ref"])
params.data_validation(7, 3, 7, 3, {"validate": "list", "source": ["eu-crr-standard"]})

# Factor table
FT = 11
params.write(FT - 1, 2, "Factor table - C 80.00 and C 81.00 rows", F["section"])
FHEAD = ["Template", "Row", "ID", "Item", "Kind", "Amount cells", "Factor cells",
         "Reg <6M", "Reg 6M-1Y", "Reg >=1Y", "Reg HQLA", "Applied <6M", "Applied 6M-1Y", "Applied >=1Y",
         "Applied HQLA", "Reference", "Children"]
for c, h in enumerate(FHEAD):
    params.write(FT, 2 + c, h, F["hdr_in"])
name("hdrFactors", params, FT, 2)
params.set_row(FT, 30)
for i, row in enumerate(ROWS):
    rr = FT + 1 + i
    applied = OVR.get(f"{row['template']}|{row['id']}", row["regulatory"])
    vals = [row["template"], row["row"], row["id"], row["item"], row["kind"], row["amount_buckets"],
            row["factor_buckets"]]
    for c, v in enumerate(vals):
        params.write_string(rr, 2 + c, v, F["desc_b"] if (row["kind"] == "AGG" and c == 5) else
                            (F["code"] if c < 3 else F["desc"]))
    for b in range(4):
        cell_ok = "123H"[b] in row["amount_buckets"]
        rv = row["regulatory"][b]
        if rv is not None:
            params.write_number(rr, 9 + b, rv, F["fix_pct"])
        elif row["kind"] in ("LEAF", "APPL") and cell_ok:
            params.write_blank(rr, 9 + b, None, F["fix_pct"])
        else:
            params.write_blank(rr, 9 + b, None, F["grey"])
        av = applied[b]
        if row["kind"] in ("LEAF", "APPL") and cell_ok:
            if av is None:
                params.write_blank(rr, 13 + b, None, F["in_pct"])
            else:
                params.write_number(rr, 13 + b, av, F["in_pct"])
        else:
            params.write_blank(rr, 13 + b, None, F["grey"])
    params.write_string(rr, 17, row["reference"], F["ref"])
    params.write_string(rr, 18, ",".join(row["children"]), F["ref"])
FT_END = FT + len(ROWS)
params.data_validation(FT + 1, 13, FT_END, 16, {"validate": "decimal", "criteria": "between", "minimum": 0,
                                                "maximum": 1, "error_title": "Factor",
                                                "error_message": "Enter a decimal from 0 to 1 (5% = 0.05)."})
params.conditional_format(FT + 1, 13, FT_END, 16, {
    "type": "formula", "criteria": f"=AND(ISNUMBER(J{FT + 2}),ISNUMBER(N{FT + 2}),ABS(N{FT + 2}-J{FT + 2})>1E-12)",
    "format": F["cf_warn"]})

# Category table
CT = FT_END + 4
params.write(CT - 1, 2, "Category table - classification of positions", F["section"])
CHEAD = ["Category", "Description", "Template", "Bucket rule", "Undated bucket", "Row unencumbered or <6M",
         "Row encumbered 6M-1Y", "Row encumbered >=1Y", "Memo row", "Reference"]
CCOLS = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11]
for c, h in zip(CCOLS, CHEAD):
    params.write(CT, c, h, F["hdr_in"])
params.set_row(CT, 30)
name("hdrCategories", params, CT, 2)
for i, cat in enumerate(CATS):
    rr = CT + 1 + i
    vals = [cat["category"], cat["description"], cat["template"], cat["bucket_rule"], cat["undated"],
            cat["row_ids"][0], cat["row_ids"][1], cat["row_ids"][2], cat["memo_id"], cat["reference"]]
    for c, v in zip(CCOLS, vals):
        params.write_string(rr, c, v, F["code"] if c == 2 else F["desc"])
CT_END = CT + len(CATS)
wb.define_name("CategoryList", f"='Params'!$C${CT + 2}:$C${CT_END + 1}")

# Summary map
ST = CT_END + 4
params.write(ST - 1, 2, "Summary map - C 84.00", F["section"])
SHEAD = ["Row", "ID", "Item", "Measure", "Plus (Template|ID; ...)", "Minus"]
for c, h in enumerate(SHEAD):
    params.write(ST, 2 + c, h, F["hdr_in"])
name("hdrSummary", params, ST, 2)
for i, s in enumerate(SUMM):
    rr = ST + 1 + i
    for c, v in enumerate([s["row"], s["id"], s["item"], s["measure"], ";".join(s["plus"]), ";".join(s["minus"])]):
        params.write_string(rr, 2 + c, v, F["code"] if c < 2 else F["desc"])
params.freeze_panes(FT + 1, 5)

# ----------------------------------------------------------------------------
# POSITIONS
# ----------------------------------------------------------------------------
pos = wb.add_worksheet("Positions")
pos.set_tab_color(TAB_IN)
banner(pos, "Positions", 11)
PHEAD = ["PositionID", "Description", "Category", "Amount", "Currency", "MaturityDate", "EncumberedUntil",
         "BucketOverride", "FactorOverride", "Comment"]
PW = [11, 44, 22, 17, 8, 12, 13, 10, 10, 34]
for c, (h, w) in enumerate(zip(PHEAD, PW)):
    pos.set_column(2 + c, 2 + c, w)
    pos.write(4, 2 + c, h, F["hdr_in"])
pos.set_row(4, 28)
name("hdrPositions", pos, 4, 2)
pos.write(2, 2, "One row per position, amounts in units of the reporting currency (>= 0). Category: see Params. "
          "Blank maturity = the category's undated bucket.", F["note"])
buttons(pos, 3, 7, [("Run benchmark", "RunNSFR")])
pos.set_row(3, 28)
PF = ["in_txt", "in_txt", "in_txt", "in_amt", "in_ctr", "in_date", "in_date", "in_ctr", "in_pct", "in_txt"]
for rr in range(5, 5 + INPUT_ROWS):
    for c, f in enumerate(PF):
        pos.write_blank(rr, 2 + c, None, F[f])
for i, p in enumerate(DEMO["positions"]):
    rr = 5 + i
    for c, v in enumerate(p):
        if c in (5, 6):
            if v:
                pos.write_datetime(rr, 2 + c, iso(v), F["in_date"])
        elif c == 3:
            pos.write_number(rr, 2 + c, v, F["in_amt"])
        elif c == 8:
            if v is not None:
                pos.write_number(rr, 2 + c, v, F["in_pct"])
        elif v:
            pos.write_string(rr, 2 + c, v, F[PF[c]])
pos.data_validation(5, 4, 5 + INPUT_ROWS, 4, {"validate": "list", "source": "=CategoryList",
                                              "error_title": "Category", "error_message": "Choose a category of Params."})
pos.data_validation(5, 9, 5 + INPUT_ROWS, 9, {"validate": "list", "source": ["<6M", "6M-1Y", ">=1Y"]})
pos.freeze_panes(5, 3)
pos.autofilter(4, 2, 4 + len(DEMO["positions"]), 11)

# ----------------------------------------------------------------------------
# BANK CELLS
# ----------------------------------------------------------------------------
bank = wb.add_worksheet("BankCells")
bank.set_tab_color(TAB_IN)
banner(bank, "Bank figures", 8)
for c, w in {2: 16, 3: 14, 4: 10, 5: 22, 6: 40, 7: 4, 8: 4}.items():
    bank.set_column(c, c, w)
bank.write(2, 2, "The bank's own C 80.00 / C 81.00 / C 84.00 figures, one cell per row. Import them or paste "
           "them; the demo holds figures with deliberate differences.", F["note"])
BSET = [("AmountScale", P["AmountScale"], "Units of the reporting currency per bank amount: 1, 1000 (thousands)..."),
        ("AbsTolerance", P["AbsTolerance"], "In units of the reporting currency."),
        ("RelTolerance", P["RelTolerance"], "Of the absolute bank figure (not applied when it is 0)."),
        ("FactorTolerance", P["FactorTolerance"], "Absolute, for factors and the ratio.")]
for i, (code, val, desc) in enumerate(BSET):
    rr = 4 + i
    bank.write(rr, 2, code, F["code"])
    bank.write_number(rr, 3, val, F["in_gen"])
    bank.write(rr, 4, desc, F["note"])
    name(code, bank, rr, 3)
buttons(bank, 3, 5, [("Import bank workbook", "ImportBankWorkbook"), ("Import bank CSV", "ImportBankCsv")])
buttons(bank, 5, 5, [("Clear bank figures", "ClearBankCells"), ("Run benchmark", "RunNSFR")])
BH = 9
for c, h in enumerate(["Template", "Row", "Column", "Value", "Source"]):
    bank.write(BH, 2 + c, h, F["hdr_in"])
name("hdrBankCells", bank, BH, 2)
for rr in range(BH + 1, BH + 1 + INPUT_ROWS):
    for c, f in enumerate(["in_ctr", "in_ctr", "in_ctr", "in_amt", "in_txt"]):
        bank.write_blank(rr, 2 + c, None, F[f])
for i, (t, rw, col, val) in enumerate(BANK["cells"]):
    rr = BH + 1 + i
    bank.write_string(rr, 2, t, F["in_ctr"])
    bank.write_string(rr, 3, rw, F["in_ctr"])
    bank.write_string(rr, 4, col, F["in_ctr"])
    bank.write_number(rr, 5, val, F["in_amt"])
    bank.write_string(rr, 6, "demo", F["in_txt"])
bank.freeze_panes(BH + 1, 0)
bank.autofilter(BH, 2, BH + len(BANK["cells"]), 6)

# ----------------------------------------------------------------------------
# RESULTS
# ----------------------------------------------------------------------------
res = wb.add_worksheet("Results")
res.set_tab_color(TAB_OUT)
banner(res, "Results", 9)
res.set_column(2, 2, 22)
res.set_column(3, 3, 22)
res.set_column(4, 4, 62)
res.set_column(5, 8, 17)
res.write(3, 2, "No results yet: press Run benchmark.", F["runinfo"])
name("resRunInfo", res, 3, 2)
buttons(res, 2, 5, [("Run benchmark", "RunNSFR"), ("Clear outputs", "ClearOutputs")])
res.set_row(2, 28)
LABELS = ["Status", "Run class", "Reporting date", "Perimeter", "Available stable funding (ASF)",
          "Required stable funding (RSF)", "NSFR", "Positions accepted", "Positions rejected"]
RFMT = ["out_bigtxt", "out_bigtxt", "out_date", "out_bigtxt", "out_big", "out_big", "out_pct2", "out_big", "out_big"]
for i, (lbl, f) in enumerate(zip(LABELS, RFMT)):
    res.write(5 + i, 2, lbl, F["lbl"])
    res.write_blank(5 + i, 3, None, F[f])
name("resValues", res, 5, 3, 13, 3)
res.conditional_format(5, 3, 5, 3, {"type": "text", "criteria": "begins with", "value": "VALID",
                                     "format": F["cf_pass"]})
res.conditional_format(5, 3, 5, 3, {"type": "text", "criteria": "begins with", "value": "INVALID",
                                     "format": F["cf_fail"]})
res.conditional_format(5, 3, 5, 3, {"type": "text", "criteria": "begins with", "value": "STOPPED",
                                     "format": F["cf_fail"]})
res.conditional_format(6, 3, 6, 3, {"type": "text", "criteria": "begins with", "value": "WHAT-IF",
                                     "format": F["cf_warn"]})
SH = 16
res.write(SH - 1, 2, "C 84.00 - NSFR summary (benchmark)", F["section"])
for c, h in enumerate(["Row", "ID", "Item", "Amount (0010)", "RSF (0020)", "ASF (0030)", "Ratio (0040)"]):
    res.write(SH, 2 + c, h, F["hdr_out"])
for i, s in enumerate(SUMM):
    rr = SH + 1 + i
    bold = s["id"] in ("1", "2", "3")
    res.write_string(rr, 2, s["row"], F["code"])
    res.write_string(rr, 3, s["id"], F["code"])
    res.write_string(rr, 4, s["item"], F["desc_b"] if bold else F["desc"])
    for c in range(4):
        res.write_blank(rr, 5 + c, None, F["out_pct2"] if c == 3 else (F["out_amt_b"] if bold else F["out_amt"]))
name("resSummaryValues", res, SH + 1, 5, SH + len(SUMM), 8)
res.freeze_panes(3, 0)

# ----------------------------------------------------------------------------
# C80 / C81 replicas
# ----------------------------------------------------------------------------
def replica(sheet_name: str, template: str, buckets: list[str], title: str, value_name: str) -> None:
    ws = wb.add_worksheet(sheet_name)
    ws.set_tab_color(TAB_OUT)
    nb = len(buckets)
    last = 2 + 3 + 3 * nb + 1
    banner(ws, title, last)
    ws.set_column(2, 3, 8)
    ws.set_column(4, 4, 58)
    ws.set_column(5, 4 + nb, 15)
    ws.set_column(5 + nb, 4 + 3 * nb, 8)
    ws.set_column(5 + 3 * nb, 5 + 3 * nb, 16)
    ws.write(2, 2, f"Benchmark replica of {template}: amounts and stable funding of the last positions run; "
             "applied factor = stable funding / amount where amounts exist.", F["note"])
    groups = [("Amount", 5, 4 + nb), ("Regulatory factor", 5 + nb, 4 + 2 * nb),
              ("Applied factor", 5 + 2 * nb, 4 + 3 * nb)]
    for label, c1, c2 in groups:
        ws.merge_range(4, c1, 4, c2, label, F["hdr_grp"])
    ws.merge_range(4, 5 + 3 * nb, 5, 5 + 3 * nb, "Stable funding", F["hdr_grp"])
    for c, h in enumerate(["Row", "ID", "Item"]):
        ws.merge_range(4, 2 + c, 6, 2 + c, h, F["hdr_out"])
    acodes = [f"{10 * (b + 1):04d}" for b in range(nb)]
    first_factor = 9 if template == "C 80.00" else 7
    fcodes_std = [f"{10 * (nb + 1 + b):04d}" for b in range(nb)]
    fcodes_app = [f"{10 * (first_factor + b):04d}" for b in range(nb)]
    scode = "0130" if template == "C 80.00" else "0100"
    for b in range(nb):
        ws.write(5, 5 + b, buckets[b], F["hdr_out"])
        ws.write(5, 5 + nb + b, buckets[b], F["hdr_out"])
        ws.write(5, 5 + 2 * nb + b, buckets[b], F["hdr_out"])
        ws.write(6, 5 + b, acodes[b], F["hdr_out"])
        ws.write(6, 5 + nb + b, fcodes_std[b], F["hdr_out"])
        ws.write(6, 5 + 2 * nb + b, fcodes_app[b], F["hdr_out"])
    ws.write(6, 5 + 3 * nb, scode, F["hdr_out"])
    rows = [r_ for r_ in ROWS if r_["template"] == template]
    for i, row in enumerate(rows):
        rr = 7 + i
        bold = row["kind"] == "AGG" and row["id"].count(".") <= 1
        ws.write_string(rr, 2, row["row"], F["code"])
        ws.write_string(rr, 3, row["id"], F["code"])
        indent = "   " * row["id"].count(".")
        ws.write_string(rr, 4, indent + row["item"], F["desc_b"] if bold else F["desc"])
        for b in range(nb):
            has = "123H"[b] in row["amount_buckets"]
            ws.write_blank(rr, 5 + b, None, (F["out_amt_b"] if bold else F["out_amt"]) if has else F["grey"])
            fac = row["kind"] in ("LEAF", "APPL", "MEMO") and has
            ws.write_blank(rr, 5 + nb + b, None, F["out_pct"] if fac else F["grey"])
            ws.write_blank(rr, 5 + 2 * nb + b, None, F["out_pct"] if fac else F["grey"])
        ws.write_blank(rr, 5 + 3 * nb, None, F["out_amt_b"] if bold else F["out_amt"])
    name(value_name, ws, 7, 5, 6 + len(rows), 5 + 3 * nb)
    ws.freeze_panes(7, 5)


replica("C80", "C 80.00", ["<6M", "6M-1Y", ">=1Y", "HQLA"], "C 80.00 - Required stable funding", "c80Values")
replica("C81", "C 81.00", ["<6M", "6M-1Y", ">=1Y"], "C 81.00 - Available stable funding", "c81Values")

# ----------------------------------------------------------------------------
# POSITION CALC
# ----------------------------------------------------------------------------
pc = wb.add_worksheet("PositionCalc")
pc.set_tab_color(TAB_OUT)
banner(pc, "Position calculation", 13)
POH = ["PositionID", "Category", "Template", "Row", "Row ID", "Bucket", "Encumbrance", "Amount", "Factor",
       "Stable funding", "Memo row", "Status"]
POW = [11, 22, 9, 7, 10, 8, 18, 17, 8, 17, 10, 70]
POF = ["out_txt", "out_txt", "out_txt", "out_txt", "out_txt", "out_txt", "out_txt", "out_amt", "out_pct",
       "out_amt", "out_txt", "out_txt"]
for c, (h, w, f) in enumerate(zip(POH, POW, POF)):
    pc.set_column(2 + c, 2 + c, w, F[f])
    pc.write(4, 2 + c, h, F["hdr_out"])
name("hdrPositionCalc", pc, 4, 2)
pc.write(2, 2, "Classification and weighting of every position of the last run.", F["note"])
pc.freeze_panes(5, 3)
pc.autofilter(4, 2, 4 + INPUT_ROWS, 13)
pc.conditional_format(5, 13, 5 + INPUT_ROWS, 13, {"type": "text", "criteria": "begins with", "value": "REJECTED",
                                                  "format": F["cf_fail"]})

# ----------------------------------------------------------------------------
# RECONCILIATION
# ----------------------------------------------------------------------------
rc = wb.add_worksheet("Reconciliation")
rc.set_tab_color(TAB_OUT)
banner(rc, "Reconciliation - bank vs benchmark", 14)
RW = [24, 9, 7, 10, 46, 8, 17, 17, 16, 10, 10, 15]
for c, w in enumerate(RW):
    rc.set_column(2 + c, 2 + c, w)
rc.write(2, 2, "No reconciliation yet: enter or import bank figures on BankCells, then run.", F["runinfo"])
name("recStatus", rc, 2, 2)
buttons(rc, 2, 9, [("Run benchmark", "RunNSFR")])
rc.set_row(2, 26)
# Attribution: levels and effects in the wide columns I:K
rc.write(4, 2, "Level", F["hdr_out"])
rc.merge_range(4, 3, 4, 7, "Meaning", F["hdr_out"])
for c, h in enumerate(["ASF", "RSF", "NSFR"]):
    rc.write(4, 8 + c, h, F["hdr_out"])
LEVELS = [("Reported", "The bank's C 84.00 (else its C 80.00 / C 81.00 totals)"),
          ("Re-performed", "Bank amounts x bank factors, summed bottom-up"),
          ("Restated", "Bank amounts x CRR factors"),
          ("Benchmark", "Positions run of this workbook"),
          ("Arithmetic", "Re-performed - reported"),
          ("Factor effect", "Restated - re-performed"),
          ("Amounts / classification", "Benchmark - restated"),
          ("Total difference", "Benchmark - reported")]
for i, (lvl, meaning) in enumerate(LEVELS):
    rr = 5 + i
    rc.write(rr, 2, lvl, F["lbl"])
    rc.merge_range(rr, 3, rr, 7, meaning, F["desc"])
    rc.write_blank(rr, 8, None, F["out_amt_b"])
    rc.write_blank(rr, 9, None, F["out_amt_b"])
    rc.write_blank(rr, 10, None, F["out_pct2"])
name("recHeadline", rc, 5, 8, 12, 10)
# Counts by check and result, from the lines below
CK = 14
RH = 21
rc.write(CK, 2, "Check", F["hdr_out"])
rc.merge_range(CK, 3, CK, 7, "What it compares", F["hdr_out"])
for c, h in enumerate(["MATCH", "DIFF", "MISSING", "INFO"]):
    rc.write(CK, 8 + c, h, F["hdr_out"])
CHECKS = [("ARITHMETIC", "Bank stable funding vs its amounts x factors; parents vs children; C 84.00"),
          ("FACTOR", "Bank applicable factor vs CRR factor, where an amount is reported"),
          ("AMOUNT", "Bank amount vs benchmark amount (positions run VALID)"),
          ("STABLE FUNDING", "Bank stable funding vs benchmark, per row")]
for i, (chk, what) in enumerate(CHECKS):
    rr = CK + 1 + i
    rc.write(rr, 2, chk, F["lbl"])
    rc.merge_range(rr, 3, rr, 7, what, F["desc"])
    for c, res_ in enumerate(["MATCH", "DIFF", "MISSING", "INFO"]):
        rc.write_formula(rr, 8 + c, f'=COUNTIFS($C${RH + 2}:$C${RH + 1 + 20000},$C{rr + 1},'
                         f'$M${RH + 2}:$M${RH + 1 + 20000},"{res_}")', F["out_amt"], 0)
RHEAD = ["Check", "Template", "Row", "ID", "Item", "Column", "Bank", "Expected / benchmark", "Difference",
         "Relative", "Result", "SF impact"]
RF = ["out_txt", "out_txt", "out_txt", "out_txt", "out_txt", "out_txt", "out_amt", "out_amt", "out_amt",
      "out_pct", "out_txt", "out_amt"]
for c, (h, f) in enumerate(zip(RHEAD, RF)):
    rc.write(RH, 2 + c, h, F["hdr_out"])
rc.set_row(RH, 28)
for c, f in enumerate(RF):
    rc.set_column(2 + c, 2 + c, RW[c])
name("hdrRecon", rc, RH, 2)
rc.write(RH - 1, 2, "Factors are shown as decimals in Bank / Expected for FACTOR lines (0.05 = 5%); "
         "SF impact = amount x (CRR factor - bank factor).", F["note"])
for rr in range(RH + 1, RH + 1 + 2000):
    for c, f in enumerate(RF):
        rc.write_blank(rr, 2 + c, None, F[f])
rc.autofilter(RH, 2, RH + 20000, 13)
rc.conditional_format(RH + 1, 12, RH + 20000, 12, {"type": "cell", "criteria": "==", "value": '"DIFF"',
                                                   "format": F["cf_fail"]})
rc.conditional_format(RH + 1, 12, RH + 20000, 12, {"type": "cell", "criteria": "==", "value": '"MISSING"',
                                                   "format": F["cf_warn"]})
rc.conditional_format(RH + 1, 12, RH + 20000, 12, {"type": "cell", "criteria": "==", "value": '"MATCH"',
                                                   "format": F["cf_pass"]})
rc.freeze_panes(RH + 1, 4)

# ----------------------------------------------------------------------------
# CHECKS
# ----------------------------------------------------------------------------
ck = wb.add_worksheet("Checks")
ck.set_tab_color(TAB_OUT)
banner(ck, "Checks", 7)
for c, (h, w) in enumerate(zip(["Severity", "Sheet", "Row", "Reference", "Message"], [10, 12, 6, 22, 110])):
    ck.set_column(2 + c, 2 + c, w, F["out_txt"] if c != 2 else F["out_amt"])
    ck.write(4, 2 + c, h, F["hdr_out"])
name("hdrChecks", ck, 4, 2)
ck.write(2, 2, "ERROR rejects a position (totals withheld) or stops the run; WARNING needs a look; INFO "
         "documents a WHAT-IF factor or a competent-authority factor. Row = row of the input table.", F["note"])
ck.conditional_format(5, 2, 5 + INPUT_ROWS, 2, {"type": "cell", "criteria": "==", "value": '"ERROR"',
                                                "format": F["cf_fail"]})
ck.conditional_format(5, 2, 5 + INPUT_ROWS, 2, {"type": "cell", "criteria": "==", "value": '"WARNING"',
                                                "format": F["cf_warn"]})
ck.freeze_panes(5, 0)
ck.autofilter(4, 2, 4 + INPUT_ROWS, 6)

# ----------------------------------------------------------------------------
# TEST CATALOGUE
# ----------------------------------------------------------------------------
tc = wb.add_worksheet("TestCatalogue")
tc.set_tab_color(TAB_TEST)
banner(tc, "Test catalogue", 8)
for c, w in {2: 9, 3: 52, 4: 20, 5: 20, 6: 10, 7: 34}.items():
    tc.set_column(c, c, w)
tc.write(2, 2, "Demo results against the independent Python reference (tests/derivations/prototype/"
         "nsfr_reference.py). Valid only for the demo inputs as delivered.", F["note"])
tc.write(3, 2, "Summary", F["lbl"])
for c, h in enumerate(["Test", "Description", "Expected", "Actual", "Result", "Source"]):
    tc.write(5, 2 + c, h, F["hdr_in"])
tests: list[tuple[str, Any, str, str]] = []   # description, expected, actual formula, kind
E = EXPECTED
tests.append(("Status", E["status"], "=Results!$D$6", "text"))
tests.append(("Run class", E["run_class"], "=Results!$D$7", "text"))
tests.append(("ASF", E["ASF"], "=Results!$D$10", "num"))
tests.append(("RSF", E["RSF"], "=Results!$D$11", "num"))
tests.append(("NSFR", E["NSFR"], "=Results!$D$12", "num"))
for i, s in enumerate(SUMM):
    if s["measure"] == "NSFR":
        tests.append((f"C 84.00 r{s['row']} ratio", E["NSFR"], f"=Results!$I${SH + 2 + i}", "num"))
        continue
    amt, val = E["summary"][s["row"]]
    col = "G" if s["measure"] == "RSF" else "H"
    tests.append((f"C 84.00 r{s['row']} amount", amt, f"=Results!$F${SH + 2 + i}", "num"))
    tests.append((f"C 84.00 r{s['row']} {s['measure']}", val, f"=Results!${col}${SH + 2 + i}", "num"))
for template, sheet, scol in (("C 80.00", "C80", "R"), ("C 81.00", "C81", "O")):
    rows = [r_ for r_ in ROWS if r_["template"] == template]
    for i, row in enumerate(rows):
        exp = E["rows"][f"{template}|{row['row']}"]
        tests.append((f"{template} r{row['row']} ({row['id']}) stable funding", exp, f"={sheet}!${scol}${8 + i}",
                      "num"))
for pid, (bucket, rowcode, factor, sfv) in E["positions"].items():
    if factor is None:
        continue
    tests.append((f"{pid} factor", factor,
                  f'=INDEX(PositionCalc!$K$6:$K${5 + INPUT_ROWS},MATCH("{pid}",PositionCalc!$C$6:$C${5 + INPUT_ROWS},0))',
                  "num"))
    tests.append((f"{pid} bucket", bucket,
                  f'=INDEX(PositionCalc!$H$6:$H${5 + INPUT_ROWS},MATCH("{pid}",PositionCalc!$C$6:$C${5 + INPUT_ROWS},0))',
                  "text"))
for li, lvl in enumerate(["reported", "reperformed", "restated", "benchmark"]):
    for ci, m in enumerate(["ASF", "RSF", "NSFR"]):
        tests.append((f"Reconciliation {lvl} {m}", E["headline"][lvl][ci],
                      f"=Reconciliation!${'IJK'[ci]}${6 + li}", "num"))
for i, chk in enumerate(["ARITHMETIC", "FACTOR", "AMOUNT", "STABLE FUNDING"]):
    for c, res_ in enumerate(["MATCH", "DIFF", "MISSING", "INFO"]):
        tests.append((f"Reconciliation count {chk} {res_}", E["counts"].get(f"{chk}|{res_}", 0),
                      f"=Reconciliation!${'IJKL'[c]}${CK + 2 + i}", "num"))
for k, (desc, exp, formula, kind) in enumerate(tests):
    rr = 6 + k
    tc.write(rr, 2, f"T-{k + 1:03d}", F["code"])
    tc.write(rr, 3, desc, F["desc"])
    if kind == "num":
        tc.write_number(rr, 4, exp, F["out_amt"] if abs(exp) >= 2 else fmt(font_size=9, num_format="0.000000",
                                                                            border=1, border_color=GREY))
        tc.write_formula(rr, 5, formula, F["out_amt"] if abs(exp) >= 2 else fmt(font_size=9, num_format="0.000000",
                                                                                 border=1, border_color=GREY))
        tc.write_formula(rr, 6, f'=IF(ISNUMBER(F{rr + 1}),IF(ABS(F{rr + 1}-E{rr + 1})<=1E-6*MAX(1,ABS(E{rr + 1})),'
                         f'"PASS","FAIL"),"FAIL")', F["out_txt"])
    else:
        tc.write_string(rr, 4, exp, F["out_txt"])
        tc.write_formula(rr, 5, formula, F["out_txt"])
        tc.write_formula(rr, 6, f'=IF(F{rr + 1}=E{rr + 1},"PASS","FAIL")', F["out_txt"])
    tc.write(rr, 7, "nsfr_reference.py", F["ref"])
last = 6 + len(tests)
tc.write_formula(3, 3, f'=COUNTIF(G7:G{last},"PASS")&" of {len(tests)} PASS"', F["desc_b"])
tc.conditional_format(6, 6, last, 6, {"type": "cell", "criteria": "==", "value": '"PASS"', "format": F["cf_pass"]})
tc.conditional_format(6, 6, last, 6, {"type": "cell", "criteria": "==", "value": '"FAIL"', "format": F["cf_fail"]})
tc.freeze_panes(6, 0)

wb.close()
print(f"written {OUT.relative_to(ROOT)}: {len(ROWS)} factor rows, {len(CATS)} categories, "
      f"{len(DEMO['positions'])} positions, {len(BANK['cells'])} bank cells, {len(tests)} tests")
