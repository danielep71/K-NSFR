#!/usr/bin/env python3
"""Build the macro-free K-NSFR workbook template, byte for byte reproducible.

This Source Code Form is subject to the terms of the Mozilla Public
License, v. 2.0. If a copy of the MPL was not distributed with this
file, You can obtain one at https://mozilla.org/MPL/2.0/.

Writes templates/NSFR_Template.xlsx with the layout of every sheet of the
non-admitted prototype (prototype/tools/build_template.py): the same sheets,
headers, formats, input regions and validations, located through the names of
their first header cells (hdr*). The workbook names differ where
docs/SOURCE_CONTROL.md (gaps against the contract and the prototype) says,
and the layout is not yet contract v1. It leaves out what is
not admitted: the regulatory factors and template rows, the demo data and
bank figures, the expected results, and the macro buttons. The parameter set
(PROD-01) and the admitted modules (#137) bring them.

The file is reproducible: XlsxWriter writes every ZIP member with a fixed
date, and the document properties carry the fixed date CREATED. With --check
the template is built in a temporary folder and compared, byte for byte, with
templates/NSFR_Template.xlsx in this file's worktree; the `template` gate of
tools/check.py runs it, and with --ci that worktree must be clean, so the file
compared is the committed one.

Usage: python tools/build_template.py [--output PATH | --check]
"""

from __future__ import annotations

import argparse
import datetime as dt
import sys
import tempfile
import zipfile
from collections.abc import Sequence
from pathlib import Path
from typing import Any

import xlsxwriter  # type: ignore[import-untyped]

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "templates" / "NSFR_Template.xlsx"
# The document date, fixed so that every build gives the same bytes.
CREATED = dt.datetime(2026, 11, 10)
SHEETS = ("README", "Params", "Positions", "BankCells", "Results", "C80", "C81", "PositionCalc",
          "Reconciliation", "Checks", "TestCatalogue")
NAMES = ("ReportingDate", "ReportingCurrency", "Perimeter", "ProfileID", "hdrFactors", "hdrCategories",
         "CategoryList", "hdrSummary", "hdrPositions", "AmountScale", "AbsTolerance", "RelTolerance",
         "FactorTolerance", "hdrBankCells", "resRunInfo", "resValues", "hdrPositionCalc", "recStatus",
         "recHeadline", "hdrRecon", "hdrChecks", "hdrTests")

NAVY, BLUE, BAND, GOLD, INPUT, GREY = "#1F3864", "#2E75B6", "#265A99", "#FFC000", "#FFF2CC", "#BFBFBF"
TAB_DOC, TAB_IN, TAB_OUT, TAB_TEST = "#1F3864", "#BF9000", "#2E75B6", "#548235"
FONT = "Aptos"
INPUT_ROWS = 5000
CATEGORY_ROWS = 200   # the blank category rows that CategoryList covers
RECON_ROWS = 2000


class Builder:
    """One workbook and its formats; each sheet is added by its own method."""

    def __init__(self, path: Path) -> None:
        self.wb: Any = xlsxwriter.Workbook(str(path))
        self.wb.set_properties({"title": "K-NSFR template", "author": "", "company": "", "comments": "",
                                "created": CREATED})
        self.wb.formats[0].set_font_name(FONT)
        self.wb.formats[0].set_font_size(10)
        self._cache: dict[tuple[tuple[str, Any], ...], Any] = {}
        self.f = self._formats()

    def fmt(self, **kw: Any) -> Any:
        key = tuple(sorted(kw.items()))
        if key not in self._cache:
            base: dict[str, Any] = {"font_name": FONT, "valign": "vcenter"}
            base.update(kw)
            self._cache[key] = self.wb.add_format(base)
        return self._cache[key]

    def _formats(self) -> dict[str, Any]:
        f = self.fmt
        cell = {"border": 1, "border_color": GREY}
        hdr = {"font_size": 9, "bold": True, "text_wrap": True, "align": "center", **cell}
        inp = {"font_size": 9, "bg_color": INPUT, **cell}
        return {
            "brand": f(font_size=14, font_color=GOLD, bg_color="#000000", align="left", indent=1),
            "band1": f(bg_color="#000000"),
            "title": f(font_size=18, bold=True, font_color="#FFFFFF", bg_color=BAND, align="left", indent=1),
            "band2": f(bg_color=BAND),
            "note": f(font_size=10, italic=True, font_color="#595959"),
            "runinfo": f(font_size=10, italic=True, font_color="#7F6000"),
            "section": f(font_size=14, bold=True, font_color=NAVY, bottom=2, bottom_color=NAVY),
            "hdr_in": f(font_color="#FFFFFF", bg_color=NAVY, **hdr),
            "hdr_out": f(font_color="#FFFFFF", bg_color=BLUE, **hdr),
            "hdr_grp": f(font_color=NAVY, bg_color="#D9E1F2", **hdr),
            "code": f(font_size=9, bold=True, num_format="@"),
            "desc": f(font_size=9),
            "in_txt": f(num_format="@", **inp),
            "in_ctr": f(num_format="@", align="center", **inp),
            "in_amt": f(num_format="#,##0.00", **inp),
            "in_gen": f(num_format="General", align="center", **inp),
            "in_pct": f(num_format="0.0%", align="center", **inp),
            "in_date": f(num_format="yyyy-mm-dd", align="center", **inp),
            "out_amt": f(font_size=9, num_format="#,##0;-#,##0;-", **cell),
            "out_amt_b": f(font_size=9, num_format="#,##0;-#,##0;-", bold=True, **cell),
            "out_pct": f(font_size=9, num_format="0.0%", align="center", **cell),
            "out_pct2": f(font_size=10, num_format="0.00%", bold=True, align="right", **cell),
            "out_txt": f(font_size=9, num_format="@", **cell),
            "out_big": f(font_size=11, bold=True, num_format="#,##0", align="right", **cell),
            "out_bigtxt": f(font_size=11, bold=True, align="right", **cell),
            "out_date": f(font_size=10, num_format="yyyy-mm-dd", align="right", **cell),
            "lbl": f(font_size=10, bold=True, bg_color="#F2F2F2", **cell),
            "rm_h": f(font_size=14, bold=True, font_color=NAVY, bottom=2, bottom_color=NAVY, text_wrap=True),
            "rm_k": f(font_size=10, bold=True, text_wrap=True, bottom=3, bottom_color=GREY),
            "rm_v": f(font_size=9, text_wrap=True, bottom=3, bottom_color=GREY),
            "cf_bad": self.wb.add_format({"bg_color": "#F8CBAD"}),
            "cf_warn": self.wb.add_format({"bg_color": "#FFE699"}),
            "cf_pass": self.wb.add_format({"bg_color": "#C6EFCE", "font_color": "#006100"}),
            "cf_fail": self.wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"}),
        }

    # -- helpers ----------------------------------------------------------------------------------

    def sheet(self, name: str, tab: str, title: str, last_col: int) -> Any:
        """A sheet with the brand and title bands over B:last_col."""
        ws = self.wb.add_worksheet(name)
        ws.set_tab_color(tab)
        ws.hide_gridlines(2)
        ws.set_column(0, 0, 2)
        ws.set_column(1, 1, 2.5)
        ws.set_row(0, 24)
        ws.set_row(1, 26.1)
        for c in range(1, last_col + 1):
            ws.write_blank(0, c, None, self.f["band1"])
            ws.write_blank(1, c, None, self.f["band2"])
        ws.write(0, 1, "K-NSFR", self.f["brand"])
        ws.write(1, 1, title, self.f["title"])
        return ws

    def name(self, nm: str, ws: Any, row: int, col: int, row2: int | None = None,
             col2: int | None = None) -> None:
        ref = cell_ref(row, col)
        if row2 is not None and col2 is not None:
            ref += ":" + cell_ref(row2, col2)
        self.wb.define_name(nm, f"='{ws.name}'!{ref}")

    def headers(self, ws: Any, row: int, col: int, labels: Sequence[str], style: str) -> None:
        for c, h in enumerate(labels):
            ws.write(row, col + c, h, self.f[style])

    def blanks(self, ws: Any, first: int, count: int, col: int, styles: Sequence[str]) -> None:
        for rr in range(first, first + count):
            for c, s in enumerate(styles):
                ws.write_blank(rr, col + c, None, self.f[s])

    # -- sheets -----------------------------------------------------------------------------------

    def readme(self) -> None:
        ws = self.sheet("README", TAB_DOC, "ReadMe", 4)
        ws.set_column(2, 2, 34)
        ws.set_column(3, 3, 150)
        ws.write(2, 2, "K-NSFR workbook template: the layout only. It holds no calculation, factor, "
                 "position or result.", self.f["note"])
        rows = [
            ("Template", None),
            ("What it is", "The macro-free layout of the K-NSFR workbook, generated by tools/build_template.py. "
             "Every table is located through the workbook name of its first header cell (hdr*), every "
             "setting through the workbook name equal to its code."),
            ("What it leaves out", "Regulatory factors and template rows, demo data, bank figures, expected "
             "results and macro buttons. The verified parameter set and the admitted VBA modules bring them."),
            ("How to use", None),
            ("Build a workbook", "Save a copy as an Excel Macro-Enabled Workbook (.xlsm) outside the "
             "repository, and import the VBA source components as docs/INSTALLATION.md describes. Never commit "
             "the .xlsm."),
            ("Status", None),
            ("Not for use", "Repository setup: no NSFR calculation exists. Do not use this workbook for "
             "regulatory reporting."),
        ]
        r = 3
        for key, value in rows:
            if value is None:
                ws.set_row(r, 8)
                r += 1
                ws.set_row(r, 24)
                ws.write(r, 2, key, self.f["rm_h"])
                ws.write_blank(r, 3, None, self.f["rm_h"])
            else:
                ws.set_row(r, 15 * max(1, (len(value) // 165) + 1) + 4)
                ws.write(r, 2, key, self.f["rm_k"])
                ws.write(r, 3, value, self.f["rm_v"])
            r += 1
        ws.freeze_panes(3, 0)

    def params(self) -> None:
        ws = self.sheet("Params", TAB_IN, "Parameters", 19)
        for c, w in {2: 10, 3: 8, 4: 11, 5: 52, 6: 7, 7: 8, 8: 8}.items():
            ws.set_column(c, c, w)
        ws.set_column(9, 16, 8.5)
        ws.set_column(17, 17, 40)
        ws.set_column(18, 18, 24)
        ws.write(3, 2, "Run settings", self.f["section"])
        settings = [("ReportingDate", "Reporting (as-of) date", "in_date"),
                    ("ReportingCurrency", "Reporting currency", "in_ctr"),
                    ("Perimeter", "Reporting perimeter", "in_txt"),
                    ("ProfileID", "Calculation profile", "in_ctr")]
        for i, (code, desc, style) in enumerate(settings):
            rr = 4 + i
            ws.write(rr, 2, code, self.f["code"])
            ws.merge_range(rr, 3, rr, 4, "", self.f[style])
            ws.write(rr, 5, desc, self.f["desc"])
            self.name(code, ws, rr, 3)
        ws.write(8, 2, "Version", self.f["code"])
        ws.data_validation(7, 3, 7, 3, {"validate": "list", "source": ["eu-crr-standard"]})
        ft = 11
        ws.write(ft - 1, 2, "Factor table - C 80.00 and C 81.00 rows (from the parameter set)", self.f["section"])
        self.headers(ws, ft, 2, ["Template", "Row", "ID", "Item", "Kind", "Amount cells", "Factor cells",
                                 "Reg <6M", "Reg 6M-1Y", "Reg >=1Y", "Reg HQLA", "Applied <6M", "Applied 6M-1Y",
                                 "Applied >=1Y", "Applied HQLA", "Reference", "Children"], "hdr_in")
        ws.set_row(ft, 30)
        self.name("hdrFactors", ws, ft, 2)
        ct = ft + 4
        ws.write(ct - 1, 2, "Category table - classification of positions", self.f["section"])
        self.headers(ws, ct, 2, ["Category", "Description", "Template", "Bucket rule", "Undated bucket",
                                 "Row unencumbered or <6M", "Row encumbered 6M-1Y", "Row encumbered >=1Y",
                                 "Memo row", "Reference"], "hdr_in")
        ws.set_row(ct, 30)
        self.name("hdrCategories", ws, ct, 2)
        self.wb.define_name("CategoryList", f"='Params'!$C${ct + 2}:$C${ct + 1 + CATEGORY_ROWS}")
        st = ct + CATEGORY_ROWS + 4
        ws.write(st - 1, 2, "Summary map - C 84.00", self.f["section"])
        self.headers(ws, st, 2, ["Row", "ID", "Item", "Measure", "Plus (Template|ID; ...)", "Minus"], "hdr_in")
        self.name("hdrSummary", ws, st, 2)
        ws.freeze_panes(ft + 1, 5)

    def positions(self) -> None:
        ws = self.sheet("Positions", TAB_IN, "Positions", 11)
        heads = ["PositionID", "Description", "Category", "Amount", "Currency", "MaturityDate",
                 "EncumberedUntil", "BucketOverride", "FactorOverride", "Comment"]
        for c, (h, w) in enumerate(zip(heads, [11, 44, 22, 17, 8, 12, 13, 10, 10, 34], strict=True)):
            ws.set_column(2 + c, 2 + c, w)
            ws.write(4, 2 + c, h, self.f["hdr_in"])
        ws.set_row(4, 28)
        self.name("hdrPositions", ws, 4, 2)
        ws.write(2, 2, "One row per position, amounts in units of the reporting currency (>= 0). Category: see "
                 "Params. Blank maturity = the category's undated bucket.", self.f["note"])
        self.blanks(ws, 5, INPUT_ROWS, 2, ["in_txt", "in_txt", "in_txt", "in_amt", "in_ctr", "in_date",
                                           "in_date", "in_ctr", "in_pct", "in_txt"])
        ws.data_validation(5, 4, 4 + INPUT_ROWS, 4, {"validate": "list", "source": "=CategoryList",
                                                     "error_title": "Category",
                                                     "error_message": "Choose a category of Params."})
        ws.data_validation(5, 9, 4 + INPUT_ROWS, 9, {"validate": "list", "source": ["<6M", "6M-1Y", ">=1Y"]})
        ws.freeze_panes(5, 3)
        ws.autofilter(4, 2, 4 + INPUT_ROWS, 11)

    def bank_cells(self) -> None:
        ws = self.sheet("BankCells", TAB_IN, "Bank figures", 8)
        for c, w in {2: 16, 3: 14, 4: 10, 5: 22, 6: 40, 7: 4, 8: 4}.items():
            ws.set_column(c, c, w)
        ws.write(2, 2, "The bank's own C 80.00 / C 81.00 / C 84.00 figures, one cell per row.", self.f["note"])
        settings = [("AmountScale", "Units of the reporting currency per bank amount: 1, 1000 (thousands)..."),
                    ("AbsTolerance", "In units of the reporting currency."),
                    ("RelTolerance", "Of the absolute bank figure (not applied when it is 0)."),
                    ("FactorTolerance", "Absolute, for factors and the ratio.")]
        for i, (code, desc) in enumerate(settings):
            rr = 4 + i
            ws.write(rr, 2, code, self.f["code"])
            ws.write_blank(rr, 3, None, self.f["in_gen"])
            ws.write(rr, 4, desc, self.f["note"])
            self.name(code, ws, rr, 3)
        bh = 9
        self.headers(ws, bh, 2, ["Template", "Row", "Column", "Value", "Source"], "hdr_in")
        self.name("hdrBankCells", ws, bh, 2)
        self.blanks(ws, bh + 1, INPUT_ROWS, 2, ["in_ctr", "in_ctr", "in_ctr", "in_amt", "in_txt"])
        ws.freeze_panes(bh + 1, 0)
        ws.autofilter(bh, 2, bh + INPUT_ROWS, 6)

    def results(self) -> None:
        ws = self.sheet("Results", TAB_OUT, "Results", 9)
        ws.set_column(2, 3, 22)
        ws.set_column(4, 4, 62)
        ws.set_column(5, 8, 17)
        ws.write(3, 2, "No results.", self.f["runinfo"])
        self.name("resRunInfo", ws, 3, 2)
        labels = ["Status", "Run class", "Reporting date", "Perimeter", "Available stable funding (ASF)",
                  "Required stable funding (RSF)", "NSFR", "Positions accepted", "Positions rejected"]
        styles = ["out_bigtxt", "out_bigtxt", "out_date", "out_bigtxt", "out_big", "out_big", "out_pct2",
                  "out_big", "out_big"]
        for i, (lbl, s) in enumerate(zip(labels, styles, strict=True)):
            ws.write(5 + i, 2, lbl, self.f["lbl"])
            ws.write_blank(5 + i, 3, None, self.f[s])
        self.name("resValues", ws, 5, 3, 13, 3)
        for value, style in (("VALID", "cf_pass"), ("INVALID", "cf_fail"), ("STOPPED", "cf_fail")):
            ws.conditional_format(5, 3, 5, 3, {"type": "text", "criteria": "begins with", "value": value,
                                               "format": self.f[style]})
        ws.conditional_format(6, 3, 6, 3, {"type": "text", "criteria": "begins with", "value": "WHAT-IF",
                                           "format": self.f["cf_warn"]})
        sh = 16
        ws.write(sh - 1, 2, "C 84.00 - NSFR summary", self.f["section"])
        self.headers(ws, sh, 2, ["Row", "ID", "Item", "Amount (0010)", "RSF (0020)", "ASF (0030)",
                                 "Ratio (0040)"], "hdr_out")
        ws.freeze_panes(3, 0)

    def replica(self, sheet_name: str, template: str, buckets: Sequence[str], title: str) -> None:
        nb = len(buckets)
        ws = self.sheet(sheet_name, TAB_OUT, title, 2 + 3 + 3 * nb + 1)
        ws.set_column(2, 3, 8)
        ws.set_column(4, 4, 58)
        ws.set_column(5, 4 + nb, 15)
        ws.set_column(5 + nb, 4 + 3 * nb, 8)
        ws.set_column(5 + 3 * nb, 5 + 3 * nb, 16)
        ws.write(2, 2, f"Replica of {template}: amounts and stable funding of the last run.", self.f["note"])
        for label, c1, c2 in (("Amount", 5, 4 + nb), ("Regulatory factor", 5 + nb, 4 + 2 * nb),
                              ("Applied factor", 5 + 2 * nb, 4 + 3 * nb)):
            ws.merge_range(4, c1, 4, c2, label, self.f["hdr_grp"])
        ws.merge_range(4, 5 + 3 * nb, 5, 5 + 3 * nb, "Stable funding", self.f["hdr_grp"])
        for c, h in enumerate(["Row", "ID", "Item"]):
            ws.merge_range(4, 2 + c, 6, 2 + c, h, self.f["hdr_out"])
        first_factor = 9 if template == "C 80.00" else 7
        for b in range(nb):
            for g, code in enumerate((10 * (b + 1), 10 * (nb + 1 + b), 10 * (first_factor + b))):
                ws.write(5, 5 + g * nb + b, buckets[b], self.f["hdr_out"])
                ws.write(6, 5 + g * nb + b, f"{code:04d}", self.f["hdr_out"])
        ws.write(6, 5 + 3 * nb, "0130" if template == "C 80.00" else "0100", self.f["hdr_out"])
        ws.freeze_panes(7, 5)

    def position_calc(self) -> None:
        ws = self.sheet("PositionCalc", TAB_OUT, "Position calculation", 13)
        heads = ["PositionID", "Category", "Template", "Row", "Row ID", "Bucket", "Encumbrance", "Amount",
                 "Factor", "Stable funding", "Memo row", "Status"]
        widths = [11, 22, 9, 7, 10, 8, 18, 17, 8, 17, 10, 70]
        styles = ["out_txt"] * 7 + ["out_amt", "out_pct", "out_amt", "out_txt", "out_txt"]
        for c, (h, w, s) in enumerate(zip(heads, widths, styles, strict=True)):
            ws.set_column(2 + c, 2 + c, w, self.f[s])
            ws.write(4, 2 + c, h, self.f["hdr_out"])
        self.name("hdrPositionCalc", ws, 4, 2)
        ws.write(2, 2, "Classification and weighting of every position of the last run.", self.f["note"])
        ws.freeze_panes(5, 3)
        ws.autofilter(4, 2, 4 + INPUT_ROWS, 13)
        ws.conditional_format(5, 13, 5 + INPUT_ROWS, 13, {"type": "text", "criteria": "begins with",
                                                          "value": "REJECTED", "format": self.f["cf_fail"]})

    def reconciliation(self) -> None:
        ws = self.sheet("Reconciliation", TAB_OUT, "Reconciliation - bank vs benchmark", 14)
        widths = [24, 9, 7, 10, 46, 8, 17, 17, 16, 10, 10, 15]
        for c, w in enumerate(widths):
            ws.set_column(2 + c, 2 + c, w)
        ws.write(2, 2, "No reconciliation.", self.f["runinfo"])
        self.name("recStatus", ws, 2, 2)
        ws.set_row(2, 26)
        ws.write(4, 2, "Level", self.f["hdr_out"])
        ws.merge_range(4, 3, 4, 7, "Meaning", self.f["hdr_out"])
        self.headers(ws, 4, 8, ["ASF", "RSF", "NSFR"], "hdr_out")
        levels = [("Reported", "The bank's C 84.00 (else its C 80.00 / C 81.00 totals)"),
                  ("Re-performed", "Bank amounts x bank factors, summed bottom-up"),
                  ("Restated", "Bank amounts x the parameter set's factors"),
                  ("Benchmark", "Positions run of this workbook"),
                  ("Arithmetic", "Re-performed - reported"),
                  ("Factor effect", "Restated - re-performed"),
                  ("Amounts / classification", "Benchmark - restated"),
                  ("Total difference", "Benchmark - reported")]
        for i, (level, meaning) in enumerate(levels):
            rr = 5 + i
            ws.write(rr, 2, level, self.f["lbl"])
            ws.merge_range(rr, 3, rr, 7, meaning, self.f["desc"])
            for c, s in enumerate(("out_amt_b", "out_amt_b", "out_pct2")):
                ws.write_blank(rr, 8 + c, None, self.f[s])
        self.name("recHeadline", ws, 5, 8, 12, 10)
        self._recon_lines(ws)

    def _recon_lines(self, ws: Any) -> None:
        ck, rh = 14, 21
        ws.write(ck, 2, "Check", self.f["hdr_out"])
        ws.merge_range(ck, 3, ck, 7, "What it compares", self.f["hdr_out"])
        results = ["MATCH", "DIFF", "MISSING", "INFO"]
        self.headers(ws, ck, 8, results, "hdr_out")
        checks = [("ARITHMETIC", "Bank stable funding vs its amounts x factors; parents vs children; C 84.00"),
                  ("FACTOR", "Bank applicable factor vs the parameter set's factor, where an amount is reported"),
                  ("AMOUNT", "Bank amount vs benchmark amount (positions run VALID)"),
                  ("STABLE FUNDING", "Bank stable funding vs benchmark, per row")]
        for i, (check, what) in enumerate(checks):
            rr = ck + 1 + i
            ws.write(rr, 2, check, self.f["lbl"])
            ws.merge_range(rr, 3, rr, 7, what, self.f["desc"])
            for c, result in enumerate(results):
                ws.write_formula(rr, 8 + c, f'=COUNTIFS($C${rh + 2}:$C${rh + 1 + 20000},$C{rr + 1},'
                                 f'$M${rh + 2}:$M${rh + 1 + 20000},"{result}")', self.f["out_amt"], 0)
        styles = ["out_txt"] * 6 + ["out_amt", "out_amt", "out_amt", "out_pct", "out_txt", "out_amt"]
        ws.write(rh - 1, 2, "Factors are shown as decimals in Bank / Expected for FACTOR lines (0.05 = 5%); "
                 "SF impact = amount x (parameter-set factor - bank factor).", self.f["note"])
        self.headers(ws, rh, 2, ["Check", "Template", "Row", "ID", "Item", "Column", "Bank",
                                 "Expected / benchmark", "Difference", "Relative", "Result", "SF impact"], "hdr_out")
        ws.set_row(rh, 28)
        self.name("hdrRecon", ws, rh, 2)
        self.blanks(ws, rh + 1, RECON_ROWS, 2, styles)
        ws.autofilter(rh, 2, rh + 20000, 13)
        for value, style in (('"DIFF"', "cf_fail"), ('"MISSING"', "cf_warn"), ('"MATCH"', "cf_pass")):
            ws.conditional_format(rh + 1, 12, rh + 20000, 12, {"type": "cell", "criteria": "==", "value": value,
                                                               "format": self.f[style]})
        ws.freeze_panes(rh + 1, 4)

    def checks(self) -> None:
        ws = self.sheet("Checks", TAB_OUT, "Checks", 7)
        for c, (h, w) in enumerate(zip(["Severity", "Sheet", "Row", "Reference", "Message"],
                                       [10, 12, 6, 22, 110], strict=True)):
            ws.set_column(2 + c, 2 + c, w, self.f["out_txt"] if c != 2 else self.f["out_amt"])
            ws.write(4, 2 + c, h, self.f["hdr_out"])
        self.name("hdrChecks", ws, 4, 2)
        ws.write(2, 2, "ERROR rejects a position (totals withheld) or stops the run; WARNING needs a look; INFO "
                 "documents an applied factor that differs. Row = row of the input table.", self.f["note"])
        for value, style in (('"ERROR"', "cf_fail"), ('"WARNING"', "cf_warn")):
            ws.conditional_format(5, 2, 5 + INPUT_ROWS, 2, {"type": "cell", "criteria": "==", "value": value,
                                                            "format": self.f[style]})
        ws.freeze_panes(5, 0)
        ws.autofilter(4, 2, 4 + INPUT_ROWS, 6)

    def test_catalogue(self) -> None:
        ws = self.sheet("TestCatalogue", TAB_TEST, "Test catalogue", 8)
        for c, w in {2: 9, 3: 52, 4: 20, 5: 20, 6: 10, 7: 34}.items():
            ws.set_column(c, c, w)
        ws.write(2, 2, "Registered fixtures and their results, once fixtures are admitted.", self.f["note"])
        ws.write(3, 2, "Summary", self.f["lbl"])
        self.headers(ws, 5, 2, ["Test", "Description", "Expected", "Actual", "Result", "Source"], "hdr_in")
        self.name("hdrTests", ws, 5, 2)
        ws.freeze_panes(6, 0)

    def build(self) -> None:
        self.readme()
        self.params()
        self.positions()
        self.bank_cells()
        self.results()
        self.replica("C80", "C 80.00", ["<6M", "6M-1Y", ">=1Y", "HQLA"], "C 80.00 - Required stable funding")
        self.replica("C81", "C 81.00", ["<6M", "6M-1Y", ">=1Y"], "C 81.00 - Available stable funding")
        self.position_calc()
        self.reconciliation()
        self.checks()
        self.test_catalogue()
        self.wb.close()


def cell_ref(row: int, col: int) -> str:
    """Absolute A1 reference of a zero-based cell."""
    letters = ""
    col += 1
    while col:
        col, rem = divmod(col - 1, 26)
        letters = chr(65 + rem) + letters
    return f"${letters}${row + 1}"


def build(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    Builder(path).build()


def macro_findings(path: Path) -> list[str]:
    """A macro-free package holds no VBA project and no macro-enabled content type."""
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        types = z.read("[Content_Types].xml").decode("utf-8")
    found = [f"{path.name}: holds {n}" for n in names if n.lower().endswith("vbaproject.bin")]
    if "macroEnabled" in types:
        found.append(f"{path.name}: has a macro-enabled content type")
    return found


def check(committed: Path) -> list[str]:
    """Rebuild the template and compare it with the file at the given path."""
    if not committed.is_file():
        return [f"{committed.relative_to(ROOT)} is missing: run python tools/build_template.py"]
    with tempfile.TemporaryDirectory() as tmp:
        fresh = Path(tmp) / committed.name
        build(fresh)
        findings = macro_findings(committed)
        if fresh.read_bytes() != committed.read_bytes():
            findings.append(f"{committed.relative_to(ROOT)} differs from the generator's output: "
                            "run python tools/build_template.py and commit the result")
    return findings


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--output", type=Path, default=TEMPLATE, help="where to write the template")
    group.add_argument("--check", action="store_true", help="compare a fresh build with the committed template")
    args = parser.parse_args(argv)
    if args.check:
        findings = check(TEMPLATE)
        for finding in findings:
            print(finding)
        print(f"template: {'PASS' if not findings else 'FAIL'} ({len(findings)} findings)")
        return 1 if findings else 0
    build(args.output)
    print(f"written {args.output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
