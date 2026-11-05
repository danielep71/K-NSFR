"""Independent Python reference of the K-NSFR prototype (engine and reconciliation).

Written from docs/methodology/PROTOTYPE.md, separately from the VBA, to derive
the expected values of the demo workbook. It is a test oracle for the
prototype, not a second NSFR engine: it reads the same regulatory tables
(regulatory-tables.json) and the synthetic demo (demo-positions.json).

Usage:
    python tests/derivations/prototype/nsfr_reference.py            # print the demo results
    python tests/derivations/prototype/nsfr_reference.py --write    # also write demo-bank-cells.json
                                                                    # and demo.expected.json
"""

from __future__ import annotations

import argparse
import calendar
import datetime as dt
import json
from pathlib import Path
from typing import Any

HERE = Path(__file__).resolve().parent
TABLES = json.loads((HERE / "regulatory-tables.json").read_text())
DEMO = json.loads((HERE / "demo-positions.json").read_text())

BUCKETS = ["<6M", "6M-1Y", ">=1Y", "HQLA"]
LETTERS = "123H"
# Column codes of the amount, applicable-factor and stable-funding cells.
COLS = {"C 80.00": (["0010", "0020", "0030", "0040"], ["0090", "0100", "0110", "0120"], "0130"),
        "C 81.00": (["0010", "0020", "0030"], ["0070", "0080", "0090"], "0100")}


def add_months(d: dt.date, n: int) -> dt.date:
    """Calendar months, day clamped to the month's end (VBA DateAdd, Excel EDATE)."""
    m = d.month - 1 + n
    y, m = d.year + m // 12, m % 12 + 1
    return dt.date(y, m, min(d.day, calendar.monthrange(y, m)[1]))


def maturity_bucket(rd: dt.date, mat: dt.date) -> int:
    """0: < 6 months, 1: 6 months to < 1 year, 2: 1 year or more."""
    if mat < add_months(rd, 6):
        return 0
    if mat < add_months(rd, 12):
        return 1
    return 2


def iso(text: str) -> dt.date | None:
    return dt.date.fromisoformat(text) if text else None


class Tables:
    def __init__(self, tables: dict[str, Any], applied_overrides: dict[str, list[Any]]):
        self.rows = tables["rows"]
        self.by_id = {(r["template"], r["id"]): r for r in self.rows}
        self.by_row = {(r["template"], r["row"]): r for r in self.rows}
        self.cats = {c["category"]: c for c in tables["categories"]}
        self.summary = tables["summary"]
        self.applied: dict[tuple[str, str], list[Any]] = {}
        for r in self.rows:
            key = (r["template"], r["id"])
            self.applied[key] = list(applied_overrides.get(f"{r['template']}|{r['id']}", r["regulatory"]))

    def children(self, r: dict[str, Any]) -> list[dict[str, Any]]:
        return [self.by_row[(r["template"], c)] for c in r["children"]]


# --------------------------------------------------------------------------
# Engine
# --------------------------------------------------------------------------
Key = tuple[str, str]
Grid = dict[Key, list[float]]


def whatif_cells(t: Tables) -> list[str]:
    """Cells whose applied factor differs from the regulatory one."""
    out = []
    for r in t.rows:
        reg, app = r["regulatory"], t.applied[(r["template"], r["id"])]
        for b in range(4):
            if reg[b] is not None and app[b] is not None and abs(reg[b] - app[b]) > 1e-12:
                out.append(f"{r['template']} {r['id']} {BUCKETS[b]}")
    return out


def bucket_of(c: dict[str, Any], rd: dt.date, mat: str, override: str) -> int | None:
    """Amount bucket 0-3 of a position, None when its maturity is required."""
    rule = c["bucket_rule"]
    if rule == "HQLA":
        return 3
    if rule == "FIXED1":
        return 0
    if rule == "FIXED3":
        return 2
    if override:
        return BUCKETS.index(override)
    if mat:
        return maturity_bucket(rd, iso(mat))  # type: ignore[arg-type]
    if c["undated"]:
        return int(c["undated"]) - 1
    return None


def classify(t: Tables, rd: dt.date, c: dict[str, Any], mat: str, enc: str, override: str
             ) -> tuple[Key, int] | str:
    """Row key and bucket of a non-derivative position, or the reason it is rejected."""
    b = bucket_of(c, rd, mat, override)
    if b is None:
        return "maturity required"
    e = 0
    if enc:
        if c["template"] == "C 81.00":
            return "encumbrance on a liability"
        ed = iso(enc)
        if ed is not None and ed > rd:
            e = maturity_bucket(rd, ed)
    key = (c["template"], c["row_ids"][e])
    if LETTERS[b] not in t.by_id[key]["amount_buckets"]:
        return "cell not reportable"
    return key, b


def book(t: Tables, amounts: Grid, sf: Grid, c: dict[str, Any], key: Key, b: int, amount: float,
         f: float) -> None:
    """Add a position to its row and, where the bucket exists, to its memo row."""
    amounts[key][b] += amount
    sf[key][b] += amount * f
    if c["memo_id"]:
        mkey = (c["template"], c["memo_id"])
        mrow = t.by_id[mkey]
        if LETTERS[b] in mrow["amount_buckets"]:
            amounts[mkey][b] += amount
            if mrow["factor_buckets"]:
                sf[mkey][b] += amount * f


def book_derivatives(t: Tables, drv: dict[str, float], amounts: Grid, sf: Grid) -> tuple[float, float]:
    """Net the derivatives [Art. 428d, 428s(2), 428ah(2), 428k(3)(a)]."""
    da = max(0.0, drv["DRV_FV_ASSET"] - drv["DRV_VM_RECEIVED"])
    dl = max(0.0, drv["DRV_FV_LIAB"] - drv["DRV_VM_POSTED"])
    for key, amt in ((("C 80.00", "1.7.1"), drv["DRV_FV_LIAB"]), (("C 80.00", "1.7.2"), max(0.0, da - dl)),
                     (("C 81.00", "2.7"), max(0.0, dl - da))):
        if amt:
            amounts[key][0] += amt
            sf[key][0] += amt * t.applied[key][0]
    return da, dl


def aggregate(t: Tables, amounts: Grid, sf: Grid) -> None:
    """Fill every AGG row with the sum of its children, deepest first."""
    for r in sorted(t.rows, key=lambda x: -x["id"].count(".")):
        if r["kind"] == "AGG":
            key = (r["template"], r["id"])
            for ch in t.children(r):
                ck = (ch["template"], ch["id"])
                for b in range(4):
                    amounts[key][b] += amounts[ck][b]
                    sf[key][b] += sf[ck][b]


def run_engine(t: Tables, params: dict[str, Any], positions: list[list[Any]]) -> dict[str, Any]:
    rd = iso(params["ReportingDate"])
    assert rd is not None
    errors: list[str] = []
    whatif = whatif_cells(t)
    amounts: Grid = {(r["template"], r["id"]): [0.0] * 4 for r in t.rows}
    sf: Grid = {(r["template"], r["id"]): [0.0] * 4 for r in t.rows}
    drv = {"DRV_FV_ASSET": 0.0, "DRV_FV_LIAB": 0.0, "DRV_VM_RECEIVED": 0.0, "DRV_VM_POSTED": 0.0}
    calc: list[dict[str, Any]] = []
    seen: set[str] = set()
    for pid, _desc, catcode, amount, _ccy, mat, enc, override, fovr, _c in positions:
        c = t.cats.get(catcode)
        if pid in seen or c is None or amount is None or amount < 0:
            errors.append(f"{pid}: invalid")
            continue
        seen.add(pid)
        if c["bucket_rule"] == "DERIVATIVE":
            drv[catcode] += amount
            calc.append({"id": pid, "bucket": "", "row": "", "factor": None, "sf": None})
            continue
        where = classify(t, rd, c, mat, enc, override)
        if isinstance(where, str):
            errors.append(f"{pid}: {where}")
            continue
        key, b = where
        f = fovr if fovr is not None else t.applied[key][b]
        if f is None:
            errors.append(f"{pid}: no factor")
            continue
        if fovr is not None:
            whatif.append(f"{pid} override")
        book(t, amounts, sf, c, key, b, amount, f)
        calc.append({"id": pid, "bucket": BUCKETS[b], "row": t.by_id[key]["row"], "factor": f,
                     "sf": amount * f})
    da, dl = book_derivatives(t, drv, amounts, sf)
    aggregate(t, amounts, sf)
    rows_out: dict[str, dict[str, Any]] = {}
    for r in t.rows:
        key = (r["template"], r["id"])
        rows_out[f"{r['template']}|{r['row']}"] = {"amounts": amounts[key], "sf": sum(sf[key]),
                                                  "sf_buckets": sf[key]}
    summary = summarise(t, {k: (sum(v["amounts"]), v["sf"]) for k, v in rows_out.items()})
    status = "VALID" if not errors else f"INVALID: {len(errors)} position error(s)"
    if not errors and summary["RSF"] <= 0:
        status = "INVALID: RSF is zero"
    return {"status": status, "run_class": "WHAT-IF" if whatif else "REFERENCE", "errors": errors,
            "whatif": whatif, "rows": rows_out, "summary": summary, "calc": calc,
            "derivatives": {"assets": da, "liabilities": dl}}


def summarise(t: Tables, row_totals: dict[str, tuple[float, float]]) -> dict[str, Any]:
    """C 84.00 lines from per-row (amount, stable funding) keyed 'template|row'."""
    id_key = {f"{r['template']}|{r['id']}": f"{r['template']}|{r['row']}" for r in t.rows}
    lines = {}
    for s in t.summary:
        if s["measure"] == "NSFR":
            continue
        amt = sum(row_totals[id_key[p]][0] for p in s["plus"]) - sum(row_totals[id_key[m]][0] for m in s["minus"])
        val = sum(row_totals[id_key[p]][1] for p in s["plus"]) - sum(row_totals[id_key[m]][1] for m in s["minus"])
        lines[s["row"]] = {"amount": amt, "value": val, "measure": s["measure"]}
    rsf, asf = lines["0010"]["value"], lines["0120"]["value"]
    return {"lines": lines, "RSF": rsf, "ASF": asf, "NSFR": asf / rsf if rsf > 0 else None}


# --------------------------------------------------------------------------
# Demo bank figures: the benchmark's results with documented perturbations
# --------------------------------------------------------------------------
def demo_bank_cells(t: Tables, eng: dict[str, Any]) -> list[list[Any]]:
    amounts = {k: list(v["amounts"]) for k, v in eng["rows"].items()}
    factors: dict[str, list[Any]] = {}
    for r in t.rows:
        k = f"{r['template']}|{r['row']}"
        factors[k] = [t.applied[(r["template"], r["id"])][b] if LETTERS[b] in r["factor_buckets"] else None
                      for b in range(4)]
    # A memo row takes the factor its amounts were weighted with on the main row.
    for r in t.rows:
        if r["kind"] == "MEMO" and r["factor_buckets"]:
            k = f"{r['template']}|{r['row']}"
            e = eng["rows"][k]
            factors[k] = [e["sf_buckets"][b] / e["amounts"][b] if e["amounts"][b] else None for b in range(4)]
    # 1 classification: 100m of SME loans (RW > 35%) reported in r0770 instead of r0820, >=1Y
    amounts["C 80.00|0820"][2] -= 100_000_000
    amounts["C 80.00|0770"][2] += 100_000_000
    # 2 factor: other assets >=1Y at 50%
    factors["C 80.00|1030"][2] = 0.5
    # 4 amount: stable retail deposits <6M
    amounts["C 81.00|0090"][0] += 20_000_000
    reported: dict[str, float] = {}
    for r in sorted(t.rows, key=lambda x: -x["id"].count(".")):
        k = f"{r['template']}|{r['row']}"
        if r["kind"] == "AGG":
            amounts[k] = [sum(amounts[f"{r['template']}|{c}"][b] for c in r["children"]) for b in range(4)]
            reported[k] = sum(reported[f"{r['template']}|{c}"] for c in r["children"])
        else:
            # 5 template: an amount cell without a factor cell takes 0% in the template formula
            reported[k] = sum(amounts[k][b] * (factors[k][b] or 0.0) for b in range(4))
            if k == "C 80.00|1020":  # 3 arithmetic: the bank's own formula omits 10m
                reported[k] -= 10_000_000
    cells: list[list[Any]] = []
    for r in t.rows:
        k = f"{r['template']}|{r['row']}"
        acols, fcols, scol = COLS[r["template"]]
        any_amount = False
        for b, col in enumerate(acols):
            if LETTERS[b] in r["amount_buckets"] and amounts[k][b]:
                cells.append([r["template"], r["row"], col, amounts[k][b]])
                any_amount = True
                if factors[k][b] is not None:
                    cells.append([r["template"], r["row"], fcols[b], factors[k][b]])
        if any_amount:
            cells.append([r["template"], r["row"], scol, reported[k]])
    summ = summarise(t, {k: (sum(amounts[k]), reported[k]) for k in reported})
    for row, line in summ["lines"].items():
        cells.append(["C 84.00", row, "0010", line["amount"]])
        cells.append(["C 84.00", row, "0020" if line["measure"] == "RSF" else "0030", line["value"]])
    cells.append(["C 84.00", "0220", "0040", summ["NSFR"]])
    return cells


# --------------------------------------------------------------------------
# Reconciliation
# --------------------------------------------------------------------------
class Recon:
    """Bank figures of one reconciliation, the checks and their lines."""

    def __init__(self, t: Tables, params: dict[str, Any], bank_cells: list[list[Any]]):
        self.t = t
        self.scale = params["AmountScale"]
        self.abs_tol, self.rel_tol = params["AbsTolerance"], params["RelTolerance"]
        self.f_tol = params["FactorTolerance"]
        self.lines: list[dict[str, Any]] = []
        self.grid: dict[str, dict[str, float]] = {}
        for tpl, row, col, val in bank_cells:
            self.grid.setdefault(f"{tpl}|{row}", {})[col] = val
        self.g84 = {(row, col): v for (tpl, row, col, v) in bank_cells if tpl == "C 84.00"}
        # Scaled bank amounts, factors as the template applies them, reported SF.
        self.amt: dict[str, list[float]] = {}
        self.eff: dict[str, list[float | None]] = {}
        self.sf: dict[str, float | None] = {}
        for r in t.rows:
            k = f"{r['template']}|{r['row']}"
            g = self.grid.get(k, {})
            acols, fcols, scol = COLS[r["template"]]
            self.amt[k] = [g.get(acols[b], 0.0) * self.scale if b < len(acols) else 0.0 for b in range(4)]
            self.eff[k] = [0.0 if b >= len(acols) or LETTERS[b] not in r["factor_buckets"] else g.get(fcols[b])
                           for b in range(4)]
            self.sf[k] = g[scol] * self.scale if scol in g else None
        self.reperf: dict[str, float] = {}
        self.restat: dict[str, float] = {}
        self.derived_amt: dict[str, list[float]] = {}

    def result(self, diff: float, base: float) -> str:
        if abs(diff) <= self.abs_tol or (base != 0 and abs(diff) / abs(base) <= self.rel_tol):
            return "MATCH"
        return "DIFF"

    def add(self, check: str, tpl: str, row: str, col: str, bank: float | None, other: float | None,
            res: str, impact: float | None = None) -> None:
        self.lines.append({"check": check, "template": tpl, "row": row, "column": col, "bank": bank,
                           "other": other, "result": res, "impact": impact})

    def agg_row(self, r: dict[str, Any], k: str) -> float:
        """Parent: re-performed and restated sums; amounts against the children's."""
        acols = COLS[r["template"]][0]
        kids = [f"{r['template']}|{c}" for c in r["children"]]
        self.reperf[k] = sum(self.reperf[c] for c in kids)
        self.restat[k] = sum(self.restat[c] for c in kids)
        # A parent without a cell in a bucket carries its children's amounts
        # upwards (C 80.00 row 1.7 has no HQLA cell, its child 1.7.3 has one).
        for b in range(4):
            exp = sum(self.derived_amt[c][b] for c in kids)
            if b >= len(acols) or acols[b] not in self.grid.get(k, {}):
                self.derived_amt[k][b] = exp
            else:
                self.add("ARITHMETIC", r["template"], r["row"], acols[b], self.amt[k][b], exp,
                         self.result(exp - self.amt[k][b], self.amt[k][b]))
        return sum(s if (s := self.sf[c]) is not None else self.reperf[c] for c in kids)

    def factor_line(self, r: dict[str, Any], b: int, a: float, f: float | None) -> float:
        """FACTOR line of a leaf bucket; returns the factor used for the restatement."""
        acols, fcols, _ = COLS[r["template"]]
        reg = r["regulatory"][b]
        if f is None:
            self.add("FACTOR", r["template"], r["row"], fcols[b], None, reg, "MISSING",
                     a * reg if reg is not None else None)
            return reg if reg is not None else 0.0
        if reg is None:
            self.add("FACTOR", r["template"], r["row"], fcols[b] if b < len(fcols) else "", f, None, "INFO")
            return f
        col = fcols[b] if LETTERS[b] in r["factor_buckets"] else acols[b]
        self.add("FACTOR", r["template"], r["row"], col, f, reg,
                 "MATCH" if abs(f - reg) <= self.f_tol else "DIFF", a * (reg - f))
        return float(reg)

    def leaf_row(self, r: dict[str, Any], k: str) -> float:
        """Leaf or memo row: factor checks, re-performance and restatement."""
        fcols = COLS[r["template"]][1]
        total = rtotal = 0.0
        for b in range(4):
            a = self.amt[k][b]
            if not a:
                continue
            f = self.eff[k][b]
            if r["kind"] == "MEMO":
                if f is None:  # a memo row's factor feeds C 84.00 (2.3 / 2.4)
                    self.add("FACTOR", r["template"], r["row"], fcols[b], None, None, "MISSING", None)
                total += a * (f or 0.0)
                rtotal += a * (f or 0.0)
            else:
                rtotal += a * self.factor_line(r, b, a, f)
                total += a * (f or 0.0)
        self.reperf[k] = total
        self.restat[k] = rtotal
        return total

    def rows(self) -> None:
        """Every row, deepest first. A parent's reported figures are checked against
        its children's reported figures, so that an error is flagged where it arises;
        re-performance itself never uses the bank's subtotals."""
        for r in sorted(self.t.rows, key=lambda x: -x["id"].count(".")):
            k = f"{r['template']}|{r['row']}"
            self.derived_amt[k] = list(self.amt[k])
            expected = self.agg_row(r, k) if r["kind"] == "AGG" else self.leaf_row(r, k)
            bsf = self.sf[k]
            if bsf is not None:
                self.add("ARITHMETIC", r["template"], r["row"], COLS[r["template"]][2], bsf, expected,
                         self.result(expected - bsf, bsf))

    def summary(self) -> None:
        """C 84.00 against the bank's own rows."""
        rows_rep = {}
        for r in self.t.rows:
            k = f"{r['template']}|{r['row']}"
            rows_rep[k] = (sum(self.derived_amt[k]), s if (s := self.sf[k]) is not None else self.reperf[k])
        derived = summarise(self.t, rows_rep)
        for row, line in derived["lines"].items():
            col = "0020" if line["measure"] == "RSF" else "0030"
            for c, exp in (("0010", line["amount"]), (col, line["value"])):
                if (row, c) in self.g84:
                    b = self.g84[(row, c)] * self.scale
                    self.add("ARITHMETIC", "C 84.00", row, c, b, exp, self.result(exp - b, b))
        if ("0220", "0040") in self.g84 and derived["NSFR"] is not None:
            b = self.g84[("0220", "0040")]
            self.add("ARITHMETIC", "C 84.00", "0220", "0040", b, derived["NSFR"],
                     "MATCH" if abs(derived["NSFR"] - b) <= self.f_tol else "DIFF")

    def headline(self) -> dict[str, tuple[Any, ...]]:
        g84, scale = self.g84, self.scale
        rep_rsf = g84[("0010", "0020")] * scale if ("0010", "0020") in g84 else self.sf["C 80.00|0010"]
        rep_asf = g84[("0120", "0030")] * scale if ("0120", "0030") in g84 else self.sf["C 81.00|0010"]
        rep_nsfr = g84.get(("0220", "0040"), rep_asf / rep_rsf if rep_asf is not None and rep_rsf else None)
        rp_rsf, rp_asf = self.reperf["C 80.00|0010"], self.reperf["C 81.00|0010"]
        rs_rsf, rs_asf = self.restat["C 80.00|0010"], self.restat["C 81.00|0010"]
        return {"reported": (rep_asf, rep_rsf, rep_nsfr), "reperformed": (rp_asf, rp_rsf, rp_asf / rp_rsf),
                "restated": (rs_asf, rs_rsf, rs_asf / rs_rsf)}

    def benchmark(self, eng: dict[str, Any]) -> None:
        for r in self.t.rows:
            k = f"{r['template']}|{r['row']}"
            acols, _, scol = COLS[r["template"]]
            e = eng["rows"][k]
            for b in range(len(acols)):
                if LETTERS[b] in r["amount_buckets"] and (self.amt[k][b] or e["amounts"][b]):
                    self.add("AMOUNT", r["template"], r["row"], acols[b], self.amt[k][b], e["amounts"][b],
                             self.result(e["amounts"][b] - self.amt[k][b], self.amt[k][b]))
            bsf = s if (s := self.sf[k]) is not None else 0.0
            if bsf or e["sf"]:
                self.add("STABLE FUNDING", r["template"], r["row"], scol, bsf, e["sf"],
                         self.result(e["sf"] - bsf, bsf))


def reconcile(t: Tables, params: dict[str, Any], bank_cells: list[list[Any]], eng: dict[str, Any] | None
              ) -> dict[str, Any]:
    rc = Recon(t, params, bank_cells)
    rc.rows()
    rc.summary()
    head = rc.headline()
    if eng is not None and eng["status"] == "VALID":
        rc.benchmark(eng)
        s = eng["summary"]
        head["benchmark"] = (s["ASF"], s["RSF"], s["NSFR"])
    counts: dict[str, int] = {}
    for ln in rc.lines:
        key = f"{ln['check']}|{ln['result']}"
        counts[key] = counts.get(key, 0) + 1
    return {"lines": rc.lines, "headline": head, "counts": counts}


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--write", action="store_true")
    args = ap.parse_args()
    params = DEMO["params"]
    t = Tables(TABLES, params["applied_overrides"])
    eng = run_engine(t, params, DEMO["positions"])
    bank = demo_bank_cells(t, eng)
    rec = reconcile(t, params, bank, eng)
    s = eng["summary"]
    print("status", eng["status"], eng["run_class"], eng["errors"])
    print(f"ASF {s['ASF']:,.2f}  RSF {s['RSF']:,.2f}  NSFR {s['NSFR']:.6%}")
    for name, (a, r, n) in rec["headline"].items():
        print(f"{name:12s} ASF {a:18,.2f} RSF {r:18,.2f} NSFR {n:.6%}")
    print(rec["counts"])
    for ln in rec["lines"]:
        if ln["result"] not in ("MATCH",):
            print(ln)
    if args.write:
        (HERE / "demo-bank-cells.json").write_text(json.dumps(
            {"description": "Bank figures of the demo: the benchmark's results with the perturbations listed in "
                            "demo-positions.json. Generated by nsfr_reference.py --write.",
             "columns": ["Template", "Row", "Column", "Value"], "cells": bank}, indent=1) + "\n")
        exp = {"status": eng["status"], "run_class": eng["run_class"],
               "ASF": s["ASF"], "RSF": s["RSF"], "NSFR": s["NSFR"],
               "summary": {k: [v["amount"], v["value"]] for k, v in s["lines"].items()},
               "rows": {k: v["sf"] for k, v in eng["rows"].items()},
               "row_amounts": {k: v["amounts"] for k, v in eng["rows"].items()},
               "positions": {c["id"]: [c["bucket"], c["row"], c["factor"], c["sf"]] for c in eng["calc"]},
               "headline": rec["headline"], "counts": rec["counts"]}
        (HERE / "demo.expected.json").write_text(json.dumps(exp, indent=1) + "\n")
        print("written")


if __name__ == "__main__":
    main()
