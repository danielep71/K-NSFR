"""Headless regression of the K-NSFR prototype in LibreOffice (VBA compatibility mode).

Not a substitute for Excel: LibreOffice runs the exported VBA with
`Option VBASupport 1`, which is close to, but not, the VBE. It is the check
that runs without Excel: it loads templates/NSFR_Template.xlsx, inserts the
modules of src/ into the document's Basic library and runs three cases:

1. demo      NSFR_Host_Main.RunNSFR_Silent on the demo; Results, C80, C81,
             PositionCalc and Reconciliation against
             tests/derivations/prototype/demo.expected.json, and every
             TestCatalogue row PASS;
2. invalid   the demo with one position given an unknown category: the run
             must be INVALID, its totals withheld, and the bank checks run;
3. csv       ImportBankCsvFrom on a CSV written from the demo bank cells in
             the export format, then a run: the same headline as case 1;
4. defect    the demo bank figures with a memo-row factor removed and a
             non-numeric C 84.00 NSFR: a FACTOR MISSING line, an ERROR and
             a WARNING. A separate unit case (scan) checks that the import
             keeps a worksheet error value as "ERROR".

Harness-only shims, applied to an in-memory copy of the source, never to the
files: `Attribute VB_*` lines are dropped, `vbObjectError + n` constants are
replaced by their value, `Option Private Module` is removed and
`Debug.Print` is dropped.

Usage: python tools/lo_regression.py [--keep DIR] [--bank-workbook PATH]
--bank-workbook additionally empties Positions, imports a bank's reporting
workbook through ImportBankWorkbookFrom and runs the bank-only checks on it
(case 4, printed, not compared); its data never leave the temporary copy.
Needs LibreOffice (soffice), its Python UNO bridge and openpyxl.
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from typing import Any

import openpyxl  # type: ignore[import-untyped,import-not-found,unused-ignore]
import uno  # type: ignore[import-not-found,unused-ignore]
from com.sun.star.beans import PropertyValue  # type: ignore[import-not-found,unused-ignore]

ROOT = Path(__file__).resolve().parents[1]
MODULES = ["src/core/NSFR_Core_Config.bas", "src/core/NSFR_Core_Util.bas", "src/core/NSFR_Core_Engine.bas",
           "src/core/NSFR_Core_Reconcile.bas", "src/core/NSFR_Core_Import.bas", "src/api/NSFR_API_Functions.bas",
           "src/host/NSFR_Host_Main.bas"]
DER = ROOT / "tests/derivations/prototype"
EXPECTED = json.loads((DER / "demo.expected.json").read_text())
BANK = json.loads((DER / "demo-bank-cells.json").read_text())


def prop(name: str, value: Any) -> Any:
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


def shim(source: str) -> str:
    lines = source.replace("\r\n", "\n").split("\n")
    source = "\n".join(line for line in lines if not line.startswith("Attribute VB_"))
    source = re.sub(r"vbObjectError \+ (\d+)", lambda m: str(-2147221504 + int(m.group(1))), source)
    source = source.replace("Option Private Module", "")
    source = re.sub(r"^\s*Debug\.Print .*$", "", source, flags=re.M)
    return "Option VBASupport 1\n" + source


class Office:
    def __init__(self) -> None:
        self.port = 2003
        self.profile = Path(tempfile.mkdtemp(prefix="knsfr_lo_"))
        self.proc = subprocess.Popen(["soffice", "--headless", "--invisible", "--norestore", "--nologo",
                                      f"--accept=socket,host=localhost,port={self.port};urp;",
                                      f"-env:UserInstallation={self.profile.as_uri()}"],
                                     stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
        for _ in range(120):
            try:
                self.ctx = resolver.resolve(f"uno:socket,host=localhost,port={self.port};urp;"
                                            "StarOffice.ComponentContext")
                break
            except Exception:
                time.sleep(0.5)
        else:
            raise RuntimeError("LibreOffice did not start")
        self.desktop = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)

    def open(self, path: Path, extra_module: str = "") -> Any:
        doc = self.desktop.loadComponentFromURL(uno.systemPathToFileUrl(str(path)), "_blank", 0,
                                                (prop("Hidden", True), prop("MacroExecutionMode", 4)))
        lib = doc.BasicLibraries.getByName("Standard") if doc.BasicLibraries.hasByName("Standard") \
            else doc.BasicLibraries.createLibrary("Standard")
        for module in MODULES:
            source = (ROOT / module).read_text(encoding="ascii")
            name = re.search(r'Attribute VB_Name = "([^"]+)"', source)
            assert name is not None, module
            lib.insertByName(name.group(1), shim(source))
        lib.insertByName("TEST_LO", "Option VBASupport 1\n" + extra_module)
        return doc

    def call(self, doc: Any, proc: str, *args: Any) -> str:
        script = doc.getScriptProvider().getScript(
            f"vnd.sun.star.script:Standard.TEST_LO.{proc}?language=Basic&location=document")
        return str(script.invoke(tuple(args), (), ())[0])

    def save(self, doc: Any, out: Path) -> None:
        doc.calculateAll()
        doc.storeToURL(uno.systemPathToFileUrl(str(out)), (prop("FilterName", "Calc Office Open XML"),))
        doc.close(True)

    def close(self) -> None:
        try:
            self.desktop.terminate()
        except Exception:
            pass
        self.proc.wait(timeout=60)


DRIVER = """Option Explicit
Public Function RunEngine() As String
    RunEngine = NSFR_Host_Main.RunNSFR_Silent()
End Function
Public Function BreakPosition() As String
    ThisComponent.Sheets.getByName("Positions").getCellRangeByName("E8").setString("NOT_A_CATEGORY")
    BreakPosition = "ok"
End Function
Public Function ImportCsv(ByVal Path As String) As String
    ImportCsv = NSFR_Host_Main.ImportBankCsvFrom(Path)
End Function
Public Function ScanErrorValue() As String
    Dim V(1 To 4, 1 To 6) As Variant
    Dim Cells() As Variant
    Dim N As Long
    Dim P As String
    Dim T As String
    Dim Item As Variant
    V(1, 2) = "C 81.00 - NSFR - AVAILABLE STABLE FUNDING"
    V(2, 5) = "0010"
    V(2, 6) = "0100"
    V(3, 2) = "0150"
    V(3, 5) = 16102549
    V(3, 6) = CVErr(2042)
    T = NSFR_Core_Import.TemplateOfSheet(V)
    NSFR_Core_Import.ScanTemplateSheet V, T, "test", Cells, N, P
    Item = Cells(2)
    ScanErrorValue = T & "|" & N & "|" & Item(4) & "|" & P
End Function
Public Function InjectBankDefect() As String
    Dim Sh As Object
    Dim r As Long
    Dim Hits As Long
    Sh = ThisComponent.Sheets.getByName("BankCells")
    For r = 10 To 600
        If Sh.getCellByPosition(2, r).getString() = "C 81.00" And Sh.getCellByPosition(3, r).getString() = "0150" _
           And Sh.getCellByPosition(4, r).getString() = "0010" Then
            Sh.getCellByPosition(5, r).setValue(16102549)
            Hits = Hits + 1
        End If
        If Sh.getCellByPosition(2, r).getString() = "C 81.00" And Sh.getCellByPosition(3, r).getString() = "0150" _
           And Sh.getCellByPosition(4, r).getString() = "0070" Then
            Sh.getCellByPosition(5, r).setFormula("")
            Hits = Hits + 1
        End If
        If Sh.getCellByPosition(2, r).getString() = "C 84.00" And Sh.getCellByPosition(3, r).getString() = "0220" _
           And Sh.getCellByPosition(4, r).getString() = "0040" Then
            Sh.getCellByPosition(5, r).setString("#N/A")
            Hits = Hits + 1
        End If
    Next r
    InjectBankDefect = CStr(Hits)
End Function
Public Function ClearPositions() As String
    ThisComponent.Sheets.getByName("Positions").getCellRangeByName("C6:L5005").clearContents(7)
    ClearPositions = "ok"
End Function
Public Function ImportBook(ByVal Path As String) As String
    ImportBook = NSFR_Host_Main.ImportBankWorkbookFrom(Path)
End Function
"""


def near(a: Any, b: float) -> bool:
    return isinstance(a, (int, float)) and abs(a - b) <= 1e-6 * max(1.0, abs(b))


def check_demo(out: Path) -> list[str]:
    wb = openpyxl.load_workbook(out, data_only=True)
    problems: list[str] = []
    res = wb["Results"]
    if res["D6"].value != EXPECTED["status"]:
        problems.append(f"status {res['D6'].value!r}")
    for cell, key in (("D10", "ASF"), ("D11", "RSF"), ("D12", "NSFR")):
        if not near(res[cell].value, EXPECTED[key]):
            problems.append(f"{key}: {res[cell].value!r} expected {EXPECTED[key]!r}")
    tc = wb["TestCatalogue"]
    fails = [f"{tc.cell(r, 3).value} {tc.cell(r, 4).value}: expected {tc.cell(r, 5).value!r}, got "
             f"{tc.cell(r, 6).value!r}" for r in range(7, tc.max_row + 1)
             if tc.cell(r, 7).value not in ("PASS", None)]
    problems += fails
    print("   TestCatalogue:", tc["D4"].value)
    return problems


def check_invalid(out: Path) -> list[str]:
    wb = openpyxl.load_workbook(out, data_only=True)
    res, rc = wb["Results"], wb["Reconciliation"]
    problems = []
    if not str(res["D6"].value).startswith("INVALID"):
        problems.append(f"status {res['D6'].value!r}, expected INVALID")
    if res["D12"].value != "WITHHELD":
        problems.append(f"NSFR {res['D12'].value!r}, expected WITHHELD")
    if rc["I9"].value is not None:
        problems.append("benchmark level written although the run is INVALID")
    if not near(rc["K6"].value, EXPECTED["headline"]["reported"][2]):
        problems.append(f"reported NSFR {rc['K6'].value!r}")
    if rc["I18"].value != 0:
        problems.append(f"AMOUNT lines {rc['I18'].value!r}, expected none")
    pc = wb["PositionCalc"]
    if not str(pc["N8"].value).startswith("REJECTED"):
        problems.append(f"PositionCalc row 8 status {pc['N8'].value!r}")
    return problems


def check_defect(out: Path) -> list[str]:
    wb = openpyxl.load_workbook(out, data_only=True)
    rc, ck = wb["Reconciliation"], wb["Checks"]
    problems = []
    missing = [(rc.cell(r, 4).value, rc.cell(r, 5).value) for r in range(23, rc.max_row + 1)
               if rc.cell(r, 3).value == "FACTOR" and rc.cell(r, 13).value == "MISSING"]
    if ("C 81.00", "0150") not in missing:
        problems.append(f"no FACTOR MISSING line for C 81.00 r0150: {missing}")
    msgs = [(ck.cell(r, 3).value, ck.cell(r, 6).value) for r in range(6, ck.max_row + 1) if ck.cell(r, 3).value]
    if not any(sev == "ERROR" and ref and "r0220" in ref for sev, ref in msgs):
        problems.append("no ERROR for the non-numeric C 84.00 NSFR")
    if not any(sev == "WARNING" and ref and "r0220" in ref for sev, ref in msgs):
        problems.append("no WARNING on the reported level")
    return problems


def write_csv(path: Path) -> None:
    lines = ["RowColumn_Index,Suffix,Amount"]
    for tpl, row, col, val in BANK["cells"]:
        lines.append(f"{tpl} c{col} r{row},a,{val!r}")
    path.write_text("\r\n".join(lines) + "\r\n")


def check_csv(out: Path) -> list[str]:
    wb = openpyxl.load_workbook(out, data_only=True)
    rc = wb["Reconciliation"]
    problems = []
    for li, lvl in enumerate(["reported", "reperformed", "restated", "benchmark"]):
        for ci in range(3):
            got = rc.cell(6 + li, 9 + ci).value
            if not near(got, EXPECTED["headline"][lvl][ci]):
                problems.append(f"{lvl} {'ASF RSF NSFR'.split()[ci]}: {got!r}")
    src = wb["BankCells"]["G11"].value
    if not str(src).endswith("line 2"):
        problems.append(f"BankCells source {src!r}, expected the CSV line")
    return problems


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--keep", type=Path, help="directory to keep the calculated workbooks in")
    ap.add_argument("--bank-workbook", type=Path, help="also import and reconcile this bank workbook")
    args = ap.parse_args()
    work = Path(tempfile.mkdtemp(prefix="knsfr_"))
    # A copy, so that LibreOffice's lock file never lands next to the template.
    template = work / "template.xlsx"
    shutil.copy(ROOT / "templates/NSFR_Template.xlsx", template)
    office = Office()
    failed = False
    try:
        doc = office.open(template, DRIVER)
        scan = office.call(doc, "ScanErrorValue")
        doc.close(True)
        print("[scan] TemplateOfSheet|cells|value of an error cell|problem:", scan)
        if scan != "C 81.00|2|ERROR|":
            print("[scan] FAIL: expected 'C 81.00|2|ERROR|'")
            failed = True
        else:
            print("[scan] PASS")
        cases = [("demo", None), ("invalid", "BreakPosition"), ("csv", "csv"), ("defect", "InjectBankDefect")]
        for case, step in cases:
            doc = office.open(template, DRIVER)
            if step in ("BreakPosition", "InjectBankDefect"):
                print(f"[{case}] {step}:", office.call(doc, step))
            if step == "csv":
                csv = work / "bank-export.csv"
                write_csv(csv)
                imported = office.call(doc, "ImportCsv", str(csv))
                print(f"[{case}] import:", (imported.splitlines() or [repr(imported)])[0])
            outcome = office.call(doc, "RunEngine")
            print(f"[{case}] run:", outcome)
            out = work / f"{case}.xlsx"
            office.save(doc, out)
            problems = {"demo": check_demo, "invalid": check_invalid, "csv": check_csv,
                        "defect": check_defect}[case](out)
            if not outcome.startswith("OK"):
                problems.insert(0, outcome)
            print(f"[{case}]", "PASS" if not problems else f"FAIL ({len(problems)})")
            for p in problems[:25]:
                print("    ", p)
            failed |= bool(problems)
        if args.bank_workbook:
            doc = office.open(template, DRIVER)
            office.call(doc, "ClearPositions")
            print("[bank] import:", office.call(doc, "ImportBook", str(args.bank_workbook.resolve())))
            print("[bank] run:", office.call(doc, "RunEngine"))
            office.save(doc, work / "bank.xlsx")
            print("[bank] saved", work / "bank.xlsx")
    finally:
        office.close()
    if args.keep:
        args.keep.mkdir(parents=True, exist_ok=True)
        for f in work.glob("*.xlsx"):
            shutil.copy(f, args.keep / f.name)
    print("REGRESSION", "FAIL" if failed else "PASS")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
