"""Tests of the workbook template generator (tools/build_template.py)."""

from __future__ import annotations

import re
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

import build_template as bt


def build_to(folder: Path, name: str = "t.xlsx") -> Path:
    path = folder / name
    bt.build(path)
    return path


def workbook_xml(path: Path) -> str:
    with zipfile.ZipFile(path) as z:
        return z.read("xl/workbook.xml").decode("utf-8")


class TemplateTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.folder = Path(self.tmp.name)

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def test_two_builds_give_the_same_bytes(self) -> None:
        first = build_to(self.folder, "a.xlsx").read_bytes()
        second = build_to(self.folder, "b.xlsx").read_bytes()
        self.assertEqual(first, second)

    def test_the_committed_template_is_the_generators_output(self) -> None:
        self.assertEqual(bt.check(bt.TEMPLATE), [])

    def test_every_sheet_and_name_of_the_layout_is_present(self) -> None:
        xml = workbook_xml(build_to(self.folder))
        sheets = re.findall(r'<sheet name="([^"]+)"', xml)
        self.assertEqual(tuple(sheets), bt.SHEETS)
        # Autofilters add Excel's own hidden _xlnm names; the layout's names are the others.
        names = {n for n in re.findall(r'<definedName name="([^"]+)"', xml) if not n.startswith("_xlnm.")}
        self.assertEqual(names, set(bt.NAMES))

    def test_the_template_is_macro_free_and_holds_no_number(self) -> None:
        path = build_to(self.folder)
        self.assertEqual(bt.macro_findings(path), [])
        with zipfile.ZipFile(path) as z:
            sheets = [z.read(n).decode("utf-8") for n in z.namelist() if n.startswith("xl/worksheets/sheet")]
            self.assertFalse([n for n in z.namelist() if "vmlDrawing" in n or "ctrlProp" in n],
                             "no button or form control")
        for xml in sheets:
            # Every cell is text, a blank or a formula: no factor, amount or result is written.
            self.assertNotRegex(xml, r'<c r="[A-Z]+\d+"(?: s="\d+")?><v>')
        self.assertTrue(any("COUNTIFS" in xml for xml in sheets))

    def test_the_input_filters_cover_every_input_row(self) -> None:
        path = build_to(self.folder)
        with zipfile.ZipFile(path) as z:
            xml = {n: z.read(n).decode("utf-8") for n in z.namelist() if n.startswith("xl/worksheets/sheet")}
        filters = {re.search(r'<autoFilter ref="([^"]+)"', x).group(1)  # type: ignore[union-attr]
                   for x in xml.values() if "<autoFilter" in x}
        self.assertIn(f"C5:L{5 + bt.INPUT_ROWS}", filters)    # Positions
        self.assertIn(f"C10:G{10 + bt.INPUT_ROWS}", filters)  # BankCells

    def test_a_vba_project_or_macro_content_type_is_reported(self) -> None:
        path = self.folder / "m.xlsm"
        with zipfile.ZipFile(path, "w") as z:
            z.writestr("[Content_Types].xml", "application/vnd.ms-excel.sheet.macroEnabled.main+xml")
            z.writestr("xl/vbaProject.bin", b"\x00")
        found = bt.macro_findings(path)
        self.assertEqual(len(found), 2)

    def test_a_drifted_or_missing_template_fails_the_check(self) -> None:
        drifted = build_to(self.folder)
        with zipfile.ZipFile(drifted, "a") as z:
            z.writestr("extra.txt", "x")
        with mock.patch.object(bt, "ROOT", self.folder):
            self.assertTrue(any("differs" in f for f in bt.check(drifted)))
            self.assertTrue(any("missing" in f for f in bt.check(self.folder / "absent.xlsx")))

    def test_the_main_entry_point_reports_the_check(self) -> None:
        with mock.patch.object(bt, "check", return_value=[]):
            self.assertEqual(bt.main(["--check"]), 0)
        with mock.patch.object(bt, "check", return_value=["drift"]):
            self.assertEqual(bt.main(["--check"]), 1)
        out = self.folder / "out.xlsx"
        self.assertEqual(bt.main(["--output", str(out)]), 0)
        self.assertTrue(out.is_file())


if __name__ == "__main__":
    unittest.main()
