"""Regression fixtures for the VBA static checks (tools/check_vba.py)."""

from __future__ import annotations

import contextlib
import io
import json
import tempfile
import unittest
from pathlib import Path

import check_vba

CORE = "src/core/NSFR_Core_Example.bas"
HOST = "src/host/NSFR_Host_Example.bas"
TEST = "tests/core/Test_NSFR_Example.bas"
CORE_OPTIONS = ("Option Explicit", "Option Private Module")
HOST_OPTIONS = ("Option Explicit",)


def source(path: str, *lines: str, options: tuple[str, ...] | None = None) -> bytes:
    """An exported module: VB_Name, options, then the given lines, with CRLF."""
    stem = Path(path).stem
    chosen = options if options is not None else (HOST_OPTIONS if "/host/" in path else CORE_OPTIONS)
    text = [f'Attribute VB_Name = "{stem}"', *chosen, *lines]
    return "\r\n".join(text).encode("ascii") + b"\r\n"


def findings(path: str, *lines: str, classes: frozenset[str] = frozenset(),
             options: tuple[str, ...] | None = None) -> list[check_vba.Finding]:
    return check_vba.check_source(path, source(path, *lines, options=options), classes)


def rules(path: str, *lines: str, **options: object) -> list[str]:
    return [finding.rule for finding in findings(path, *lines, **options)]  # type: ignore[arg-type]


def exceptions(*entries: dict[str, object]) -> str:
    return json.dumps({"schema_version": 1, "exceptions": list(entries)})


def waiver(path: str = HOST, rule: str = "declare", contains: str = "GetTickCount") -> dict[str, object]:
    return {"path": path, "rule": rule, "contains": contains,
            "reason": "Timer for the host progress display", "issue": 13}


class CleanModuleTests(unittest.TestCase):
    def test_house_style_core_module_passes(self) -> None:
        self.assertEqual(rules(
            CORE,
            "Private Const MAX_RECORDS As Long = 100  'Upper bound",
            "",
            "Public Function ValidateRecord( _",
            "    ByVal Amount As Currency) _",
            "    As Boolean",
            "        ValidateRecord = (Amount >= 0)",
            "End Function",
        ), [])


class CommentAndContinuationTests(unittest.TestCase):
    def test_comments_hide_keywords(self) -> None:
        self.assertEqual(rules(CORE, "' GoTo Done", "Rem GoTo Done", "x = 1 ' On Error Resume Next"), [])

    def test_rem_after_a_colon_ends_the_line(self) -> None:
        self.assertEqual(rules(CORE, "x = 1: Rem GoTo Done"), [])
        self.assertEqual(rules(CORE, "x = 1: Rem a: GoTo Done"), [])
        self.assertEqual(rules(CORE, "Rem: GoTo Done"), [])

    def test_a_continued_comment_stays_a_comment(self) -> None:
        self.assertEqual(rules(CORE, "' This comment continues _", "    GoTo Done"), [])

    def test_continued_code_is_joined_and_reported_on_its_first_line(self) -> None:
        found = findings(HOST, "    Result = Application. _", "        Run(\"Macro\")")
        self.assertEqual([(f.rule, f.line) for f in found], [("dynamic-call", 3)])


class StringTests(unittest.TestCase):
    def test_keywords_inside_strings_are_ignored(self) -> None:
        self.assertEqual(rules(
            CORE, 's = "GoTo CreateObject(""x"") : On Error Resume Next ''"'), [])


class StatementSplittingTests(unittest.TestCase):
    def test_a_statement_after_a_colon_is_checked(self) -> None:
        self.assertEqual(rules(CORE, "x = 1: GoTo Done"), ["goto"])

    def test_a_label_is_not_a_statement(self) -> None:
        self.assertEqual(rules(CORE, "Done: x = 2", "Finish:", "10: x = 3"), [])

    def test_a_keyword_before_a_colon_is_a_statement_not_a_label(self) -> None:
        self.assertEqual(rules(CORE, "Return: x = 1"), ["gosub"])
        self.assertEqual(rules(CORE, "Resume: x = 1"), ["error-handler"])

    def test_named_arguments_do_not_split_statements(self) -> None:
        self.assertEqual(rules(CORE, "Report Value:=1"), [])
        self.assertEqual(rules(CORE, "Report Value:=1: GoSub Helper"), ["gosub"])

    def test_single_line_if_is_checked(self) -> None:
        self.assertEqual(rules(CORE, "If x Then GoTo Done Else x = 1"), ["goto"])


class MemberNameTests(unittest.TestCase):
    def test_members_named_like_keywords_are_not_statements(self) -> None:
        self.assertEqual(rules(
            HOST,
            "Application.Goto Reference:=Target",
            "With Target: .Goto: End With",
            "Item.Run",
            "Item.Evaluate",
            "Set Made = Factory.CreateObject(\"A\")",
        ), [])


class TypeAndEnumTests(unittest.TestCase):
    def test_member_declarations_are_not_calls(self) -> None:
        self.assertEqual(rules(
            CORE,
            "Private Type TRecord",
            "    Evaluate As Boolean",
            "    CallByName As String",
            "End Type",
            "Public Enum NsfrRunKind",
            "    rkRun = 1",
            "End Enum",
            "Record.Evaluate = True",
        ), [])

    def test_procedures_and_variables_named_like_builtins_are_not_calls(self) -> None:
        self.assertEqual(rules(
            CORE,
            "Public Function Evaluate(ByVal Amount As Long) As Long",
            "        Evaluate = Amount",
            "End Function",
            "Private Sub Report(ByVal CallByName As String)",
            "    Dim CreateObject As Boolean",
            "    Total = Evaluate(3)",
            "End Sub",
        ), [])

    def test_later_uses_of_a_variable_named_like_a_builtin_are_not_calls(self) -> None:
        self.assertEqual(rules(
            CORE,
            "Private Sub Report()",
            "    Dim Evaluate As Long",
            "    Evaluate = 3",
            "    Debug.Print Evaluate",
            "End Sub",
            "Private Const CallByName = 1",
            "Private Sub Use()",
            "    Total = CallByName + 1",
            "End Sub",
        ), [])

    def test_a_named_argument_is_not_a_call(self) -> None:
        self.assertEqual(rules(CORE, "Report Evaluate:=3, CallByName:=\"x\""), [])
        self.assertEqual(rules(CORE, "Report Value:=Evaluate(\"1+1\")"), ["dynamic-call"])

    def test_a_procedure_declared_in_another_module_is_not_a_builtin(self) -> None:
        procedures = frozenset({"evaluate"})
        found = check_vba.check_source(CORE, source(CORE, "Total = Evaluate(3)"),
                                       procedures=procedures)
        self.assertEqual(found, [])

    def test_the_excel_function_stays_forbidden_when_qualified(self) -> None:
        self.assertEqual(rules(
            CORE, "Public Function Evaluate() As Long", "End Function",
            'Total = Application.Evaluate("1+1")'), ["dynamic-call"])

    def test_code_after_the_block_is_checked(self) -> None:
        self.assertEqual(rules(
            CORE, "Private Type TRecord", "    Evaluate As Boolean", "End Type",
            "Result = Evaluate(\"1+1\")"), ["dynamic-call"])


class CompileEnvironmentTests(unittest.TestCase):
    DECLARES = (
        "#If VBA7 Then",
        '    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long',
        "#Else",
        '    Private Declare Function GetTickCount Lib "kernel32" () As Long',
        "#End If",
    )

    def test_every_branch_is_checked(self) -> None:
        found = findings(HOST, *self.DECLARES)
        self.assertEqual(sorted((f.rule, f.line) for f in found),
                         [("declare", 4), ("declare", 6), ("declare-ptrsafe", 6)])

    def test_a_reviewed_exception_waives_declare_but_never_ptrsafe(self) -> None:
        kept = check_vba.apply_exceptions(findings(HOST, *self.DECLARES),
                                          check_vba.load_exceptions(exceptions(waiver()))[0])
        self.assertEqual([(f.rule, f.line) for f in kept], [("declare-ptrsafe", 6)])

    def test_host_conditions_may_test_only_vba7_and_win64(self) -> None:
        self.assertEqual(rules(HOST, "#If VBA7 And Win64 Then", "#ElseIf Not VBA7 Then", "#End If"), [])
        self.assertEqual(rules(HOST, "#If VBA7 = True Then", "#ElseIf Win64 = False Then",
                               "#End If"), [])
        self.assertEqual(rules(HOST, "#If Win64 = &H0 Then", "#ElseIf VBA7 <> &O1& Then",
                               "#ElseIf Win64 = 0 Then", "#End If"), [])
        self.assertEqual(rules(HOST, "#If Mac Then", "#End If"), ["conditional-compilation"])

    def test_conditional_compilation_is_host_only(self) -> None:
        self.assertEqual(rules(CORE, "#If VBA7 Then", "#End If"), ["conditional-compilation"])

    def test_compile_constants_are_not_allowed(self) -> None:
        self.assertEqual(rules(HOST, "#Const DEBUG_MODE = 1"), ["compile-constant"])


class ExternalObjectTests(unittest.TestCase):
    def test_create_and_get_object_with_and_without_parentheses(self) -> None:
        self.assertEqual(rules(
            CORE,
            'Set Engine = CreateObject("Scripting.Dictionary")',
            'CreateObject "Excel.Application"',
            'Set Engine = VBA.CreateObject("Scripting.Dictionary")',
            'Set App = GetObject(, "Excel.Application")',
        ), ["com-object"] * 4)

    def test_new_on_a_qualified_type_is_an_external_object(self) -> None:
        self.assertEqual(rules(CORE, "Set Items = New Scripting.Dictionary"), ["com-object"])
        self.assertEqual(rules(CORE, "Dim Items As New Scripting.Dictionary"), ["com-object"])

    def test_new_on_builtins_and_project_classes_is_allowed(self) -> None:
        classes = frozenset({"cfundingrecord"})
        self.assertEqual(rules(CORE, "Dim Items As New Collection",
                               "Set Items = New VBA.Collection",
                               "Set Record = New CFundingRecord", classes=classes), [])

    def test_new_on_an_unknown_class_is_reported(self) -> None:
        self.assertEqual(rules(CORE, "Set Items = New Dictionary"), ["com-object"])

    def test_a_host_exception_waives_an_external_object(self) -> None:
        found = findings(HOST, 'Set Shell = CreateObject("WScript.Shell")')
        entry = waiver(rule="com-object", contains="WScript.Shell")
        kept = check_vba.apply_exceptions(found, check_vba.load_exceptions(exceptions(entry))[0])
        self.assertEqual(kept, [])


class DynamicCallTests(unittest.TestCase):
    def test_dynamic_calls_are_reported(self) -> None:
        self.assertEqual(rules(
            CORE,
            'Result = CallByName(Target, "Run", VbMethod)',
            'Application.Run "Macro"',
            'Result = Application.Evaluate("1+1")',
            'Result = Evaluate("1+1")',
            'Result = ExecuteExcel4Macro("GET.WORKSPACE(1)")',
        ), ["dynamic-call"] * 5)


class LayerTests(unittest.TestCase):
    def test_error_handlers_only_at_the_host_and_in_tests(self) -> None:
        self.assertEqual(rules(CORE, "On Error GoTo Failed"), ["error-handler"])
        self.assertEqual(rules("src/api/NSFR_API_Example.bas", "Resume Next"), ["error-handler"])
        self.assertEqual(rules(HOST, "On Error GoTo Failed", "Resume CleanUp", "On Error GoTo 0"), [])
        self.assertEqual(rules(TEST, "On Error GoTo Failed", "Resume Next"), [])

    def test_resume_next_on_error_and_plain_goto_are_never_allowed(self) -> None:
        self.assertEqual(rules(HOST, "On Error Resume Next"), ["on-error-resume-next"])
        self.assertEqual(rules(HOST, "GoTo Done"), ["goto"])
        self.assertEqual(rules(HOST, "On Index GoTo First, Second"), ["goto"])
        self.assertEqual(rules(HOST, "GoSub Helper", "Return"), ["gosub", "gosub"])

    def test_declare_outside_the_host_is_reported(self) -> None:
        found = findings(CORE, 'Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long')
        self.assertEqual([f.rule for f in found], ["declare"])
        self.assertIn("only in src/host/", found[0].message)

    def test_vba_outside_the_layout_is_reported(self) -> None:
        self.assertIn("vba-location", rules("misc/Loose.bas"))


class ModuleRuleTests(unittest.TestCase):
    def test_option_explicit_is_required(self) -> None:
        self.assertEqual(rules(HOST, "x = 1", options=()), ["option-explicit"])

    def test_option_private_module_is_required_in_core_and_tests(self) -> None:
        self.assertEqual(rules(CORE, options=("Option Explicit",)), ["option-private-module"])
        self.assertEqual(rules(TEST, options=("Option Explicit",)), ["option-private-module"])
        self.assertEqual(rules(HOST, options=("Option Explicit",)), [])

    def test_vb_name_matches_the_file_name(self) -> None:
        raw = source(CORE).replace(b"NSFR_Core_Example", b"NSFR_Core_Other", 1)
        self.assertEqual([f.rule for f in check_vba.check_source(CORE, raw)], ["vb-name"])
        raw = b"Option Explicit\r\nOption Private Module\r\n"
        self.assertEqual([f.rule for f in check_vba.check_source(CORE, raw)], ["vb-name"])

    def test_non_ascii_bytes_are_reported_with_line_and_column(self) -> None:
        raw = source(CORE, 'Label = "Caf?"').replace(b"Caf?", b"Caf\xe9")
        found = check_vba.check_source(CORE, raw)
        self.assertEqual([(f.rule, f.line) for f in found], [("ascii", 4)])
        self.assertIn("column 13", found[0].message)


class HeaderTests(unittest.TestCase):
    def test_class_and_form_headers_are_skipped(self) -> None:
        cls = ("VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\n"
               'Attribute VB_Name = "CFundingRecord"\r\nOption Explicit\r\n').encode("ascii")
        self.assertEqual(check_vba.check_source("src/api/CFundingRecord.cls", cls), [])
        frm = ("VERSION 5.00\r\nBegin VB.UserForm frmRunOptions\r\n   Caption = \"GoTo\"\r\n"
               "   BeginProperty Font\r\n      Name = \"Arial\"\r\n   EndProperty\r\n"
               "   Begin VB.CommandButton cmdRun\r\n   End\r\nEnd\r\n"
               'Attribute VB_Name = "frmRunOptions"\r\nOption Explicit\r\nGoTo Done\r\n').encode("ascii")
        found = check_vba.check_source("src/host/frmRunOptions.frm", frm)
        self.assertEqual([(f.rule, f.line) for f in found], [("goto", 12)])


class ExceptionFileTests(unittest.TestCase):
    def problems(self, raw: str) -> list[str]:
        return [f.message for f in check_vba.load_exceptions(raw)[1]]

    def test_the_repository_file_is_valid(self) -> None:
        raw = (Path(__file__).resolve().parents[1] / check_vba.EXCEPTIONS_FILE).read_text("utf-8")
        self.assertEqual(check_vba.load_exceptions(raw), ([], []))

    def test_exceptions_are_limited_to_host_boundaries(self) -> None:
        self.assertIn("src/host/", self.problems(exceptions(waiver(path=CORE)))[0])
        self.assertIn("rule must be", self.problems(exceptions(waiver(rule="goto")))[0])
        self.assertIn("rule must be", self.problems(exceptions(waiver(rule="declare-ptrsafe")))[0])

    def test_malformed_files_are_reported(self) -> None:
        self.assertEqual(len(self.problems("{")), 1)
        self.assertEqual(len(self.problems('{"schema_version": 2, "exceptions": []}')), 1)
        entry = waiver()
        del entry["reason"]
        self.assertIn("exactly", self.problems(exceptions(entry))[0])
        self.assertIn("issue", self.problems(exceptions({**waiver(), "issue": True}))[0])

    def test_an_unused_exception_is_reported(self) -> None:
        kept = check_vba.apply_exceptions([], check_vba.load_exceptions(exceptions(waiver()))[0])
        self.assertEqual([f.rule for f in kept], ["exception-unused"])


class CommandLineTests(unittest.TestCase):
    def run_main(self, files: dict[str, bytes], waivers: str | None = exceptions()) -> tuple[int, str]:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for name, data in files.items():
                (root / name).parent.mkdir(parents=True, exist_ok=True)
                (root / name).write_bytes(data)
            if waivers is not None:
                (root / check_vba.EXCEPTIONS_FILE).parent.mkdir(parents=True, exist_ok=True)
                (root / check_vba.EXCEPTIONS_FILE).write_text(waivers, encoding="utf-8")
            output = io.StringIO()
            with contextlib.redirect_stdout(output), contextlib.redirect_stderr(output):
                code = check_vba.main(["--root", str(root), *files])
        return code, output.getvalue()

    def test_findings_are_reported_with_file_and_line(self) -> None:
        code, output = self.run_main({CORE: source(CORE, "GoTo Done")})
        self.assertEqual(code, 1)
        self.assertIn(f"{CORE}:4: [goto] GoTo is allowed only as On Error GoTo", output)

    def test_a_clean_tree_passes(self) -> None:
        code, output = self.run_main({CORE: source(CORE, "x = 1")})
        self.assertEqual(code, 0, output)
        self.assertIn("1 VBA file(s), 0 finding(s)", output)

    def test_a_missing_exception_file_cannot_run(self) -> None:
        code, _ = self.run_main({CORE: source(CORE)}, waivers=None)
        self.assertEqual(code, 2)

    def test_only_the_prototype_directory_is_not_checked(self) -> None:
        self.assertFalse(check_vba.admitted("prototype/src/core/NSFR_Core_Engine.bas"))
        for path in ("src/core/NSFR_Core_Engine.bas", "prototypes/Loose.bas", "tests/prototype/Loose.bas",
                     "src/prototype/Loose.bas"):
            with self.subTest(path=path):
                self.assertTrue(check_vba.admitted(path))
        code, output = self.run_main({"prototype/src/core/Loose.bas": b"GoTo Done\n",
                                      CORE: source(CORE, "x = 1")})
        self.assertEqual(code, 0, output)
        self.assertIn("1 VBA file(s), 0 finding(s); 1 not admitted in prototype/, not checked", output)

    def test_vba_beside_the_prototype_directory_is_still_checked(self) -> None:
        code, output = self.run_main({"prototypes/Loose.bas": source("prototypes/Loose.bas")})
        self.assertEqual(code, 1)
        self.assertIn("[vba-location]", output)

    def test_the_repository_passes(self) -> None:
        output = io.StringIO()
        with contextlib.redirect_stdout(output):
            self.assertEqual(check_vba.main([]), 0, output.getvalue())



class ManifestTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        (self.root / "src" / "host").mkdir(parents=True)
        (self.root / "src" / "host" / "NSFR_Host_A.bas").write_bytes(b'Attribute VB_Name = "NSFR_Host_A"\r\n')
        (self.root / "src" / "host" / "NSFR_Host_B.bas").write_bytes(b'Attribute VB_Name = "NSFR_Host_B"\r\n')
        self.paths = ["src/host/NSFR_Host_A.bas", "src/host/NSFR_Host_B.bas"]

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def found(self, manifest: str | None) -> list[str]:
        if manifest is not None:
            (self.root / "src" / "manifest.txt").write_text(manifest, encoding="ascii")
        return [f.message for f in check_vba.manifest_findings(self.root, self.paths)]

    def test_a_complete_manifest_passes(self) -> None:
        self.assertEqual(self.found("# comment\n\nNSFR_Host_A src/host/NSFR_Host_A.bas\n"
                                    "NSFR_Host_B src/host/NSFR_Host_B.bas\n"), [])

    def test_a_missing_manifest_is_reported(self) -> None:
        self.assertEqual(self.found(None), ["the source manifest is missing"])

    def test_an_unlisted_file_and_a_wrong_name_are_reported(self) -> None:
        self.assertEqual(self.found("NSFR_Host_X src/host/NSFR_Host_A.bas\n"),
                         ["src/host/NSFR_Host_A.bas does not declare NSFR_Host_X",
                          "src/host/NSFR_Host_B.bas is not listed"])

    def test_a_bad_line_a_duplicate_and_a_foreign_path_are_reported(self) -> None:
        found = self.found("NSFR_Host_A src/host/NSFR_Host_A.bas\nNSFR_Host_A src/host/NSFR_Host_A.bas\n"
                           "NSFR_Host_B src/host/NSFR_Host_B.bas extra\nNSFR_Core_Z src/core/Z.bas\n")
        self.assertIn("src/host/NSFR_Host_A.bas is listed twice", found)
        self.assertIn("a line must be 'name path'", found)
        self.assertIn("src/core/Z.bas is not an admitted VBA file", found)


if __name__ == "__main__":
    unittest.main()
