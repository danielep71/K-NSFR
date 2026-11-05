"""Tests for the repository integrity runner (tools/check.py)."""

from __future__ import annotations

import contextlib
import io
import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import check

MISSING_TOOL = "k-nsfr-tool-that-does-not-exist"


def gate(name: str, code: int) -> check.Gate:
    return check.Gate(name, "python", (sys.executable, "-c", f"raise SystemExit({code})"))


def result(status: str) -> check.Result:
    return check.Result("gate", status, "")


class OverallStatusTests(unittest.TestCase):
    def status(self, *statuses: str, ci: bool = False, dirty: bool = False,
               allow_missing: bool = False) -> str:
        return check.overall_status([result(s) for s in statuses], ci=ci, dirty=dirty,
                                    allow_missing=allow_missing)

    def test_all_pass(self) -> None:
        self.assertEqual(self.status("pass", "pass"), "pass")

    def test_any_failure_fails(self) -> None:
        self.assertEqual(self.status("pass", "fail"), "fail")
        self.assertEqual(self.status("fail", "not-run", allow_missing=True), "fail")

    def test_not_run_is_never_pass(self) -> None:
        self.assertEqual(self.status("pass", "not-run"), "fail")
        self.assertEqual(self.status("pass", "not-run", allow_missing=True), "incomplete")

    def test_ci_requires_every_tool_and_a_clean_tree(self) -> None:
        self.assertEqual(self.status("pass", "not-run", ci=True, allow_missing=True), "fail")
        self.assertEqual(self.status("pass", ci=True, dirty=True), "fail")
        self.assertEqual(self.status("pass", ci=False, dirty=True), "pass")


class RunGatesTests(unittest.TestCase):
    def test_outcomes_and_logs(self) -> None:
        gates = [gate("ok", 0), gate("bad", 3), check.Gate("absent", MISSING_TOOL, (MISSING_TOOL,))]
        with tempfile.TemporaryDirectory() as directory, contextlib.redirect_stdout(io.StringIO()):
            output = Path(directory)
            results = check.run_gates(gates, output, output)
            self.assertEqual([r.status for r in results], ["pass", "fail", "not-run"])
            self.assertEqual(results[1].detail, "exit 3")
            self.assertIn("not installed", (output / "absent.log").read_text(encoding="utf-8"))
            self.assertEqual(sorted(p.name for p in output.iterdir()),
                             ["absent.log", "bad.log", "ok.log"])

    def test_probe_outcome_is_used(self) -> None:
        probe = check.Gate("probe", "python", probe=lambda root, scratch: (1, "rejected"))
        with tempfile.TemporaryDirectory() as directory, contextlib.redirect_stdout(io.StringIO()):
            output = Path(directory)
            self.assertEqual(check.run_gates([probe], output, output)[0].status, "fail")
            self.assertEqual((output / "probe.log").read_text(encoding="utf-8"), "rejected")


class MainTests(unittest.TestCase):
    def setUp(self) -> None:
        self.directory = tempfile.TemporaryDirectory()
        self.root = Path(self.directory.name)
        for args in (("init", "-q"), ("config", "user.email", "test@example.invalid"),
                     ("config", "user.name", "Test"), ("config", "commit.gpgsign", "false")):
            subprocess.run(["git", "-C", str(self.root), *args], check=True)
        (self.root / ".gitignore").write_text("test-results/\n", encoding="utf-8")
        subprocess.run(["git", "-C", str(self.root), "add", ".gitignore"], check=True)
        subprocess.run(["git", "-C", str(self.root), "commit", "-q", "-m", "init"], check=True)

    def tearDown(self) -> None:
        self.directory.cleanup()

    def main(self, *argv: str, gates: list[check.Gate]) -> int:
        with contextlib.redirect_stdout(io.StringIO()):
            return check.main(["--root", str(self.root), *argv], gates)

    def report(self) -> dict[str, object]:
        path = self.root / "test-results" / "repository-integrity.json"
        value: dict[str, object] = json.loads(path.read_text(encoding="utf-8"))
        return value

    def test_report_binds_the_candidate(self) -> None:
        self.assertEqual(self.main("--ci", gates=[gate("ok", 0)]), 0)
        report = self.report()
        head = subprocess.run(["git", "-C", str(self.root), "rev-parse", "HEAD"],
                              capture_output=True, text=True, check=True).stdout.strip()
        self.assertEqual(report["candidate_sha"], head)
        self.assertEqual((report["status"], report["mode"], report["dirty"]), ("pass", "ci", False))
        markdown = (self.root / "test-results" / "repository-integrity.md").read_text("utf-8")
        self.assertIn("Status: **PASS**", markdown)
        self.assertIn(head, markdown)

    def test_failing_gate_fails_the_run(self) -> None:
        self.assertEqual(self.main(gates=[gate("ok", 0), gate("bad", 1)]), 1)
        self.assertEqual(self.report()["status"], "fail")

    def test_missing_tool_fails_unless_allowed(self) -> None:
        absent = check.Gate("absent", MISSING_TOOL, (MISSING_TOOL,))
        self.assertEqual(self.main(gates=[absent]), 1)
        self.assertEqual(self.main("--allow-missing", gates=[absent]), 0)
        self.assertEqual(self.report()["status"], "incomplete")
        self.assertEqual(self.report()["tools"], {MISSING_TOOL: "not installed"})

    def test_ci_rejects_a_dirty_tree(self) -> None:
        (self.root / "untracked.txt").write_text("x\n", encoding="utf-8")
        self.assertEqual(self.main("--ci", gates=[gate("ok", 0)]), 1)
        self.assertEqual(self.report()["dirty"], True)

    def test_ci_and_allow_missing_are_exclusive(self) -> None:
        with contextlib.redirect_stderr(io.StringIO()), self.assertRaises(SystemExit):
            self.main("--ci", "--allow-missing", gates=[])


class DefaultGatesTests(unittest.TestCase):
    def test_names_are_unique_and_cover_every_tool(self) -> None:
        gates = check.default_gates()
        self.assertEqual(len({g.name for g in gates}), len(gates))
        self.assertEqual({g.tool for g in gates},
                         {"python", "python-libraries", "node", "ruff", "mypy", "actionlint",
                          "pyyaml"})
        for g in gates:
            self.assertTrue(bool(g.command) != (g.probe is not None), g.name)
        self.assertIn("workflow-invariants", {g.name for g in gates})
        self.assertIn("contracts", {g.name for g in gates})

    def test_each_gate_names_what_it_needs(self) -> None:
        tools = {g.name: g.tool for g in check.default_gates()}
        # These import PyYAML and markdown-it-py; issue-config runs the Node validator,
        # and so do the issue tests in tool-tests.
        self.assertEqual(tools["tool-tests"], "python-libraries")
        tool_tests = next(g for g in check.default_gates() if g.name == "tool-tests")
        self.assertEqual(tool_tests.tools(), ("python-libraries", "node"))
        self.assertEqual(tools["content"], "python-libraries")
        self.assertEqual(tools["issue-config"], "node")
        self.assertEqual(tools["workflow-invariants"], "pyyaml")
        mypy = next(g for g in check.default_gates() if g.name == "mypy")
        self.assertEqual(mypy.tools(), ("mypy", "python-libraries"))


class LibraryAvailabilityTests(unittest.TestCase):
    def local(self) -> dict[str, str]:
        return {key: value for key, value in os.environ.items() if key != "GITHUB_ACTIONS"}

    def test_any_missing_library_skips_the_gate_locally(self) -> None:
        def find_spec(name: str) -> object | None:
            return None if name == "markdown_it" else object()

        with mock.patch("importlib.util.find_spec", side_effect=find_spec), \
                mock.patch.dict(os.environ, self.local(), clear=True):
            self.assertFalse(check.available("python-libraries"))
            self.assertTrue(check.available("pyyaml"))

    def test_gate_runs_on_github_actions(self) -> None:
        with mock.patch("importlib.util.find_spec", return_value=None), \
                mock.patch.dict(os.environ, {"GITHUB_ACTIONS": "true"}):
            self.assertTrue(check.available("python-libraries"))

    def test_missing_library_is_not_run_and_the_run_incomplete(self) -> None:
        gates = [gate("ok", 0),
                 check.Gate("content", "python-libraries", (sys.executable, "-c", "pass"))]
        with tempfile.TemporaryDirectory() as directory, \
                contextlib.redirect_stdout(io.StringIO()), \
                mock.patch("importlib.util.find_spec", return_value=None), \
                mock.patch.dict(os.environ, self.local(), clear=True):
            output = Path(directory)
            results = check.run_gates(gates, output, output)
        self.assertEqual([r.status for r in results], ["pass", "not-run"])
        self.assertEqual(check.overall_status(results, ci=False, dirty=False, allow_missing=True),
                         "incomplete")

    def test_versions_name_each_library(self) -> None:
        with mock.patch("importlib.metadata.version", return_value="1.0"):
            self.assertEqual(check.library_versions("python-libraries"),
                             "PyYAML 1.0, markdown-it-py 1.0")


if __name__ == "__main__":
    unittest.main()
