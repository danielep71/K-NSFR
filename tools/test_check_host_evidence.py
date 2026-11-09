"""Tests of check_host_evidence.py against a temporary Git repository."""

from __future__ import annotations

import copy
import datetime as dt
import hashlib
import io
import json
import tempfile
import unittest
from pathlib import Path
from typing import Any
from unittest import mock

import check_host_evidence as he
from check_source_control import git

EXPECTED = ["SUITE HARNESS-SELF-TEST", "H-001 PASS", "H-002 PASS",
            "COUNTS declared=2 executed=2 passed=2 failed=0 skipped=0 missing=0 "
            "undeclared=0 cleanup_failed=0 problems=0", "RESULT PASS"]
FAILED = ["SUITE HARNESS-SELF-TEST", "H-001 PASS", "H-002 FAIL line 3 differs",
          "COUNTS declared=2 executed=2 passed=1 failed=1 skipped=0 missing=0 "
          "undeclared=0 cleanup_failed=0 problems=0", "RESULT FAIL"]
MODULE = b"Attribute VB_Name = \"Test_Example\"\nOption Explicit\n"
ATTRIBUTES = b"* text=auto eol=lf\n*.bas text eol=crlf\n"
TODAY = dt.date(2099, 1, 1)


class HostEvidenceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        git("init", "-q", cwd=self.root)
        git("config", "user.name", "Fixture", cwd=self.root)
        git("config", "user.email", "fixture@example.invalid", cwd=self.root)
        self.write("tests/core/Test_Example.bas", MODULE)
        self.write(".gitattributes", ATTRIBUTES)
        readme = ("# Harness\n\nWhen the harness behaves as documented, the self-test prints:\n\n"
                  "```text\n" + "\n".join(EXPECTED) + "\n```\n")
        self.write("tests/README.md", readme.encode())
        git("add", "-A", cwd=self.root)
        git("commit", "-qm", "fixture", cwd=self.root)
        self.sha = git("rev-parse", "HEAD", cwd=self.root).decode().strip()
        git("update-ref", "refs/remotes/origin/main", self.sha, cwd=self.root)

    def write(self, name: str, data: bytes) -> None:
        path = self.root / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(data)

    def record(self) -> dict[str, Any]:
        return {
            "schema": he.SCHEMA,
            "candidate_sha": self.sha,
            "recorded_on": "2098-12-31",
            "worktree_clean": True,
            "host": {"os": "Windows 11 Pro 24H2, OS build 26100.2314",
                     "excel": "Microsoft 365 Apps for enterprise, Version 2410 "
                              "(Build 18129.20158 Click-to-Run)",
                     "bitness": "64-bit"},
            "fingerprints": {
                "tests/core/Test_Example.bas": hashlib.sha256(MODULE).hexdigest(),
                ".gitattributes": hashlib.sha256(ATTRIBUTES).hexdigest()},
            "steps": {"import": "PASS", "compile": "PASS", "run": "PASS", "cleanup": "PASS"},
            "result": "PASS",
            "output": list(EXPECTED),
        }

    def findings(self, record: dict[str, Any], name: str | None = None,
                 candidate: str | None = None) -> list[str]:
        self.write(f"tests/evidence/{name or self.sha + '.json'}",
                   json.dumps(record).encode())
        return he.check(self.root, candidate, TODAY)[1]

    def assertFinding(self, fragment: str, record: dict[str, Any], **kwargs: Any) -> None:
        findings = self.findings(record, **kwargs)
        self.assertTrue(any(fragment in f for f in findings),
                        f"no finding containing {fragment!r} in {findings}")

    def test_a_valid_pass_record(self) -> None:
        self.assertEqual(self.findings(self.record(), candidate=self.sha), [])

    def test_no_record_passes_but_a_candidate_needs_one(self) -> None:
        self.assertEqual(he.check(self.root, None, TODAY), (0, []))
        self.assertEqual(he.check(self.root, self.sha, TODAY)[1],
                         [f"candidate {self.sha}: no host evidence (missing)"])

    def test_a_record_for_another_commit_is_stale(self) -> None:
        self.findings(self.record())
        other = "f" * 40
        self.assertIn(f"candidate {other}: no host evidence (stale: records exist only for "
                      "other commits)", he.check(self.root, other, TODAY)[1])

    def test_a_failed_record_does_not_certify_the_candidate(self) -> None:
        record = self.record()
        record.update(steps={**record["steps"], "run": "FAIL"}, result="FAIL",
                      output=FAILED, reason="H-002 failed")
        self.assertEqual(self.findings(record), [])
        self.assertIn(f"candidate {self.sha}: its host evidence is not PASS",
                      self.findings(record, candidate=self.sha))

    def test_a_run_stopped_by_an_error_has_empty_output(self) -> None:
        record = self.record()
        record.update(steps={**record["steps"], "run": "FAIL"}, result="FAIL", output=[],
                      reason="Run-time error 9: Subscript out of range")
        self.assertEqual(self.findings(record), [])
        record = self.record()
        record["output"] = []
        self.assertFinding("output: a passing run prints the self-test summary", record)

    def test_not_run_needs_a_reason_and_no_output(self) -> None:
        record = self.record()
        del record["output"]
        record.update(steps=dict.fromkeys(he.STEPS, "NOT_RUN"), result="NOT_RUN",
                      reason="No supported host available")
        self.assertEqual(self.findings(record), [])
        del record["reason"]
        self.assertFinding("reason: required", record)
        record.update(reason="x", output=EXPECTED)
        self.assertFinding("output: given although the run did not happen", record)

    def test_the_result_must_follow_the_steps(self) -> None:
        record = self.record()
        record["steps"]["cleanup"] = "FAIL"
        self.assertFinding("result: the steps give FAIL", record)
        record = self.record()
        record["steps"].update(compile="FAIL")
        record.update(result="FAIL", reason="compile error")
        self.assertFinding("steps: run is PASS although compile is FAIL", record)

    def test_an_unknown_or_unreachable_candidate_is_rejected(self) -> None:
        record = self.record()
        record["candidate_sha"] = "a" * 40
        self.assertFinding("is not in this clone", record, name="a" * 40 + ".json")
        record["candidate_sha"] = self.sha[:12]
        self.assertFinding("must be a full 40-character", record)
        git("checkout", "-q", "--orphan", "other", cwd=self.root)
        git("commit", "-qm", "unrelated", cwd=self.root)
        self.assertFinding("is not an ancestor of HEAD", self.record())

    def test_task_evidence_fails_in_synthetic_merge_and_after_squash(self) -> None:
        base = self.sha
        git("checkout", "-qb", "task", cwd=self.root)
        self.write("task.txt", b"task change\n")
        git("add", "task.txt", cwd=self.root)
        git("commit", "-qm", "task", cwd=self.root)
        self.sha = git("rev-parse", "HEAD", cwd=self.root).decode().strip()
        task_record = self.record()
        git("checkout", "-qb", "synthetic", base, cwd=self.root)
        git("merge", "--no-ff", "-qm", "synthetic merge", "task", cwd=self.root)
        self.assertFinding("not in published main/release history", task_record)
        (self.root / he.EVIDENCE / f"{self.sha}.json").unlink()
        git("checkout", "-qb", "release/0.1.0", base, cwd=self.root)
        git("merge", "--squash", "task", cwd=self.root)
        git("commit", "-qm", "squash task", cwd=self.root)
        git("update-ref", "refs/remotes/origin/release/0.1.0", "HEAD", cwd=self.root)
        self.assertFinding("not an ancestor of HEAD", task_record)
        (self.root / he.EVIDENCE / f"{self.sha}.json").unlink()
        self.sha = base
        self.assertEqual(self.findings(self.record()), [])
        self.sha = git("rev-parse", "HEAD", cwd=self.root).decode().strip()
        self.assertEqual(self.findings(self.record()), [])

    def test_local_or_task_refs_do_not_replace_published_history(self) -> None:
        git("update-ref", "-d", "refs/remotes/origin/main", cwd=self.root)
        git("update-ref", "refs/heads/main", self.sha, cwd=self.root)
        git("update-ref", "refs/remotes/origin/task", self.sha, cwd=self.root)
        git("update-ref", "refs/remotes/origin/main-extra", self.sha, cwd=self.root)
        self.assertFinding("not in published main/release history", self.record())

    def test_the_file_name_binds_the_candidate(self) -> None:
        self.assertFinding("file name must be", self.record(), name="evidence.json")

    def test_a_dirty_worktree_is_rejected(self) -> None:
        record = self.record()
        record["worktree_clean"] = False
        self.assertFinding("worktree_clean", record)

    def test_fingerprints_must_match_the_checked_out_files(self) -> None:
        record = self.record()
        record["fingerprints"][".gitattributes"] = "0" * 64
        self.assertFinding("fingerprints: .gitattributes differs", record)
        record = self.record()
        # The digest is of the bytes stored in Git, not of the CRLF checkout.
        record["fingerprints"]["tests/core/Test_Example.bas"] = hashlib.sha256(
            MODULE.replace(b"\n", b"\r\n")).hexdigest()
        self.assertFinding("fingerprints: tests/core/Test_Example.bas differs", record)
        record = self.record()
        del record["fingerprints"][".gitattributes"]
        self.assertFinding("fingerprints: must cover exactly", record)

    def test_the_host_must_be_the_supported_one(self) -> None:
        for key, value, fragment in (("bitness", "32-bit", "bitness must be 64-bit"),
                                     ("os", "Windows 10 Pro", "os must be Windows 11"),
                                     ("os", "Windows 11", "the Windows version and OS build"),
                                     ("os", "Windows 11 Pro 24H2", "the Windows version and OS build"),
                                     ("os", "Windows 11 Pro, OS build 26100.2314",
                                      "the Windows version and OS build"),
                                     ("excel", "Microsoft 365", "with its version and build")):
            with self.subTest(key=key):
                record = self.record()
                record["host"][key] = value
                self.assertFinding(fragment, record)

    def test_a_pass_output_must_equal_the_documented_output(self) -> None:
        record = self.record()
        record["output"] = EXPECTED[:-1] + ["RESULT FAIL"]
        self.assertFinding("output: differs from the self-test output", record)
        del record["output"]
        self.assertFinding("output: the complete Immediate-window output is required", record)

    def test_counts_must_be_finite_and_consistent(self) -> None:
        cases = {
            "declared=nan": "finite non-negative integer",
            "declared=-1": "finite non-negative integer",
            "declared=3": "declared=3 but the lines give 2",
            "executed=1": "executed=1 but the lines give 2",
        }
        for replace, fragment in cases.items():
            with self.subTest(replace=replace):
                output = copy.copy(FAILED)
                name = replace.split("=")[0]
                output[3] = " ".join(replace if part.startswith(name + "=") else part
                                     for part in output[3].split())
                self.assertTrue(any(fragment in f for f in he.summary_findings(output)),
                                he.summary_findings(output))
        self.assertIn("output: RESULT PASS contradicts the counts",
                      he.summary_findings(FAILED[:-1] + ["RESULT PASS"]))
        self.assertEqual(he.summary_findings(EXPECTED), [])

    def test_private_machine_details_are_rejected(self) -> None:
        for text, fragment in ((r"Imported from C:\Users\someone\src", "user folder path"),
                               ("see /home/someone/k", "user folder path"),
                               (r"\\fileserver\share", "network share path"),
                               ("ask someone@example.com", "e-mail address"),
                               ("host 192.168.1.10", "IP address")):
            with self.subTest(fragment=fragment):
                record = self.record()
                record.update(steps={**record["steps"], "cleanup": "FAIL"}, result="FAIL",
                              reason=text)
                self.assertFinding(fragment, record)

    def test_unknown_fields_and_dates(self) -> None:
        record = self.record()
        record["machine_name"] = "WORKSTATION"
        self.assertFinding("unknown fields ['machine_name']", record)
        record = self.record()
        record["recorded_on"] = "2000-01-01"
        self.assertFinding("recorded_on: 2000-01-01 must fall between", record)

    def test_the_branch_verifier_runs_the_candidate_recorder_in_a_worktree(self) -> None:
        script = (Path(__file__).resolve().parent / "host" / "Verify-NSFR-Branch.ps1").read_text(
            encoding="ascii")
        for needed in ("'worktree', 'add', '--detach', $checkout, $sha",
                       "worktree remove --force $checkout",
                       "Join-Path $checkout 'tools\\host\\Record-HostEvidence.ps1'",
                       "tests\\evidence\\$sha.json",
                       "[Parameter(Mandatory)][string]$Branch",
                       "$exitCode -ne 0", "if ($result -ne 'PASS') { exit 1 }"):
            self.assertIn(needed, script)
        # Like the recorder, the verifier never changes Trust Center settings.
        for forbidden in ("AccessVBOM", "Set-ItemProperty", "New-ItemProperty", "Remove-ItemProperty",
                          "AutomationSecurity"):
            self.assertNotIn(forbidden, script)

    def test_the_recording_script_writes_the_record_the_gate_reads(self) -> None:
        script = (Path(__file__).resolve().parent / "host" / "Record-HostEvidence.ps1").read_text(
            encoding="ascii")
        self.assertIn(f"$Schema = '{he.SCHEMA}'", script)
        for key in sorted(he.KEYS - {"output", "reason"}):
            self.assertRegex(script, rf"\n\s+{key}\s+=", key)
        self.assertIn("$record.output", script)
        self.assertIn("$record.reason", script)
        steps = "[ordered]@{ " + "; ".join(f"{s} = 'NOT_RUN'" for s in he.STEPS) + " }"
        self.assertIn(steps, script)
        self.assertIn("ConvertTo-Json -InputObject $record", script)
        self.assertIn("New-Object Text.UTF8Encoding $false", script)  # no BOM
        # The script never changes Trust Center settings.
        for forbidden in ("AccessVBOM", "Set-ItemProperty", "New-ItemProperty", "Remove-ItemProperty",
                          "AutomationSecurity"):
            self.assertNotIn(forbidden, script)

    def test_the_inventory_prints_the_expected_fingerprints(self) -> None:
        with mock.patch("sys.stdout", new_callable=io.StringIO) as stdout:
            code = he.main(["--root", str(self.root), "--candidate", self.sha, "--inventory"])
        self.assertEqual((code, json.loads(stdout.getvalue())),
                         (0, self.record()["fingerprints"]))

    def test_the_repository_records_pass(self) -> None:
        root = Path(__file__).resolve().parents[1]
        self.assertEqual(he.check(root)[1], [])


if __name__ == "__main__":
    unittest.main()
