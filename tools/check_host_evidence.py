#!/usr/bin/env python3
"""Validate the host evidence records in tests/evidence/ against Git history.

This Source Code Form is subject to the terms of the Mozilla Public
License, v. 2.0. If a copy of the MPL was not distributed with this
file, You can obtain one at https://mozilla.org/MPL/2.0/.

A record states what one run on the supported host did with one exact commit
(docs/validation/HOST_EVIDENCE.md). This checker reads each record and
rejects one that is not bound to its candidate: an unknown or unreachable
commit, a dirty worktree, a fingerprint that differs from the commit's files,
an unsupported host, an inconsistent or non-finite summary, a result that the
steps do not support, or a machine detail that must not be published. With
--candidate SHA it also requires a PASS record for that commit, and reports a
missing or stale one. Read-only; it never runs Excel.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import re
import subprocess
import sys
from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import Any

EVIDENCE = Path("tests/evidence")
SCHEMA = "k-nsfr-host-evidence/1"
HARNESS_DIR = "tests/core/"
CONFIG = (".gitattributes",)
VBA_SUFFIXES = (".bas", ".cls", ".frm")  # checked out with CRLF (.gitattributes)
EXPECTED_SOURCE = "tests/README.md"

PASS, FAIL, NOT_RUN = "PASS", "FAIL", "NOT_RUN"
STEPS = ("import", "compile", "run", "cleanup")
KEYS = {"schema", "candidate_sha", "recorded_on", "worktree_clean", "host", "fingerprints",
        "steps", "result", "output", "reason"}
REQUIRED = KEYS - {"output", "reason"}
HOST_KEYS = {"os", "excel", "bitness"}

SHA = re.compile(r"^[0-9a-f]{40}$")
WINDOWS_DETAILS = re.compile(r"\b\d{2}H\d\b.*\bOS build \d+\.\d+\b")
SHA256 = re.compile(r"^[0-9a-f]{64}$")
CASE = re.compile(r"^([A-Z0-9-]{1,32}) (PASS|FAIL|SKIPPED|MISSING)(?: .*)?$")
COUNT_NAMES = ("declared", "executed", "passed", "failed", "skipped", "missing",
               "undeclared", "cleanup_failed", "problems")
COUNTS = re.compile(r"^COUNTS " + " ".join(rf"{name}=(\S+)" for name in COUNT_NAMES) + "$")
# Details that identify a person or a machine never belong in a public record.
PRIVATE = (
    (re.compile(r"[A-Za-z]:\\+(?:Users|Documents and Settings)\\", re.I), "a user folder path"),
    (re.compile(r"(?:^|[\s(\"'])/(?:Users|home)/", re.I), "a user folder path"),
    (re.compile(r"\\\\[A-Za-z0-9._-]+\\"), "a network share path"),
    (re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}"), "an e-mail address"),
    (re.compile(r"\b(?:\d{1,3}\.){3}\d{1,3}\b"), "an IP address"),
)


def git(*args: str, root: Path) -> bytes:
    return subprocess.check_output(["git", *args], cwd=root, stderr=subprocess.DEVNULL)


def git_ok(*args: str, root: Path) -> bool:
    return subprocess.run(["git", *args], cwd=root, capture_output=True).returncode == 0


def strings(value: Any) -> list[str]:
    if isinstance(value, str):
        return [value]
    if isinstance(value, Mapping):
        return [s for item in value.values() for s in strings(item)]
    if isinstance(value, list):
        return [s for item in value for s in strings(item)]
    return []


def checkout_bytes(path: str, blob: bytes) -> bytes:
    """The file as Git checks it out on Windows: VBA source with CRLF, the rest with LF."""
    return blob.replace(b"\n", b"\r\n") if path.endswith(VBA_SUFFIXES) else blob


def expected_fingerprints(sha: str, root: Path) -> dict[str, str]:
    names = git("ls-tree", "-r", "--name-only", sha, "--", HARNESS_DIR, *CONFIG,
                root=root).decode().split()
    modules = [n for n in names if n.startswith(HARNESS_DIR) and n.endswith(VBA_SUFFIXES)]
    return {name: hashlib.sha256(checkout_bytes(name, git("cat-file", "blob", f"{sha}:{name}",
                                                          root=root))).hexdigest()
            for name in [*modules, *[n for n in CONFIG if n in names]]}


def expected_output(sha: str, root: Path) -> list[str] | None:
    """The self-test summary that tests/README.md documents at the candidate."""
    try:
        text = git("show", f"{sha}:{EXPECTED_SOURCE}", root=root).decode()
    except subprocess.CalledProcessError:
        return None
    match = re.search(r"the self-test prints:\s*```text\n(.*?)\n```", text, re.S)
    return match.group(1).split("\n") if match else None


def summary_findings(output: Sequence[str]) -> list[str]:
    """Counts that are not finite non-negative integers, or that contradict the lines."""
    counts_lines = [line for line in output if line.startswith("COUNTS ")]
    if len(counts_lines) != 1:
        return [f"output: expected one COUNTS line, found {len(counts_lines)}"]
    match = COUNTS.match(counts_lines[0])
    if not match:
        return ["output: the COUNTS line does not list the nine counts in order"]
    if not all(value.isdigit() and value.isascii() for value in match.groups()):
        return ["output: every count must be a finite non-negative integer"]
    n = dict(zip(COUNT_NAMES, (int(v) for v in match.groups()), strict=True))
    # The case lines come after SUITE and before the cleanup, problem and count lines.
    end = next(i for i, line in enumerate(output)
               if line.startswith(("CLEANUP FAILED ", "PROBLEM ", "COUNTS ")))
    body = output[1:end] if output[0].startswith("SUITE ") else output[:end]
    cases = [m.group(2) for m in map(CASE.match, body) if m]
    if len(cases) != len(body):
        findings_head = [f"output: {len(body) - len(cases)} case lines are not 'ID OUTCOME'"]
    else:
        findings_head = []
    tally = {outcome: cases.count(outcome) for outcome in ("PASS", "FAIL", "SKIPPED", "MISSING")}
    findings = findings_head
    expected = {
        "declared": len(cases),
        "passed": tally["PASS"], "failed": tally["FAIL"],
        "skipped": tally["SKIPPED"], "missing": tally["MISSING"],
        "executed": n["passed"] + n["failed"],
        "cleanup_failed": sum(line.startswith("CLEANUP FAILED ") for line in output),
        "problems": sum(line.startswith("PROBLEM ") for line in output),
    }
    for name, value in expected.items():
        if n[name] != value:
            findings.append(f"output: {name}={n[name]} but the lines give {value}")
    passed = (n["declared"] > 0 and n["passed"] == n["declared"]
              and n["cleanup_failed"] == 0 and n["problems"] == 0)
    result = output[-1] if output else ""
    if result not in ("RESULT PASS", "RESULT FAIL"):
        findings.append("output: the last line must be RESULT PASS or RESULT FAIL")
    elif (result == "RESULT PASS") != passed:
        findings.append(f"output: {result} contradicts the counts")
    return findings


def step_findings(record: Mapping[str, Any]) -> list[str]:
    steps = record.get("steps")
    if not isinstance(steps, Mapping) or set(steps) != set(STEPS) or any(
            steps[s] not in (PASS, FAIL, NOT_RUN) for s in STEPS):
        return [f"steps: must give {', '.join(STEPS)}, each PASS, FAIL or NOT_RUN"]
    findings = []
    for before, after in zip(STEPS, STEPS[1:3], strict=False):
        if steps[before] != PASS and steps[after] != NOT_RUN:
            findings.append(f"steps: {after} is {steps[after]} although {before} is {steps[before]}")
    if steps["import"] == NOT_RUN and steps["cleanup"] != NOT_RUN:
        findings.append("steps: cleanup cannot run without an import")
    values = set(steps.values())
    result = PASS if values == {PASS} else FAIL if FAIL in values else NOT_RUN
    if record.get("result") != result:
        findings.append(f"result: the steps give {result}, the record says {record.get('result')}")
    if result != PASS and not str(record.get("reason", "")).strip():
        findings.append("reason: required when a step failed or did not run")
    return findings


def output_findings(record: Mapping[str, Any], expected: list[str] | None) -> list[str]:
    run = (record.get("steps") or {}).get("run") if isinstance(record.get("steps"), Mapping) else None
    output = record.get("output")
    if run == NOT_RUN:
        return [] if output is None else ["output: given although the run did not happen"]
    if not isinstance(output, list) or not all(isinstance(s, str) for s in output):
        return ["output: the complete Immediate-window output is required, one string per line"]
    if not output:
        # A run stopped by a runtime error returns no summary at all.
        return [] if run == FAIL else ["output: a passing run prints the self-test summary"]
    if expected is None:
        return [f"output: no documented self-test output in {EXPECTED_SOURCE} at the candidate"]
    if run == PASS and output != expected:
        return [f"output: differs from the self-test output documented in {EXPECTED_SOURCE}"]
    if run == FAIL and output == expected:
        return ["output: equals the documented passing output, so the run did not fail"]
    return summary_findings(output) if run == PASS or any(s.startswith("COUNTS ")
                                                           for s in output) else []


def host_findings(host: Any) -> list[str]:
    if not isinstance(host, Mapping) or set(host) != HOST_KEYS or not all(
            isinstance(host[k], str) and host[k].strip() for k in HOST_KEYS):
        return [f"host: must give exactly {sorted(HOST_KEYS)} as text"]
    findings = []
    if not host["os"].startswith("Windows 11"):
        findings.append("host: os must be Windows 11, the supported host")
    elif not WINDOWS_DETAILS.search(host["os"]):
        findings.append("host: os must give the Windows version and OS build, "
                        "as Settings > System > About shows them (24H2, OS build 26100.2314)")
    if "Microsoft 365" not in host["excel"] or not re.search(r"Version \S+ \(Build [\d.]+",
                                                              host["excel"]):
        findings.append("host: excel must be Microsoft 365 with its version and build, "
                        "as About Excel shows them")
    if host["bitness"] != "64-bit":
        findings.append("host: bitness must be 64-bit, the supported host")
    return findings


def candidate_findings(record: Mapping[str, Any], name: str, root: Path,
                       today: dt.date) -> tuple[list[str], str | None]:
    sha = record.get("candidate_sha")
    if not isinstance(sha, str) or not SHA.match(sha):
        return ["candidate_sha: must be a full 40-character lower-case SHA"], None
    findings = []
    if name != f"{sha}.json":
        findings.append(f"file name must be {sha}.json")
    if not git_ok("cat-file", "-e", f"{sha}^{{commit}}", root=root):
        return findings + [f"candidate_sha: commit {sha} is not in this clone; "
                           "fetch the full history"], None
    if not git_ok("merge-base", "--is-ancestor", sha, "HEAD", root=root):
        findings.append(f"candidate_sha: {sha} is not an ancestor of HEAD; test a commit of "
                        "the release branch or main, never one that a squash merge discards")
    recorded = record.get("recorded_on")
    committed = dt.date.fromisoformat(git("show", "-s", "--format=%cs", sha, root=root)
                                      .decode().strip())
    try:
        day = dt.date.fromisoformat(str(recorded))
    except ValueError:
        findings.append("recorded_on: must be an ISO date YYYY-MM-DD")
    else:
        if not committed <= day <= today:
            findings.append(f"recorded_on: {day} must fall between the commit date "
                            f"{committed} and today")
    return findings, sha


def record_findings(path: Path, root: Path, today: dt.date) -> list[str]:
    try:
        record = json.loads(path.read_text(encoding="utf-8"))
    except (ValueError, UnicodeDecodeError) as error:
        return [f"not valid UTF-8 JSON: {error}"]
    if not isinstance(record, dict):
        return ["must be a JSON object"]
    findings = []
    if set(record) - KEYS:
        findings.append(f"unknown fields {sorted(set(record) - KEYS)}")
    if REQUIRED - set(record):
        return findings + [f"missing fields {sorted(REQUIRED - set(record))}"]
    if record["schema"] != SCHEMA:
        findings.append(f"schema: must be {SCHEMA}")
    if record["worktree_clean"] is not True:
        findings.append("worktree_clean: the modules must come from a clean checkout")
    for text in strings(record):
        for pattern, label in PRIVATE:
            if pattern.search(text):
                findings.append(f"contains {label}; remove machine and personal details")
    findings += host_findings(record["host"]) + step_findings(record)
    more, sha = candidate_findings(record, path.name, root, today)
    findings += more
    if sha is None:
        return findings
    fingerprints = record["fingerprints"]
    expected = expected_fingerprints(sha, root)
    if not isinstance(fingerprints, Mapping) or set(fingerprints) != set(expected):
        findings.append(f"fingerprints: must cover exactly {sorted(expected)}")
    else:
        for name, digest in sorted(fingerprints.items()):
            if not isinstance(digest, str) or not SHA256.match(digest.lower()):
                findings.append(f"fingerprints: {name} is not a SHA-256")
            elif digest.lower() != expected[name]:
                findings.append(f"fingerprints: {name} differs from the file at {sha[:12]}")
    return findings + output_findings(record, expected_output(sha, root))


def check(root: Path, candidate: str | None = None,
          today: dt.date | None = None) -> tuple[int, list[str]]:
    """The number of records and every finding; with a candidate, its own as well."""
    today = today or dt.date.today()
    paths = sorted((root / EVIDENCE).glob("*.json"))
    findings = [f"{path.relative_to(root).as_posix()}: {finding}"
                for path in paths for finding in record_findings(path, root, today)]
    if candidate is not None:
        own = root / EVIDENCE / f"{candidate}.json"
        if not own.exists():
            kind = "stale: records exist only for other commits" if paths else "missing"
            findings.append(f"candidate {candidate}: no host evidence ({kind})")
        else:
            try:
                result = json.loads(own.read_text(encoding="utf-8")).get("result")
            except (ValueError, AttributeError):
                result = None
            if result != PASS:
                findings.append(f"candidate {candidate}: its host evidence is not PASS")
    return len(paths), findings


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Validate the host evidence records.")
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--candidate", help="also require a PASS record for this full SHA")
    args = parser.parse_args(argv)
    count, findings = check(args.root, args.candidate)
    for finding in findings:
        print(finding)
    print(f"host-evidence: {'FAIL' if findings else 'PASS'} "
          f"({count} records, {len(findings)} findings)")
    return 1 if findings else 0


if __name__ == "__main__":
    sys.exit(main())
