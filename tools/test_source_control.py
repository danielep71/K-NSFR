"""Exercise committed-content checks and real Git checkout round trips."""

import json
from pathlib import Path
import subprocess
import tempfile
import unittest

from check_source_control import check, git


class SourceControlTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.repo = Path(self.temp.name) / "source"
        self.repo.mkdir()
        git("init", "-q", cwd=self.repo)
        git("config", "user.name", "Fixture", cwd=self.repo)
        git("config", "user.email", "fixture@example.invalid", cwd=self.repo)
        git("config", "core.autocrlf", "false", cwd=self.repo)
        self.write("tools/binary-allowlist.json", b"[]\n")
        attributes = Path(__file__).resolve().parents[1] / ".gitattributes"
        self.write(".gitattributes", attributes.read_bytes())

    def write(self, name, data):
        path = self.repo / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(data)

    def commit(self):
        git("add", "-f", ".", cwd=self.repo)
        git("commit", "-qm", "fixture", cwd=self.repo)

    def test_fresh_checkout_round_trip(self):
        original = b'Attribute VB_Name = "Example"\nOption Explicit\n'
        self.write("src/core/Example.bas", original)
        self.write("docs/example.md", "Synthetic café\n".encode())
        self.commit()
        for setting in ("false", "true", "input"):
            with self.subTest(autocrlf=setting):
                clone = Path(self.temp.name) / setting
                git("clone", "-q", "-c", f"core.autocrlf={setting}",
                    str(self.repo), str(clone))
                self.assertEqual((clone / "src/core/Example.bas").read_bytes(),
                                 original.replace(b"\n", b"\r\n"))
                self.assertNotIn(b"\r", (clone / "docs/example.md").read_bytes())
                self.assertEqual(check(cwd=clone)[2], [])
                git("add", "--renormalize", ".", cwd=clone)
                self.assertEqual(git("diff", "--cached", "--name-only", cwd=clone), b"")

    def test_dirty_checkout_does_not_change_committed_result(self):
        self.write("example.txt", b"clean\n")
        self.commit()
        self.write("example.txt", b"dirty  \r\n")
        self.assertEqual(check(cwd=self.repo)[2], [])

    def test_bad_committed_text(self):
        self.write("bad.txt", b"trailing \nmissing-newline")
        self.write("bad-encoding.txt", b"\xff\n")
        self.write("bom.txt", b"\xef\xbb\xbfbom\n")
        self.commit()
        findings = "\n".join(check(cwd=self.repo)[2])
        for expected in ("trailing whitespace", "missing final newline", "invalid UTF-8", "BOM"):
            self.assertIn(expected, findings)

    def test_committed_crlf_is_rejected(self):
        self.write("bad.txt", b"bad\r\n")
        self.commit()
        # Write an unnormalized blob directly to exercise committed-byte checks.
        oid = subprocess.check_output(["git", "hash-object", "-w", "--stdin"],
                                      cwd=self.repo, input=b"bad\r\n").decode().strip()
        git("update-index", "--cacheinfo", "100644", oid, "bad.txt", cwd=self.repo)
        git("commit", "-qm", "malformed committed blob", cwd=self.repo)
        self.assertIn("committed content must use LF", "\n".join(check(cwd=self.repo)[2]))

    def test_binary_requires_exact_existing_entry(self):
        self.write("templates/example.xlsx", b"PK\x00\xff")
        self.commit()
        self.assertTrue(check(cwd=self.repo)[2])
        self.write("tools/binary-allowlist.json",
                   json.dumps(["templates/example.xlsx"]).encode() + b"\n")
        self.commit()
        self.assertEqual(check(cwd=self.repo)[2], [])
        self.write("tools/binary-allowlist.json", b'["templates/*.xlsx"]\n')
        self.commit()
        with self.assertRaises(ValueError):
            check(cwd=self.repo)

    def test_stale_allowlist_is_rejected(self):
        self.write("tools/binary-allowlist.json", b'["missing.png"]\n')
        self.commit()
        self.assertIn("does not exist", "\n".join(check(cwd=self.repo)[2]))

    def test_ignored_outputs_and_trackable_sources(self):
        ignore = Path(__file__).resolve().parents[1] / ".gitignore"
        self.write(".gitignore", ignore.read_bytes())
        for name in ("test-output/run.json", "build/example.xlsm", "~$example.xlsx",
                     ".env.local", "private.pem", "tools/__pycache__/test.pyc"):
            with self.subTest(path=name):
                self.assertEqual(git("check-ignore", name, cwd=self.repo).decode().strip(), name)
        tracked = subprocess.run(
            ["git", "check-ignore", "--non-matching", "-v", "src/core/Example.bas",
             "tests/fixtures/example.json", "docs/report.md", "src/form.frx"],
            cwd=self.repo, capture_output=True, text=True,
        )
        self.assertEqual(tracked.returncode, 1)
        self.assertEqual(len(tracked.stdout.splitlines()), 4)
        self.assertTrue(all(line.startswith("::\t") for line in tracked.stdout.splitlines()))

    def test_markdown_breaks_and_host_encoding_are_preserved(self):
        self.write("docs/example.md", b"line  \nnext\n")
        self.write("src/example.bas", b"' synthetic accented byte: \xe9\n")
        self.commit()
        self.assertEqual(check(cwd=self.repo)[2], [])


if __name__ == "__main__":
    unittest.main()
