"""Tests of check_links.py with a fake fetcher; no network access."""

from __future__ import annotations

import socket
import tempfile
import unittest
import urllib.error
from collections.abc import Callable
from pathlib import Path
from unittest import mock

import check_links as cl

REPO = "owner/project"


def fetcher(answers: dict[str, list[int | Exception]]) -> tuple[cl.Fetch, dict[str, int]]:
    """A fetcher that returns each URL's answers in order, and counts the calls."""
    calls: dict[str, int] = {}

    def fetch(url: str) -> int:
        calls[url] = calls.get(url, 0) + 1
        answer = answers[url][min(calls[url], len(answers[url])) - 1]
        if isinstance(answer, Exception):
            raise answer
        return answer

    return fetch, calls


def no_sleep(_: float) -> None:
    pass


class ClassificationTests(unittest.TestCase):
    def check(self, *answers: int | Exception) -> tuple[cl.Result, int]:
        fetch, calls = fetcher({"https://example.org/": list(answers)})
        return cl.check_url("https://example.org/", fetch, no_sleep), calls["https://example.org/"]

    def test_2xx_is_verified(self) -> None:
        for status in (200, 204, 299):
            with self.subTest(status=status):
                self.assertEqual(self.check(status)[0].outcome, cl.VERIFIED)

    def test_other_statuses_are_broken_at_once(self) -> None:
        for status in (301, 400, 404, 410, 451):
            with self.subTest(status=status):
                result, calls = self.check(status)
                self.assertEqual((result.outcome, result.detail, calls),
                                 (cl.BROKEN, f"HTTP {status}", 1))

    def test_403_is_not_verifiable_and_not_retried(self) -> None:
        result, calls = self.check(403)
        self.assertEqual((result.outcome, calls), (cl.NOT_VERIFIABLE, 1))

    def test_429_and_5xx_are_retried_then_not_verifiable(self) -> None:
        for status in (429, 500, 503, 599):
            with self.subTest(status=status):
                result, calls = self.check(status)
                self.assertEqual(result.outcome, cl.NOT_VERIFIABLE)
                self.assertEqual(calls, cl.RETRIES + 1)
                self.assertIn(f"after {cl.RETRIES + 1} attempts", result.detail)

    def test_network_errors_are_not_verifiable(self) -> None:
        for error in (urllib.error.URLError("refused"), TimeoutError("timed out"),
                      socket.gaierror("no such host"), ConnectionResetError("reset")):
            with self.subTest(error=type(error).__name__):
                result, calls = self.check(error)
                self.assertEqual((result.outcome, calls), (cl.NOT_VERIFIABLE, cl.RETRIES + 1))
                self.assertIn(type(error).__name__, result.detail)

    def test_a_retry_can_verify_or_break(self) -> None:
        self.assertEqual(self.check(503, 200), (cl.Result("https://example.org/", cl.VERIFIED,
                                                          "HTTP 200"), 2))
        self.assertEqual(self.check(TimeoutError("slow"), 404)[0].outcome, cl.BROKEN)

    def test_backoff_doubles(self) -> None:
        waits: list[float] = []
        fetch, _ = fetcher({"https://example.org/": [500]})
        cl.check_url("https://example.org/", fetch, waits.append)
        self.assertEqual(waits, [cl.BACKOFF * 2 ** i for i in range(cl.RETRIES)])

    def test_limits_are_small_and_bounded(self) -> None:
        self.assertTrue(0 < cl.TIMEOUT <= 30 and 0 < cl.RETRIES <= 3 and 0 < cl.WORKERS <= 8)
        self.assertEqual(cl.LimitedRedirects.max_redirections, cl.MAX_REDIRECTS)


class UrlTests(unittest.TestCase):
    def test_links_images_and_html_are_collected_without_fragments(self) -> None:
        documents = {
            "b.md": "[x](https://example.org/a#part) ![i](http://example.org/img.png)\n",
            "a.md": ('<a href="https://example.org/a">same</a> [mail](mailto:x@example.org) '
                     "[local](docs/x.md) `https://example.org/code`\n"),
        }
        self.assertEqual(cl.external_urls(documents, REPO), {
            "https://example.org/a": ["a.md", "b.md"],
            "http://example.org/img.png": ["b.md"],
        })

    def test_this_repository_is_skipped(self) -> None:
        documents = {"a.md": (
            "[1](https://github.com/owner/project) [2](https://github.com/Owner/Project/issues/3) "
            "[3](https://github.com/owner/project-other) [4](https://github.com/owner) "
            "[5](https://example.org/owner/project)\n")}
        self.assertEqual(sorted(cl.external_urls(documents, REPO)), [
            "https://example.org/owner/project", "https://github.com/owner",
            "https://github.com/owner/project-other"])

    def test_the_repository_comes_from_argument_environment_or_origin(self) -> None:
        def remote(url: str) -> Callable[[], str]:
            return lambda: url
        environment = {"GITHUB_REPOSITORY": "env/repo"}
        self.assertEqual(cl.repository_from("arg/repo", environment, remote("")), "arg/repo")
        self.assertEqual(cl.repository_from(None, environment, remote("")), "env/repo")
        for url in ("https://github.com/o/r.git\n", "git@github.com:o/r.git",
                    "http://proxy@127.0.0.1:1234/git/o/r", "https://github.com/o/r/"):
            with self.subTest(url=url):
                self.assertEqual(cl.repository_from(None, {}, remote(url)), "o/r")
        with self.assertRaises(ValueError):
            cl.repository_from(None, {}, remote("not a url"))


class ReportTests(unittest.TestCase):
    def run_main(self, answers: dict[str, list[int | Exception]]) -> tuple[int, str, str]:
        documents = {"README.md": " ".join(f"[l]({url})" for url in answers) + "\n"}
        fetch, _ = fetcher(answers)
        with tempfile.TemporaryDirectory() as directory, \
                mock.patch.object(cl, "committed_files", return_value=set(documents)), \
                mock.patch.object(cl, "blob", side_effect=lambda _, path: documents[path]), \
                mock.patch("sys.stdout") as stdout:
            summary = Path(directory) / "summary.md"
            code = cl.main(["--repository", REPO, "--summary", str(summary)], fetch, no_sleep)
            printed = "".join(call.args[0] for call in stdout.write.call_args_list)
            return code, printed, summary.read_text(encoding="utf-8")

    def test_all_verified_passes(self) -> None:
        code, printed, summary = self.run_main({"https://example.org/": [200]})
        self.assertEqual(code, 0)
        self.assertIn("external-links: PASS (1 links: 1 verified, 0 not verifiable, 0 broken)",
                      printed)
        self.assertNotIn("| Outcome |", summary)

    def test_not_verifiable_is_a_warning_never_a_pass(self) -> None:
        code, printed, summary = self.run_main({"https://example.org/": [200],
                                                "https://example.net/": [429]})
        self.assertEqual(code, 0)
        self.assertIn("external-links: WARN", printed)
        self.assertNotIn("PASS", printed)
        self.assertIn("::warning::not verifiable: https://example.net/", printed)
        self.assertIn("| not verifiable | <https://example.net/> |", summary)

    def test_broken_fails(self) -> None:
        code, printed, summary = self.run_main({"https://example.org/gone": [404],
                                                "https://example.net/": [503]})
        self.assertEqual(code, 1)
        self.assertIn("external-links: FAIL", printed)
        self.assertIn("::error::broken: https://example.org/gone (HTTP 404) in README.md", printed)
        rows = [line for line in summary.splitlines() if line.startswith("| broken")
                or line.startswith("| not verifiable")]
        self.assertTrue(rows[0].startswith("| broken |"))

    def test_cells_cannot_break_the_table(self) -> None:
        self.assertEqual(cl.cell("a|b\\c\nd"), "a\\|b\\\\c d")


if __name__ == "__main__":
    unittest.main()
