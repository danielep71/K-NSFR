#!/usr/bin/env python3
"""Check the external links of the tracked Markdown files over the network.

This Source Code Form is subject to the terms of the Mozilla Public
License, v. 2.0. If a copy of the MPL was not distributed with this
file, You can obtain one at https://mozilla.org/MPL/2.0/.

Run weekly and on demand by the External links workflow, never by Repository
integrity, which stays offline (tools/README.md, external links). Each http or
https URL found in a link, image or HTML href/src of a committed Markdown file is fetched
once, with GET and redirects followed. Its outcome is:

- verified: a 2xx answer after the redirects;
- not verifiable: 403, 429, a 5xx answer, a timeout or another network error;
  all but 403 are retried RETRIES times; a warning, never a pass;
- broken: any other answer outside 2xx.

Absolute URLs to this repository are skipped: their targets are checked
offline by the content gate, or are not files at all. Exit status: 1 when a
link is broken, otherwise 0; the last line says PASS only when every checked
link is verified.
"""

from __future__ import annotations

import argparse
import http.client
import os
import re
import subprocess
import sys
import time
import urllib.error
import urllib.request
from collections.abc import Callable, Iterable, Mapping, Sequence
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from urllib.parse import urlsplit, urlunsplit

from check_content import blob, committed_files, navigation

TIMEOUT = 15  # seconds, for the connection and for each read of the answer
RETRIES = 2  # further attempts after a 429, a 5xx or a network error
BACKOFF = 5  # seconds before the first retry, doubled for each further one
WORKERS = 4  # URLs fetched at the same time
MAX_REDIRECTS = 10
USER_AGENT = "K-NSFR-link-check (+https://github.com/danielep71/K-NSFR)"

VERIFIED = "verified"
NOT_VERIFIABLE = "not verifiable"
BROKEN = "broken"

GITHUB_HOSTS = {"github.com", "www.github.com"}
REPOSITORY = re.compile(r"(?:[:/])([A-Za-z0-9_.-]+)/([A-Za-z0-9_.-]+?)(?:\.git)?/?$")


@dataclass(frozen=True)
class Result:
    url: str
    outcome: str
    detail: str


Fetch = Callable[[str], int]


def classify(status: int) -> str:
    if 200 <= status < 300:
        return VERIFIED
    if status in (403, 429) or status >= 500:
        return NOT_VERIFIABLE
    return BROKEN


def repository_from(argument: str | None, environment: Mapping[str, str],
                    remote: Callable[[], str]) -> str:
    """owner/name from the argument, GITHUB_REPOSITORY or the origin remote."""
    if argument:
        return argument
    if environment.get("GITHUB_REPOSITORY"):
        return environment["GITHUB_REPOSITORY"]
    match = REPOSITORY.search(remote().strip())
    if not match:
        raise ValueError("cannot identify this repository: pass --repository owner/name")
    return f"{match.group(1)}/{match.group(2)}"


def origin_url() -> str:
    return subprocess.check_output(["git", "remote", "get-url", "origin"], text=True)


def is_this_repository(url: str, repository: str) -> bool:
    parts = urlsplit(url)
    if parts.hostname not in GITHUB_HOSTS:
        return False
    path = parts.path.rstrip("/").lower()
    root = "/" + repository.lower()
    return path == root or path.startswith(root + "/")


def external_urls(documents: Mapping[str, str], repository: str) -> dict[str, list[str]]:
    """Each external URL, without its fragment, and the files that link to it."""
    found: dict[str, list[str]] = {}
    for path, text in sorted(documents.items()):
        for target in navigation(text)[1]:
            parts = urlsplit(target.strip())
            if parts.scheme.lower() not in ("http", "https") or not parts.netloc:
                continue
            url = urlunsplit((parts.scheme.lower(), parts.netloc, parts.path, parts.query, ""))
            if is_this_repository(url, repository):
                continue
            files = found.setdefault(url, [])
            if path not in files:
                files.append(path)
    return found


class LimitedRedirects(urllib.request.HTTPRedirectHandler):
    max_redirections = MAX_REDIRECTS


def http_status(url: str) -> int:
    """The final status of a GET that follows redirects; raises on network errors."""
    opener = urllib.request.build_opener(LimitedRedirects())
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT, "Accept": "*/*"})
    try:
        with opener.open(request, timeout=TIMEOUT) as response:
            response.read(1024)
            return int(response.status)
    except urllib.error.HTTPError as error:
        return int(error.code)


def check_url(url: str, fetch: Fetch = http_status,
              sleep: Callable[[float], None] = time.sleep) -> Result:
    detail = ""
    for attempt in range(RETRIES + 1):
        if attempt:
            sleep(BACKOFF * 2 ** (attempt - 1))
        try:
            status = fetch(url)
        except (OSError, ValueError, http.client.HTTPException) as error:
            # URLError, timeouts, refused or reset connections and malformed answers
            status, detail = 0, f"{type(error).__name__}: {error}"
        else:
            detail = f"HTTP {status}"
        outcome = classify(status) if status else NOT_VERIFIABLE
        if outcome != NOT_VERIFIABLE or status == 403:
            # A refusal is not retried: it rarely changes within seconds.
            return Result(url, outcome, detail)
    return Result(url, NOT_VERIFIABLE, f"{detail} after {RETRIES + 1} attempts")


def check_all(urls: Iterable[str], fetch: Fetch = http_status,
              sleep: Callable[[float], None] = time.sleep) -> list[Result]:
    with ThreadPoolExecutor(max_workers=WORKERS) as pool:
        return list(pool.map(lambda url: check_url(url, fetch, sleep), sorted(urls)))


def cell(value: str) -> str:
    return " ".join(value.split()).replace("\\", "\\\\").replace("|", "\\|")


def report(results: Sequence[Result], files: Mapping[str, list[str]]) -> tuple[str, list[str]]:
    """The status line and a Markdown table of every link that is not verified."""
    counts = {outcome: sum(r.outcome == outcome for r in results)
              for outcome in (VERIFIED, NOT_VERIFIABLE, BROKEN)}
    status = "FAIL" if counts[BROKEN] else "WARN" if counts[NOT_VERIFIABLE] else "PASS"
    line = (f"external-links: {status} ({len(results)} links: {counts[VERIFIED]} verified, "
            f"{counts[NOT_VERIFIABLE]} not verifiable, {counts[BROKEN]} broken)")
    rows = [f"| {r.outcome} | <{cell(r.url)}> | {cell(r.detail)} | {cell(', '.join(files[r.url]))} |"
            for r in sorted(results, key=lambda r: (r.outcome != BROKEN, r.url))
            if r.outcome != VERIFIED]
    table = ["| Outcome | URL | Detail | Linked from |", "| --- | --- | --- | --- |", *rows]
    return line, table if rows else []


def main(argv: Sequence[str] | None = None, fetch: Fetch = http_status,
         sleep: Callable[[float], None] = time.sleep) -> int:
    parser = argparse.ArgumentParser(
        description="Check the external links of the tracked Markdown files.")
    parser.add_argument("--repository", help="owner/name of this repository, to skip its URLs")
    parser.add_argument("--ref", default="HEAD", help="the commit whose Markdown is read")
    parser.add_argument("--summary", help="append the Markdown report to this file")
    args = parser.parse_args(argv)
    repository = repository_from(args.repository, os.environ, origin_url)
    documents = {path: blob(args.ref, path) for path in committed_files(args.ref)
                 if path.endswith(".md")}
    files = external_urls(documents, repository)
    results = check_all(files, fetch, sleep)
    line, table = report(results, files)
    for result in results:
        if result.outcome != VERIFIED:
            level = "error" if result.outcome == BROKEN else "warning"
            print(f"::{level}::{result.outcome}: {result.url} ({result.detail}) "
                  f"in {', '.join(files[result.url])}")
    print(line)
    if args.summary:
        with open(args.summary, "a", encoding="utf-8") as summary:
            summary.write("\n".join(["## External links", "", line, "", *table, ""]) + "\n")
    return 1 if line.startswith("external-links: FAIL") else 0


if __name__ == "__main__":
    sys.exit(main())
