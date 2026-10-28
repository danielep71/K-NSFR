"""Validate committed Markdown navigation and YAML; never fetch remote links.

This Source Code Form is subject to the terms of the Mozilla Public
License, v. 2.0. If a copy of the MPL was not distributed with this
file, You can obtain one at https://mozilla.org/MPL/2.0/.
"""

from __future__ import annotations

import posixpath
import re
import subprocess
import sys
from collections.abc import Mapping
from html.parser import HTMLParser
from pathlib import PurePosixPath
from typing import Any
from urllib.parse import unquote, urlsplit

import yaml
from markdown_it import MarkdownIt
from markdown_it.token import Token


class UniqueKeysLoader(yaml.SafeLoader):
    pass


def unique_mapping(loader: UniqueKeysLoader, node: yaml.MappingNode) -> dict[Any, Any]:
    result: dict[Any, Any] = {}
    for key_node, value_node in node.value:
        key = loader.construct_object(key_node)
        if key in result:
            raise yaml.constructor.ConstructorError(None, None, f"duplicate key: {key}", key_node.start_mark)
        result[key] = loader.construct_object(value_node)
    return result


UniqueKeysLoader.add_constructor(yaml.resolver.BaseResolver.DEFAULT_MAPPING_TAG, unique_mapping)


class HtmlNavigation(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.anchors: list[str] = []
        self.links: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        for key, value in attrs:
            if value is None:
                continue
            if key == "id" or (tag == "a" and key == "name"):
                self.anchors.append(value)
            if key == "href" or (tag == "img" and key == "src"):
                self.links.append(value)


def navigation(text: str) -> tuple[set[str], list[str], list[str]]:
    tokens = MarkdownIt("commonmark").parse(text)
    html = HtmlNavigation()
    anchors: set[str] = set()
    links: list[str] = []
    errors: list[str] = []
    for token in tokens:
        inspect_token(token, html, links)
        if token.type == "fence" and token.map:
            closing = text.splitlines()[token.map[1] - 1]
            if token.map[1] - token.map[0] < 2 or not re.fullmatch(r" {0,3}" + re.escape(token.markup[0]) + "{" + str(len(token.markup)) + r",}\s*", closing):
                errors.append("unclosed code fence")
    for anchor in html.anchors:
        if anchor in anchors:
            errors.append(f"duplicate anchor {anchor}")
        anchors.add(anchor)
    return anchors, links + html.links, errors


def inspect_token(token: Token, html: HtmlNavigation, links: list[str]) -> None:
    if token.type in ("html_block", "html_inline"):
        html.feed(token.content)
    if token.type in ("link_open", "image"):
        target = token.attrGet("href" if token.type == "link_open" else "src")
        if target:
            links.append(str(target))
    for child in token.children or []:
        inspect_token(child, html, links)


def target_error(path: str, target: str, names: set[str], anchors: Mapping[str, set[str]]) -> str | None:
    url = urlsplit(target)
    if url.scheme or url.netloc:
        return None
    raw = unquote(url.path)
    resolved = posixpath.normpath(posixpath.join(str(PurePosixPath(path).parent), raw)) if raw else path
    if resolved.startswith(("../", "/")) or resolved == "..":
        return f"link escapes repository: {target}"
    if resolved not in names:
        if not url.fragment and any(n.startswith(resolved.rstrip("/") + "/") for n in names):
            return None
        return f"missing link target {target}"
    if url.fragment and resolved.endswith(".md"):
        if unquote(url.fragment) not in anchors.get(resolved, set()):
            return f"missing section target {target}"
    return None


def markdown_errors(path: str, text: str, names: set[str],
                    documents: Mapping[str, str] | None = None) -> list[str]:
    documents = dict(documents or {})
    documents[path] = text
    anchors = {name: navigation(body)[0] for name, body in documents.items()}
    _, links, errors = navigation(text)
    for target in links:
        error = target_error(path, target, names, anchors)
        if error:
            errors.append(error)
    return [f"{path}: {error}" for error in errors]


def committed_files(ref: str = "HEAD") -> set[str]:
    names = subprocess.check_output(["git", "ls-tree", "-rz", "--name-only", ref])
    return {name.decode("utf-8") for name in names.split(b"\0") if name}


def blob(ref: str, path: str) -> str:
    return subprocess.check_output(["git", "show", f"{ref}:{path}"]).decode("utf-8")


def check(ref: str = "HEAD") -> list[str]:
    names = committed_files(ref)
    documents = {path: blob(ref, path) for path in names if path.endswith(".md")}
    errors = []
    for path, text in sorted(documents.items()):
        errors.extend(markdown_errors(path, text, names, documents))
    for path in sorted(names):
        if path.endswith((".yml", ".yaml")):
            try:
                list(yaml.load_all(blob(ref, path), Loader=UniqueKeysLoader))
            except yaml.YAMLError as exc:
                errors.append(f"{path}: invalid YAML: {exc}")
    return errors


if __name__ == "__main__":
    failures = check()
    for failure in failures:
        print(failure, file=sys.stderr)
    print(f"Committed Markdown/YAML check: {len(failures)} findings")
    sys.exit(bool(failures))
