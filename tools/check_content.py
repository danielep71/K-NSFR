"""Check committed Markdown links/fences and YAML syntax without a worktree dependency.

This Source Code Form is subject to the terms of the Mozilla Public
License, v. 2.0. If a copy of the MPL was not distributed with this
file, You can obtain one at https://mozilla.org/MPL/2.0/.
"""

import re
import subprocess
import sys
from pathlib import PurePosixPath
from urllib.parse import unquote

import yaml


def committed_files(ref="HEAD"):
    names = subprocess.check_output(["git", "ls-tree", "-rz", "--name-only", ref])
    return {name.decode("utf-8") for name in names.split(b"\0") if name}


def blob(ref, path):
    return subprocess.check_output(["git", "show", f"{ref}:{path}"]).decode("utf-8")


class UniqueKeysLoader(yaml.SafeLoader):
    pass


def unique_mapping(loader, node):
    result = {}
    for key_node, value_node in node.value:
        key = loader.construct_object(key_node)
        if key in result:
            raise yaml.constructor.ConstructorError(None, None, f"duplicate key: {key}", key_node.start_mark)
        result[key] = loader.construct_object(value_node)
    return result


UniqueKeysLoader.add_constructor(yaml.resolver.BaseResolver.DEFAULT_MAPPING_TAG, unique_mapping)


def markdown_errors(path, text, names):
    errors = []
    fence = None
    for number, line in enumerate(text.splitlines(), 1):
        match = re.match(r"^\s{0,3}(`{3,}|~{3,})", line)
        if match:
            token = match.group(1)
            if fence is None:
                fence = token
            elif token[0] == fence[0] and len(token) >= len(fence):
                fence = None
            continue
        if fence:
            continue
        for target in re.findall(r"(?<!!)\[[^\]]+\]\(([^)]+)\)", line):
            target = target.split("#", 1)[0]
            if not target or "://" in target or target.startswith(("mailto:", "#")):
                continue
            target = unquote(target).split(" ", 1)[0]
            parent = PurePosixPath(path).parent
            parts = []
            for part in (str(parent / target)).split("/"):
                if part == "..":
                    if parts:
                        parts.pop()
                elif part not in ("", "."):
                    parts.append(part)
            resolved = "/".join(parts)
            if resolved not in names and not any(n.startswith(resolved.rstrip("/") + "/") for n in names):
                errors.append(f"{path}:{number}: missing link target {target}")
    if fence:
        errors.append(f"{path}: unclosed code fence")
    return errors


def check(ref="HEAD"):
    names = committed_files(ref)
    errors = []
    for path in sorted(names):
        if path.endswith(".md"):
            errors.extend(markdown_errors(path, blob(ref, path), names))
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
