"""Validate committed blobs, independent of checkout line-ending conversion.

This Source Code Form is subject to the terms of the Mozilla Public
License, v. 2.0. If a copy of the MPL was not distributed with this
file, You can obtain one at https://mozilla.org/MPL/2.0/.
"""

import argparse
import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path, PurePosixPath


def git(*args: str, cwd: Path | None = None) -> bytes:
    return subprocess.check_output(["git", *args], cwd=cwd)


def committed_attributes(entries: dict[str, bytes]) -> dict[str, str]:
    """Resolve only committed attributes, never caller config, index or overrides."""
    env = {key: value for key, value in os.environ.items() if not key.startswith("GIT_")}
    env.update(GIT_CONFIG_NOSYSTEM="1", GIT_CONFIG_GLOBAL=os.devnull,
               GIT_ATTR_NOSYSTEM="1")
    with tempfile.TemporaryDirectory(prefix="committed-attributes-") as directory:
        root = Path(directory)
        command = ["git", "-c", f"core.attributesFile={os.devnull}"]
        subprocess.run([*command, "init", "-q", "--template=", str(root)],
                       env=env, check=True, capture_output=True)
        for name, data in entries.items():
            if PurePosixPath(name).name == ".gitattributes":
                target = root / name
                if not target.resolve().is_relative_to(root):
                    raise ValueError("attribute path escapes repository")
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(data)
        result = subprocess.check_output(
            [*command, "check-attr", "-z", "--stdin", "text"],
            cwd=root, env=env, input=b"".join(name.encode() + b"\0" for name in entries),
        ).split(b"\0")
    return {result[i].decode(): result[i + 2].decode()
            for i in range(0, len(result) - 1, 3)}


def check(ref: str = "HEAD", cwd: Path | None = None) -> tuple[str, int, list[str]]:
    commit = git("rev-parse", "--verify", f"{ref}^{{commit}}", cwd=cwd).decode().strip()
    entries = {}
    errors = []
    for record in git("ls-tree", "-rz", commit, cwd=cwd).split(b"\0"):
        if not record:
            continue
        meta, path = record.split(b"\t", 1)
        mode, kind, oid = meta.decode().split()
        name = path.decode("utf-8")
        if kind != "blob" or mode not in ("100644", "100755"):
            errors.append(f"{name}: submodules and symlinks require a policy decision")
            continue
        entries[name] = git("cat-file", "blob", oid, cwd=cwd)
    allowed = json.loads(entries["tools/binary-allowlist.json"])
    if (not isinstance(allowed, list) or
            any(not isinstance(p, str) or not p or
                PurePosixPath(p).is_absolute() or ".." in PurePosixPath(p).parts or
                any(c in p for c in "*?[]\\") for p in allowed)):
        raise ValueError("binary allowlist must contain exact repository-relative paths")
    if len(allowed) != len(set(allowed)):
        errors.append("binary allowlist: duplicate paths")
    text_attrs = committed_attributes(entries)
    for name in allowed:
        if name not in entries:
            errors.append(f"{name}: allowlisted path does not exist")
        elif text_attrs[name] != "unset":
            errors.append(f"{name}: approved binary must have binary attributes")
    for name, data in entries.items():
        if name in allowed:
            continue
        if text_attrs[name] == "unset":
            errors.append(f"{name}: binary format requires an exact allowlist entry")
            continue
        errors.extend(text_errors(name, data))
    return commit, len(entries), errors


def text_errors(name: str, data: bytes) -> list[str]:
    errors: list[str] = []
    host_text = PurePosixPath(name).suffix.lower() in (
        ".bas", ".cls", ".frm", ".bat", ".cmd", ".ps1", ".psm1",
        ".psd1", ".vbs", ".reg", ".ini",
    )
    try:
        # Latin-1 is used only to inspect host-text bytes losslessly; it is
        # not a declaration of the VBE/system code page.
        text = data.decode("latin-1" if host_text else "utf-8")
    except UnicodeDecodeError:
        errors.append(f"{name}: invalid UTF-8 or unapproved binary")
        return errors
    if "\x00" in text or data.startswith(b"\xef\xbb\xbf"):
        errors.append(f"{name}: NUL or UTF-8 BOM is not permitted")
    if "\r" in text:
        errors.append(f"{name}: committed content must use LF")
    if text and not text.endswith("\n"):
        errors.append(f"{name}: missing final newline")
    # Markdown hard breaks and the exact standard license text are retained.
    if (PurePosixPath(name).suffix.lower() != ".md" and name != "LICENSE" and
            any(line.endswith((" ", "\t")) for line in text.split("\n"))):
        errors.append(f"{name}: trailing whitespace in committed content")
    return errors


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--ref", default="HEAD")
    args = parser.parse_args()
    try:
        commit, count, errors = check(args.ref)
    except (subprocess.CalledProcessError, ValueError, KeyError, UnicodeError) as exc:
        print(f"Source-control check could not complete: {exc}", file=sys.stderr)
        return 1
    for error in errors:
        print(error, file=sys.stderr)
    print(f"Checked {count} committed blobs at {commit}: {len(errors)} findings")
    return bool(errors)


if __name__ == "__main__":
    sys.exit(main())
