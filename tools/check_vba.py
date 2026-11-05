#!/usr/bin/env python3
"""Check exported VBA source against the repository's static rules.

The rules come from the implementation conventions, the VBA house style and
ADR-008. The check reads source text only: it never compiles or runs VBA,
and docs/IMPLEMENTATION.md and tools/README.md state its limits.

This Source Code Form is subject to the terms of the Mozilla Public
License, v. 2.0. If a copy of the MPL was not distributed with this
file, You can obtain one at https://mozilla.org/MPL/2.0/.
"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from collections.abc import Iterable, Sequence
from dataclasses import dataclass, field
from pathlib import Path, PurePosixPath
from typing import Any

EXTENSIONS = (".bas", ".cls", ".frm")
EXCEPTIONS_FILE = "tools/vba-exceptions.json"
LAYERS = (("src/core/", "core"), ("src/api/", "api"), ("src/host/", "host"),
          ("tests/", "test"), ("examples/", "example"))
# Rules a reviewed exception may waive, and only in src/host/.
EXCEPTIONABLE = frozenset({"declare", "com-object", "dynamic-call"})
COMPILE_CONSTANTS = frozenset({"vba7", "win64"})
CONDITION_WORDS = frozenset({"not", "and", "or", "then"})
CREATABLE_BUILTINS = frozenset({"collection"})
QUALIFIERS = frozenset({"vba", "interaction"})
TOKEN = re.compile(r'"(?:[^"]|"")*"?|[A-Za-z_][A-Za-z0-9_]*[$%&!#@]?|\d+(?:\.\d+)?|:=|\S')
CONTINUATION = re.compile(r"(?:^|\s)_$")
LABEL = re.compile(r"^\s*(?:[A-Za-z_][A-Za-z0-9_]*|\d+)\s*$")
LINE_NUMBER = re.compile(r"^\s*\d+\s+")
REM = re.compile(r"^\s*(?:\d+\s+)?Rem\b", re.I)
# Keywords that form a statement on their own, so `Keyword:` is never a label.
STATEMENT_WORDS = frozenset({"rem", "return", "resume", "stop", "end", "else", "do", "loop",
                             "next", "wend", "exit"})
HEADER_OPEN = re.compile(r"^\s*BEGIN(?:PROPERTY)?\b", re.I)
HEADER_CLOSE = re.compile(r"^\s*END(?:PROPERTY)?\s*$", re.I)
VB_NAME = re.compile(r'^\s*Attribute\s+VB_Name\s*=\s*"([^"]*)"', re.I)
DIRECTIVE = re.compile(r"^\s*#\s*(\w+)\b(.*)$")
PROCEDURE = re.compile(
    r"^\s*(?:(?:Public|Private|Friend|Static|Global)\s+)*"
    r"(?:Sub|Function|Property\s+(?:Get|Let|Set)|Declare\s+(?:PtrSafe\s+)?(?:Sub|Function))"
    r"\s+([A-Za-z_][A-Za-z0-9_]*)", re.I)
# A name right after these words is being declared, not called.
DECLARING_WORDS = frozenset({"sub", "function", "get", "let", "event"})


@dataclass(frozen=True)
class Finding:
    path: str
    line: int
    rule: str
    message: str
    text: str = ""

    def render(self) -> str:
        return f"{self.path}:{self.line}: [{self.rule}] {self.message}"


@dataclass(frozen=True)
class Waiver:
    path: str
    rule: str
    contains: str
    reason: str
    issue: int


@dataclass
class Module:
    path: str
    layer: str
    classes: frozenset[str]
    procedures: frozenset[str] = frozenset()
    findings: list[Finding] = field(default_factory=list)
    in_block: bool = False

    def add(self, line: int, rule: str, message: str, text: str = "") -> None:
        self.findings.append(Finding(self.path, line, rule, message, text.strip()))


def layer_of(path: str) -> str:
    return next((layer for prefix, layer in LAYERS if path.startswith(prefix)), "unknown")


def strip_header(lines: list[str]) -> int:
    """Index of the first line after a .cls/.frm VERSION/BEGIN...END header."""
    first = next((i for i, line in enumerate(lines) if line.strip()), None)
    if first is None or not lines[first].lstrip().upper().startswith("VERSION"):
        return 0
    depth, opened = 0, False
    for index in range(first + 1, len(lines)):
        if HEADER_OPEN.match(lines[index]):
            depth, opened = depth + 1, True
        elif HEADER_CLOSE.match(lines[index]):
            depth -= 1
        if opened and depth == 0:
            return index + 1
    return len(lines)


def logical_lines(lines: list[str], start: int) -> Iterable[tuple[int, str]]:
    """Join continued physical lines; yield (first line number, text)."""
    index = start
    while index < len(lines):
        number, parts = index + 1, [lines[index]]
        while CONTINUATION.search(parts[-1].rstrip()) and index + 1 < len(lines):
            parts[-1] = parts[-1].rstrip()[:-1]
            index += 1
            parts.append(lines[index])
        index += 1
        yield number, " ".join(parts)


def strip_comment(code: str) -> str:
    in_string = False
    for position, char in enumerate(code):
        if char == '"':
            in_string = not in_string
        elif char == "'" and not in_string:
            return code[:position]
    return code


def split_statements(code: str) -> list[str]:
    """Split on colons outside strings and named arguments; a Rem ends the line."""
    if REM.match(code):
        return []
    parts: list[str] = []
    current: list[str] = []
    in_string = False
    for position, char in enumerate(code):
        if char == '"':
            in_string = not in_string
        if char == ":" and not in_string and code[position + 1:position + 2] != "=":
            parts.append("".join(current))
            current = []
            if REM.match(code[position + 1:]):
                return parts
            continue
        current.append(char)
    parts.append("".join(current))
    return parts


def statements(code: str) -> list[str]:
    """The statements of a logical line, without a leading label or line numbers."""
    parts = split_statements(code)
    first = parts[0].strip().lower() if parts else ""
    if len(parts) > 1 and LABEL.match(first) and first not in STATEMENT_WORDS:
        parts = parts[1:]
    return [part for part in (LINE_NUMBER.sub("", part) for part in parts) if part.strip()]


def tokens_of(statement: str) -> list[str]:
    return TOKEN.findall(statement)


def qualifier(tokens: list[str], index: int) -> str | None:
    """The object name before `.name`, '' for a bare `.name`, None when unqualified."""
    if index > 0 and tokens[index - 1] == ".":
        return tokens[index - 2].lower() if index > 1 else ""
    return None


def is_builtin_call(tokens: list[str], index: int, procedures: frozenset[str],
                    extra: frozenset[str] = frozenset()) -> bool:
    """Whether a name refers to the VBA or Excel function rather than a declaration."""
    owner = qualifier(tokens, index)
    if owner is not None:
        return owner in QUALIFIERS | extra
    previous = tokens[index - 1].lower() if index > 0 else ""
    declared = (tokens[index + 1:index + 2] and tokens[index + 1].lower() == "as") \
        or previous in DECLARING_WORDS \
        or (previous == "set" and index > 1 and tokens[index - 2].lower() == "property")
    return not declared and tokens[index].lower() not in procedures


def jump_findings(tokens: list[str], layer: str) -> list[tuple[str, str]]:
    words = [token.lower() for token in tokens]
    found: list[tuple[str, str]] = []
    handled: set[int] = set()
    for index, word in enumerate(words):
        if qualifier(tokens, index) is not None or index in handled:
            continue
        if word == "on" and words[index + 1:index + 2] == ["error"]:
            handled.update({index + 1, index + 2})
            action = words[index + 2] if index + 2 < len(words) else ""
            if action == "resume":
                found.append(("on-error-resume-next", "On Error Resume Next is not allowed"))
            elif layer in ("core", "api", "unknown"):
                found.append(("error-handler", "On Error appears only at the host boundary (ADR-008)"))
        elif word in ("gosub", "return"):
            found.append(("gosub", "GoSub and Return are not allowed"))
        elif word == "goto":
            found.append(("goto", "GoTo is allowed only as On Error GoTo"))
        elif word == "resume" and layer in ("core", "api", "unknown"):
            found.append(("error-handler", "Resume appears only at the host boundary (ADR-008)"))
    return found


def declare_findings(tokens: list[str], layer: str) -> list[tuple[str, str]]:
    words = [token.lower() for token in tokens[:3]]
    if "declare" not in words or words.index("declare") > 1:
        return []
    position = words.index("declare")
    found = []
    if tokens[position + 1:position + 2] and tokens[position + 1].lower() != "ptrsafe":
        found.append(("declare-ptrsafe", "Declare must use PtrSafe (64-bit Excel)"))
    where = ("needs a reviewed exception in " + EXCEPTIONS_FILE) if layer == "host" \
        else "is allowed only in src/host/"
    found.append(("declare", f"Declare {where}"))
    return found


def new_findings(tokens: list[str], index: int, classes: frozenset[str]) -> list[str]:
    name = tokens[index + 1] if index + 1 < len(tokens) else ""
    if tokens[index + 2:index + 3] == ["."]:
        member = tokens[index + 3] if index + 3 < len(tokens) else ""
        return [] if name.lower() == "vba" else [f"New {name}.{member} creates an external object"]
    if name.lower() not in classes | CREATABLE_BUILTINS:
        return [f"New {name}: not a class of this project"]
    return []


def com_findings(tokens: list[str], module: Module) -> list[tuple[str, str]]:
    layer, classes = module.layer, module.classes
    messages = []
    for index, token in enumerate(tokens):
        word = token.lower()
        if word in ("createobject", "getobject") and is_builtin_call(tokens, index, module.procedures):
            messages.append(f"{token} creates or binds an external object")
        elif word == "new" and qualifier(tokens, index) is None:
            messages += new_findings(tokens, index, classes)
    suffix = f"; it needs a reviewed exception in {EXCEPTIONS_FILE}" if layer == "host" \
        else "; external objects are allowed only in src/host/"
    return [("com-object", message + suffix) for message in messages]


def dynamic_findings(tokens: list[str], module: Module) -> list[tuple[str, str]]:
    layer = module.layer
    messages = []
    for index, token in enumerate(tokens):
        word = token.lower()
        if word in ("callbyname", "evaluate", "executeexcel4macro") \
                and is_builtin_call(tokens, index, module.procedures, frozenset({"application"})):
            messages.append(f"{token} builds a call or expression at run time")
        elif word == "run" and qualifier(tokens, index) == "application":
            messages.append("Application.Run calls a procedure by name")
    suffix = f"; it needs a reviewed exception in {EXCEPTIONS_FILE}" if layer == "host" \
        else "; dynamic calls are allowed only in src/host/"
    return [("dynamic-call", message + suffix) for message in messages]


def block_change(tokens: list[str], in_block: bool) -> tuple[bool, bool]:
    """(skip this statement, inside a Type/Enum block afterwards)."""
    words = [token.lower() for token in tokens]
    if in_block:
        return True, words[:2] not in (["end", "type"], ["end", "enum"])
    start = 1 if words[:1] in (["public"], ["private"], ["friend"]) else 0
    opens = words[start:start + 1] in (["type"], ["enum"]) and len(words) > start + 1
    return False, opens


def check_statement(module: Module, line: int, statement: str) -> None:
    tokens = tokens_of(statement)
    skip, module.in_block = block_change(tokens, module.in_block)
    if skip or not tokens:
        return
    for rule, message in (jump_findings(tokens, module.layer)
                          + declare_findings(tokens, module.layer)
                          + com_findings(tokens, module)
                          + dynamic_findings(tokens, module)):
        module.add(line, rule, message, statement)


def check_directive(module: Module, line: int, text: str) -> None:
    match = DIRECTIVE.match(strip_comment(text))
    if not match:
        return
    kind, condition = match.group(1).lower(), match.group(2)
    if kind == "const":
        module.add(line, "compile-constant", "#Const is not allowed; test only VBA7 and Win64", text)
    elif kind in ("if", "elseif"):
        if module.layer != "host":
            module.add(line, "conditional-compilation",
                       "Conditional compilation is allowed only in src/host/", text)
            return
        names = {word.lower() for word in re.findall(r"[A-Za-z_]\w*", condition)}
        if names - CONDITION_WORDS - COMPILE_CONSTANTS:
            module.add(line, "conditional-compilation",
                       "Conditional compilation may test only VBA7 and Win64", text)


def ascii_findings(module: Module, raw: bytes) -> None:
    for number, line in enumerate(raw.splitlines(), start=1):
        column = next((index for index, byte in enumerate(line) if byte > 0x7F), None)
        if column is not None:
            module.add(number, "ascii", f"Non-ASCII byte at column {column + 1}; VBA source is ASCII only (ADR-008)")


def module_findings(module: Module, options: set[str], name: str | None) -> None:
    stem = PurePosixPath(module.path).stem
    if "explicit" not in options:
        module.add(1, "option-explicit", "Option Explicit is missing")
    private_required = module.path.endswith(".bas") and module.layer in ("core", "test")
    if private_required and "private module" not in options:
        module.add(1, "option-private-module", "Option Private Module is missing")
    if name is None:
        module.add(1, "vb-name", "Attribute VB_Name is missing")
    elif name != stem:
        module.add(1, "vb-name", f'Attribute VB_Name "{name}" differs from the file name "{stem}"')


def source_statements(raw: bytes) -> Iterable[tuple[int, str]]:
    """(line, statement) for the code of a module, after its header and attributes."""
    lines = raw.decode("ascii", errors="replace").splitlines()
    for line, text in logical_lines(lines, strip_header(lines)):
        if re.match(r"^\s*(?:Attribute\s|#)", text, re.I):
            continue
        for statement in statements(strip_comment(text)):
            yield line, statement


def declared_procedures(raw: bytes) -> frozenset[str]:
    """Lower-case names of the procedures a module declares."""
    names = (PROCEDURE.match(statement) for _, statement in source_statements(raw))
    return frozenset(match.group(1).lower() for match in names if match)


def check_source(path: str, raw: bytes, classes: frozenset[str] = frozenset(),
                 procedures: frozenset[str] = frozenset()) -> list[Finding]:
    """All findings for one exported module, before exceptions are applied."""
    module = Module(path, layer_of(path), classes, procedures | declared_procedures(raw))
    if module.layer == "unknown":
        module.add(1, "vba-location", "VBA source belongs in src/core, src/api, src/host, tests or examples")
    ascii_findings(module, raw)
    lines = raw.decode("ascii", errors="replace").splitlines()
    start = strip_header(lines)
    options: set[str] = set()
    name = None
    for line, text in logical_lines(lines, start):
        match = VB_NAME.match(text)
        if match:
            name = match.group(1)
            continue
        if re.match(r"^\s*Attribute\s", text, re.I):
            continue
        if text.lstrip().startswith("#"):
            check_directive(module, line, text)
            continue
        for statement in statements(strip_comment(text)):
            option = re.match(r"^\s*Option\s+(Explicit|Private\s+Module)\b", statement, re.I)
            if option:
                options.add(" ".join(option.group(1).lower().split()))
            check_statement(module, line, statement)
    module_findings(module, options, name)
    return module.findings


def load_exceptions(raw: str) -> tuple[list[Waiver], list[Finding]]:
    problems: list[Finding] = []
    try:
        data: Any = json.loads(raw)
    except json.JSONDecodeError as error:
        return [], [Finding(EXCEPTIONS_FILE, 1, "exception-invalid", f"invalid JSON: {error}")]
    entries = data.get("exceptions") if isinstance(data, dict) else None
    if not isinstance(data, dict) or data.get("schema_version") != 1 or not isinstance(entries, list):
        return [], [Finding(EXCEPTIONS_FILE, 1, "exception-invalid",
                            "expected {\"schema_version\": 1, \"exceptions\": [...]}")]
    result = []
    for number, entry in enumerate(entries, start=1):
        problem = exception_error(entry)
        if problem:
            problems.append(Finding(EXCEPTIONS_FILE, 1, "exception-invalid",
                                    f"exception {number}: {problem}"))
        else:
            result.append(Waiver(entry["path"], entry["rule"], entry["contains"],
                                     entry["reason"], entry["issue"]))
    return result, problems


def exception_error(entry: Any) -> str | None:
    if not isinstance(entry, dict) or set(entry) != {"path", "rule", "contains", "reason", "issue"}:
        return "needs exactly path, rule, contains, reason and issue"
    if not isinstance(entry["path"], str) or not entry["path"].startswith("src/host/"):
        return "path must be under src/host/ (a host boundary)"
    if entry["rule"] not in EXCEPTIONABLE:
        return f"rule must be one of {sorted(EXCEPTIONABLE)}"
    if not all(isinstance(entry[key], str) and entry[key].strip() for key in ("contains", "reason")):
        return "contains and reason must be non-empty text"
    if not isinstance(entry["issue"], int) or isinstance(entry["issue"], bool) or entry["issue"] < 1:
        return "issue must be an issue number"
    return None


def apply_exceptions(findings: list[Finding], exceptions: list[Waiver]) -> list[Finding]:
    used: set[int] = set()
    kept = []
    for finding in findings:
        match = next((index for index, exception in enumerate(exceptions)
                      if exception.path == finding.path and exception.rule == finding.rule
                      and exception.contains in finding.text), None)
        if match is None:
            kept.append(finding)
        else:
            used.add(match)
    kept += [Finding(EXCEPTIONS_FILE, 1, "exception-unused",
                     f"exception {index + 1} ({exception.path}, {exception.rule}) matches no finding")
             for index, exception in enumerate(exceptions) if index not in used]
    return kept


def tracked_sources(root: Path) -> list[str]:
    output = subprocess.run(["git", "ls-files", "-z", "--", *(f"*{ext}" for ext in EXTENSIONS)],
                            cwd=root, capture_output=True, check=True).stdout
    return sorted(name for name in output.decode().split("\0") if name)


def check_files(root: Path, paths: Sequence[str], exceptions_raw: str) -> list[Finding]:
    classes = frozenset(PurePosixPath(path).stem.lower() for path in paths
                        if path.lower().endswith((".cls", ".frm")))
    sources = {path: (root / path).read_bytes() for path in paths}
    procedures = frozenset().union(*(declared_procedures(raw) for raw in sources.values()))
    findings: list[Finding] = []
    for path, raw in sources.items():
        findings += check_source(path, raw, classes, procedures)
    exceptions, problems = load_exceptions(exceptions_raw)
    return problems + apply_exceptions(findings, exceptions)


def parse_args(argv: Sequence[str] | None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0] if __doc__ else None)
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("paths", nargs="*",
                        help="repository-relative files; default: every tracked .bas, .cls and .frm")
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    root = args.root.resolve()
    paths = [PurePosixPath(path).as_posix() for path in args.paths] or tracked_sources(root)
    exceptions_path = root / EXCEPTIONS_FILE
    if not exceptions_path.is_file():
        print(f"check_vba: {EXCEPTIONS_FILE} is missing", file=sys.stderr)
        return 2
    findings = check_files(root, paths, exceptions_path.read_text(encoding="utf-8"))
    for finding in findings:
        print(finding.render())
    print(f"check_vba: {len(paths)} VBA file(s), {len(findings)} finding(s)")
    return 1 if findings else 0


if __name__ == "__main__":
    raise SystemExit(main())
