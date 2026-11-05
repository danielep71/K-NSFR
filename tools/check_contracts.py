#!/usr/bin/env python3
"""Validate the NSFR input contract, its documentation and its synthetic examples.

This is the reference validation of docs/contracts/CONTRACTS.md: it checks the
form of an input, never a regulatory result. It computes reconciliation totals
of the input amounts and the input fingerprint; it applies no factor.

This Source Code Form is subject to the terms of the Mozilla Public
License, v. 2.0. If a copy of the MPL was not distributed with this
file, You can obtain one at https://mozilla.org/MPL/2.0/.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import io
import json
import re
import sys
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, localcontext
from pathlib import Path
from typing import Any

CONTRACTS = Path("docs/contracts")
SCHEMA = CONTRACTS / "nsfr-contract-v1.json"
DOCUMENT = CONTRACTS / "CONTRACTS.md"
EXAMPLES = CONTRACTS / "examples"
IDENTIFIER = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]{0,63}$")
CURRENCY = re.compile(r"^[A-Z]{3}$")
TIMESTAMP = re.compile(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$")
SCALE = Decimal("0.0001")
LIMIT = Decimal(10) ** 15  # an amount has at most 15 integer digits
PRECISION = 60  # exact for any sum of contract amounts
WIDTH = "#cells"  # marks a row whose cell count differs from its header
SIDES = {"own_funds": "asf", "liability": "asf", "asset": "rsf", "off_balance_sheet": "rsf"}

Row = dict[str, str]


@dataclass(frozen=True, order=True)
class Diagnostic:
    code: str
    record_id: str
    field: str

    def as_dict(self) -> dict[str, str]:
        return {"code": self.code, "record_id": self.record_id, "field": self.field}


def load_schema(root: Path) -> dict[str, Any]:
    data: dict[str, Any] = json.loads((root / SCHEMA).read_text(encoding="utf-8"))
    return data


def read_table(path: Path) -> tuple[list[str], list[Row]]:
    with path.open(encoding="utf-8", newline="") as handle:
        reader = csv.reader(handle)
        rows = list(reader)
    if not rows:
        return [], []
    header, body = rows[0], rows[1:]
    table = []
    for values in body:
        row = dict(zip(header, values, strict=False))
        if len(values) != len(header):
            row[WIDTH] = str(len(values))
        table.append(row)
    return header, table


def parse_date(value: str) -> date | None:
    try:
        return date.fromisoformat(value) if re.fullmatch(r"\d{4}-\d{2}-\d{2}", value) else None
    except ValueError:
        return None


def timestamp_ok(value: str) -> bool:
    if not TIMESTAMP.match(value):
        return False
    try:
        datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return False
    return True


def amount_code(value: str, pattern: str) -> str | None:
    """The diagnostic for an amount, or None when it is valid."""
    try:
        number = Decimal(value)
    except InvalidOperation:
        return "NSFR-E004"
    if not number.is_finite():
        return "NSFR-E004"
    if number < 0:
        return "NSFR-E005"
    if number >= LIMIT:
        return "NSFR-E016"
    exponent = number.as_tuple().exponent
    if isinstance(exponent, int) and exponent < -4:
        return "NSFR-E006"
    return None if re.match(pattern, value) else "NSFR-E002"


def classification_code(value: str, side: str, schema: Mapping[str, Any]) -> str | None:
    match = re.match(schema["classification"]["pattern"], value)
    if not match:
        return "NSFR-E002"
    articles = schema["classification"]["articles"].get(side)
    if articles is not None and match.group(1) not in articles:
        return "NSFR-E010"
    return None


def value_code(spec: Mapping[str, Any], value: str, row: Row, schema: Mapping[str, Any]) -> str | None:
    """The diagnostic for one present value, judged by its declared type."""
    kind = spec["type"]
    if kind == "enum":
        if value in spec["values"]:
            return None
        return "NSFR-E012" if value in spec.get("unsupported", []) else "NSFR-E002"
    if kind == "amount":
        return amount_code(value, schema["amount"]["pattern"])
    if kind == "classification":
        return classification_code(value, row.get("side", ""), schema)
    checks = {
        "identifier": lambda: bool(IDENTIFIER.match(value)),
        "integer": lambda: value.isdigit() and int(value) > 0,
        "date": lambda: parse_date(value) is not None,
        "timestamp": lambda: timestamp_ok(value),
        "currency": lambda: bool(CURRENCY.match(value)),
        "text": lambda: bool(value.strip()),
    }
    return None if checks[kind]() else "NSFR-E002"


CONDITIONS = {
    "maturity_date": ("maturity_kind", "dated"),
    "encumbrance_end_date": ("encumbered", "true"),
}


def field_diagnostics(spec: Mapping[str, Any], row: Row, record_id: str,
                      schema: Mapping[str, Any]) -> list[Diagnostic]:
    name = spec["name"]
    value = row.get(name, "")
    if spec["required"] is not True:
        controller, expected = CONDITIONS[name]
        if (row.get(controller) == expected) != bool(value):
            return [Diagnostic("NSFR-E011", record_id, name)]
        if not value:
            return []
    elif not value:
        return [Diagnostic("NSFR-E001", record_id, name)]
    code = value_code(spec, value, row, schema)
    return [Diagnostic(code, record_id, name)] if code else []


def header_diagnostics(header: list[str], fields: Sequence[Mapping[str, Any]], table: str) -> list[Diagnostic]:
    expected = [spec["name"] for spec in fields]
    return [] if header == expected else [Diagnostic("NSFR-E013", table, "header")]


def cross_diagnostics(row: Row, record_id: str, run: Row) -> list[Diagnostic]:
    found = []
    if CURRENCY.match(row.get("currency", "")) and row.get("currency") != run.get("reporting_currency"):
        found.append(Diagnostic("NSFR-E007", record_id, "currency"))
    as_of = parse_date(run.get("as_of_date", ""))
    for name in ("maturity_date", "encumbrance_end_date"):
        when = parse_date(row.get(name, ""))
        if as_of and when and when < as_of:
            found.append(Diagnostic("NSFR-E008", record_id, name))
    item_side = SIDES.get(row.get("item_type", ""))
    if item_side and row.get("side") in ("asf", "rsf") and item_side != row.get("side"):
        found.append(Diagnostic("NSFR-E009", record_id, "item_type"))
    return found


def validate(schema: Mapping[str, Any], run_table: tuple[list[str], list[Row]],
             positions_table: tuple[list[str], list[Row]]) -> list[Diagnostic]:
    """Every contract-stage diagnostic for one input, sorted."""
    tables = schema["tables"]
    run_header, run_rows = run_table
    header, positions = positions_table
    found = header_diagnostics(run_header, tables["run"]["fields"], "run")
    found += header_diagnostics(header, tables["positions"]["fields"], "positions")
    if len(run_rows) != 1:
        found.append(Diagnostic("NSFR-E013", "run", "rows"))
    run = run_rows[0] if run_rows else {}
    found += [Diagnostic("NSFR-E013", "run", "cells") for row in run_rows if WIDTH in row]
    for spec in tables["run"]["fields"]:
        found += field_diagnostics(spec, run, "run", schema)
    seen: set[str] = set()
    for row in positions:
        record_id = row.get("record_id", "") or "?"
        if record_id in seen:
            found.append(Diagnostic("NSFR-E003", record_id, "record_id"))
        seen.add(record_id)
        if WIDTH in row:
            found.append(Diagnostic("NSFR-E013", record_id, "cells"))
        for spec in tables["positions"]["fields"]:
            found += field_diagnostics(spec, row, record_id, schema)
        found += cross_diagnostics(row, record_id, run)
    return sorted(set(found))


def money(value: Decimal) -> str:
    return str(value.quantize(SCALE))


def reconciliation(positions: Sequence[Row]) -> dict[str, Any]:
    """Counts and amount totals by side and by classification; no factor applied."""
    with localcontext() as context:
        context.prec = PRECISION
        return exact_reconciliation(positions)


def exact_reconciliation(positions: Sequence[Row]) -> dict[str, Any]:
    sides: dict[str, list[Any]] = {"asf": [0, Decimal(0)], "rsf": [0, Decimal(0)]}
    groups: dict[tuple[str, str], list[Any]] = {}
    for row in positions:
        amount = Decimal(row["amount"])
        for bucket in (sides[row["side"]], groups.setdefault((row["side"], row["classification"]),
                                                              [0, Decimal(0)])):
            bucket[0] += 1
            bucket[1] += amount
    return {
        "record_count": len(positions),
        "sides": {side: {"count": count, "amount": money(total)}
                  for side, (count, total) in sides.items()},
        "classifications": [{"side": side, "classification": key, "count": count,
                             "amount": money(total)}
                            for (side, key), (count, total) in sorted(groups.items())],
    }


def canonical(schema: Mapping[str, Any], run: Row, positions: Sequence[Row]) -> str:
    """The canonical text of a valid input: schema column order, records by ID, 4 decimals."""
    buffer = io.StringIO()
    writer = csv.writer(buffer, lineterminator="\n")
    for table, rows in (("run", [run]), ("positions", sorted(positions, key=lambda r: r["record_id"]))):
        names = [spec["name"] for spec in schema["tables"][table]["fields"]]
        buffer.write(f"[{table}]\n")
        writer.writerow(names)
        for row in rows:
            writer.writerow([money(Decimal(row[name])) if name == "amount" else row.get(name, "")
                             for name in names])
    return buffer.getvalue()


def fingerprint(schema: Mapping[str, Any], run: Row, positions: Sequence[Row]) -> str:
    with localcontext() as context:
        context.prec = PRECISION
        text = canonical(schema, run, positions)
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def evaluate_example(schema: Mapping[str, Any], directory: Path) -> dict[str, Any]:
    run_table = read_table(directory / "run.csv")
    positions_table = read_table(directory / "positions.csv")
    diagnostics = validate(schema, run_table, positions_table)
    if diagnostics:
        return {"diagnostics": [d.as_dict() for d in diagnostics]}
    run, positions = run_table[1][0], positions_table[1]
    return {"diagnostics": [], "reconciliation": reconciliation(positions),
            "input_fingerprint": fingerprint(schema, run, positions)}


def document_problems(schema: Mapping[str, Any], text: str) -> list[str]:
    """Every field and diagnostic code is documented in exactly one table row."""
    names = [spec["name"] for table in schema["tables"].values() for spec in table["fields"]]
    problems = []
    for name in [*names, *schema["diagnostics"]]:
        count = len(re.findall(rf"^\| `{re.escape(name)}` \|", text, re.M))
        if count != 1:
            problems.append(f"{DOCUMENT}: `{name}` must head exactly one table row, found {count}")
    return problems


def check(root: Path) -> list[str]:
    schema = load_schema(root)
    problems = document_problems(schema, (root / DOCUMENT).read_text(encoding="utf-8"))
    directories = sorted(path for path in (root / EXAMPLES).iterdir() if path.is_dir())
    if not directories:
        problems.append(f"{EXAMPLES}: no examples")
    for directory in directories:
        expected = json.loads((directory / "expected.json").read_text(encoding="utf-8"))
        actual = evaluate_example(schema, directory)
        if actual != expected:
            name = directory.relative_to(root).as_posix()
            problems.append(f"{name}: result differs from expected.json:\n"
                            f"{json.dumps(actual, indent=2, sort_keys=True)}")
    return problems


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0] if __doc__ else None)
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--show", type=Path, help="print the result of one example directory")
    args = parser.parse_args(argv)
    root = args.root.resolve()
    if args.show:
        print(json.dumps(evaluate_example(load_schema(root), args.show), indent=2, sort_keys=True))
        return 0
    problems = check(root)
    for problem in problems:
        print(problem, file=sys.stderr)
    print(f"check_contracts: {len(problems)} problem(s)")
    return 1 if problems else 0


if __name__ == "__main__":
    raise SystemExit(main())
