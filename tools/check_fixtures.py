#!/usr/bin/env python3
"""Check the registered numerical fixtures against their schema and an independent evaluation.

This implements the fixture rules of docs/validation/VALIDATION.md. For every
fixture under tests/fixtures it checks the metadata, re-derives the
hand-written derivation, and evaluates the input with exact rational
arithmetic: contract validation, the engine-stage diagnostics, the stable
funding totals and the ratio. It never runs the engine, so a fixture it
accepts is registered, not executed.

This Source Code Form is subject to the terms of the Mozilla Public
License, v. 2.0. If a copy of the MPL was not distributed with this
file, You can obtain one at https://mozilla.org/MPL/2.0/.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from collections.abc import Mapping, Sequence
from decimal import Decimal, InvalidOperation
from fractions import Fraction
from pathlib import Path
from typing import Any

import check_contracts as contracts

SCHEMA = Path("docs/validation/fixture-v1.json")
DOCUMENT = Path("docs/validation/VALIDATION.md")
FIXTURES = Path("tests/fixtures")
FIXTURE_FILE = "fixture.json"
FIXTURE_ID = re.compile(r"^FX-[0-9]{4}$")
DATE = re.compile(r"^[0-9]{4}-[0-9]{2}-[0-9]{2}$")
ISSUE = re.compile(r"^#[1-9][0-9]*$")
SOURCE_ID = re.compile(r"^S-[0-9]{2}$")  # factors cite normative sources only
FACTOR = re.compile(r"^(0(\.[0-9]{1,4})?|1(\.0{1,4})?)$")
FUNDING_PLACES = 8
RATIO_PLACES = 6
RUN = "run"
REGISTER = Path("docs/methodology/SOURCE_REGISTER.md")
MISSING = object()  # a field absent from a result, unlike an explicit null
TOLERABLE = ("available_stable_funding", "required_stable_funding", "ratio")


def load(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def fixed(value: Fraction, places: int) -> str:
    """Round a non-negative exact value half-even to `places` decimals, as text."""
    scaled = value * 10 ** places
    whole, remainder = divmod(scaled.numerator, scaled.denominator)
    twice = 2 * remainder
    if twice > scaled.denominator or (twice == scaled.denominator and whole % 2 == 1):
        whole += 1
    digits = str(whole).rjust(places + 1, "0")
    return f"{digits[:-places]}.{digits[-places:]}"


def exact(text: str) -> Fraction:
    return Fraction(Decimal(text))


def diagnostic(code: str, record_id: str, field: str) -> dict[str, str]:
    return {"code": code, "record_id": record_id, "field": field}


def parameter_diagnostics(run: Mapping[str, str], parameters: Mapping[str, Any]) -> list[dict[str, str]]:
    """NSFR-E014 when the parameter set does not match the run's identity."""
    pairs = (("parameter_set_id", parameters.get("parameter_set_id")),
             ("parameter_set_version", str(parameters.get("parameter_set_version"))),
             ("profile_id", parameters.get("profile_id")),
             ("register_version", str(parameters.get("register_version"))))
    return [diagnostic("NSFR-E014", RUN, name) for name, value in pairs if run.get(name) != value]


def factor_table(parameters: Mapping[str, Any]) -> dict[str, Fraction]:
    return {entry["classification"]: exact(entry["factor"]) for entry in parameters.get("factors", [])}


def evaluate(schema: Mapping[str, Any], directory: Path, inputs: Mapping[str, str]) -> dict[str, Any]:
    """The independent result for one fixture input, in the reported form."""
    run_table = contracts.read_table(directory / inputs["run"])
    positions_table = contracts.read_table(directory / inputs["positions"])
    found = contracts.validate(schema, run_table, positions_table)
    if found:
        return {"status": "rejected", "diagnostics": [d.as_dict() for d in found]}
    run, positions = run_table[1][0], positions_table[1]
    fingerprint = contracts.fingerprint(schema, run, positions)
    parameters = load(directory / inputs["parameters"])
    errors = parameter_diagnostics(run, parameters)
    factors = factor_table(parameters)
    errors += [diagnostic("NSFR-E015", row["record_id"], "classification")
               for row in sorted(positions, key=lambda r: r["record_id"])
               if row["classification"] not in factors]
    if errors:
        ordered = sorted(errors, key=lambda d: (d["code"], d["record_id"], d["field"]))
        return {"status": "rejected", "input_fingerprint": fingerprint, "diagnostics": ordered}
    totals = {"asf": Fraction(0), "rsf": Fraction(0)}
    for row in positions:
        totals[row["side"]] += exact(row["amount"]) * factors[row["classification"]]
    result: dict[str, Any] = {
        "status": "ok", "input_fingerprint": fingerprint,
        "available_stable_funding": fixed(totals["asf"], FUNDING_PLACES),
        "required_stable_funding": fixed(totals["rsf"], FUNDING_PLACES),
        "ratio": None, "diagnostics": []}
    if totals["rsf"] == 0:
        result["diagnostics"] = [diagnostic("NSFR-W001", RUN, "required_stable_funding")]
    else:
        result["ratio"] = fixed(totals["asf"] / totals["rsf"], RATIO_PLACES)
    return result


def missing_keys(data: Mapping[str, Any], keys: Sequence[str], where: str) -> list[str]:
    return [f"{where}: missing `{key}`" for key in keys if key not in data]


def enum_problems(data: Mapping[str, Any], enums: Mapping[str, Sequence[str]], where: str) -> list[str]:
    return [f"{where}: `{key}` must be one of {', '.join(values)}"
            for key, values in enums.items() if key in data and data[key] not in values]


def metadata_problems(fixture: Mapping[str, Any], spec: Mapping[str, Any], where: str) -> list[str]:
    """Required sections, enumerations and identifiers of fixture.json."""
    problems = missing_keys(fixture, spec["required"], where)
    if problems:
        return problems
    for section, keys in spec["sections"].items():
        problems += missing_keys(fixture[section], keys, f"{where} {section}")
    if problems:
        return problems
    problems += enum_problems(fixture, spec["enums"]["fixture"], where)
    for section in ("provenance", "derivation", "review", "execution"):
        problems += enum_problems(fixture[section], spec["enums"].get(section, {}), f"{where} {section}")
    if fixture["fixture_schema_version"] != spec["fixture_schema_version"]:
        problems.append(f"{where}: unsupported fixture_schema_version")
    if not FIXTURE_ID.match(str(fixture["fixture_id"])):
        problems.append(f"{where}: fixture_id must match FX-nnnn")
    if fixture["derivation"]["independent_of_implementation"] is not True:
        problems.append(f"{where}: the derivation must be independent of the implementation under test")
    problems += review_problems(fixture["review"], where)
    return problems + execution_problems(fixture["execution"], spec["execution_run"], where)


def review_problems(review: Mapping[str, Any], where: str) -> list[str]:
    if review.get("status") == "draft":
        return []
    if not review.get("reviewer") or not DATE.match(str(review.get("date"))):
        return [f"{where}: a reviewed or approved fixture names its reviewer and date"]
    return []


def execution_problems(execution: Mapping[str, Any], keys: Sequence[str], where: str) -> list[str]:
    runs = execution.get("runs", [])
    if (execution.get("status") == "executed") != bool(runs):
        return [f"{where}: an executed fixture lists its runs, and only an executed one does"]
    return [problem for run in runs for problem in missing_keys(run, keys, f"{where} execution run")]


def verified_sources(text: str) -> frozenset[str]:
    """Normative register entries whose status is verified and has nothing left to verify."""
    found = set()
    for line in text.splitlines():
        match = re.match(r"^\| (S-[0-9]{2}) \|", line)
        status = line.rstrip().rstrip("|").rsplit("|", 1)[-1].lower()
        if match and "verified" in status and "to verify" not in status:
            found.add(match.group(1))
    return frozenset(found)


def factor_problems(entry: Mapping[str, Any], synthetic: bool, verified: frozenset[str], where: str) -> list[str]:
    key = entry.get("classification")
    problems = []
    if not FACTOR.match(str(entry.get("factor"))):
        problems.append(f"{where}: factor of {key} must be 0 to 1, at most 4 decimals")
    if synthetic:
        if (entry.get("source_id"), entry.get("locator")) != ("synthetic", "none"):
            problems.append(f"{where}: a synthetic factor has source_id `synthetic` and locator `none`")
    elif not (SOURCE_ID.match(str(entry.get("source_id"))) and entry.get("locator")):
        problems.append(f"{where}: factor of {key} must cite a register entry")
    elif entry["source_id"] not in verified:
        problems.append(f"{where}: factor of {key} cites {entry['source_id']}, "
                        "which is not a verified entry of the source register")
    return problems


def provenance_problems(fixture: Mapping[str, Any], parameters: Mapping[str, Any],
                        verified: frozenset[str], where: str) -> list[str]:
    """Synthetic parameters belong to neutral examples; others cite verified register entries."""
    problems = []
    synthetic = parameters.get("review_status") == "synthetic"
    if synthetic != (fixture["kind"] == "neutral-example"):
        problems.append(f"{where}: synthetic parameters are allowed in, and required by, a neutral example")
    for entry in parameters.get("factors", []):
        problems += factor_problems(entry, synthetic, verified, where)
    keys = [entry.get("classification") for entry in parameters.get("factors", [])]
    if len(keys) != len(set(keys)):
        problems.append(f"{where}: a classification has more than one factor")
    sources = fixture["provenance"]["sources"]
    if not synthetic and not sources:
        problems.append(f"{where}: a validation fixture lists its register sources")
    for source in sources:
        if str(source).split(" ", 1)[0] not in verified:
            problems.append(f"{where}: provenance source `{source}` does not start with a verified "
                            "normative register entry")
    return problems


def derivation_problems(fixture: Mapping[str, Any], directory: Path, where: str) -> list[str]:
    """Re-derive every hand-written weighted amount and total."""
    derivation, expected = fixture["derivation"], fixture["expected"]
    if expected["status"] != "ok":
        return []
    positions = {row["record_id"]: row for row in contracts.read_table(
        directory / fixture["inputs"]["positions"])[1]}
    factors = factor_table(load(directory / fixture["inputs"]["parameters"]))
    steps = derivation.get("weighted", [])
    problems = []
    if sorted(step.get("record_id") for step in steps) != sorted(positions):
        problems.append(f"{where}: the derivation must weight every position exactly once")
        return problems
    totals = {"asf": Fraction(0), "rsf": Fraction(0)}
    for step in steps:
        row = positions[step["record_id"]]
        try:
            amount, factor, weighted = (exact(step[key]) for key in ("amount", "factor", "weighted"))
        except (KeyError, InvalidOperation):
            problems.append(f"{where}: derivation step {step.get('record_id')} is incomplete")
            continue
        if amount != exact(row["amount"]) or factor != factors.get(row["classification"]):
            problems.append(f"{where}: derivation step {step['record_id']} does not match the inputs")
        if weighted != amount * factor:
            problems.append(f"{where}: derivation step {step['record_id']} is not amount x factor")
        totals[row["side"]] += weighted
    for side, field in (("asf", "available_stable_funding"), ("rsf", "required_stable_funding")):
        if derivation.get(field) != expected.get(field) or exact(str(expected.get(field))) != totals[side]:
            problems.append(f"{where}: derivation {field} does not add up to the expected value")
    return problems


def compare(expected: Mapping[str, Any], actual: Mapping[str, Any],
            exceptions: Sequence[Mapping[str, Any]], where: str) -> list[str]:
    """Exact comparison, except for fields with an approved tolerance."""
    tolerated = {entry["field"]: entry for entry in exceptions}
    problems = []
    for field in sorted(set(expected) | set(actual)):
        want, got = expected.get(field, MISSING), actual.get(field, MISSING)
        if field in tolerated and isinstance(want, str) and isinstance(got, str):
            if not within(want, got, tolerated[field]):
                problems.append(f"{where}: `{field}` {got} is outside the tolerance of {want}")
        elif want is MISSING or got is MISSING or want != got:
            problems.append(f"{where}: `{field}` expected {shown(want)}, got {shown(got)}")
    return problems


def shown(value: object) -> str:
    return "no field" if value is MISSING else json.dumps(value)


def within(want: str, got: str, rule: Mapping[str, Any]) -> bool:
    """A tolerance comparison of two finite decimals; anything non-finite fails."""
    try:
        values = [Decimal(text) for text in (want, got, rule["tolerance"], rule.get("floor", "0"))]
    except InvalidOperation:
        return False
    if not all(value.is_finite() for value in values):
        return False
    a, b, tolerance, floor = values
    allowed = tolerance if rule["mode"] == "absolute" else max(tolerance * abs(a), floor)
    return abs(a - b) <= allowed


def exception_problems(comparison: Mapping[str, Any], spec: Mapping[str, Any], where: str) -> list[str]:
    problems = []
    if comparison.get("default") != "exact":
        problems.append(f"{where}: the default comparison is exact")
    for entry in comparison.get("exceptions", []):
        problems += missing_keys(entry, spec["exception"], f"{where} comparison exception")
        if entry.get("field") not in TOLERABLE:
            problems.append(f"{where}: only {', '.join(TOLERABLE)} may have a tolerance")
        if entry.get("mode") not in ("absolute", "relative"):
            problems.append(f"{where}: a tolerance mode is absolute or relative")
        if entry.get("mode") == "relative" and "floor" not in entry:
            problems.append(f"{where}: a relative tolerance needs an absolute floor near zero")
        if not ISSUE.match(str(entry.get("approved_in"))):
            problems.append(f"{where}: a tolerance names the issue that approved it")
    return problems


def check_fixture(schema: Mapping[str, Any], spec: Mapping[str, Any], verified: frozenset[str],
                  directory: Path, where: str) -> list[str]:
    fixture = load(directory / FIXTURE_FILE)
    problems = metadata_problems(fixture, spec, where)
    if problems:
        return problems
    if fixture["contract_version"] != schema["contract_version"]:
        problems.append(f"{where}: contract_version differs from the contract")
    problems += exception_problems(fixture["comparison"], spec, where)
    parameters = load(directory / fixture["inputs"]["parameters"])
    problems += provenance_problems(fixture, parameters, verified, where)
    actual = evaluate(schema, directory, fixture["inputs"])
    problems += compare(fixture["expected"], actual, fixture["comparison"].get("exceptions", []), where)
    return problems + derivation_problems(fixture, directory, where)


def document_problems(spec: Mapping[str, Any], text: str) -> list[str]:
    """Every fixture field is documented in exactly one table row of VALIDATION.md."""
    names = [*spec["required"], *(f"{section}.{key}" for section, keys in spec["sections"].items()
                                  for key in keys)]
    problems = []
    for name in names:
        count = len(re.findall(rf"^\| `{re.escape(name)}` \|", text, re.M))
        if count != 1:
            problems.append(f"{DOCUMENT}: `{name}` must head exactly one table row, found {count}")
    return problems


def constant_problems(spec: Mapping[str, Any]) -> list[str]:
    """The rules this checker applies are the ones fixture-v1.json declares."""
    declared = {"funding_places": (spec["expected"]["funding_places"], FUNDING_PLACES),
                "ratio_places": (spec["expected"]["ratio_places"], RATIO_PLACES),
                "rounding": (spec["expected"]["rounding"], "half-even"),
                "factor_pattern": (spec["parameters"]["factor_pattern"], FACTOR.pattern)}
    return [f"{SCHEMA}: `{name}` differs from the checker" for name, (value, applied) in declared.items()
            if value != applied]


def check(root: Path) -> list[str]:
    schema = contracts.load_schema(root)
    spec = load(root / SCHEMA)
    verified = verified_sources((root / REGISTER).read_text(encoding="utf-8"))
    problems = constant_problems(spec)
    problems += document_problems(spec, (root / DOCUMENT).read_text(encoding="utf-8"))
    directories = sorted(path.parent for path in (root / FIXTURES).glob(f"*/{FIXTURE_FILE}"))
    if not directories:
        problems.append(f"{FIXTURES}: no registered fixture")
    for directory in directories:
        problems += check_fixture(schema, spec, verified, directory, directory.relative_to(root).as_posix())
    return problems


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0] if __doc__ else None)
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--show", type=Path, help="print the independent result of one fixture directory")
    args = parser.parse_args(argv)
    root = args.root.resolve()
    if args.show:
        fixture = load(args.show / FIXTURE_FILE)
        result = evaluate(contracts.load_schema(root), args.show, fixture["inputs"])
        print(json.dumps(result, indent=2, sort_keys=True))
        return 0
    problems = check(root)
    for problem in problems:
        print(problem, file=sys.stderr)
    print(f"check_fixtures: {len(problems)} problem(s)")
    return 1 if problems else 0


if __name__ == "__main__":
    raise SystemExit(main())
