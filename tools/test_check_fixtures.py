"""Tests for the numerical fixture checker (tools/check_fixtures.py)."""

from __future__ import annotations

import contextlib
import copy
import csv
import io
import json
import shutil
import tempfile
import unittest
from collections.abc import Callable
from fractions import Fraction
from pathlib import Path
from typing import Any

import check_contracts as cc
import check_fixtures as cf

ROOT = Path(__file__).resolve().parents[1]
NEUTRAL = ROOT / cf.FIXTURES / "neutral-synthetic"
SCHEMA = cc.load_schema(ROOT)
SPEC = cf.load(ROOT / cf.SCHEMA)
FIXTURE = cf.load(NEUTRAL / cf.FIXTURE_FILE)
VERIFIED = cf.verified_sources((ROOT / cf.REGISTER).read_text(encoding="utf-8"))


class Copy:
    """A writable copy of the neutral fixture."""

    def __init__(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.path = Path(self.temp.name) / "fixture"
        shutil.copytree(NEUTRAL, self.path)

    def __enter__(self) -> Copy:
        return self

    def __exit__(self, *args: object) -> None:
        self.temp.cleanup()

    def edit(self, name: str, change: Callable[[Any], object]) -> None:
        data = cf.load(self.path / name)
        change(data)
        (self.path / name).write_text(json.dumps(data), encoding="utf-8")

    def rows(self, change: Callable[[list[dict[str, str]]], list[dict[str, str]]]) -> None:
        header, rows = cc.read_table(self.path / "positions.csv")
        with (self.path / "positions.csv").open("w", encoding="utf-8", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=header, lineterminator="\n")
            writer.writeheader()
            writer.writerows(change(rows))

    def result(self) -> dict[str, Any]:
        return cf.evaluate(SCHEMA, self.path, FIXTURE["inputs"])

    def problems(self) -> list[str]:
        return cf.check_fixture(SCHEMA, SPEC, VERIFIED, self.path, "copy")


def codes(result: dict[str, Any]) -> list[tuple[str, str, str]]:
    return [(d["code"], d["record_id"], d["field"]) for d in result["diagnostics"]]


class RepositoryTests(unittest.TestCase):
    def test_the_registered_fixtures_pass(self) -> None:
        self.assertEqual(cf.check(ROOT), [])

    def test_the_neutral_example_matches_its_hand_derivation(self) -> None:
        result = cf.evaluate(SCHEMA, NEUTRAL, FIXTURE["inputs"])
        self.assertEqual(result["ratio"], "6.766915")
        self.assertEqual(result, FIXTURE["expected"])

    def test_show_prints_the_independent_result(self) -> None:
        out = io.StringIO()
        with contextlib.redirect_stdout(out):
            self.assertEqual(cf.main(["--root", str(ROOT), "--show", str(NEUTRAL)]), 0)
        self.assertEqual(json.loads(out.getvalue()), FIXTURE["expected"])


class RoundingTests(unittest.TestCase):
    def test_ties_go_to_the_even_digit(self) -> None:
        cases = {"1.2345675": "1.234568", "1.2345665": "1.234566", "0.0000005": "0.000000",
                 "0.0000015": "0.000002", "0": "0.000000", "2.9999995": "3.000000"}
        for value, text in cases.items():
            with self.subTest(value=value):
                self.assertEqual(cf.fixed(Fraction(value), 6), text)

    def test_near_ties_round_to_the_nearer_value(self) -> None:
        tie = Fraction("1.2345665")
        tiny = Fraction(1, 10 ** 40)
        self.assertEqual(cf.fixed(tie + tiny, 6), "1.234567")
        self.assertEqual(cf.fixed(tie - tiny, 6), "1.234566")

    def test_an_exact_quotient_is_rounded_once(self) -> None:
        self.assertEqual(cf.fixed(Fraction(2, 3), 6), "0.666667")
        self.assertEqual(cf.fixed(Fraction(1, 8), 8), "0.12500000")


class EvaluationTests(unittest.TestCase):
    def test_a_parameter_set_that_differs_from_the_run_is_rejected(self) -> None:
        with Copy() as fixture:
            fixture.edit("parameters.json", lambda p: p.update(parameter_set_version=2, profile_id="other"))
            result = fixture.result()
        self.assertEqual(result["status"], "rejected")
        self.assertEqual(codes(result), [("NSFR-E014", "run", "parameter_set_version"),
                                         ("NSFR-E014", "run", "profile_id")])
        self.assertIn("input_fingerprint", result)
        self.assertNotIn("ratio", result)

    def test_a_classification_without_a_factor_is_rejected(self) -> None:
        with Copy() as fixture:
            fixture.edit("parameters.json", lambda p: p["factors"].pop())
            result = fixture.result()
        self.assertEqual(codes(result), [("NSFR-E015", "r-002", "classification")])

    def test_zero_required_funding_gives_no_ratio_and_a_warning(self) -> None:
        def zero(parameters: dict[str, Any]) -> None:
            for entry in parameters["factors"]:
                if entry["classification"] in ("428ah", "428r", "428s"):
                    entry["factor"] = "0"
        with Copy() as fixture:
            fixture.edit("parameters.json", zero)
            result = fixture.result()
        self.assertEqual((result["status"], result["ratio"]), ("ok", None))
        self.assertEqual(result["required_stable_funding"], "0.00000000")
        self.assertEqual(codes(result), [("NSFR-W001", "run", "required_stable_funding")])

    def test_zero_available_funding_gives_a_zero_ratio(self) -> None:
        with Copy() as fixture:
            fixture.rows(lambda rows: [row for row in rows if row["side"] == "rsf"])
            result = fixture.result()
        self.assertEqual((result["available_stable_funding"], result["ratio"]), ("0.00000000", "0.000000"))

    def test_a_contract_error_stops_before_the_parameters(self) -> None:
        with Copy() as fixture:
            fixture.rows(lambda rows: [{**rows[0], "currency": "USD"}, *rows[1:]])
            result = fixture.result()
        self.assertEqual(result, {"status": "rejected",
                                  "diagnostics": [{"code": "NSFR-E007", "record_id": "r-003",
                                                   "field": "currency"}]})

    def test_the_order_of_positions_does_not_change_the_result(self) -> None:
        with Copy() as fixture:
            fixture.rows(lambda rows: list(reversed(rows)))
            self.assertEqual(fixture.result(), FIXTURE["expected"])

    def test_totals_at_the_amount_limit_stay_exact(self) -> None:
        with Copy() as fixture:
            fixture.rows(lambda rows: [{**row, "amount": "999999999999999.9999"} for row in rows])
            result = fixture.result()
        self.assertEqual(result["available_stable_funding"], "1499999999999999.99985000")


class MetadataTests(unittest.TestCase):
    def test_a_missing_section_is_reported(self) -> None:
        with Copy() as fixture:
            fixture.edit(cf.FIXTURE_FILE, lambda f: f.pop("review"))
            self.assertEqual(fixture.problems(), ["copy: missing `review`"])

    def test_the_oracle_cannot_be_the_implementation(self) -> None:
        with Copy() as fixture:
            fixture.edit(cf.FIXTURE_FILE, lambda f: f["derivation"].update(
                method="implementation", independent_of_implementation=False))
            problems = fixture.problems()
        self.assertIn("copy derivation: `method` must be one of hand, independent-tool, published", problems)
        self.assertTrue(any("independent of the implementation" in p for p in problems))

    def test_a_reviewed_fixture_names_its_reviewer(self) -> None:
        with Copy() as fixture:
            fixture.edit(cf.FIXTURE_FILE, lambda f: f["review"].update(status="reviewed"))
            self.assertEqual(fixture.problems(),
                             ["copy: a reviewed or approved fixture names its reviewer and date"])

    def test_execution_needs_recorded_runs(self) -> None:
        with Copy() as fixture:
            fixture.edit(cf.FIXTURE_FILE, lambda f: f["execution"].update(status="executed"))
            self.assertEqual(len(fixture.problems()), 1)
            fixture.edit(cf.FIXTURE_FILE, lambda f: f["execution"].update(runs=[{"engine_commit": "abc"}]))
            self.assertEqual(len(fixture.problems()), 4)

    def test_the_fixture_identifier_has_a_fixed_form(self) -> None:
        with Copy() as fixture:
            fixture.edit(cf.FIXTURE_FILE, lambda f: f.update(fixture_id="F-1"))
            self.assertEqual(fixture.problems(), ["copy: fixture_id must match FX-nnnn"])


class ProvenanceTests(unittest.TestCase):
    def test_synthetic_factors_belong_to_a_neutral_example(self) -> None:
        with Copy() as fixture:
            fixture.edit(cf.FIXTURE_FILE, lambda f: f.update(kind="validation"))
            problems = fixture.problems()
        self.assertIn("copy: synthetic parameters are allowed in, and required by, a neutral example",
                      problems)

    def validation_copy(self, fixture: Copy, source: str = "S-01") -> None:
        def cite(parameters: dict[str, Any]) -> None:
            parameters["review_status"] = "reviewed"
            for entry in parameters["factors"]:
                entry.update(source_id=source, locator="Article 428 (synthetic test)")
        fixture.edit(cf.FIXTURE_FILE, lambda f: (f.update(kind="validation"),
                                                 f["provenance"].update(sources=[f"{source} Article 428"])))
        fixture.edit("parameters.json", cite)

    def test_validation_factors_cite_normative_register_entries(self) -> None:
        with Copy() as fixture:
            self.validation_copy(fixture)
            self.assertEqual(fixture.problems(), [])
            fixture.edit("parameters.json", lambda p: p["factors"][0].update(source_id="R-01"))
            fixture.edit(cf.FIXTURE_FILE, lambda f: f["provenance"].update(sources=[]))
            self.assertEqual(fixture.problems(), ["copy: factor of 428k must cite a register entry",
                                                  "copy: a validation fixture lists its register sources"])

    def test_citations_must_exist_in_the_register(self) -> None:
        with Copy() as fixture:
            self.validation_copy(fixture, "S-99")
            problems = fixture.problems()
        self.assertIn("copy: factor of 428k cites S-99, which is not a verified entry of the source register",
                      problems)
        self.assertIn("copy: provenance source `S-99 Article 428` does not start with a verified normative "
                      "register entry", problems)

    def test_only_verified_register_entries_count(self) -> None:
        text = ("| S-01 | Act | Version | Scope | Link | OJ reference: Verified (EUR-Lex) |\n"
                "| S-02 | Act | Version | Scope | Link | To verify on EUR-Lex |\n"
                "| S-03 | Act | Version | Scope | Link | Verified (EUR-Lex); article to verify |\n"
                "| R-01 | Act | Version | Scope | Link | Verified (EUR-Lex) |\n")
        self.assertEqual(cf.verified_sources(text), frozenset({"S-01"}))
        self.assertTrue({"S-01", "S-02", "S-03", "S-04"} <= VERIFIED)

    def test_factors_are_decimals_from_zero_to_one(self) -> None:
        for value in ("1.5", "0.12345", "-0.1", "1.0001"):
            def change(parameters: dict[str, Any], factor: str = value) -> None:
                parameters["factors"][0]["factor"] = factor
            with self.subTest(value=value), Copy() as fixture:
                fixture.edit("parameters.json", change)
                self.assertTrue(any("must be 0 to 1" in p for p in fixture.problems()))

    def test_a_classification_has_one_factor(self) -> None:
        with Copy() as fixture:
            fixture.edit("parameters.json", lambda p: p["factors"].append(copy.deepcopy(p["factors"][0])))
            self.assertIn("copy: a classification has more than one factor", fixture.problems())


class DerivationTests(unittest.TestCase):
    def test_a_wrong_weighted_amount_is_reported(self) -> None:
        with Copy() as fixture:
            fixture.edit(cf.FIXTURE_FILE, lambda f: f["derivation"]["weighted"][1].update(weighted="125000.2"))
            problems = fixture.problems()
        self.assertIn("copy: derivation step a-002 is not amount x factor", problems)
        self.assertIn("copy: derivation available_stable_funding does not add up to the expected value",
                      problems)

    def test_every_position_is_weighted_once(self) -> None:
        with Copy() as fixture:
            fixture.edit(cf.FIXTURE_FILE, lambda f: f["derivation"]["weighted"].pop())
            self.assertEqual(fixture.problems(),
                             ["copy: the derivation must weight every position exactly once"])

    def test_a_wrong_expected_ratio_is_reported(self) -> None:
        with Copy() as fixture:
            fixture.edit(cf.FIXTURE_FILE, lambda f: f["expected"].update(ratio="6.766916"))
            self.assertEqual(fixture.problems(), ['copy: `ratio` expected "6.766916", got "6.766915"'])


class ComparisonTests(unittest.TestCase):
    def test_a_missing_field_differs_from_an_explicit_null(self) -> None:
        self.assertEqual(cf.compare({"ratio": None}, {"ratio": None}, [], "x"), [])
        self.assertEqual(cf.compare({}, {"ratio": None}, [], "x"), ["x: `ratio` expected no field, got null"])
        self.assertEqual(cf.compare({"ratio": None}, {}, [], "x"), ["x: `ratio` expected null, got no field"])

    def test_tolerances_are_approved_exceptions(self) -> None:
        exception = {"field": "status", "mode": "relative", "tolerance": "0.1",
                     "reason": "test", "approved_in": "soon"}
        with Copy() as fixture:
            fixture.edit(cf.FIXTURE_FILE, lambda f: f["comparison"]["exceptions"].append(exception))
            problems = fixture.problems()
        self.assertIn("copy: only available_stable_funding, required_stable_funding, ratio may have a "
                      "tolerance", problems)
        self.assertIn("copy: a relative tolerance needs an absolute floor near zero", problems)
        self.assertIn("copy: a tolerance names the issue that approved it", problems)

    def test_an_approved_tolerance_admits_a_close_value(self) -> None:
        exception = {"field": "ratio", "mode": "absolute", "tolerance": "0.000001",
                     "reason": "test", "approved_in": "#10"}
        with Copy() as fixture:
            fixture.edit(cf.FIXTURE_FILE, lambda f: (f["comparison"]["exceptions"].append(exception),
                                                     f["expected"].update(ratio="6.766916")))
            self.assertEqual(fixture.problems(), [])

    def test_within_applies_the_mode_and_the_floor(self) -> None:
        relative = {"mode": "relative", "tolerance": "0.01", "floor": "0.001"}
        self.assertTrue(cf.within("100", "100.9", relative))
        self.assertFalse(cf.within("100", "101.1", relative))
        self.assertTrue(cf.within("0", "0.001", relative))
        self.assertFalse(cf.within("0", "0.002", relative))
        self.assertFalse(cf.within("NaN", "NaN", {"mode": "absolute", "tolerance": "1"}))
        self.assertFalse(cf.within("1", "Infinity", {"mode": "absolute", "tolerance": "1"}))
        self.assertFalse(cf.within("1", "x", {"mode": "absolute", "tolerance": "1"}))

    def test_the_default_comparison_is_exact(self) -> None:
        with Copy() as fixture:
            fixture.edit(cf.FIXTURE_FILE, lambda f: f["comparison"].update(default="tolerant"))
            self.assertIn("copy: the default comparison is exact", fixture.problems())


class SchemaTests(unittest.TestCase):
    def test_the_schema_declares_the_rules_the_checker_applies(self) -> None:
        self.assertEqual(cf.constant_problems(SPEC), [])
        changed = copy.deepcopy(SPEC)
        changed["expected"]["ratio_places"] = 4
        self.assertEqual(cf.constant_problems(changed), [f"{cf.SCHEMA}: `ratio_places` differs from the checker"])


class DocumentTests(unittest.TestCase):
    def test_every_required_field_heads_one_row(self) -> None:
        text = (ROOT / cf.DOCUMENT).read_text(encoding="utf-8")
        self.assertEqual(cf.document_problems(SPEC, text), [])
        missing = text.replace("| `review.date` |", "| review date |")
        self.assertEqual(cf.document_problems(SPEC, missing),
                         [f"{cf.DOCUMENT}: `review.date` must head exactly one table row, found 0"])


if __name__ == "__main__":
    unittest.main()
