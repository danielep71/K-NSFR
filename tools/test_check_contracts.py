"""Tests for the NSFR contract checker (tools/check_contracts.py)."""

from __future__ import annotations

import contextlib
import io
import json
import shutil
import tempfile
import unittest
from pathlib import Path
from typing import Any

import check_contracts as cc

ROOT = Path(__file__).resolve().parents[1]
SCHEMA = cc.load_schema(ROOT)
RUN_FIELDS = [spec["name"] for spec in SCHEMA["tables"]["run"]["fields"]]
POSITION_FIELDS = [spec["name"] for spec in SCHEMA["tables"]["positions"]["fields"]]
RUN = {"contract_version": "1.0.0", "profile_id": "eu-crr-standard", "profile_version": "1",
       "register_version": "1", "parameter_set_id": "synthetic-parameters",
       "parameter_set_version": "1", "as_of_date": "2026-06-30",
       "reporting_entity_id": "SYNTHETIC-BANK-A", "basis": "individual",
       "reporting_currency": "EUR", "amount_unit": "units", "source_system": "ledger",
       "extracted_at": "2026-07-01T08:00:00Z"}
POSITION = {"record_id": "p-001", "side": "asf", "item_type": "liability",
            "description": "Synthetic liability", "amount": "100.5", "currency": "EUR",
            "maturity_kind": "dated", "maturity_date": "2027-06-30", "classification": "428l",
            "encumbered": "false", "encumbrance_end_date": "", "counterparty_type": "other",
            "transaction_type": "none", "source_record": "ledger:1"}


def codes(*positions: dict[str, str], run: dict[str, str] | None = None) -> list[tuple[str, str]]:
    found = cc.validate(SCHEMA, (RUN_FIELDS, [run or RUN]), (POSITION_FIELDS, list(positions)))
    return [(d.code, d.field) for d in found]


def position(**changes: str) -> dict[str, str]:
    return {**POSITION, **changes}


class ValidInputTests(unittest.TestCase):
    def test_a_valid_position_has_no_diagnostic(self) -> None:
        self.assertEqual(codes(POSITION), [])

    def test_conditional_dates_follow_their_controller(self) -> None:
        self.assertEqual(codes(position(maturity_kind="no_stated_maturity", maturity_date="")), [])
        self.assertEqual(codes(position(encumbered="true", encumbrance_end_date="2026-06-30")), [])

    def test_classification_paragraphs_and_points(self) -> None:
        self.assertEqual(codes(position(classification="428l(2)(a)(i)")), [])


class AmountTests(unittest.TestCase):
    def test_each_amount_problem_has_its_code(self) -> None:
        cases = {"": "NSFR-E001", "NaN": "NSFR-E004", "Infinity": "NSFR-E004", "abc": "NSFR-E004",
                 "-1": "NSFR-E005", "1.00001": "NSFR-E006", "1e3": "NSFR-E002",
                 "+1": "NSFR-E002", "1,000": "NSFR-E004"}
        for value, code in cases.items():
            with self.subTest(value=value):
                self.assertEqual(codes(position(amount=value)), [(code, "amount")])

    def test_four_decimals_and_zero_are_accepted(self) -> None:
        self.assertEqual(codes(position(amount="0")), [])
        self.assertEqual(codes(position(amount="1.0000")), [])

    def test_amounts_are_limited_to_fifteen_integer_digits(self) -> None:
        self.assertEqual(codes(position(amount="999999999999999.9999")), [])
        self.assertEqual(codes(position(amount="1000000000000000")), [("NSFR-E016", "amount")])

    def test_sums_at_the_limit_do_not_overflow(self) -> None:
        rows = [position(record_id=f"r-{n}", amount="999999999999999.9999") for n in range(1000)]
        totals = cc.reconciliation(rows)
        self.assertEqual(totals["sides"]["asf"]["amount"], "999999999999999999.9000")
        self.assertEqual(len(cc.fingerprint(SCHEMA, RUN, rows)), 64)


class FieldTests(unittest.TestCase):
    def test_enumerations_reject_unknown_and_unsupported_values(self) -> None:
        self.assertEqual(codes(position(side="both")), [("NSFR-E002", "side")])
        self.assertEqual(codes(position(item_type="derivative")), [("NSFR-E012", "item_type")])
        self.assertEqual(codes(position(transaction_type="netted")), [("NSFR-E012", "transaction_type")])

    def test_dates_and_identifiers_are_validated(self) -> None:
        self.assertEqual(codes(position(maturity_date="2027-02-30")), [("NSFR-E002", "maturity_date")])
        self.assertEqual(codes(position(record_id="bad id")), [("NSFR-E002", "record_id")])

    def test_conditional_values_must_match_their_controller(self) -> None:
        self.assertEqual(codes(position(maturity_date="")), [("NSFR-E011", "maturity_date")])
        self.assertEqual(codes(position(maturity_kind="no_stated_maturity")),
                         [("NSFR-E011", "maturity_date")])
        self.assertEqual(codes(position(encumbered="true")), [("NSFR-E011", "encumbrance_end_date")])

    def test_cross_field_rules(self) -> None:
        self.assertEqual(codes(position(currency="USD")), [("NSFR-E007", "currency")])
        self.assertEqual(codes(position(maturity_date="2026-06-29")), [("NSFR-E008", "maturity_date")])
        self.assertEqual(codes(position(item_type="asset")), [("NSFR-E009", "item_type")])
        self.assertEqual(codes(position(classification="428r")), [("NSFR-E010", "classification")])
        self.assertEqual(codes(position(side="rsf", item_type="asset", classification="428l")),
                         [("NSFR-E010", "classification")])
        self.assertEqual(codes(position(classification="413")), [("NSFR-E002", "classification")])

    def test_duplicates_are_reported_on_each_repeat(self) -> None:
        self.assertEqual(codes(POSITION, POSITION, POSITION),
                         [("NSFR-E003", "record_id")])  # one diagnostic per (code, record, field)


class TableTests(unittest.TestCase):
    def test_run_must_have_one_row_and_exact_columns(self) -> None:
        found = cc.validate(SCHEMA, (RUN_FIELDS, [RUN, RUN]), (POSITION_FIELDS, [POSITION]))
        self.assertEqual([(d.code, d.record_id, d.field) for d in found], [("NSFR-E013", "run", "rows")])
        found = cc.validate(SCHEMA, (RUN_FIELDS, [RUN]), (POSITION_FIELDS[:-1], [POSITION]))
        self.assertIn(("NSFR-E013", "positions", "header"), [(d.code, d.record_id, d.field) for d in found])

    def test_rows_with_more_or_fewer_cells_are_reported(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "positions.csv"
            values = [POSITION[name] for name in POSITION_FIELDS]
            path.write_text("\n".join([",".join(POSITION_FIELDS), ",".join([*values, "EXTRA"]),
                                       ",".join(values[:-1]).replace("p-001", "p-002")]) + "\n",
                            encoding="utf-8")
            found = cc.validate(SCHEMA, (RUN_FIELDS, [RUN]), cc.read_table(path))
        self.assertIn(("NSFR-E013", "p-001", "cells"), [(d.code, d.record_id, d.field) for d in found])
        self.assertIn(("NSFR-E013", "p-002", "cells"), [(d.code, d.record_id, d.field) for d in found])

    def test_consolidated_basis_is_not_supported(self) -> None:
        self.assertEqual(codes(POSITION, run={**RUN, "basis": "consolidated"}), [("NSFR-E012", "basis")])

    def test_run_values_are_validated(self) -> None:
        self.assertEqual(codes(POSITION, run={**RUN, "extracted_at": "2026-07-01 08:00"}),
                         [("NSFR-E002", "extracted_at")])
        self.assertEqual(codes(POSITION, run={**RUN, "parameter_set_version": "0"}),
                         [("NSFR-E002", "parameter_set_version")])


class ReconciliationTests(unittest.TestCase):
    def test_totals_are_exact_and_grouped(self) -> None:
        rows = [position(record_id="a", amount="0.1"), position(record_id="b", amount="0.2"),
                position(record_id="c", side="rsf", item_type="asset", classification="428r",
                         amount="3")]
        result = cc.reconciliation(rows)
        self.assertEqual(result["sides"]["asf"], {"count": 2, "amount": "0.3000"})
        self.assertEqual(result["record_count"], 3)
        self.assertEqual(result["classifications"][1],
                         {"side": "rsf", "classification": "428r", "count": 1, "amount": "3.0000"})

    def test_fingerprint_ignores_row_order_and_trailing_zeros(self) -> None:
        first = [position(record_id="a", amount="1.5"), position(record_id="b")]
        second = [position(record_id="b"), position(record_id="a", amount="1.5000")]
        self.assertEqual(cc.fingerprint(SCHEMA, RUN, first), cc.fingerprint(SCHEMA, RUN, second))
        changed = [position(record_id="a", amount="1.6"), position(record_id="b")]
        self.assertNotEqual(cc.fingerprint(SCHEMA, RUN, first), cc.fingerprint(SCHEMA, RUN, changed))


class RepositoryTests(unittest.TestCase):
    def test_the_repository_contract_and_examples_pass(self) -> None:
        self.assertEqual(cc.check(ROOT), [])

    def copy(self, directory: str) -> Path:
        target = Path(directory)
        shutil.copytree(ROOT / cc.CONTRACTS, target / cc.CONTRACTS)
        return target

    def test_an_undocumented_field_is_reported(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = self.copy(directory)
            document = root / cc.DOCUMENT
            document.write_text(document.read_text("utf-8").replace("| `source_record` |", "| source |"),
                                "utf-8")
            self.assertTrue(any("`source_record`" in problem for problem in cc.check(root)))

    def test_a_changed_expectation_is_reported(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = self.copy(directory)
            expected = root / cc.EXAMPLES / "valid-basic" / "expected.json"
            data: dict[str, Any] = json.loads(expected.read_text("utf-8"))
            data["reconciliation"]["record_count"] = 4
            expected.write_text(json.dumps(data), "utf-8")
            self.assertTrue(any("valid-basic" in problem for problem in cc.check(root)))

    def test_main_reports_success(self) -> None:
        output = io.StringIO()
        with contextlib.redirect_stdout(output):
            self.assertEqual(cc.main([]), 0)
        self.assertIn("0 problem(s)", output.getvalue())


if __name__ == "__main__":
    unittest.main()
