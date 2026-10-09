"""Strict manual-start acceptance, using original synthetic PDFs only."""

from hashlib import sha256
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from pypdf import PdfWriter


ROOT = Path(__file__).resolve().parents[2]


def load_module(name, path):
    spec = importlib.util.spec_from_file_location(name, path.resolve())
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


engine = load_module("wbs_manual_test_engine", ROOT / "engine/winbooksplit_engine.py")
fixtures = load_module("wbs_manual_test_fixtures", ROOT / "tests/fixtures/generate_pdf_fixtures.py")
MANUAL_ORACLES = json.loads(
    (ROOT / "docs/codex-v1.0.0/PLAN_ORACLES.json").read_text(encoding="utf-8")
)["manual"]


class ManualPlanTests(unittest.TestCase):
    def assert_plan_error(self, data, pages, code="invalid_start_pages"):
        with self.assertRaises(engine.ManualPlanError) as raised:
            engine.plan_manual_starts(data, pages)
        self.assertEqual(raised.exception.code, code)
        self.assertTrue(str(raised.exception), "Rejected input needs a diagnostic")

    def test_all_22_predeclared_manual_oracles(self):
        self.assertEqual([case["id"] for case in MANUAL_ORACLES],
                         [f"MAN-{number:02d}" for number in range(1, 23)])
        for case in MANUAL_ORACLES:
            with self.subTest(oracle=case["id"]):
                if "expected_error" in case:
                    self.assert_plan_error(case["input"], case["pages"], case["expected_error"])
                else:
                    plan = engine.plan_manual_starts(case["input"], case["pages"])
                    self.assertEqual(plan["ranges"], [tuple(item) for item in case["expected_ranges"]])
                    self.assertEqual(plan["starts"], [item[0] + 1 for item in case["expected_ranges"]])

    def test_ac013_one_section_has_explicit_notice(self):
        plan = engine.plan_manual_starts("1", 10)
        self.assertEqual(plan["ranges"], [(0, 10)])
        self.assertRegex(" ".join(plan["notices"]).lower(), r"one section.*no internal split")

    def test_ac014_implicit_first_page_notice(self):
        implicit = engine.plan_manual_starts("4,7", 10)
        explicit = engine.plan_manual_starts("1,4,7", 10)
        self.assertEqual(implicit["ranges"], [(0, 3), (3, 6), (6, 10)])
        self.assertEqual(implicit["ranges"], explicit["ranges"])
        self.assertRegex(" ".join(implicit["notices"]).lower(), r"page 1.*added")
        self.assertEqual(explicit["notices"], [])

    def test_ac015_sort_and_duplicate_notices(self):
        plan = engine.plan_manual_starts("7, 04,4, 1", 10)
        self.assertEqual(plan["starts"], [1, 4, 7])
        self.assertEqual(plan["ranges"], [(0, 3), (3, 6), (6, 10)])
        notices = " ".join(plan["notices"]).lower()
        self.assertIn("sort", notices)
        self.assertIn("duplicate", notices)

    def test_surrounding_whitespace_and_leading_zeros_are_accepted(self):
        plan = engine.plan_manual_starts(" \t001\n,\t04 , 007\r\n", 10)
        self.assertEqual(plan["starts"], [1, 4, 7])
        self.assertEqual(plan["ranges"], [(0, 3), (3, 6), (6, 10)])

    def test_valid_long_leading_zero_token_is_independent_of_integer_digit_limit(self):
        plan = engine.plan_manual_starts("0" * 5000 + "4,7", 10)
        self.assertEqual(plan["starts"], [1, 4, 7])
        self.assertEqual(plan["ranges"], [(0, 3), (3, 6), (6, 10)])

    def test_invalid_long_and_mixed_tokens_are_rejected_entirely(self):
        for data in ("9" * 5000, "0" * 5000, "1," + "9" * 5000 + ",7",
                     "1," + "0" * 5000 + "x,7", "1,4 7", "1,4\n7"):
            with self.subTest(input_kind=(data[:20], len(data))):
                self.assert_plan_error(data, 10)

    def test_comma_zero_and_non_ascii_glyphs_are_rejected(self):
        for data in (",", ",,", "1, ,4", "0", "000", "１,４", "1,４,7", "١,٤", "1,٤,7"):
            with self.subTest(data=data):
                self.assert_plan_error(data, 10)

    def test_missing_or_non_string_data_is_rejected(self):
        for data in (None, 1, [], [1, 4, 7]):
            with self.subTest(data=data):
                self.assert_plan_error(data, 10)

    def test_zero_or_invalid_document_page_count_is_rejected(self):
        for pages in (0, -1, None, "10", 10.0, True):
            with self.subTest(pages=pages):
                self.assert_plan_error("1", pages, "invalid_document")


class ManualExecutionRejectionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.owned = tempfile.TemporaryDirectory(prefix="wbs-manual-unit-")
        cls.addClassCleanup(cls.owned.cleanup)
        cls.work = Path(cls.owned.name).resolve()
        cls.input_path = fixtures.generate_fixtures(cls.work / "fixtures")["simple10"]
        cls.zero_path = cls.work / "zero-pages.pdf"
        with cls.zero_path.open("xb") as stream:
            PdfWriter().write(stream)

    def test_invalid_manual_requests_never_reach_writer_or_change_input_and_neighbors(self):
        output = self.work / "rejected-output"
        output.mkdir()
        neighbor = output / "owned-neighbor.txt"
        neighbor.write_bytes(b"Synthetic neighbor must remain unchanged\n")
        before_input = sha256(self.input_path.read_bytes()).hexdigest()
        before_neighbor = neighbor.read_bytes()
        invalid = [case["input"] for case in MANUAL_ORACLES
                   if case.get("expected_error") == "invalid_start_pages"]
        invalid += [None, ",", "0", "１,４", "1,４,7", "9" * 5000,
                    "1," + "9" * 5000 + ",7"]
        for data in invalid:
            with self.subTest(input_kind=(str(data)[:20], len(str(data)))):
                with patch.object(engine, "write_slice") as writer, patch.object(engine, "log") as log:
                    with self.assertRaises(SystemExit) as raised:
                        engine.split_pdf(str(self.input_path), str(output), "manual", data)
                self.assertNotEqual(raised.exception.code, 0)
                writer.assert_not_called()
                self.assertIn("invalid_start_pages", " ".join(str(call.args[0]) for call in log.call_args_list))
                self.assertEqual({path.name for path in output.iterdir()}, {neighbor.name})
                self.assertEqual(neighbor.read_bytes(), before_neighbor)
                self.assertEqual(sha256(self.input_path.read_bytes()).hexdigest(), before_input)

    def test_zero_page_document_never_reaches_writer(self):
        output = self.work / "rejected-zero-output"
        output.mkdir()
        before = self.zero_path.read_bytes()
        with patch.object(engine, "write_slice") as writer, patch.object(engine, "log") as log:
            with self.assertRaises(SystemExit) as raised:
                engine.split_pdf(str(self.zero_path), str(output), "manual", "1")
        self.assertNotEqual(raised.exception.code, 0)
        writer.assert_not_called()
        self.assertIn("invalid_document", " ".join(str(call.args[0]) for call in log.call_args_list))
        self.assertEqual(list(output.iterdir()), [])
        self.assertEqual(self.zero_path.read_bytes(), before)


if __name__ == "__main__":
    unittest.main(verbosity=2)
