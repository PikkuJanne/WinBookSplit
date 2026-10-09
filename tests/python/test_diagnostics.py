"""Structured diagnostics and applicable choices, using synthetic PDFs only."""

import copy
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, TextStringObject


ROOT = Path(__file__).resolve().parents[2]


def load_module(name, path):
    spec = importlib.util.spec_from_file_location(name, path.resolve())
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


engine = load_module("wbs_diagnostics_test_engine", ROOT / "engine/winbooksplit_engine.py")
fixtures = load_module("wbs_diagnostics_test_fixtures", ROOT / "tests/fixtures/generate_pdf_fixtures.py")


class DiagnosticTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.owned = tempfile.TemporaryDirectory(prefix="wbs-diagnostics-unit-")
        cls.addClassCleanup(cls.owned.cleanup)
        cls.work = Path(cls.owned.name).resolve()
        cls.inputs = fixtures.generate_fixtures(cls.work / "fixtures")
        cls.flat = cls.work / "flat.pdf"
        cls.flat.write_bytes(fixtures._pdf_bytes({"pages": 10, "title": "Synthetic flat PDF", "outline": []}))
        cls.zero = cls.work / "zero.pdf"
        with cls.zero.open("xb") as stream:
            PdfWriter().write(stream)
        cls.empty = cls.work / "empty.pdf"
        cls.empty.write_bytes(b"")
        cls.corrupt = cls.work / "corrupt.pdf"
        cls.corrupt.write_bytes(b"%PDF-1.7\nOriginal synthetic corrupt bytes\n%%EOF\n")
        cls.truncated = cls.work / "truncated.pdf"
        cls.truncated.write_bytes(cls.inputs["simple10"].read_bytes()[:128])
        for name, with_valid in (("unusable", False), ("warning-success", True)):
            writer = PdfWriter()
            for page in PdfReader(cls.flat).pages:
                writer.add_page(page)
            invalid = writer.add_outline_item("Synthetic unresolvable named bookmark", 0).get_object()
            invalid["/A"][NameObject("/D")] = TextStringObject("Missing synthetic target")
            if with_valid:
                writer.add_outline_item("Valid A", 3)
            path = cls.work / (name + ".pdf")
            with path.open("xb") as stream:
                writer.write(stream)
            setattr(cls, name.replace("-", "_"), path)

    def assert_failure_shape(self, result, status, code, exit_code, fallbacks):
        self.assertEqual(result["protocol"], "winbooksplit.result")
        self.assertEqual(result["version"], 1)
        self.assertEqual(result["status"], status)
        self.assertEqual(result["code"], code)
        self.assertEqual(result["exit_code"], exit_code)
        self.assertEqual(result["fallback_modes"], fallbacks)
        self.assertEqual(result["written_count"], 0)
        self.assertIsNone(result["execution"])
        self.assertTrue(result["message"])
        self.assertIsInstance(result["warnings"], tuple)

    def rejected_run(self, name, source, mode, manual_data=None):
        output = self.work / name
        output.mkdir()
        neighbor = output / "owned-neighbor.txt"
        neighbor.write_bytes(b"Original synthetic neighbor preserved\n")
        before = source.read_bytes() if source.exists() else None
        with patch.object(engine, "write_slice") as writer:
            result = engine.run_split(str(source), str(output), mode, manual_data)
        writer.assert_not_called()
        self.assertEqual({path.name for path in output.iterdir()}, {neighbor.name})
        self.assertEqual(neighbor.read_bytes(), b"Original synthetic neighbor preserved\n")
        if before is not None:
            self.assertEqual(source.read_bytes(), before)
        return result

    def test_ac030_valid_flat_zero_page_and_unreadable_inputs_have_distinct_categories(self):
        flat = self.rejected_run("flat-result", self.flat, "1")
        self.assert_failure_shape(flat, "no_plan", "no_bookmarks", 5, ("manual",))
        zero = self.rejected_run("zero-result", self.zero, "1")
        self.assert_failure_shape(zero, "invalid_input", "invalid_document", 6, ())
        for name, source in (("empty", self.empty), ("corrupt", self.corrupt), ("truncated", self.truncated)):
            with self.subTest(input=name):
                result = self.rejected_run(name + "-result", source, "1")
                self.assert_failure_shape(result, "read_error", "unreadable_document", 6, ())

    def test_ac031_fallback_modes_depend_on_the_requested_level_and_usable_parents(self):
        cases = [("flat-level2", self.flat, "2", "no_bookmarks", ("manual",)),
                 ("unusable-level1", self.unusable, "1", "no_usable_bookmarks", ("manual",)),
                 ("unusable-level2", self.unusable, "2", "no_usable_bookmarks", ("manual",)),
                 ("no-level2", self.inputs["simple10"], "2", "no_bookmarks_at_level", ("1", "manual"))]
        for name, source, mode, code, fallbacks in cases:
            with self.subTest(case=name):
                result = self.rejected_run(name, source, mode)
                self.assert_failure_shape(result, "no_plan", code, 5, fallbacks)
                self.assertEqual(result["mode"], mode)
                if code == "no_bookmarks_at_level":
                    self.assertIn("Level 2", result["message"])
                if code == "no_usable_bookmarks":
                    self.assertTrue(result["warnings"])

    def test_invalid_manual_input_or_mode_offers_no_fallback_and_does_not_probe_for_invalid_mode(self):
        manual = self.rejected_run("bad-manual", self.flat, "manual", "1,abc,7")
        self.assert_failure_shape(manual, "invalid_input", "invalid_start_pages", 2, ())
        for mode in ("unsupported", None, []):
            with self.subTest(mode=mode):
                with patch.object(engine, "PdfReader") as probe, patch.object(engine, "write_slice") as writer:
                    result = engine.run_split("unused-synthetic-source.pdf", "unused-synthetic-output", mode)
                self.assert_failure_shape(result, "invalid_input", "invalid_mode", 2, ())
                probe.assert_not_called()
                writer.assert_not_called()

    def test_missing_file_is_path_validation_error_without_a_manual_or_level1_offer(self):
        missing = self.work / "never-created-source.pdf"
        result = self.rejected_run("missing-file-result", missing, "1")
        self.assert_failure_shape(result, "invalid_input", "input_invalid", 2, ())
        self.assertFalse(missing.exists())

    def test_malformed_outline_has_diagnostics_and_no_fallback_before_writer(self):
        node = {"/Title": "Synthetic cyclic outline"}
        node["/Next"] = node

        class Reader:
            pages = [None] * 10
            root_object = {"/Outlines": {"/First": node}}

            @property
            def outline(self):
                raise AssertionError("Raw cycle must be rejected before pypdf outline retrieval")

        with patch.object(engine, "PdfReader", return_value=Reader()), patch.object(engine, "write_slice") as writer:
            result = engine.run_split("unused-synthetic-source.pdf", "unused-synthetic-output", "1")
        self.assert_failure_shape(result, "invalid_input", "invalid_outline", 6, ())
        self.assertTrue(result["warnings"])
        writer.assert_not_called()

    def test_invalid_execution_plan_cannot_be_a_success_or_offer_fallback(self):
        reader = PdfReader(self.inputs["simple10"])
        malformed = copy.deepcopy(engine.plan_level1(reader, 10))
        malformed["entries"] = []
        malformed["ranges"] = []
        output = self.work / "invalid-plan-result"
        output.mkdir()
        with patch.object(engine, "plan_level1", return_value=malformed), patch.object(engine, "write_slice") as writer:
            result = engine.run_split(str(self.inputs["simple10"]), str(output), "1")
        self.assert_failure_shape(result, "error", "invalid_plan", 6, ())
        writer.assert_not_called()
        self.assertEqual(list(output.iterdir()), [])

    def test_first_writer_failure_returns_write_error_without_success_or_fallback(self):
        output = self.work / "failed-writer-result"
        output.mkdir()
        before = self.inputs["simple10"].read_bytes()
        with patch.object(engine, "write_slice", side_effect=OSError("Synthetic first slice write failed")) as writer:
            result = engine.run_split(str(self.inputs["simple10"]), str(output), "manual", "1")
        self.assert_failure_shape(result, "write_error", "output_write_failed", 6, ())
        self.assertEqual(writer.call_count, 1)
        self.assertIn("Synthetic first slice write failed", result["message"])
        diagnostic = result["diagnostic"]
        self.assertTrue(diagnostic["cleanup_complete"])
        self.assertIsNone(diagnostic["retained_staging"])
        self.assertTrue(Path(diagnostic["record_path"]).is_file())
        self.assertEqual(list(output.glob("*.pdf")), [])
        self.assertEqual(self.inputs["simple10"].read_bytes(), before)

    def test_zero_section_execution_result_cannot_be_reported_as_success(self):
        output = self.work / "empty-execution-result"
        output.mkdir()
        before = self.inputs["simple10"].read_bytes()
        with patch.object(engine, "execute_split", return_value={"written_count": 0, "outputs": ()}), \
                patch.object(engine, "write_slice") as writer:
            result = engine.run_split(str(self.inputs["simple10"]), str(output), "manual", "1")
        self.assert_failure_shape(result, "error", "invalid_execution", 6, ())
        writer.assert_not_called()
        self.assertEqual(list(output.iterdir()), [])
        self.assertEqual(self.inputs["simple10"].read_bytes(), before)

    def test_valid_manual_execution_on_flat_pdf_is_explicit_success_with_real_output_identity(self):
        output = self.work / "flat-manual-success"
        output.mkdir()
        result = engine.run_split(str(self.flat), str(output), "manual", "1")
        self.assertEqual(result["status"], "success")
        self.assertEqual(result["exit_code"], 0)
        self.assertEqual(result["fallback_modes"], ())
        self.assertEqual(result["written_count"], 1)
        execution = result["execution"]
        self.assertEqual(execution["coverage"], {"complete": True, "covered_pages": 10, "section_count": 1})
        self.assertEqual(fixtures.page_ids(Path(execution["final_directory"]) / execution["outputs"][0]["filename"]),
                         list(range(1, 11)))

    def test_recoverable_invalid_bookmark_warning_survives_success_without_losing_pages(self):
        output = self.work / "warning-success-result"
        output.mkdir()
        before = self.warning_success.read_bytes()
        result = engine.run_split(str(self.warning_success), str(output), "1")
        self.assertEqual(result["status"], "success")
        self.assertEqual(result["written_count"], 2)
        self.assertTrue(any(warning["code"] == "invalid_destination" for warning in result["warnings"]))
        self.assertEqual([page for entry in result["execution"]["outputs"]
                          for page in fixtures.page_ids(Path(result["execution"]["final_directory"]) / entry["filename"])],
                         list(range(1, 11)))
        self.assertEqual(self.warning_success.read_bytes(), before)

    def test_results_and_nested_warning_metadata_are_immutable_and_plain_json_serializable(self):
        result = self.rejected_run("serialize-result", self.unusable, "1")
        with self.assertRaises(TypeError):
            result["status"] = "success"
        with self.assertRaises(TypeError):
            result["warnings"][0]["message"] = "Changed diagnostic"
        serialized = engine.serialize_result(result)
        self.assertNotIn("\n", serialized)
        decoded = json.loads(serialized)
        self.assertEqual(decoded["protocol"], "winbooksplit.result")
        self.assertEqual(decoded["code"], "no_usable_bookmarks")
        self.assertEqual(decoded["fallback_modes"], ["manual"])
        self.assertEqual(decoded["warnings"][0]["code"], result["warnings"][0]["code"])

    def test_legacy_split_emits_a_single_structured_result_without_the_old_sentinel(self):
        output = self.work / "legacy-no-level2-result"
        output.mkdir()
        with patch.object(engine, "write_slice") as writer, patch.object(engine, "log") as log:
            with self.assertRaises(SystemExit) as raised:
                engine.split_pdf(str(self.inputs["simple10"]), str(output), "2")
        self.assertEqual(raised.exception.code, 5)
        writer.assert_not_called()
        messages = [str(call.args[0]) for call in log.call_args_list]
        self.assertTrue(all("[NO_BOOKMARKS_FOUND]" not in message for message in messages))
        payloads = [json.loads(message) for message in messages if message.startswith("{")]
        self.assertEqual(len(payloads), 1)
        self.assertEqual(payloads[0]["code"], "no_bookmarks_at_level")
        self.assertEqual(payloads[0]["fallback_modes"], ["1", "manual"])
        self.assertEqual(list(output.iterdir()), [])


if __name__ == "__main__":
    unittest.main(verbosity=2)
