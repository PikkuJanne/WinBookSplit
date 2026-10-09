"""M3 outcome contract and actual engine failure boundaries, synthetic inputs only."""

import importlib.util
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

from pypdf import PdfWriter


ROOT = Path(__file__).resolve().parents[2]
ENGINE = ROOT / "engine/winbooksplit_engine.py"
SPEC = importlib.util.spec_from_file_location("wbs_outcome_test_engine", ENGINE)
engine = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(engine)


class OutcomeTests(unittest.TestCase):
    def setUp(self):
        self.owned = tempfile.TemporaryDirectory(prefix="wbs-outcome-unit-")
        self.addCleanup(self.owned.cleanup)
        self.work = Path(self.owned.name).resolve()
        self.source = self.work / "synthetic-flat.pdf"
        writer = PdfWriter()
        for _ in range(3):
            writer.add_blank_page(width=100, height=100)
        with self.source.open("xb") as stream:
            writer.write(stream)
        self.before = self.source.read_bytes()
        self.base = self.work / "output"
        self.base.mkdir()
        self.neighbor = self.base / "synthetic-neighbor.txt"
        self.neighbor.write_bytes(b"Preserve this unrelated synthetic neighbor\n")

    def assert_preserved(self):
        self.assertEqual(self.source.read_bytes(), self.before)
        self.assertEqual(self.neighbor.read_bytes(), b"Preserve this unrelated synthetic neighbor\n")

    def assert_failed(self, result, code, exit_code, status=None):
        self.assertEqual(result["code"], code)
        self.assertEqual(result["exit_code"], exit_code)
        self.assertNotEqual(result["status"], "success")
        if status:
            self.assertEqual(result["status"], status)
        self.assertEqual(result["written_count"], 0)
        self.assertIsNone(result["execution"])
        self.assert_preserved()

    def test_ac050_documented_categories_are_pinned_independently_of_the_shipped_map(self):
        expected = {
            0: ("split_complete", "preview_complete"),
            2: ("invalid_arguments", "invalid_mode", "invalid_start_pages", "input_invalid",
                "output_base_invalid", "output_path_too_long", "output_destination_changed", "source_output_alias"),
            3: ("dependency_missing", "runtime_invalid", "runtime_not_found", "converter_invalid",
                "converter_not_found", "processor_start_failed"),
            4: ("conversion_start_failed", "conversion_failed", "conversion_output_invalid",
                "conversion_cleanup_failed", "conversion_source_changed", "conversion_ownership_failed"),
            5: ("no_bookmarks", "no_usable_bookmarks", "no_bookmarks_at_level"),
            6: ("invalid_document", "invalid_outline", "unreadable_document", "invalid_plan",
                "invalid_prepared_split", "source_changed", "output_exists", "invalid_execution",
                "output_write_failed", "output_validation_failed", "output_ownership_failed",
                "output_manifest_failed", "output_publish_failed", "output_handle_close_failed",
                "processor_protocol_failed", "console_setup_failed", "console_finalize_failed", "processor_failed"),
            7: ("unsupported_document",),
            130: ("output_cancelled", "conversion_cancelled", "conversion_timeout", "processing_cancelled",
                  "processor_cancelled", "processor_timeout", "cancelled"),
        }
        self.assertEqual(set(engine.EXIT_CODES), {code for codes in expected.values() for code in codes})
        for exit_code, codes in expected.items():
            for code in codes:
                with self.subTest(code=code):
                    self.assertEqual(engine.result_exit_code(code), exit_code)
        with self.assertRaises(TypeError):
            engine.EXIT_CODES["invalid_arguments"] = 0

    def test_unknown_diagnostic_code_cannot_acquire_a_success_or_ad_hoc_exit(self):
        for code in ("unknown_outcome", "Split_Complete", None):
            with self.subTest(code=code), self.assertRaises(ValueError):
                engine.result_exit_code(code)

    def test_ac050_actual_cli_native_and_result_codes_agree_for_no_plan_and_bad_manual(self):
        cases = (("1", "", "no_bookmarks", 5), ("manual", "1,bad", "invalid_start_pages", 2))
        for mode, data, code, expected in cases:
            with self.subTest(code=code):
                child = subprocess.run([sys.executable, "-I", "-B", "-X", "utf8", str(ENGINE),
                    str(self.source), str(self.base), mode, data], cwd=self.work,
                    capture_output=True, text=True, encoding="utf-8", timeout=20)
                self.assertEqual(child.returncode, expected, child.stdout + child.stderr)
                records = [json.loads(line) for line in child.stdout.splitlines() if line.startswith("{")]
                self.assertEqual(len(records), 1)
                self.assert_failed(records[0], code, expected)
                self.assertNotIn("[Writing]", child.stdout)
        self.assertEqual({path.name for path in self.base.iterdir()}, {self.neighbor.name})

    def test_ac050_actual_cli_argument_failure_emits_one_code2_record(self):
        child = subprocess.run([sys.executable, "-I", "-B", str(ENGINE)], cwd=self.work,
            capture_output=True, text=True, encoding="utf-8", timeout=20)
        self.assertEqual(child.returncode, 2)
        records = [json.loads(line) for line in child.stdout.splitlines() if line.startswith("{")]
        self.assertEqual(len(records), 1)
        self.assert_failed(records[0], "invalid_arguments", 2)

    def test_ac050_actual_direct_engine_missing_dependency_is_code3(self):
        # -S intentionally excludes this developer venv's pypdf; this is an
        # actual import rejection, not a claim about ordinary runtime discovery.
        child = subprocess.run([sys.executable, "-I", "-B", "-S", str(ENGINE),
            str(self.source), str(self.base), "manual", "1"], cwd=self.work,
            capture_output=True, text=True, encoding="utf-8", timeout=20)
        self.assertEqual(child.returncode, 3, child.stdout + child.stderr)
        records = [json.loads(line) for line in child.stdout.splitlines() if line.startswith("{")]
        self.assertEqual(len(records), 1)
        self.assert_failed(records[0], "dependency_missing", 3)
        self.assertNotIn("[Writing]", child.stdout)

    def test_ac050_invalid_destination_is_validation_code2_before_output(self):
        absent = self.work / "never-created-base"
        with patch.object(engine, "write_slice") as writer:
            result = engine.run_split(self.source, absent, "manual", "1")
        self.assert_failed(result, "output_base_invalid", 2)
        writer.assert_not_called()
        self.assertFalse(absent.exists())

    def test_ac051_planning_interrupt_returns_cancelled_without_output(self):
        with patch.object(engine, "prepare_split", side_effect=KeyboardInterrupt), \
                patch.object(engine, "write_slice") as writer:
            result = engine.run_split(self.source, self.base, "manual", "1")
        self.assert_failed(result, "processing_cancelled", 130, "cancelled")
        writer.assert_not_called()
        self.assertEqual({path.name for path in self.base.iterdir()}, {self.neighbor.name})

    def test_ac051_interrupt_after_planning_before_execution_returns_cancelled(self):
        with patch.object(engine, "log", side_effect=KeyboardInterrupt), \
                patch.object(engine, "execute_split") as writer:
            result = engine.run_split(self.source, self.base, "manual", "1")
        self.assert_failed(result, "processing_cancelled", 130, "cancelled")
        writer.assert_not_called()

    def test_ac050_converter_failures_and_ac051_cancel_timeout_keep_distinct_categories(self):
        ebook = self.work / "synthetic.epub"
        ebook.write_bytes(b"Authored ebook boundary control\n")
        cases = (("converter_not_found", 3, "error"), ("conversion_failed", 4, "error"),
                 ("conversion_output_invalid", 4, "error"), ("conversion_cancelled", 130, "cancelled"),
                 ("conversion_timeout", 130, "timeout"))
        for code, exit_code, status in cases:
            with self.subTest(code=code), \
                    patch.object(engine, "prepare_ebook", side_effect=engine.OutputError(code, "Authored conversion boundary failure")), \
                    patch.object(engine, "write_slice") as writer:
                result = engine.run_split(ebook, self.base, "manual", "1")
                self.assert_failed(result, code, exit_code, status)
                writer.assert_not_called()
        self.assertEqual(ebook.read_bytes(), b"Authored ebook boundary control\n")

    def test_ac052_cleanup_failure_preserves_primary_conversion_cancel_and_timeout(self):
        for primary, expected, status in (("conversion_cancelled", 130, "cancelled"),
                                         ("conversion_timeout", 130, "timeout"),
                                         ("conversion_failed", 4, "error")):
            with self.subTest(primary=primary):
                error = engine.OutputError("conversion_cleanup_failed", "Retained synthetic conversion stage",
                    {"primary_code": primary, "cleanup_complete": False, "retained_staging": str(self.work)})
                with patch.object(engine, "prepare_ebook", side_effect=error), patch.object(engine, "write_slice") as writer:
                    result = engine.run_split(self.work / "synthetic.epub", self.base, "manual", "1")
                self.assert_failed(result, "conversion_cleanup_failed", expected, status)
                self.assertEqual(result["diagnostic"]["primary_code"], primary)
                self.assertFalse(result["diagnostic"]["cleanup_complete"])
                writer.assert_not_called()

    def test_ac052_real_partial_cancel_and_failed_cleanup_retains_only_owned_stage(self):
        original = engine.PdfWriter.write
        calls = 0
        def cancel_with_extra_member(writer, stream):
            nonlocal calls
            calls += 1
            if calls == 2:
                stream.write(b"Authored partial output\n")
                Path(stream.name).parent.joinpath("unknown-authored-member.txt").write_bytes(b"Preserve this unregistered fixture member")
                raise KeyboardInterrupt
            return original(writer, stream)
        with patch.object(engine.PdfWriter, "write", autospec=True, side_effect=cancel_with_extra_member), patch.object(engine, "log"):
            result = engine.run_split(self.source, self.base, "manual", "1,2")
        self.assertEqual(calls, 2)
        self.assert_failed(result, "output_cancelled", 130, "cancelled")
        self.assertFalse(result["diagnostic"]["cleanup_complete"])
        retained = Path(result["diagnostic"]["retained_staging"])
        self.assertEqual(retained.parent, self.base)
        self.assertTrue(retained.name.startswith(".WinBookSplit-stage-"))
        self.assertEqual(retained.joinpath("unknown-authored-member.txt").read_bytes(), b"Preserve this unregistered fixture member")
        self.assertFalse(any(path.is_dir() and not path.name.startswith(".WinBookSplit-") for path in self.base.iterdir()))


if __name__ == "__main__":
    unittest.main(verbosity=2)
