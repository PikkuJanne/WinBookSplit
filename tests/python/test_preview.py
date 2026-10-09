"""Plan-only engine CLI and writer traps; synthetic inputs exclusively."""

from hashlib import sha256
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


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


engine = load("wbs_preview_unit_engine", ENGINE)
fixtures = load("wbs_preview_unit_fixtures", ROOT / "tests/fixtures/generate_pdf_fixtures.py")


class PreviewTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.owned = tempfile.TemporaryDirectory(prefix="wbs-preview-unit-")
        cls.addClassCleanup(cls.owned.cleanup)
        cls.work = Path(cls.owned.name).resolve()
        cls.inputs = fixtures.generate_fixtures(cls.work / "fixtures")

    def setUp(self):
        self.case = Path(tempfile.mkdtemp(prefix="case-", dir=self.work))
        self.base = self.case / "output"
        self.base.mkdir()
        self.neighbor = self.base / "synthetic-neighbor.txt"
        self.neighbor.write_bytes(b"Preserve this unrelated synthetic neighbor\n")
        self.flat = self.case / "flat.pdf"
        writer = PdfWriter()
        for _ in range(3):
            writer.add_blank_page(width=100, height=100)
        with self.flat.open("xb") as stream:
            writer.write(stream)

    def assert_unchanged(self, source, before):
        self.assertEqual(source.read_bytes(), before)
        self.assertEqual(self.neighbor.read_bytes(), b"Preserve this unrelated synthetic neighbor\n")
        self.assertEqual({path.name for path in self.base.iterdir()}, {self.neighbor.name})

    def assert_preview(self, result, source, before, pages):
        self.assertEqual(result["protocol"], "winbooksplit.result")
        self.assertEqual(result["version"], 1)
        self.assertEqual(result["status"], "preview")
        self.assertEqual(result["code"], "preview_complete")
        self.assertEqual(result["exit_code"], 0)
        self.assertEqual(result["written_count"], 0)
        self.assertIsNone(result["execution"])
        self.assertEqual(result["fallback_modes"], ())
        plan = result["plan"]
        self.assertEqual(plan["total_pages"], pages)
        self.assertEqual(plan["coverage"], {"complete": True, "covered_pages": pages,
                                            "section_count": len(plan["entries"])})
        self.assertEqual(tuple((entry["start"], entry["end"]) for entry in plan["entries"]), plan["ranges"])
        self.assertEqual([page for start, end in plan["ranges"] for page in range(start, end)], list(range(pages)))
        self.assertEqual(plan["source_identity"]["sha256"], sha256(before).hexdigest())
        self.assertEqual(plan["output_naming"]["resolved_base"], str(self.base.resolve()))
        self.assertEqual(result["warnings"], plan["warnings"])
        self.assert_unchanged(source, before)

    def test_ac056_every_pdf_mode_returns_its_same_prepared_plan_without_output_or_replanning(self):
        for mode, data, fixture, pages in (("manual", "4,7", "simple10", 10),
                                           ("1", None, "simple10", 10), ("2", None, "nested12", 12)):
            with self.subTest(mode=mode):
                source = self.inputs[fixture]
                before = source.read_bytes()
                prepared = engine.prepare_split(source, mode, data, output_base=self.base)
                with patch.object(engine, "prepare_split", return_value=prepared) as prepare, \
                        patch.object(engine, "preview_plan", wraps=engine.preview_plan) as preview, \
                        patch.object(engine, "PdfReader", side_effect=AssertionError("Preview cannot reopen source")), \
                        patch.object(engine, "execute_split", side_effect=AssertionError("Preview cannot execute")) as execute, \
                        patch.object(engine, "write_slice", side_effect=AssertionError("Preview cannot write")) as write, \
                        patch.object(engine, "OutputRun", side_effect=AssertionError("PDF preview cannot reserve output")) as reserve:
                    result = engine.run_split(source, self.base, mode, data, preview=True)
                prepare.assert_called_once_with(source, mode, data, output_base=self.base)
                preview.assert_called_once_with(prepared)
                execute.assert_not_called(); write.assert_not_called(); reserve.assert_not_called()
                self.assert_preview(result, source, before, pages)
                self.assertEqual(result["plan"], prepared.plan)
                with self.assertRaises(TypeError):
                    result["plan"]["entries"][0]["start"] = 1

    def test_ac056_pdf_preview_prepares_an_actual_source_and_matches_manual_normalization(self):
        before = self.flat.read_bytes()
        with patch.object(engine, "execute_split") as execute, patch.object(engine, "write_slice") as write:
            result = engine.run_split(self.flat, self.base, "manual", "3,3", preview=True)
        execute.assert_not_called(); write.assert_not_called()
        self.assert_preview(result, self.flat, before, 3)
        self.assertEqual(result["plan"]["normalized_inputs"]["starts"], (1, 3))
        self.assertIn("Duplicate start pages removed.", result["plan"]["notices"])
        self.assertIn("Physical page 1 added to preserve opening pages.", result["plan"]["notices"])

    def test_ac056_ebook_preview_reuses_conversion_preparation_and_preserves_original_generated_identities(self):
        # This authored boundary tests reuse; real Calibre/cleanup is a separate
        # required application acceptance route, not claimed by this unit.
        source = self.case / "synthetic.epub"
        source.write_bytes(b"Synthetic original ebook conversion boundary\n")
        original = {"path": str(source), "resolved_path": str(source.resolve()), "binding": "ebook_snapshot",
                    "sha256": sha256(source.read_bytes()).hexdigest(), "size_bytes": source.stat().st_size}
        data = self.flat.read_bytes()
        generated_path = self.case / "removed-owned-conversion.pdf"
        generated = {"path": str(generated_path), "resolved_path": str(generated_path), "binding": "reader_snapshot",
                     "sha256": sha256(data).hexdigest(), "size_bytes": len(data), "page_count": 3}
        conversion = {"generated_pdf_identity": generated, "original_source_identity": original,
                      "workspace_cleanup": {"cleanup_complete": True, "retained_staging": None}}
        prepared = engine.prepare_split(generated_path, "manual", "1,3", output_base=self.base, _pdf_bytes=data,
            _conversion_metadata={"original_ebook_identity": original, "conversion": conversion, "keep_converted_pdf": False})
        before = source.read_bytes()
        with patch.object(engine, "prepare_ebook", return_value=prepared) as prepare, \
                patch.object(engine, "execute_split", side_effect=AssertionError("Ebook preview cannot write")) as execute:
            result = engine.run_split(source, self.base, "manual", "1,3", calibre_path="owned-calibre.exe", preview=True)
        prepare.assert_called_once_with(source, "manual", "1,3", output_base=self.base,
            calibre_path="owned-calibre.exe", keep_converted_pdf=False, conversion_timeout=1800)
        execute.assert_not_called()
        self.assertEqual(result["status"], "preview")
        self.assertEqual(result["plan"]["original_ebook_identity"], original)
        self.assertEqual(result["plan"]["conversion"]["generated_pdf_identity"], generated)
        self.assertTrue(result["plan"]["conversion"]["workspace_cleanup"]["cleanup_complete"])
        self.assertFalse(generated_path.exists())
        self.assert_unchanged(source, before)

    def test_ac056_preview_failures_keep_existing_no_plan_invalid_manual_and_cancellation_categories(self):
        for mode, data, code, exit_code in (("1", None, "no_bookmarks", 5),
                                           ("manual", "1,bad", "invalid_start_pages", 2)):
            with self.subTest(code=code), patch.object(engine, "execute_split") as execute:
                result = engine.run_split(self.flat, self.base, mode, data, preview=True)
            self.assertEqual((result["code"], result["exit_code"]), (code, exit_code))
            self.assertNotIn("plan", result)
            execute.assert_not_called()
        with patch.object(engine, "prepare_split", side_effect=KeyboardInterrupt), patch.object(engine, "execute_split") as execute:
            result = engine.run_split(self.flat, self.base, "manual", "1", preview=True)
        self.assertEqual((result["status"], result["exit_code"]), ("cancelled", 130))
        execute.assert_not_called()

    def test_ac054_preview_rejects_retention_or_non_boolean_before_conversion_or_writes(self):
        for preview, keep in ((True, True), ("true", False), (1, False)):
            with self.subTest(preview=preview, keep=keep), \
                    patch.object(engine, "prepare_ebook") as prepare, patch.object(engine, "execute_split") as execute:
                result = engine.run_split(self.case / "unused.epub", self.base, "manual", "1",
                    preview=preview, keep_converted_pdf=keep)
            self.assertEqual((result["code"], result["exit_code"]), ("invalid_arguments", 2))
            prepare.assert_not_called(); execute.assert_not_called()

    def test_ac056_preview_rejects_absent_destination_without_creating_it(self):
        absent = self.case / "never-created-output"
        with patch.object(engine, "execute_split") as execute:
            result = engine.run_split(self.flat, absent, "manual", "1", preview=True)
        self.assertEqual((result["code"], result["exit_code"]), ("output_base_invalid", 2))
        self.assertFalse(absent.exists())
        execute.assert_not_called()

    def test_ac056_actual_engine_preview_cli_emits_one_plan_and_never_an_error_or_writer_record(self):
        before = self.flat.read_bytes()
        child = subprocess.run([sys.executable, "-I", "-B", "-X", "utf8", str(ENGINE), str(self.flat),
            str(self.base), "manual", "1,3", "--preview"], cwd=self.case, capture_output=True,
            text=True, encoding="utf-8", timeout=20)
        self.assertEqual(child.returncode, 0, child.stdout + child.stderr)
        records = [json.loads(line) for line in child.stdout.splitlines() if line.startswith("{")]
        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]["status"], "preview")
        self.assertEqual(records[0]["plan"]["ranges"], [[0, 2], [2, 3]])
        self.assertEqual(records[0]["written_count"], 0)
        self.assertIsNone(records[0]["execution"])
        self.assertNotIn("[ERROR]", child.stdout)
        self.assertNotIn("[Writing]", child.stdout)
        self.assert_unchanged(self.flat, before)

    def test_ac054_preview_switch_cannot_replace_literal_manual_token_and_duplicates_fail(self):
        for manual, extras, code in (("--preview", [], "invalid_start_pages"),
                                     ("1", ["--preview", "--preview"], "invalid_arguments")):
            with self.subTest(manual=manual, extras=extras):
                child = subprocess.run([sys.executable, "-I", "-B", str(ENGINE), str(self.flat), str(self.base),
                    "manual", manual, *extras], cwd=self.case, capture_output=True, text=True, encoding="utf-8", timeout=20)
                self.assertEqual(child.returncode, 2, child.stdout + child.stderr)
                records = [json.loads(line) for line in child.stdout.splitlines() if line.startswith("{")]
                self.assertEqual(len(records), 1)
                self.assertEqual(records[0]["code"], code)
                self.assertNotIn("[Writing]", child.stdout)


if __name__ == "__main__":
    unittest.main()
