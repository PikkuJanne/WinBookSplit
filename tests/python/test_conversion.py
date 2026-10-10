"""Owned fake converters exercise actual native jobs, validation and cleanup."""

from dataclasses import FrozenInstanceError
from hashlib import sha256
import importlib.util
from io import BytesIO
import json
import os
from pathlib import Path
import stat
import sys
import tempfile
import unittest
from unittest.mock import patch

from pypdf import PdfReader, PdfWriter
from pypdf.generic import DecodedStreamObject, NameObject, NumberObject


ROOT = Path(__file__).resolve().parents[2]


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


engine = load("wbs_conversion_unit_engine", ROOT / "engine/winbooksplit_engine.py")
conversion = load("wbs_conversion_unit_helper", ROOT / "engine/winbooksplit_conversion.py")


def pdf_bytes(pages=3, encrypted=False):
    writer = PdfWriter()
    for number in range(1, pages + 1):
        page = writer.add_blank_page(width=100, height=100)
        page[NameObject("/SyntheticPageID")] = NumberObject(number)
        stream = DecodedStreamObject()
        stream.set_data(f"q\n1 0 0 1 {number} {number} cm\nQ\n".encode("ascii"))
        page[NameObject("/Contents")] = writer._add_object(stream)
    if encrypted:
        writer.encrypt(user_password="", owner_password="synthetic-owner-password")
    result = BytesIO()
    writer.write(result)
    return result.getvalue()


FAKE_CONVERTER = '''import json, os
from pathlib import Path
import sys, threading, time
source, output = Path(sys.argv[1]), Path(sys.argv[2])
source.read_bytes()
action = os.environ["WBS_SYNTHETIC_ACTION"]
template = Path(os.environ["WBS_SYNTHETIC_PDF"]).read_bytes()
if action == "unknown":
    (output.parent / "unexpected-synthetic-member.txt").write_bytes(b"Owned fixture, unregistered converter member")
if action == "replace":
    output.unlink()
if action == "corrupt":
    template = b"Owned synthetic corrupt PDF"
if action != "missing":
    with output.open("wb") as stream:
        stream.write(b"" if action == "empty" else template)
if action == "burst":
    def write(fd, label):
        for _ in range(48):
            os.write(fd, label * 8192)
        os.write(fd, ("FINAL_" + label.decode() + "_尾").encode("utf-8"))
    threads = [threading.Thread(target=write,args=(1,b"O")), threading.Thread(target=write,args=(2,b"E"))]
    for thread in threads: thread.start()
    for thread in threads: thread.join()
else:
    os.write(1,json.dumps({"args":sys.argv[1:],"cwd":os.getcwd()},ensure_ascii=True).encode("ascii"))
    os.write(2,b"Owned synthetic stderr final tail without newline")
if action == "wait":
    os.write(1,b"CONVERTER_READY")
    time.sleep(30)
raise SystemExit(7 if action == "nonzero" else 0)
'''


class ConversionHelperTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.owned = tempfile.TemporaryDirectory(prefix="wbs-c-")
        cls.addClassCleanup(cls.owned.cleanup)
        cls.work = Path(cls.owned.name).resolve()
        cls.script = cls.work / "owned_fake_converter.py"
        cls.script.write_text(FAKE_CONVERTER, encoding="utf-8")
        cls.serial = 0
        cls.jobs = load("wbs_conversion_unit_jobs", ROOT / "engine/winbooksplit_job.py")

    def setUp(self):
        type(self).serial += 1
        self.case = self.work / f"c{self.serial:02d}"
        self.case.mkdir()
        self.base = self.case / "out"
        self.base.mkdir()
        self.source = self.case / "Book [%!] O'Brien & (日本).epub"
        self.source.write_bytes(b"Owned synthetic ebook input; fake child does not interpret format.\n")
        self.neighbor_pdf = self.source.with_suffix(".pdf")
        self.neighbor_pdf.write_bytes(b"Owned unrelated same-basename PDF neighbor\n")
        self.other = self.base / "unrelated-synthetic-neighbor.txt"
        self.other.write_bytes(b"Owned output-base neighbor\n")
        self.source_before = self.source.read_bytes()
        self.source_attributes = os.stat(self.source).st_file_attributes
        self.template = self.case / "synthetic-output-template.pdf"
        self.template.write_bytes(pdf_bytes())
        self.runs, self.argv = [], []
        self.action = "valid"
        self.cancel = False

    def new_run(self, base, source):
        run = engine.OutputRun(base, source)
        self.runs.append(run)
        return run

    def factory(self, argv, *, cwd, env):
        self.argv.append(tuple(argv))
        actual_env = {**env, "WBS_SYNTHETIC_ACTION": self.action, "WBS_SYNTHETIC_PDF": str(self.template)}
        job = self.jobs.launch_job([sys.executable, "-I", "-B", str(self.script), *argv[1:]], cwd=cwd, env=actual_env)
        if self.cancel:
            original = job.poll
            calls = 0
            def cancel_once():
                nonlocal calls
                calls += 1
                if calls == 8:
                    raise KeyboardInterrupt()
                return original()
            job.poll = cancel_once
        return job

    def convert(self, **kwargs):
        return conversion.convert_ebook(self.source, sys.executable, self.base,
            new_run=self.new_run, process_factory=self.factory, record_failure=engine._failed_run, **kwargs)

    def assert_preserved(self):
        self.assertEqual(self.source.read_bytes(), self.source_before)
        self.assertEqual(os.stat(self.source).st_file_attributes, self.source_attributes)
        self.assertEqual(self.neighbor_pdf.read_bytes(), b"Owned unrelated same-basename PDF neighbor\n")
        self.assertEqual(self.other.read_bytes(), b"Owned output-base neighbor\n")

    def assert_cleaned(self):
        self.assertTrue(self.runs)
        self.assertTrue(all(not Path(run.stage).exists() for run in self.runs))
        self.assertFalse(any(path.name.startswith(".WinBookSplit-stage-") for path in self.base.iterdir()))
        self.assert_preserved()

    def test_ac040_real_fake_child_uses_literal_args_owned_target_and_returns_immutable_bytes(self):
        result = self.convert()
        self.assertEqual(result.pdf_bytes, self.template.read_bytes())
        self.assertEqual(result.original_source_identity["sha256"], sha256(self.source_before).hexdigest())
        self.assertEqual(result.original_source_identity["binding"], "ebook_snapshot")
        self.assertEqual(result.generated_pdf_identity["sha256"], sha256(result.pdf_bytes).hexdigest())
        self.assertEqual(result.generated_pdf_identity["binding"], "reader_snapshot")
        reader = PdfReader(BytesIO(result.pdf_bytes))
        self.assertEqual([int(page["/SyntheticPageID"]) for page in reader.pages], [1, 2, 3])
        self.assertEqual(result.generated_pdf_identity["page_content_sha256"],
                         tuple(sha256(page.get_contents().get_data()).hexdigest() for page in reader.pages))
        self.assertEqual(result.generated_pdf_identity["page_count"], 3)
        self.assertEqual(result.conversion["original_source_identity"], result.original_source_identity)
        self.assertEqual(result.conversion["generated_pdf_identity"], result.generated_pdf_identity)
        self.assertEqual(result.conversion["workspace_cleanup"],
                         {"cleanup_complete": True, "retained_staging": None})
        observed = json.loads(result.conversion["stdout_tail"])
        self.assertEqual(observed["args"], list(self.argv[0][1:]))
        self.assertEqual(self.argv[0][-2:], ("--output-profile", "tablet"))
        self.assertEqual(Path(self.argv[0][2]).parent, Path(self.runs[0].stage))
        self.assertEqual(Path(self.argv[0][2]).name, "WinBookSplit_Converted.pdf")
        self.assertEqual(observed["cwd"], self.runs[0].stage)
        self.assertIn("without newline", result.conversion["stderr_tail"])
        with self.assertRaises(TypeError):
            result.original_source_identity["sha256"] = "changed"
        with self.assertRaises(TypeError):
            result.conversion["argv"][0] = "changed"
        with self.assertRaises(FrozenInstanceError):
            result.pdf_bytes = b"changed"
        self.assertEqual(list(self.base.iterdir()), [self.other])
        self.assert_cleaned()

    def test_ac040_read_only_ebook_and_same_basename_neighbor_are_preserved(self):
        os.chmod(self.source, stat.S_IREAD)
        self.source_attributes = os.stat(self.source).st_file_attributes
        try:
            result = self.convert()
            self.assertEqual(result.generated_pdf_identity["page_count"], 3)
            self.assert_cleaned()
        finally:
            os.chmod(self.source, stat.S_IREAD | stat.S_IWRITE)

    def test_ac041_zero_exit_missing_empty_corrupt_zero_page_and_encrypted_outputs_rejected(self):
        for action, template in (("missing", pdf_bytes()), ("empty", pdf_bytes()), ("corrupt", pdf_bytes()),
                                 ("valid", pdf_bytes(0)), ("valid", pdf_bytes(encrypted=True))):
            with self.subTest(action=action, template_sha=sha256(template).hexdigest()):
                self.action = action
                self.template.write_bytes(template)
                expected_code = "unsupported_document" if (action == "valid" and PdfReader(BytesIO(template)).is_encrypted) else "conversion_output_invalid"
                with self.assertRaises(conversion.ConversionError) as caught:
                    self.convert()
                self.assertEqual(caught.exception.code, expected_code)
                self.assertEqual(caught.exception.conversion["exit_code"], 0)
                self.assertIn("without newline", caught.exception.conversion["stderr_tail"])
                self.assertEqual(caught.exception.conversion["original_source_identity"]["sha256"],
                                 sha256(self.source_before).hexdigest())
                self.assertTrue(caught.exception.diagnostic["cleanup_complete"])
                record = Path(caught.exception.diagnostic["record_path"])
                saved = json.loads(record.read_text(encoding="utf-8"))
                self.assertEqual(saved["code"], expected_code)
                self.assertEqual(saved["conversion"]["exit_code"], 0)
                self.assert_cleaned()

    def test_nonzero_converter_exit_rejects_even_a_readable_positive_pdf(self):
        self.action = "nonzero"
        with self.assertRaises(conversion.ConversionError) as caught:
            self.convert()
        self.assertEqual(caught.exception.code, "conversion_failed")
        self.assertEqual(caught.exception.conversion["exit_code"], 7)
        self.assertIn("stderr", caught.exception.conversion["stderr_tail"])
        self.assert_cleaned()

    def test_actual_both_pipe_bursts_are_bounded_with_final_utf8_tails(self):
        self.action = "burst"
        result = self.convert()
        for stream, marker in (("stdout", "FINAL_O_尾"), ("stderr", "FINAL_E_尾")):
            self.assertGreater(result.conversion[stream + "_total_bytes"], 65536)
            self.assertTrue(result.conversion[stream + "_truncated"])
            self.assertTrue(result.conversion[stream + "_tail"].endswith(marker))
            self.assertLessEqual(len(result.conversion[stream + "_tail"].encode("utf-8")), 65536)
        self.assert_cleaned()

    def test_actual_timeout_stops_child_before_owned_cleanup(self):
        self.action = "wait"
        with self.assertRaises(conversion.ConversionError) as caught:
            self.convert(timeout_seconds=0.4)
        self.assertEqual(caught.exception.code, "conversion_timeout")
        self.assertTrue(caught.exception.tree_stopped)
        self.assertTrue(caught.exception.diagnostic["cleanup_complete"])
        self.assert_cleaned()

    def test_actual_keyboard_interrupt_stops_child_before_owned_cleanup(self):
        self.action = "wait"
        self.cancel = True
        with self.assertRaises(conversion.ConversionError) as caught:
            self.convert(timeout_seconds=30)
        self.assertEqual(caught.exception.code, "conversion_cancelled")
        self.assertTrue(caught.exception.tree_stopped)
        self.assertTrue(caught.exception.diagnostic["cleanup_complete"])
        self.assert_cleaned()

    def test_start_failure_cleans_parent_created_placeholder(self):
        unusable = self.case / "synthetic-not-executable.exe"
        unusable.write_bytes(b"Owned fixture: an ordinary file which cannot be executed")
        with self.assertRaises(conversion.ConversionError) as caught:
            conversion.convert_ebook(self.source, unusable, self.base,
                new_run=self.new_run, record_failure=engine._failed_run)
        self.assertEqual(caught.exception.code, "conversion_start_failed")
        self.assert_cleaned()

    def test_missing_relative_none_or_directory_converter_rejects_before_allocation(self):
        for converter in (None, "relative.exe", self.case / "nonexistent.exe", self.case):
            with self.subTest(converter=converter):
                with patch.object(self, "new_run") as reserve, patch.object(self, "factory") as child:
                    with self.assertRaises(conversion.ConversionError) as caught:
                        conversion.convert_ebook(self.source, converter, self.base,
                            new_run=self.new_run, process_factory=self.factory)
                self.assertEqual(caught.exception.code, "converter_not_found")
                reserve.assert_not_called()
                child.assert_not_called()
        self.assertEqual(list(self.base.iterdir()), [self.other])
        self.assert_preserved()

    def test_unproved_startup_shutdown_refuses_workspace_cleanup_callback(self):
        error = OSError("Owned injected startup shutdown is unproved")
        error.converter_tree_stopped = False
        with patch.object(self, "factory", side_effect=error), \
                patch.object(engine, "_failed_run") as cleanup, \
                patch.object(engine.OutputRun, "cleanup", autospec=True) as remove:
            with self.assertRaises(conversion.ConversionError) as caught:
                self.convert()
        self.assertEqual(caught.exception.code, "conversion_cleanup_failed")
        self.assertEqual(caught.exception.diagnostic["primary_code"], "conversion_start_failed")
        self.assertFalse(caught.exception.tree_stopped)
        cleanup.assert_not_called()
        remove.assert_not_called()
        self.assertTrue(Path(caught.exception.diagnostic["retained_staging"]).exists())
        self.assert_preserved()

    def test_converter_replacement_and_unknown_members_are_preserved_not_adopted(self):
        for action in ("replace", "unknown"):
            with self.subTest(action=action):
                self.action = action
                with self.assertRaises(conversion.ConversionError) as caught:
                    self.convert()
                self.assertEqual(caught.exception.code, "conversion_cleanup_failed")
                self.assertEqual(caught.exception.diagnostic["primary_code"], "conversion_ownership_failed")
                self.assertFalse(caught.exception.diagnostic["cleanup_complete"])
                retained = Path(caught.exception.diagnostic["retained_staging"])
                self.assertTrue(retained.exists())
                self.assertEqual(retained.parent, self.base)
                self.assertTrue((retained / ".WinBookSplit-owner.json").exists())
                self.assert_preserved()

    def test_constructor_cancellation_requires_explicit_tree_stop_proof(self):
        for stopped in (False, True):
            with self.subTest(tree_stopped=stopped):
                error = KeyboardInterrupt()
                error.converter_tree_stopped = stopped
                with patch.object(self, "factory", side_effect=error), \
                        patch.object(engine, "_failed_run", wraps=engine._failed_run) as cleanup:
                    with self.assertRaises(conversion.ConversionError) as caught:
                        self.convert()
                self.assertEqual(caught.exception.tree_stopped, stopped)
                self.assertIsNone(caught.exception.conversion["exit_code"])
                if stopped:
                    self.assertEqual(caught.exception.code, "conversion_cancelled")
                    self.assertTrue(caught.exception.diagnostic["cleanup_complete"])
                    self.assertFalse(Path(self.runs[-1].stage).exists())
                    cleanup.assert_called_once()
                else:
                    self.assertEqual(caught.exception.code, "conversion_cleanup_failed")
                    self.assertEqual(caught.exception.diagnostic["primary_code"], "conversion_cancelled")
                    self.assertFalse(caught.exception.diagnostic["cleanup_complete"])
                    self.assertTrue(Path(caught.exception.diagnostic["retained_staging"]).exists())
                    cleanup.assert_not_called()
                self.assert_preserved()

    def prepared_conversion(self, converted, keep):
        return engine.prepare_split(converted.generated_pdf_identity["path"], "manual", "2",
            output_base=self.base, _pdf_bytes=converted.pdf_bytes,
            _conversion_metadata={"original_ebook_identity": converted.original_source_identity,
                                  "conversion": converted.conversion, "keep_converted_pdf": keep})

    def test_default_conversion_preview_and_publication_use_snapshot_after_workspace_removal(self):
        converted = self.convert()
        self.assert_cleaned()
        self.assertFalse(Path(converted.generated_pdf_identity["path"]).exists())
        prepared = self.prepared_conversion(converted, False)
        preview = engine.preview_plan(prepared)
        self.assertEqual(list(self.base.iterdir()), [self.other])
        self.assertEqual(preview["source_identity"]["sha256"], sha256(converted.pdf_bytes).hexdigest())
        self.assertEqual(preview["original_ebook_identity"], converted.original_source_identity)
        self.assertEqual(preview["ranges"], ((0, 1), (1, 3)))
        with patch.object(engine, "plan_manual_starts", side_effect=AssertionError("execution must use preview")):
            result = engine.execute_split(prepared, self.base)
        final = Path(result["final_directory"])
        self.assertEqual(result["written_count"], 2)
        self.assertIsNone(result["retained_intermediate"])
        self.assertFalse((final / conversion.CONVERTED_FILENAME).exists())
        self.assertEqual({path.name for path in final.iterdir()},
                         {engine.OWNER_FILENAME, engine.MANIFEST_FILENAME, *[entry["filename"] for entry in result["outputs"]]})
        ids = [int(page["/SyntheticPageID"]) for entry in result["outputs"]
               for page in PdfReader(final / entry["filename"]).pages]
        self.assertEqual(ids, [1, 2, 3])
        self.assert_preserved()

    def test_opt_in_retained_pdf_is_exact_snapshot_separate_from_chapter_count_and_manifest(self):
        converted = self.convert()
        prepared = self.prepared_conversion(converted, True)
        result = engine.execute_split(prepared, self.base)
        final = Path(result["final_directory"])
        retained = result["retained_intermediate"]
        self.assertEqual(result["written_count"], 2)
        self.assertEqual(len(result["outputs"]), 2)
        self.assertEqual(retained["filename"], conversion.CONVERTED_FILENAME)
        self.assertEqual((final / retained["filename"]).read_bytes(), converted.pdf_bytes)
        self.assertEqual(retained["sha256"], sha256(converted.pdf_bytes).hexdigest())
        self.assertEqual(retained["size_bytes"], len(converted.pdf_bytes))
        self.assertEqual(retained["page_count"], 3)
        manifest = json.loads((final / engine.MANIFEST_FILENAME).read_text(encoding="utf-8"))
        self.assertEqual(manifest["retained_intermediate"], dict(retained))
        self.assertEqual(manifest["written_count"], 2)
        self.assertEqual(manifest["original_ebook_identity"]["sha256"], sha256(self.source_before).hexdigest())
        self.assertEqual(manifest["source_identity"]["sha256"], retained["sha256"])
        self.assertEqual({path.name for path in final.iterdir()},
                         {engine.OWNER_FILENAME, engine.MANIFEST_FILENAME, retained["filename"],
                          *[entry["filename"] for entry in result["outputs"]]})
        self.assert_preserved()

    def test_retained_snapshot_tampering_rejects_publication_and_cleans_known_files(self):
        converted = self.convert()
        prepared = self.prepared_conversion(converted, True)
        original = engine.OutputRun.check_publication
        def mutate_then_check(run):
            path = os.path.join(run.stage, conversion.CONVERTED_FILENAME)
            run.file_guards.pop(path).close()
            with open(path, "ab") as stream:
                stream.write(b"Owned synthetic same-identity tamper")
            original(run)
        with patch.object(engine.OutputRun, "check_publication", autospec=True, side_effect=mutate_then_check):
            with self.assertRaises(engine.OutputError) as caught:
                engine.execute_split(prepared, self.base)
        self.assertEqual(caught.exception.code, "output_validation_failed")
        self.assertTrue(caught.exception.diagnostic["cleanup_complete"])
        self.assertFalse(any(path.name.startswith(".WinBookSplit-stage-") for path in self.base.iterdir()))
        self.assertEqual([path for path in self.base.iterdir() if path.is_dir() and not path.name.startswith(".WinBookSplit-failed-")], [])
        self.assert_preserved()

    def test_invalid_timeouts_fail_before_workspace_or_child(self):
        for timeout in (0, -1, float("inf"), float("nan"), True, "30"):
            with self.subTest(timeout=timeout):
                with patch.object(self, "new_run") as reserve, patch.object(self, "factory") as child:
                    with self.assertRaises(conversion.ConversionError):
                        self.convert(timeout_seconds=timeout)
                reserve.assert_not_called()
                child.assert_not_called()
        self.assertEqual(list(self.base.iterdir()), [self.other])
        self.assert_preserved()

    def test_cleanup_or_handle_close_failure_prevents_returning_success_bytes(self):
        for method in ("cleanup", "close"):
            with self.subTest(method=method):
                original = getattr(engine.OutputRun, method)
                def fail_after(run):
                    original(run)
                    raise OSError("Owned synthetic cleanup/close error")
                with patch.object(engine.OutputRun, method, autospec=True, side_effect=fail_after):
                    with self.assertRaises(conversion.ConversionError) as caught:
                        self.convert()
                self.assertEqual(caught.exception.code, "conversion_cleanup_failed")
                self.assert_preserved()


if __name__ == "__main__":
    unittest.main()
