"""Owned output transactions using actual synthetic PDFs and Windows junctions."""

from concurrent.futures import ThreadPoolExecutor
from hashlib import sha256
import importlib.util
import json
import os
from pathlib import Path
import subprocess
import tempfile
import threading
from types import SimpleNamespace
import unittest
from unittest.mock import patch
import uuid

from pypdf import PdfReader, PdfWriter
from pypdf.errors import PdfReadError


ROOT = Path(__file__).resolve().parents[2]


def load_module(name, path):
    spec = importlib.util.spec_from_file_location(name, path.resolve())
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


engine = load_module("wbs_output_test_engine", ROOT / "engine/winbooksplit_engine.py")
fixtures = load_module("wbs_output_test_fixtures", ROOT / "tests/fixtures/generate_pdf_fixtures.py")


def hashes(directory):
    """Only call on a run-owned, non-reparse synthetic tree."""
    return {str(path.relative_to(directory)): sha256(path.read_bytes()).hexdigest()
            for path in directory.rglob("*") if path.is_file()}


class OutputTransactionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.owned = tempfile.TemporaryDirectory(prefix="wbs-output-unit-")
        cls.addClassCleanup(cls.owned.cleanup)
        cls.work = Path(cls.owned.name).resolve()
        cls.inputs = fixtures.generate_fixtures(cls.work / "fixtures")

    def setUp(self):
        self.base = self.work / ('case-' + uuid.uuid4().hex[:12])
        self.base.mkdir()
        self.neighbor = self.base / "unrelated-synthetic-neighbor.txt"
        self.neighbor.write_bytes(b"Original synthetic output-base neighbor\n")
        self.source = self.inputs["simple10"]
        self.source_before = self.source.read_bytes()

    def assert_preserved(self):
        self.assertEqual(self.source.read_bytes(), self.source_before)
        self.assertEqual(self.neighbor.read_bytes(), b"Original synthetic output-base neighbor\n")

    def prepared(self, source=None, mode="manual", data="4,7"):
        return engine.prepare_split(str(source or self.source), mode, data)

    def assert_complete_run(self, result, expected_ids):
        final = Path(result["final_directory"])
        self.assertTrue(final.is_absolute())
        self.assertEqual(final.parent, self.base)
        self.assertFalse(final.name.startswith(".WinBookSplit-"))
        self.assertEqual(result["written_count"], len(result["outputs"]))
        manifest_path = final / result["manifest_filename"]
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        self.assertEqual(manifest, json.loads(engine.serialize_result(result["manifest"])))
        self.assertEqual(manifest["schema_version"], 1)
        self.assertEqual(manifest["status"], "complete")
        for key in ("run_id", "final_directory", "mode", "total_pages", "source_identity", "coverage", "written_count"):
            self.assertEqual(manifest[key], json.loads(engine.serialize_result(result[key])))
        self.assertEqual(manifest["outputs"], json.loads(engine.serialize_result(result["outputs"])))
        self.assertEqual(json.loads((final / ".WinBookSplit-owner.json").read_text(encoding="utf-8")),
                         {"schema_version": 1, "kind": "run", "run_id": result["run_id"]})
        self.assertEqual({path.name for path in final.iterdir()},
                         {entry["filename"] for entry in result["outputs"]} |
                         {".WinBookSplit-owner.json", "WinBookSplit_Manifest.json"})
        ids = []
        for entry in result["outputs"]:
            path = final / entry["filename"]
            actual = fixtures.page_ids(path)
            self.assertEqual(actual, expected_ids[entry["start"]:entry["end"]])
            self.assertEqual(entry["page_count"], len(actual))
            self.assertEqual(entry["size_bytes"], path.stat().st_size)
            self.assertEqual(entry["sha256"], sha256(path.read_bytes()).hexdigest())
            ids.extend(actual)
        self.assertEqual(ids, expected_ids)
        return final

    def assert_failed_transaction(self, result, code, *, cleanup_complete=True):
        expected_exit = 130 if code == "output_cancelled" else 2 if code == "output_base_invalid" else 6
        self.assertEqual(result["exit_code"], expected_exit)
        self.assertEqual(result["code"], code)
        self.assertNotEqual(result["status"], "success")
        self.assertEqual(result["written_count"], 0)
        self.assertIsNone(result["execution"])
        self.assertEqual(result["fallback_modes"], ())
        diagnostic = result["diagnostic"]
        self.assertEqual(diagnostic["cleanup_complete"], cleanup_complete)
        if cleanup_complete:
            self.assertIsNone(diagnostic["retained_staging"])
        else:
            self.assertTrue(Path(diagnostic["retained_staging"]).is_dir())
        self.assertTrue(Path(diagnostic["record_path"]).is_file())
        self.assertLess(Path(diagnostic["record_path"]).stat().st_size, 65536)
        failed_record = json.loads(Path(diagnostic["record_path"]).read_text(encoding="utf-8"))
        self.assertEqual(failed_record["status"], "failed")
        self.assertEqual(failed_record["code"], code)
        self.assertEqual(failed_record["run_id"], diagnostic["run_id"])
        self.assertEqual(json.loads((Path(diagnostic["record_path"]).parent / ".WinBookSplit-owner.json")
                                   .read_text(encoding="utf-8")),
                         {"schema_version": 1, "kind": "failed", "run_id": diagnostic["run_id"]})
        if cleanup_complete:
            self.assertFalse(any(path.name.startswith(".WinBookSplit-stage-") for path in self.base.iterdir()))
        self.assert_preserved()

    def test_ac032_repeated_and_same_basename_runs_publish_distinct_complete_folders(self):
        sources = []
        for index, fixture in enumerate(("simple10", "nested12")):
            directory = self.work / ("same-name-source-" + str(index))
            directory.mkdir()
            source = directory / "Book.pdf"
            source.write_bytes(self.inputs[fixture].read_bytes())
            sources.append(source)
        previous = {}
        for source in (sources[0], sources[0], sources[1]):
            before = source.read_bytes()
            with patch.object(engine, "log"):
                result = engine.execute_split(self.prepared(source, data="1"), str(self.base))
            final = self.assert_complete_run(result, fixtures.page_ids(source))
            self.assertNotIn(final, previous)
            for prior, prior_hashes in previous.items():
                self.assertEqual(hashes(prior), prior_hashes)
            previous[final] = hashes(final)
            self.assertEqual(source.read_bytes(), before)
        self.assertEqual(len(previous), 3)
        self.assert_preserved()

    def test_published_manifest_and_validated_outputs_match_every_planned_mode(self):
        for mode, fixture, data in (("manual", "simple10", "4,7"), ("1", "simple10", None),
                                    ("2", "nested12", None)):
            with self.subTest(mode=mode), patch.object(engine, "log"):
                source = self.inputs[fixture]
                prepared = self.prepared(source, mode, data)
                result = engine.execute_split(prepared, str(self.base))
                self.assertEqual([(item["start"], item["end"]) for item in result["outputs"]],
                                 list(engine.preview_plan(prepared)["ranges"]))
                self.assert_complete_run(result, fixtures.page_ids(source))
                with self.assertRaises(TypeError):
                    result["manifest"]["outputs"][0]["sha256"] = "forged"
        self.assert_preserved()

    def test_ac034_simultaneous_identical_timestamp_runs_have_independent_complete_manifests(self):
        jobs = [self.prepared(), self.prepared()]
        barrier = threading.Barrier(2)
        original = engine.write_slice
        def synchronized_write(reader, start, end, path, *, on_created=None):
            if start == 0:
                barrier.wait(timeout=10)
            return original(reader, start, end, path, on_created=on_created)
        with patch.object(engine, "_timestamp", return_value="20261009-010203"), \
                patch.object(engine, "write_slice", side_effect=synchronized_write), patch.object(engine, "log"):
            with ThreadPoolExecutor(max_workers=2) as pool:
                results = list(pool.map(lambda job: engine.execute_split(job, str(self.base)), jobs))
        finals = [self.assert_complete_run(result, list(range(1, 11))) for result in results]
        self.assertEqual(len(set(finals)), 2)
        self.assertEqual(len({result["run_id"] for result in results}), 2)
        self.assertTrue(all("20261009-010203" in final.name for final in finals))
        self.assert_preserved()

    def test_atomic_stage_reservation_retries_collision_without_touching_foreign_folder(self):
        collision, fresh = uuid.uuid4().hex, uuid.uuid4().hex
        foreign = self.base / (".WinBookSplit-stage-" + collision)
        foreign.mkdir()
        (foreign / "unrelated.txt").write_bytes(b"Earlier foreign stage must remain\n")
        before = hashes(foreign)
        with patch.object(engine, "_new_run_id", side_effect=[collision, fresh]), patch.object(engine, "log"):
            result = engine.execute_split(self.prepared(), str(self.base))
        self.assertEqual(result["run_id"], fresh)
        self.assert_complete_run(result, list(range(1, 11)))
        self.assertEqual(hashes(foreign), before)
        self.assert_preserved()

    def test_late_final_collision_retries_without_replacing_the_competing_directory(self):
        original = engine._publish_run
        protected = []
        def collide(run, final_name):
            if not protected:
                path = self.base / final_name
                path.mkdir()
                (path / "unrelated.txt").write_bytes(b"Competing final directory\n")
                protected.append((path, hashes(path)))
            return original(run, final_name)
        with patch.object(engine, "_publish_run", side_effect=collide), patch.object(engine, "log"):
            result = engine.run_split(str(self.source), str(self.base), "manual", "4,7")
        self.assertEqual(result["status"], "success")
        final = self.assert_complete_run(result["execution"], list(range(1, 11)))
        self.assertEqual(len(protected), 1)
        self.assertEqual(hashes(protected[0][0]), protected[0][1])
        self.assertNotEqual(final, protected[0][0])
        self.assertFalse((protected[0][0] / "WinBookSplit_Manifest.json").exists())

    def test_ac033_failure_during_second_pdf_write_cleans_recorded_partial_and_never_publishes(self):
        original = engine.PdfWriter.write
        count = 0
        def fail_second(writer, stream):
            nonlocal count
            count += 1
            if count == 2:
                stream.write(b"Owned partial second PDF\n")
                raise OSError("Synthetic second PDF write failure")
            return original(writer, stream)
        with patch.object(engine.PdfWriter, "write", autospec=True, side_effect=fail_second), patch.object(engine, "log"):
            result = engine.run_split(str(self.source), str(self.base), "manual", "4,7")
        self.assertEqual(count, 2)
        self.assert_failed_transaction(result, "output_write_failed")
        self.assertEqual(result["status"], "write_error")
        self.assertEqual(list(self.base.rglob("*.pdf")), [])

    def test_initial_owner_marker_failure_does_not_leave_an_unmarked_stage(self):
        with patch.object(engine.OutputRun, "write_owned", side_effect=OSError("Synthetic initial marker failure")), \
                patch.object(engine, "write_slice") as writer, patch.object(engine, "log"):
            result = engine.run_split(str(self.source), str(self.base), "manual", "4,7")
        self.assertEqual(result["exit_code"], 2)
        self.assertEqual(result["code"], "output_base_invalid")
        self.assertNotEqual(result["status"], "success")
        writer.assert_not_called()
        self.assertEqual(list(self.base.glob(".WinBookSplit-stage-*")), [])
        self.assert_preserved()

    def test_stage_guard_acquisition_failure_preserves_reservation_and_records_failed_ownership(self):
        windows = engine._windows_output()
        original = windows.DirectoryGuard
        def fail_stage_guard(path, *args, **kwargs):
            if Path(path).name.startswith(".WinBookSplit-stage-"):
                raise OSError("Synthetic stage guard acquisition failure")
            return original(path, *args, **kwargs)
        proxy = SimpleNamespace(DirectoryGuard=fail_stage_guard, FileGuard=windows.FileGuard)
        with patch.object(engine, "_windows_output", return_value=proxy), \
                patch.object(engine, "write_slice") as writer, patch.object(engine, "log"):
            result = engine.run_split(str(self.source), str(self.base), "manual", "4,7")
        writer.assert_not_called()
        self.assert_failed_transaction(result, "output_base_invalid", cleanup_complete=False)
        stage = Path(result["diagnostic"]["retained_staging"])
        self.assertEqual(stage.parent, self.base)
        self.assertEqual(stage.name, ".WinBookSplit-stage-" + result["diagnostic"]["run_id"])
        self.assertEqual(list(stage.iterdir()), [])

    def test_observed_stage_change_before_owner_creation_refuses_unregistered_write(self):
        target = self.work / "synthetic-observed-stage-escape-target"
        target.mkdir()
        (target / "sentinel.txt").write_bytes(b"Owned escape-target sentinel must remain\n")
        before_target = hashes(target)
        original_write = engine.OutputRun.write_owned
        original_ordinary = engine._ordinary
        changed_stages = []
        def change_before_owner(run, name, data):
            if name == ".WinBookSplit-owner.json":
                changed_stages.append(Path(run.stage))
            return original_write(run, name, data)
        def reject_observed_change(path, directory=False):
            if Path(path) in changed_stages:
                raise engine.OutputError("output_ownership_failed", "Synthetic already-changed stage metadata")
            return original_ordinary(path, directory=directory)
        with patch.object(engine.OutputRun, "write_owned", autospec=True, side_effect=change_before_owner), \
                patch.object(engine, "_ordinary", side_effect=reject_observed_change), \
                patch.object(engine.OutputRun, "record_created", autospec=True) as registered, \
                patch.object(engine, "write_slice") as writer, patch.object(engine, "log"):
            result = engine.run_split(str(self.source), str(self.base), "manual", "4,7")
        registered.assert_not_called()
        writer.assert_not_called()
        self.assertEqual(len(changed_stages), 1)
        self.assert_failed_transaction(result, "output_ownership_failed", cleanup_complete=False)
        self.assertEqual(Path(result["diagnostic"]["retained_staging"]), changed_stages[0])
        self.assertEqual(list(changed_stages[0].iterdir()), [])
        self.assertEqual(hashes(target), before_target)

    def test_zero_wrong_count_and_unreadable_written_pdfs_are_rejected_before_publication(self):
        original = engine.PdfWriter.write
        for defect in ("zero", "wrong-count", "unreadable"):
            with self.subTest(defect=defect):
                def invalid_write(writer, stream):
                    if defect == "unreadable":
                        stream.write(b"Owned synthetic non-PDF output\n")
                        return
                    replacement = PdfWriter()
                    if defect == "wrong-count":
                        replacement.add_page(PdfReader(self.source).pages[0])
                    return original(replacement, stream)
                with patch.object(engine.PdfWriter, "write", autospec=True, side_effect=invalid_write), patch.object(engine, "log"):
                    result = engine.run_split(str(self.source), str(self.base), "manual", "4,7")
                self.assert_failed_transaction(result, "output_validation_failed")
                self.assertEqual(result["status"], "error")
                self.assertEqual(list(self.base.rglob("*.pdf")), [])

    def test_output_reopen_failure_is_distinct_from_input_read_failure_and_never_publishes(self):
        prepared = self.prepared()
        with patch.object(engine, "prepare_split", return_value=prepared), \
                patch.object(engine, "PdfReader", side_effect=PdfReadError("Synthetic staged PDF reopen failed")), \
                patch.object(engine, "log"):
            result = engine.run_split(str(self.source), str(self.base), "manual", "4,7")
        self.assert_failed_transaction(result, "output_validation_failed")
        self.assertEqual(result["status"], "error")
        self.assertEqual(list(self.base.rglob("*.pdf")), [])

    def test_failure_record_remains_bounded_for_large_native_error_text(self):
        with patch.object(engine, "write_slice", side_effect=OSError("Owned synthetic error " + "X" * 200000)), \
                patch.object(engine, "log"):
            result = engine.run_split(str(self.source), str(self.base), "manual", "4,7")
        self.assert_failed_transaction(result, "output_write_failed")
        record = json.loads(Path(result["diagnostic"]["record_path"]).read_text(encoding="utf-8"))
        self.assertLessEqual(len(record["message"]), 2048)

    def test_manifest_and_promotion_failures_never_publish_success(self):
        for seam, code in (("_write_manifest", "output_manifest_failed"), ("_publish_run", "output_publish_failed")):
            with self.subTest(seam=seam), patch.object(engine, seam, side_effect=OSError("Synthetic " + seam + " failure")), \
                    patch.object(engine, "log"):
                result = engine.run_split(str(self.source), str(self.base), "manual", "4,7")
            self.assert_failed_transaction(result, code)
            self.assertEqual(list(self.base.rglob("*.pdf")), [])

    def test_changed_pdf_bytes_at_publication_boundary_are_detected_before_final_rename(self):
        original = engine.OutputRun.check_publication
        def change_pdf(run):
            # Model the required closed-child interval at the native publication
            # boundary, changing only bytes in this run's synthetic PDF.
            for guard in run.file_guards.values():
                guard.close()
            run.file_guards.clear()
            path = Path(run.stage) / run.publication_manifest["outputs"][0]["filename"]
            data = path.read_bytes()
            path.write_bytes(b"!" + data[1:])
            return original(run)
        with patch.object(engine.OutputRun, "check_publication", autospec=True, side_effect=change_pdf), \
                patch.object(engine, "log"):
            result = engine.run_split(str(self.source), str(self.base), "manual", "4,7")
        self.assert_failed_transaction(result, "output_validation_failed")
        self.assertEqual(list(self.base.rglob("*.pdf")), [])

    def test_unknown_member_at_publication_boundary_retains_stage_and_never_publishes(self):
        original = engine.OutputRun.check_publication
        stages = []
        def add_foreign_member(run):
            stage = Path(run.stage)
            stages.append(stage)
            (stage / "foreign.txt").write_bytes(b"Foreign publication-boundary content\n")
            return original(run)
        with patch.object(engine.OutputRun, "check_publication", autospec=True, side_effect=add_foreign_member), \
                patch.object(engine, "log"):
            result = engine.run_split(str(self.source), str(self.base), "manual", "4,7")
        self.assert_failed_transaction(result, "output_ownership_failed", cleanup_complete=False)
        self.assertEqual(Path(result["diagnostic"]["retained_staging"]), stages[0])
        self.assertEqual((stages[0] / "foreign.txt").read_bytes(), b"Foreign publication-boundary content\n")

    def test_keyboard_interrupt_during_second_write_cleans_owned_partial_and_is_cancelled(self):
        original = engine.PdfWriter.write
        count = 0
        def cancel_second(writer, stream):
            nonlocal count
            count += 1
            if count == 2:
                stream.write(b"Owned cancelled second PDF\n")
                raise KeyboardInterrupt
            return original(writer, stream)
        with patch.object(engine.PdfWriter, "write", autospec=True, side_effect=cancel_second), patch.object(engine, "log"):
            result = engine.run_split(str(self.source), str(self.base), "manual", "4,7")
        self.assertEqual(count, 2)
        self.assert_failed_transaction(result, "output_cancelled")
        self.assertEqual(result["status"], "cancelled")
        self.assertEqual(list(self.base.rglob("*.pdf")), [])

    def test_post_publication_close_error_is_incomplete_with_validated_output_retained(self):
        original = engine.OutputRun.close
        closed = []
        def close_then_error(run):
            original(run)
            closed.append(run.run_id)
            raise OSError("Synthetic close error after native handles closed")
        with patch.object(engine.OutputRun, "close", autospec=True, side_effect=close_then_error), \
                patch.object(engine, "log"):
            result = engine.run_split(str(self.source), str(self.base), "manual", "4,7")
        self.assertEqual(len(closed), 1)
        self.assertEqual(result["status"], "incomplete")
        self.assertEqual(result["exit_code"], 6)
        self.assertEqual(result["code"], "output_handle_close_failed")
        self.assertEqual(result["fallback_modes"], ())
        execution = result["execution"]
        final = self.assert_complete_run(execution, list(range(1, 11)))
        warnings = execution["post_publication_warnings"]
        self.assertTrue(warnings)
        self.assertTrue(any("Synthetic close error" in str(warning) for warning in warnings))
        self.assertTrue(all(warning in result["warnings"] for warning in warnings))
        self.assertEqual({path.name for path in self.base.iterdir()}, {final.name, self.neighbor.name})
        self.assert_preserved()

    def test_post_publication_close_interrupt_retains_validated_output_as_incomplete(self):
        original = engine.OutputRun.close
        def close_then_interrupt(run):
            original(run)
            raise KeyboardInterrupt
        with patch.object(engine.OutputRun, "close", autospec=True, side_effect=close_then_interrupt), \
                patch.object(engine, "log"):
            result = engine.run_split(str(self.source), str(self.base), "manual", "4,7")
        self.assertEqual(result["status"], "incomplete")
        self.assertEqual(result["exit_code"], 6)
        self.assertEqual(result["code"], "output_handle_close_failed")
        self.assertEqual(result["fallback_modes"], ())
        final = self.assert_complete_run(result["execution"], list(range(1, 11)))
        self.assertTrue(any("interrupted" in warning["message"] for warning in result["warnings"]))
        self.assertEqual({path.name for path in self.base.iterdir()}, {final.name, self.neighbor.name})
        self.assert_preserved()

    def test_prepublication_close_error_preserves_the_primary_write_failure_and_diagnostic(self):
        original = engine.OutputRun.close
        def close_then_error(run):
            original(run)
            raise OSError("Synthetic secondary close error")
        with patch.object(engine.OutputRun, "close", autospec=True, side_effect=close_then_error), \
                patch.object(engine, "write_slice", side_effect=OSError("Synthetic primary write failure")), \
                patch.object(engine, "log"):
            result = engine.run_split(str(self.source), str(self.base), "manual", "4,7")
        self.assert_failed_transaction(result, "output_write_failed")
        self.assertEqual(result["status"], "write_error")
        self.assertIn("Synthetic primary write failure", result["message"])
        self.assertIn("Synthetic secondary close error", result["diagnostic"]["close_error"])
        record = json.loads(Path(result["diagnostic"]["record_path"]).read_text(encoding="utf-8"))
        self.assertIn("Synthetic primary write failure", record["message"])
        self.assertEqual(list(self.base.rglob("*.pdf")), [])

    def test_close_attempts_every_guard_and_collects_multiple_errors(self):
        run = engine.OutputRun(str(self.base), str(self.source))
        run.write_owned("owned.txt", b"Synthetic registered close-all file\n")
        guards = list(run.file_guards.values()) + [run.stage_guard] + list(reversed(run.guards))
        failed_ids = {id(guards[0]): "Synthetic first guard close error",
                      id(run.stage_guard): "Synthetic stage guard close error"}
        attempts = []
        for guard in guards:
            original = guard.close
            def close_guard(guard=guard, original=original):
                attempts.append(id(guard))
                original()
                message = failed_ids.pop(id(guard), None)
                if message:
                    raise OSError(message)
            guard.close = close_guard
        try:
            with self.assertRaises(engine.OutputError) as raised:
                run.close()
            self.assertEqual(raised.exception.code, "output_handle_close_failed")
            outcome = str(raised.exception)
            self.assertCountEqual(attempts, [id(guard) for guard in guards])
            self.assertTrue(all(guard._handle is None for guard in guards))
            self.assertIn("Synthetic first guard close error", str(outcome))
            self.assertIn("Synthetic stage guard close error", str(outcome))
            self.assert_preserved()
        finally:
            # All injected errors are one-shot; finish releasing any handles if
            # a broken implementation stopped at the first error.
            for guard in guards:
                if guard._handle is not None:
                    guard.close()

    def test_ac035_unregistered_member_blocks_cleanup_without_deleting_any_member(self):
        run = engine.OutputRun(str(self.base), str(self.source))
        try:
            stage = Path(run.stage)
            ours = stage / "registered.txt"
            with ours.open("xb") as stream:
                run.record_created(str(ours), stream)
                stream.write(b"Owned registered file\n")
            foreign = stage / "unregistered.txt"
            foreign.write_bytes(b"Foreign member must be preserved\n")
            before = hashes(stage)
            with self.assertRaises(engine.OutputError) as raised:
                run.cleanup()
            self.assertEqual(raised.exception.code, "output_ownership_failed")
            self.assertEqual(hashes(stage), before)
            self.assert_preserved()
        finally:
            run.close()

    def test_marker_lock_blocks_tamper_and_lost_lock_still_rejects_changed_ownership(self):
        run = engine.OutputRun(str(self.base), str(self.source))
        try:
            marker = Path(run.stage) / ".WinBookSplit-owner.json"
            foreign_owner = b'{"schema_version":1,"kind":"run","run_id":"foreign"}'
            with self.assertRaises(PermissionError):
                marker.write_bytes(foreign_owner)
            self.assertEqual(marker.read_bytes(), run.marker_bytes)
            # Inject loss of this run's own marker lock, then change its bytes.
            # Cleanup must independently reject ownership rather than deleting it.
            run.file_guards.pop(str(marker)).close()
            marker.write_bytes(foreign_owner)
            with self.assertRaises(engine.OutputError) as raised:
                run.cleanup()
            self.assertEqual(raised.exception.code, "output_ownership_failed")
            self.assertEqual(marker.read_bytes(), foreign_owner)
            self.assert_preserved()
        finally:
            run.close()

    def test_blocked_child_rename_and_unsealed_unlink_replacement_preserve_foreign_file(self):
        run = engine.OutputRun(str(self.base), str(self.source))
        try:
            path = Path(run.stage) / "registered.txt"
            with path.open("xb") as stream:
                run.record_created(str(path), stream)
                stream.write(b"Owned original file\n")
            replacement = Path(run.stage) / "replacement.tmp"
            replacement.write_bytes(b"Competing replacement must remain\n")
            with self.assertRaises(PermissionError):
                replacement.replace(path)
            self.assertEqual(path.read_bytes(), b"Owned original file\n")
            # The directory lock blocks rename, but an unsealed child can still
            # be unlinked and recreated. Cleanup must reject its new identity.
            replacement.unlink()
            path.unlink()
            with path.open("xb") as stream:
                stream.write(b"Competing replacement must remain\n")
            with self.assertRaises(engine.OutputError) as raised:
                run.cleanup()
            self.assertEqual(raised.exception.code, "output_ownership_failed")
            self.assertEqual(path.read_bytes(), b"Competing replacement must remain\n")
            self.assert_preserved()
        finally:
            run.close()

    def test_unknown_member_during_failed_execution_retains_marked_stage_and_diagnostic(self):
        original = engine.write_slice
        stages = []
        def foreign_member_then_fail(reader, start, end, path, *, on_created=None):
            original(reader, start, end, path, on_created=on_created)
            stage = Path(path).parent
            stages.append(stage)
            (stage / "foreign.txt").write_bytes(b"Foreign content must survive failed cleanup\n")
            raise OSError("Synthetic failure after unrelated member appeared")
        with patch.object(engine, "write_slice", side_effect=foreign_member_then_fail), patch.object(engine, "log"):
            result = engine.run_split(str(self.source), str(self.base), "manual", "4,7")
        self.assert_failed_transaction(result, "output_write_failed", cleanup_complete=False)
        self.assertEqual(Path(result["diagnostic"]["retained_staging"]), stages[0])
        self.assertEqual((stages[0] / "foreign.txt").read_bytes(), b"Foreign content must survive failed cleanup\n")
        self.assertTrue(result["diagnostic"]["cleanup_error"])

    def junction(self, link, target):
        command = [str(Path(os.environ["SystemRoot"]) / "System32/cmd.exe"), "/d", "/c",
                   "mklink", "/J", str(link), str(target)]
        result = subprocess.run(command, capture_output=True, text=True, timeout=10)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertTrue(link.is_junction())
        # This removes only the run-created junction itself, never its target.
        self.addCleanup(lambda: link.rmdir() if os.path.lexists(link) else None)

    @unittest.skipUnless(os.name == "nt", "Requires actual Windows junctions")
    def test_ac035_real_junction_base_or_ancestor_is_rejected_without_writing_target(self):
        target = self.work / "owned-unrelated-junction-target"
        target.mkdir()
        (target / "sentinel.txt").write_bytes(b"Unrelated synthetic junction target\n")
        before = hashes(target)
        link = self.base / "escape"
        self.junction(link, target)
        for destination in (link, link / "nested"):
            with self.subTest(destination=destination), patch.object(engine, "write_slice") as writer, \
                    patch.object(engine, "log"):
                result = engine.run_split(str(self.source), str(destination), "manual", "4,7")
            self.assertEqual(result["exit_code"], 2)
            self.assertEqual(result["code"], "output_base_invalid")
            writer.assert_not_called()
            self.assertEqual(hashes(target), before)
        self.assert_preserved()

    @unittest.skipUnless(os.name == "nt", "Requires actual Windows junctions")
    def test_real_junction_child_blocks_cleanup_and_preserves_target(self):
        target = self.work / "owned-unrelated-stage-junction-target"
        target.mkdir()
        (target / "sentinel.txt").write_bytes(b"Stage junction target must survive\n")
        before = hashes(target)
        run = engine.OutputRun(str(self.base), str(self.source))
        try:
            link = Path(run.stage) / "escape"
            self.junction(link, target)
            with self.assertRaises(engine.OutputError) as raised:
                run.cleanup()
            self.assertEqual(raised.exception.code, "output_ownership_failed")
            self.assertEqual(hashes(target), before)
            self.assertTrue(link.is_junction())
            self.assert_preserved()
        finally:
            run.close()

    @unittest.skipUnless(os.name == "nt", "Requires actual Windows directory handle protection")
    def test_held_stage_directory_cannot_be_renamed_for_junction_replacement(self):
        run = engine.OutputRun(str(self.base), str(self.source))
        try:
            stage = Path(run.stage)
            with self.assertRaises(PermissionError):
                stage.rename(self.base / "competing-moved-stage")
            self.assertTrue(stage.is_dir())
            run.assert_owned()
            self.assertTrue(run.cleanup())
            self.assertFalse(stage.exists())
            self.assert_preserved()
        finally:
            run.close()


if __name__ == "__main__":
    unittest.main(verbosity=2)
