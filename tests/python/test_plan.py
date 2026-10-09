"""Shared plan validation, immutable preview and actual synthetic page parity."""

from collections.abc import Mapping
import copy
from dataclasses import FrozenInstanceError, replace
from hashlib import sha256
import importlib.util
from pathlib import Path
import random
import tempfile
import unittest
from unittest.mock import patch

from pypdf import PdfReader, PdfWriter


ROOT = Path(__file__).resolve().parents[2]


def load_module(name, path):
    spec = importlib.util.spec_from_file_location(name, path.resolve())
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


engine = load_module("wbs_plan_test_engine", ROOT / "engine/winbooksplit_engine.py")
fixtures = load_module("wbs_plan_test_fixtures", ROOT / "tests/fixtures/generate_pdf_fixtures.py")


def raw_plan(ranges=((0, 3), (3, 6), (6, 10)), pages=10, mode="manual"):
    return {"mode": mode, "total_pages": pages, "entries": [
        {"sequence": sequence, "title": f"Section {sequence}", "start": start, "end": end,
         "parent_id": None, "reason": "manual", "filename": f"{sequence:02d} - Section {sequence}.pdf", "warnings": []}
        for sequence, (start, end) in enumerate(ranges, 1)], "notices": [], "warnings": []}


class SharedPlanValidationTests(unittest.TestCase):
    def assert_invalid(self, plan):
        with self.assertRaises(engine.PlanError) as raised:
            engine.validate_plan(plan)
        self.assertTrue(raised.exception.code)
        self.assertTrue(str(raised.exception))

    def test_ac027_complete_partitions_have_derived_ranges_and_coverage(self):
        for ranges in (((0, 10),), ((0, 1), (1, 9), (9, 10)), ((0, 3), (3, 6), (6, 10))):
            with self.subTest(ranges=ranges):
                plan = engine.validate_plan(raw_plan(ranges))
                self.assertEqual(plan["ranges"], ranges)
                self.assertEqual(plan["coverage"], {"complete": True, "covered_pages": 10, "section_count": len(ranges)})

    def test_ac027_gap_overlap_empty_negative_reversed_and_overflow_ranges_are_rejected(self):
        cases = [(), ((0, 3), (4, 10)), ((0, 4), (3, 10)), ((0, 0), (0, 10)),
                 ((-1, 10),), ((0, 5), (5, 3), (3, 10)), ((0, 11),), ((1, 10),),
                 ((0, 9),), ((3, 10), (0, 3)), ((0, 4), (3, 6), (7, 10))]
        for ranges in cases:
            with self.subTest(ranges=ranges):
                self.assert_invalid(raw_plan(ranges))

    def test_page_counts_and_boundaries_require_integers_excluding_booleans(self):
        for pages in (0, -1, True, 10.0, "10", None):
            with self.subTest(pages=pages):
                self.assert_invalid(raw_plan(pages=pages))
        for key, value in (("start", False), ("start", 0.0), ("end", True), ("end", 3.0), ("end", "3")):
            with self.subTest(key=key, value=value):
                plan = raw_plan()
                plan["entries"][0][key] = value
                self.assert_invalid(plan)

    def test_cached_ranges_cannot_disagree_with_entries(self):
        plan = raw_plan()
        plan["ranges"] = [(0, 4), (4, 6), (6, 10)]
        self.assert_invalid(plan)

    def test_sequence_and_safe_case_insensitive_unique_filenames_are_required(self):
        for sequence in (2, True, 1.0):
            with self.subTest(sequence=sequence):
                plan = raw_plan()
                plan["entries"][0]["sequence"] = sequence
                self.assert_invalid(plan)
        for filename in ("../outside.pdf", "nested/chapter.pdf", "nested\\chapter.pdf", "C:\\outside.pdf", "bad\x00name.pdf"):
            with self.subTest(filename=filename):
                plan = raw_plan()
                plan["entries"][0]["filename"] = filename
                self.assert_invalid(plan)
        plan = raw_plan()
        plan["entries"][1]["filename"] = plan["entries"][0]["filename"][:-4].upper() + ".pdf"
        self.assert_invalid(plan)

    def test_invalid_mode_or_missing_entries_cannot_validate(self):
        for mode in (None, "unsupported", 1):
            with self.subTest(mode=mode):
                self.assert_invalid(raw_plan(mode=mode))
        for plan in (None, {}, {"mode": "manual", "total_pages": 10},
                     {"mode": "manual", "total_pages": 10, "entries": "not entries"}):
            with self.subTest(plan=plan):
                self.assert_invalid(plan)

    def test_seeded_complete_partitions_preserve_each_physical_index_exactly_once(self):
        generator = random.Random(27029)
        for number in range(100):
            pages = generator.randint(1, 100)
            internal = generator.sample(range(1, pages), generator.randint(0, pages - 1))
            boundaries = [0] + sorted(internal) + [pages]
            ranges = tuple(zip(boundaries, boundaries[1:]))
            with self.subTest(case=number, pages=pages):
                plan = engine.validate_plan(raw_plan(ranges, pages))
                self.assertEqual([page for start, end in plan["ranges"] for page in range(start, end)], list(range(pages)))
                self.assertEqual(plan["coverage"]["covered_pages"], pages)

    def test_validation_detaches_and_deep_freezes_the_callers_mutable_data(self):
        original = raw_plan()
        original["normalized_inputs"] = {"starts": [1, 4, 7]}
        original["entries"][0]["warnings"].append({"code": "synthetic", "message": "Original warning"})
        validated = engine.validate_plan(original)
        original["entries"][0]["end"] = 1
        original["normalized_inputs"]["starts"].append(9)
        original["entries"][0]["warnings"][0]["message"] = "Changed by caller"
        self.assertEqual(validated["entries"][0]["end"], 3)
        self.assertEqual(validated["normalized_inputs"]["starts"], (1, 4, 7))
        self.assertEqual(validated["entries"][0]["warnings"][0]["message"], "Original warning")
        with self.assertRaises(TypeError):
            validated["entries"][0]["end"] = 1
        with self.assertRaises(TypeError):
            validated["coverage"]["complete"] = False


class PreparedPlanTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.owned = tempfile.TemporaryDirectory(prefix="wbs-plan-unit-")
        cls.addClassCleanup(cls.owned.cleanup)
        cls.work = Path(cls.owned.name).resolve()
        cls.inputs = fixtures.generate_fixtures(cls.work / "fixtures")

    def test_all_modes_reject_injected_gap_or_empty_plans_before_any_writer(self):
        for mode, function, name in (("manual", "plan_manual_starts", "simple10"),
                                     ("1", "plan_level1", "simple10"), ("2", "plan_level2", "nested12")):
            source = self.inputs[name]
            before_source = source.read_bytes()
            reader = PdfReader(source)
            pages = len(reader.pages)
            base = engine.plan_manual_starts("4,7", pages) if mode == "manual" \
                else getattr(engine, function)(reader, pages)
            for defect in ("gap", "empty"):
                with self.subTest(mode=mode, defect=defect):
                    injected = copy.deepcopy(base)
                    ranges = [(0, 3), (4, pages)] if defect == "gap" else []
                    injected["ranges"] = ranges
                    if mode == "manual":
                        injected["starts"] = [start + 1 for start, end in ranges]
                    else:
                        injected["entries"] = raw_plan(ranges, pages, mode)["entries"]
                    output = self.work / ("injected-" + mode + "-" + defect)
                    output.mkdir()
                    neighbor = output / "owned-neighbor.txt"
                    neighbor.write_bytes(b"Synthetic neighbor must remain unchanged\n")
                    before_neighbor = neighbor.read_bytes()
                    with patch.object(engine, function, return_value=injected), \
                            patch.object(engine, "write_slice") as writer, patch.object(engine, "log") as log:
                        with self.assertRaises(SystemExit) as raised:
                            engine.split_pdf(str(source), str(output), mode, "4,7")
                    self.assertEqual(raised.exception.code, 6)
                    self.assertIn("invalid_plan", " ".join(str(call.args[0]) for call in log.call_args_list))
                    writer.assert_not_called()
                    self.assertEqual({path.name for path in output.iterdir()}, {neighbor.name})
                    self.assertEqual(neighbor.read_bytes(), before_neighbor)
                    self.assertEqual(source.read_bytes(), before_source)

    def test_unknown_or_unhashable_modes_are_rejected_before_source_probe_or_writer(self):
        for mode in ("unsupported", None, 1, [], {}):
            with self.subTest(mode=mode):
                with patch.object(engine, "PdfReader") as probe, patch.object(engine, "write_slice") as writer:
                    with self.assertRaises(engine.PlanError) as raised:
                        engine.prepare_split("unused-synthetic-path.pdf", mode, "1")
                self.assertEqual(raised.exception.code, "invalid_mode")
                probe.assert_not_called()
                writer.assert_not_called()

    def test_preparation_and_preview_are_deeply_read_only_and_never_write_chapters(self):
        before = {name: path.read_bytes() for name, path in self.inputs.items()}
        members = {path.relative_to(self.work) for path in self.work.rglob("*")}
        with patch.object(engine, "write_slice") as writer:
            prepared = engine.prepare_split(str(self.inputs["simple10"]), "manual", "7,04,4,1")
            preview = engine.preview_plan(prepared)
        writer.assert_not_called()
        self.assertIsInstance(preview, Mapping)
        self.assertIs(preview, prepared.plan)
        self.assertIn("normalized_inputs", preview)
        self.assertTrue(preview["notices"])
        identity = preview["source_identity"]
        self.assertEqual(identity["path"], str(self.inputs["simple10"]))
        self.assertEqual(identity["sha256"], sha256(before["simple10"]).hexdigest())
        self.assertEqual(identity["size_bytes"], len(before["simple10"]))
        self.assertEqual(identity["binding"], "reader_snapshot")
        with self.assertRaises(TypeError):
            preview["entries"][0]["start"] = 9
        with self.assertRaises(TypeError):
            preview["source_identity"]["sha256"] = "forged"
        with self.assertRaises((FrozenInstanceError, AttributeError, TypeError)):
            prepared.plan = engine.validate_plan(raw_plan())
        self.assertEqual({path.relative_to(self.work) for path in self.work.rglob("*")}, members)
        self.assertEqual({name: path.read_bytes() for name, path in self.inputs.items()}, before)

    def test_ac028_all_modes_execute_the_previewed_entries_without_replanning(self):
        for mode, name, data in (("manual", "simple10", "4,7"), ("1", "simple10", None), ("2", "nested12", None)):
            with self.subTest(mode=mode):
                source = self.inputs[name]
                before = source.read_bytes()
                prepared = engine.prepare_split(str(source), mode, data)
                preview = engine.preview_plan(prepared)
                output = self.work / ("parity-" + mode)
                output.mkdir()
                neighbor = output / "owned-neighbor.txt"
                neighbor.write_bytes(b"Synthetic neighbor preserved\n")
                with patch.object(engine, "plan_manual_starts", side_effect=AssertionError("Execution replanned manual")), \
                        patch.object(engine, "plan_level1", side_effect=AssertionError("Execution replanned Level 1")), \
                        patch.object(engine, "plan_level2", side_effect=AssertionError("Execution replanned Level 2")), patch.object(engine, "log"):
                    result = engine.execute_split(prepared, str(output))
                self.assertEqual(result["written_count"], len(preview["entries"]))
                self.assertEqual(result["mode"], mode)
                self.assertEqual(result["total_pages"], preview["total_pages"])
                self.assertEqual(result["source_identity"], preview["source_identity"])
                self.assertEqual(result["coverage"], preview["coverage"])
                self.assertEqual([entry["filename"] for entry in result["outputs"]],
                                 [entry["filename"] for entry in preview["entries"]])
                published = Path(result["final_directory"])
                observed = []
                for entry, actual in zip(preview["entries"], result["outputs"]):
                    ids = fixtures.page_ids(published / entry["filename"])
                    self.assertEqual(ids, list(range(entry["start"] + 1, entry["end"] + 1)))
                    self.assertEqual(actual["start"], entry["start"])
                    self.assertEqual(actual["end"], entry["end"])
                    self.assertEqual(actual["page_count"], entry["end"] - entry["start"])
                    observed.extend(ids)
                self.assertEqual(observed, list(range(1, preview["total_pages"] + 1)))
                self.assertEqual({path.name for path in published.iterdir()},
                                 {entry["filename"] for entry in preview["entries"]} |
                                 {".WinBookSplit-owner.json", result["manifest_filename"]})
                self.assertEqual({path.name for path in output.iterdir()}, {published.name, neighbor.name})
                self.assertEqual(neighbor.read_bytes(), b"Synthetic neighbor preserved\n")
                self.assertEqual(source.read_bytes(), before)
                with self.assertRaises(TypeError):
                    result["written_count"] = 0

    def test_plain_or_forged_prepared_objects_are_rejected_before_writer(self):
        prepared = engine.prepare_split(str(self.inputs["simple10"]), "manual", "4,7")
        other = engine.prepare_split(str(self.inputs["simple10"]), "manual", "1")
        altered = replace(prepared, plan=other.plan)
        swapped_reader = replace(prepared, _reader=other._reader)
        output = self.work / "forged-output"
        output.mkdir()
        for forged in (None, raw_plan(), engine.preview_plan(prepared), altered, swapped_reader):
            with self.subTest(forged_type=type(forged).__name__):
                with patch.object(engine, "write_slice") as writer:
                    with self.assertRaises(engine.PlanError) as raised:
                        engine.execute_split(forged, str(output))
                self.assertEqual(raised.exception.code, "invalid_prepared_split")
                writer.assert_not_called()
                self.assertEqual(list(output.iterdir()), [])

    def test_modified_private_snapshot_is_rejected_before_preview_or_writer(self):
        prepared = engine.prepare_split(str(self.inputs["simple10"]), "manual", "1")
        stream = prepared._reader.stream
        stream.seek(0)
        stream.write(b"Changed owned snapshot bytes")
        output = self.work / "mutated-snapshot-output"
        output.mkdir()
        with self.assertRaises(engine.PlanError) as raised:
            engine.preview_plan(prepared)
        self.assertEqual(raised.exception.code, "source_changed")
        with patch.object(engine, "write_slice") as writer:
            with self.assertRaises(engine.PlanError) as raised:
                engine.execute_split(prepared, str(output))
        self.assertEqual(raised.exception.code, "source_changed")
        writer.assert_not_called()
        self.assertEqual(list(output.iterdir()), [])

    def test_ac029_replaced_or_deleted_source_paths_still_execute_original_snapshot(self):
        original = self.inputs["simple10"].read_bytes()
        writer = PdfWriter()
        for page in reversed(PdfReader(self.inputs["simple10"]).pages):
            writer.add_page(page)
        replacement_path = self.work / "reverse-replacement.pdf"
        with replacement_path.open("xb") as stream:
            writer.write(stream)
        for change in ("replace", "delete", "repoint"):
            with self.subTest(change=change):
                source = self.work / ("mutable-" + change + ".pdf")
                source.write_bytes(original)
                prepared = engine.prepare_split(str(source), "manual", "4,7")
                preview = engine.preview_plan(prepared)
                if change == "replace":
                    source.write_bytes(replacement_path.read_bytes())
                elif change == "delete":
                    source.unlink()
                else:
                    source.rename(self.work / "original-moved.pdf")
                    source.write_bytes(replacement_path.read_bytes())
                output = self.work / ("source-change-" + change)
                output.mkdir()
                with patch.object(engine, "log"):
                    result = engine.execute_split(prepared, str(output))
                self.assertEqual(result["source_identity"]["sha256"], sha256(original).hexdigest())
                observed = [page for entry in preview["entries"]
                            for page in fixtures.page_ids(Path(result["final_directory"]) / entry["filename"])]
                self.assertEqual(observed, list(range(1, 11)))
                if change != "delete":
                    self.assertEqual(source.read_bytes(), replacement_path.read_bytes())
                else:
                    self.assertFalse(source.exists())
        self.assertEqual(self.inputs["simple10"].read_bytes(), original)

    def test_existing_same_named_base_chapter_is_preserved_by_unique_run_publication(self):
        prepared = engine.prepare_split(str(self.inputs["simple10"]), "manual", "4,7")
        preview = engine.preview_plan(prepared)
        output = self.work / "existing-output"
        output.mkdir()
        sentinel = output / preview["entries"][-1]["filename"]
        sentinel.write_bytes(b"Owned existing chapter must not be overwritten\n")
        before = sentinel.read_bytes()
        with patch.object(engine, "log"):
            result = engine.execute_split(prepared, str(output))
        published = Path(result["final_directory"])
        self.assertEqual({path.name for path in output.iterdir()}, {sentinel.name, published.name})
        self.assertEqual(fixtures.page_ids(published / sentinel.name), list(range(7, 11)))
        self.assertEqual(sentinel.read_bytes(), before)

    def test_source_named_like_a_chapter_is_preserved_by_unique_run_publication(self):
        output = self.work / "source-output-alias"
        output.mkdir()
        source = output / "01 - Section (Page 1-10).pdf"
        source.write_bytes(self.inputs["simple10"].read_bytes())
        before = source.read_bytes()
        prepared = engine.prepare_split(str(source), "manual", "1")
        with patch.object(engine, "log"):
            result = engine.execute_split(prepared, str(output))
        published = Path(result["final_directory"])
        self.assertNotEqual(published / source.name, source)
        self.assertEqual(fixtures.page_ids(published / source.name), list(range(1, 11)))
        self.assertEqual(source.read_bytes(), before)
        self.assertEqual({path.name for path in output.iterdir()}, {source.name, published.name})

    def test_writer_exclusive_creation_cannot_clobber_a_file_appearing_after_preflight(self):
        prepared = engine.prepare_split(str(self.inputs["simple10"]), "manual", "1")
        output = self.work / "collision-race-output"
        output.mkdir()
        targets = []
        sentinel = b"Synthetic competing file created after preflight\n"
        original_write = engine.write_slice

        def competing_write(reader, start, end, path, *, on_created=None):
            target = Path(path)
            targets.append(target)
            self.assertEqual(target.name, engine.preview_plan(prepared)["entries"][0]["filename"])
            with Path(path).open("xb") as stream:
                stream.write(sentinel)
            return original_write(reader, start, end, path, on_created=on_created)

        with patch.object(engine, "write_slice", side_effect=competing_write), patch.object(engine, "log"):
            with self.assertRaises((engine.PlanError, FileExistsError)):
                engine.execute_split(prepared, str(output))
        self.assertEqual(len(targets), 1)
        self.assertEqual(targets[0].read_bytes(), sentinel)
        self.assertFalse(any(path.is_dir() and not path.name.startswith(".WinBookSplit-")
                             for path in output.iterdir()))


if __name__ == "__main__":
    unittest.main(verbosity=2)
