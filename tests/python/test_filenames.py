"""Literal Unicode filenames, chronological numbering and Windows path budgets."""

from hashlib import sha256
import importlib.util
import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, NumberObject


ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location("wbs_filename_test_engine", ROOT / "engine/winbooksplit_engine.py")
engine = importlib.util.module_from_spec(spec)
spec.loader.exec_module(engine)


def units(value):
    return len(str(value).encode("utf-16-le")) // 2


def generate(path, titles, level=1):
    """Actual owned pages contain stable IDs unrelated to title sanitation."""
    writer = PdfWriter()
    parent = None
    for index, title in enumerate(titles):
        page = writer.add_blank_page(width=100, height=100)
        page[NameObject("/SyntheticPageID")] = NumberObject(index + 1)
        if level == 2 and parent is None:
            parent = writer.add_outline_item("Synthetic parent", 0)
        writer.add_outline_item(title, index, parent=parent)
    with path.open("xb") as stream:
        writer.write(stream)


def raw_plan(names):
    return {"mode": "manual", "total_pages": len(names), "entries": [
        {"sequence": index, "title": "Synthetic", "start": index - 1, "end": index,
         "parent_id": None, "reason": "manual", "filename": name, "warnings": []}
        for index, name in enumerate(names, 1)]}


class FilenameValidationTests(unittest.TestCase):
    def assert_invalid(self, names):
        with self.assertRaises(engine.PlanError) as caught:
            engine.validate_plan(raw_plan(names))
        self.assertEqual(caught.exception.code, "invalid_plan")

    def test_ac037_raw_device_names_and_superscripts_are_rejected(self):
        for stem in ("CON", "prn", "AUX", "nul", "CONIN$", "conout$", "COM1", "COM9", "LPT1", "LPT9",
                     "COM¹", "COM²", "COM³", "LPT¹", "LPT²", "LPT³", "con.extra", "NUL .extra"):
            with self.subTest(stem=stem):
                self.assert_invalid([stem + ".pdf"])
        for stem in ("COM0", "COM10", "LPT0", "LPT10", "CONtent", "01 - CON"):
            with self.subTest(allowed=stem):
                self.assertEqual(engine.validate_plan(raw_plan([stem + ".pdf"]))["entries"][0]["filename"], stem + ".pdf")

    def test_ac037_controls_traversal_and_terminal_filename_spaces_are_rejected(self):
        for name in ("../outside.pdf", "nested\\outside.pdf", "C:outside.pdf", " title.pdf", "title.pdf ",
                     "bad\x01.pdf", "bad\x7f.pdf", "bad\x85.pdf", "bad\u202e.pdf", "bad\ud800.pdf"):
            with self.subTest(name=repr(name)):
                self.assert_invalid([name])

    def test_ac037_casefold_duplicate_validation_includes_unicode(self):
        self.assert_invalid(["01 - Straße.pdf", "01 - STRASSE.pdf"])
        self.assert_invalid(["same.pdf", "SAME.pdf"])

    def test_ac039_component_budget_counts_astral_utf16_units(self):
        for name in ("a" * 251 + ".pdf", "😀" * 125 + "a.pdf"):
            with self.subTest(valid=units(name)):
                self.assertEqual(units(name), 255)
                engine.validate_plan(raw_plan([name]))
        for name in ("a" * 252 + ".pdf", "😀" * 126 + ".pdf"):
            with self.subTest(invalid=units(name)):
                self.assert_invalid([name])


class ActualFilenameTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.owned = tempfile.TemporaryDirectory(prefix="wbs-n-")
        cls.addClassCleanup(cls.owned.cleanup)
        cls.work = Path(cls.owned.name).resolve()
        cls.serial = 0

    def setUp(self):
        type(self).serial += 1
        self.case = self.work / f"c{self.serial:02d}"
        self.case.mkdir()
        self.base = self.case / "out"
        self.base.mkdir()
        self.neighbor = self.base / "synthetic-neighbor.txt"
        self.neighbor.write_bytes(b"Owned unrelated neighbor\n")
        self.source = self.case / "book.pdf"

    def prepare(self, titles, mode="1", base=None, source=None):
        self.source = source or self.source
        generate(self.source, titles, level=2 if mode == "2" else 1)
        self.source_hash = sha256(self.source.read_bytes()).hexdigest()
        manual = ",".join(str(index) for index in range(1, len(titles) + 1)) if mode == "manual" else None
        return engine.prepare_split(self.source, mode, manual, output_base=base)

    def assert_preserved(self):
        self.assertEqual(sha256(self.source.read_bytes()).hexdigest(), self.source_hash)
        self.assertEqual(self.neighbor.read_bytes(), b"Owned unrelated neighbor\n")

    def assert_complete(self, prepared, expected_names):
        preview = engine.preview_plan(prepared)
        self.assertEqual([entry["filename"] for entry in preview["entries"]], expected_names)
        with patch.object(engine, "plan_level1", side_effect=AssertionError("execution replanned")), \
                patch.object(engine, "plan_level2", side_effect=AssertionError("execution replanned")), \
                patch.object(engine, "plan_manual_starts", side_effect=AssertionError("execution replanned")):
            result = engine.execute_split(prepared, self.base)
        final = Path(result["final_directory"])
        self.assertEqual([entry["filename"] for entry in result["outputs"]], expected_names)
        manifest = json.loads((final / "WinBookSplit_Manifest.json").read_text(encoding="utf-8"))
        self.assertEqual([entry["filename"] for entry in manifest["outputs"]], expected_names)
        ids = [int(page["/SyntheticPageID"]) for entry in result["outputs"]
               for page in PdfReader(final / entry["filename"]).pages]
        self.assertEqual(ids, list(range(1, preview["total_pages"] + 1)))
        self.assertEqual(result["written_count"], len(expected_names))
        self.assertEqual(final.parent, self.base)
        self.assertLessEqual(units(final), 247)
        for entry in result["outputs"]:
            self.assertLessEqual(units(final / entry["filename"]), 259)
        self.assert_preserved()
        return result

    def long_base(self, length):
        # Owned non-reparse paths only; component lengths also remain bounded.
        base = self.case / "d"
        base.mkdir(exist_ok=True)
        remaining = length - units(base) - 1
        self.assertGreater(remaining, 0)
        base = base / ("x" * remaining)
        base.mkdir()
        self.assertEqual(units(base), length)
        return base

    def test_ac037_actual_controls_forbidden_and_unicode_titles_are_safe(self):
        titles = ['A<>:"/\\|?*B\x01\x7f\x85\u202e', "  one   two .  ", "日本語 åäö 😀", "CON", "LPT¹"]
        prepared = self.prepare(titles, base=self.base)
        self.assertEqual([entry["title"] for entry in prepared.plan["entries"]], titles)
        self.assert_complete(prepared, ["01 - AB.pdf", "02 - one two.pdf", "03 - 日本語 åäö 😀.pdf",
                                       "04 - CON.pdf", "05 - LPT¹.pdf"])

    def test_ac037_punctuation_and_empty_titles_get_stable_fallbacks(self):
        prepared = self.prepare(['<>:"/\\|?*', "... !!!", "\x01\x7f", "...   "], base=self.base)
        self.assert_complete(prepared, [f"{index:02d} - Section {index}.pdf" for index in range(1, 5)])

    def test_ac037_truncation_strips_terminal_dots_and_preserves_astral_boundaries(self):
        prepared = self.prepare(["A" * 49 + ".remainder", "B" * 49 + "😀more", "😀" * 25 + "rest"], base=self.base)
        self.assert_complete(prepared, ["01 - " + "A" * 49 + ".pdf", "02 - " + "B" * 49 + ".pdf",
                                       "03 - " + "😀" * 25 + ".pdf"])

    def test_ac038_actual_120_sections_have_chronological_lexical_names_in_all_modes(self):
        for mode in ("manual", "1", "2"):
            with self.subTest(mode=mode):
                self.base = self.case / ("mode-" + mode)
                self.base.mkdir()
                self.neighbor = self.base / "synthetic-neighbor.txt"
                self.neighbor.write_bytes(b"Owned unrelated neighbor\n")
                prepared = self.prepare([f"Chapter {index}" for index in range(1, 121)], mode,
                                        base=self.base, source=self.case / (mode + ".pdf"))
                names = [entry["filename"] for entry in prepared.plan["entries"]]
                self.assertEqual(names, sorted(names))
                self.assertEqual([name.split(" - ", 1)[0] for name in names], [f"{n:03d}" for n in range(1, 121)])
                self.assertEqual(len(set(name.casefold() for name in names)), 120)
                self.assert_complete(prepared, names)

    def test_ac038_small_section_counts_keep_two_digits(self):
        prepared = self.prepare(["First", "Second"], base=self.base)
        self.assert_complete(prepared, ["01 - First.pdf", "02 - Second.pdf"])

    def test_ac039_bound_preview_shortens_names_and_stem_before_writing(self):
        self.base = self.long_base(165)
        self.neighbor = self.base / "synthetic-neighbor.txt"
        self.neighbor.write_bytes(b"Owned unrelated neighbor\n")
        prepared = self.prepare(["Unicode 日本語 title " * 8], base=self.base,
                                source=self.case / ("long-source-" * 8 + ".pdf"))
        preview = engine.preview_plan(prepared)
        naming = preview["output_naming"]
        self.assertEqual(naming["resolved_base"], os.path.realpath(self.base))
        self.assertLess(units(naming["run_stem"]), 64)
        self.assertLess(naming["filename_budget"], 59)
        with self.assertRaises(TypeError):
            naming["filename_budget"] = 255
        self.assertEqual(list(self.base.iterdir()), [self.neighbor])
        observed = []
        original = engine.write_slice
        def observe(*args, **kwargs):
            observed.append(Path(args[3]))
            return original(*args, **kwargs)
        names = [entry["filename"] for entry in preview["entries"]]
        with patch.object(engine, "write_slice", side_effect=observe):
            result = self.assert_complete(prepared, names)
        self.assertTrue(observed)
        for path in observed:
            self.assertLessEqual(units(path), 259)
            self.assertLessEqual(units(path.parent), 247)
        final = Path(result["final_directory"])
        self.assertLessEqual(units(final / engine.MANIFEST_FILENAME), 259)
        self.assertLessEqual(units(final / engine.OWNER_FILENAME), 259)

    def test_ac039_impossible_bound_base_fails_before_any_reservation(self):
        base = self.long_base(180)
        generate(self.source, ["One"])
        self.source_hash = sha256(self.source.read_bytes()).hexdigest()
        with patch.object(engine, "OutputRun") as reserve, patch.object(engine, "write_slice") as writer:
            with self.assertRaises(engine.PlanError) as caught:
                engine.prepare_split(self.source, "1", output_base=base)
        self.assertEqual(caught.exception.code, "output_path_too_long")
        self.assertIn("shorter output base", str(caught.exception))
        reserve.assert_not_called()
        writer.assert_not_called()
        self.assertEqual(list(base.iterdir()), [])
        self.assert_preserved()

    def test_ac039_legacy_unbound_long_names_are_rejected_without_rewriting(self):
        prepared = self.prepare(["Long title " * 8])
        names = [entry["filename"] for entry in prepared.plan["entries"]]
        base = self.long_base(165)
        with patch.object(engine, "OutputRun") as reserve, patch.object(engine, "write_slice") as writer:
            with self.assertRaises(engine.PlanError) as caught:
                engine.execute_split(prepared, base)
        self.assertEqual(caught.exception.code, "output_path_too_long")
        reserve.assert_not_called()
        writer.assert_not_called()
        self.assertEqual([entry["filename"] for entry in engine.preview_plan(prepared)["entries"]], names)
        self.assertEqual(list(base.iterdir()), [])
        self.assert_preserved()

    def test_ac039_bound_destination_change_is_rejected_before_allocation(self):
        prepared = self.prepare(["First"], base=self.base)
        other = self.case / "other"
        other.mkdir()
        with patch.object(engine, "OutputRun") as reserve, patch.object(engine, "write_slice") as writer:
            with self.assertRaises(engine.PlanError) as caught:
                engine.execute_split(prepared, other)
        self.assertEqual(caught.exception.code, "output_destination_changed")
        reserve.assert_not_called()
        writer.assert_not_called()
        self.assertEqual(list(other.iterdir()), [])
        self.assertEqual(list(self.base.iterdir()), [self.neighbor])
        self.assert_preserved()

    def test_ac039_run_split_uses_destination_aware_plan_and_impossible_base_diagnostic(self):
        generate(self.source, ["Long title " * 8])
        self.source_hash = sha256(self.source.read_bytes()).hexdigest()
        base = self.long_base(165)
        result = engine.run_split(self.source, base, "1")
        self.assertEqual((result["status"], result["exit_code"], result["written_count"]), ("success", 0, 1))
        self.assertLess(units(result["execution"]["outputs"][0]["filename"]), 59)
        impossible = self.long_base(180)
        with patch.object(engine, "OutputRun") as reserve:
            rejected = engine.run_split(self.source, impossible, "1")
        self.assertEqual((rejected["status"], rejected["code"], rejected["exit_code"], rejected["written_count"]),
                         ("error", "output_path_too_long", 2, 0))
        reserve.assert_not_called()
        self.assertIsNone(rejected["diagnostic"])
        self.assertEqual(list(impossible.iterdir()), [])
        self.assert_preserved()


if __name__ == "__main__":
    unittest.main()
