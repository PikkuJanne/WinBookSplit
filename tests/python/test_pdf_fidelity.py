"""Authored ordinary PDF metadata/navigation units; rendered proof is separate."""

from hashlib import sha256
import importlib.util
from io import StringIO
from contextlib import redirect_stdout
from pathlib import Path
import re
import tempfile
import unittest
from unittest.mock import patch

from pypdf import PdfReader, PdfWriter
from pypdf.generic import (ArrayObject, BooleanObject, DecodedStreamObject, DictionaryObject,
    FloatObject, IndirectObject, NameObject, NullObject, NumberObject, TextStringObject)

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location("wbs_pdf_fidelity_units", ROOT / "engine/winbooksplit_engine.py")
engine = importlib.util.module_from_spec(spec)
spec.loader.exec_module(engine)


def authored_pdf(path, *, malformed=False, metadata_bad=False, catalog=False, hidden_action=False, bad_geometry=False, page_actions=False):
    writer = PdfWriter()
    for number in range(4):
        page = writer.add_blank_page(300 + number * 20, 400 + number * 10)
        font = DictionaryObject({NameObject("/Type"): NameObject("/Font"), NameObject("/Subtype"): NameObject("/Type1"),
            NameObject("/BaseFont"): NameObject("/Helvetica")})
        page[NameObject("/Resources")] = DictionaryObject({NameObject("/Font"): DictionaryObject({NameObject("/F1"): writer._add_object(font)})})
        content = DecodedStreamObject()
        content.set_data(f"BT /F1 16 Tf 20 300 Td (AUTHORED-PAGE-{number + 1:02}) Tj ET\n".encode())
        page[NameObject("/Contents")] = writer._add_object(content)
        page.cropbox.lower_left = (10, 20)
        if number == 1:
            page.rotate(90)
    writer.add_metadata({"/Title": "Authored source Å 日本", "/Author": "Authored author Å 日本", "/Producer": "Original signed source claim"})
    if metadata_bad:
        writer._info[NameObject("/Author")] = NumberObject(7)
        writer._info[NameObject("/Title")] = TextStringObject("Å\n日本\u202e" + "X" * 300)
    parent = writer.add_outline_item("Chapter A Å", 0)
    writer.add_outline_item("Child A", 1, parent=parent)
    parent = writer.add_outline_item("Chapter B 日本", 2)
    writer.add_outline_item("Child B", 3, parent=parent)
    writer.add_named_destination("inside-A", 0)
    writer.add_named_destination("outside-A", 3)
    rect = lambda: ArrayObject([FloatObject(v) for v in (10, 10, 80, 30)])
    def link(page, name, destination, action=False):
        annotation = DictionaryObject({NameObject("/Type"): NameObject("/Annot"), NameObject("/Subtype"): NameObject("/Link"),
            NameObject("/Rect"): rect(), NameObject("/NM"): TextStringObject(name)})
        if action:
            annotation[NameObject("/A")] = DictionaryObject({NameObject("/S"): NameObject("/GoTo"), NameObject("/D"): destination})
        else:
            annotation[NameObject("/Dest")] = destination
        writer.add_annotation(page, annotation)
    link(0, "direct-intra", ArrayObject([writer.pages[1].indirect_reference, NameObject("/FitH"), FloatObject(123)]))
    link(0, "direct-cross", ArrayObject([writer.pages[2].indirect_reference, NameObject("/Fit")]))
    link(1, "named-intra", TextStringObject("inside-A"), True)
    link(1, "named-cross", TextStringObject("outside-A"), True)
    square = writer.add_annotation(1, DictionaryObject({NameObject("/Type"): NameObject("/Annot"), NameObject("/Subtype"): NameObject("/Square"),
        NameObject("/Rect"): rect(), NameObject("/NM"): TextStringObject("static-square"), NameObject("/Contents"): TextStringObject("Authored static square")}))
    square[NameObject("/P")] = writer.pages[3].indirect_reference
    if hidden_action:
        square[NameObject("/AP")] = DictionaryObject({NameObject("/N"): DictionaryObject({NameObject("/S"): NameObject("/JavaScript"),
            NameObject("/JS"): TextStringObject("Authored hidden action sentinel, never executed")})})
    if bad_geometry:
        for name, subtype, rectangle, quads in (("missing-rect", "/Square", None, None),
            ("text-rect", "/Text", ArrayObject([NumberObject(1), TextStringObject("bad"), NumberObject(3), NumberObject(4)]), None),
            ("bool-rect", "/Square", ArrayObject([BooleanObject(True), NumberObject(2), NumberObject(3), NumberObject(4)]), None),
            ("missing-quads", "/Highlight", rect(), None),
            ("short-quads", "/Highlight", rect(), ArrayObject([NumberObject(n) for n in range(7)]))):
            bad = DictionaryObject({NameObject("/Subtype"): NameObject(subtype), NameObject("/NM"): TextStringObject(name)})
            if rectangle is not None:
                bad[NameObject("/Rect")] = rectangle
            if quads is not None:
                bad[NameObject("/QuadPoints")] = quads
            writer.add_annotation(1, bad)
    if malformed:
        for name, value in (("null", NullObject()), ("unresolved", TextStringObject("missing")),
            ("negative", ArrayObject([NumberObject(-1), NameObject("/Fit")])),
            ("boolean", ArrayObject([BooleanObject(True), NameObject("/Fit")])),
            ("bad-fit", ArrayObject([writer.pages[0].indirect_reference, NameObject("/Bogus")])),
            ("bad-args", ArrayObject([writer.pages[0].indirect_reference, NameObject("/FitH"), TextStringObject("bad")]))):
            link(0, name, value)
        uri = writer.add_annotation(0, DictionaryObject({NameObject("/Subtype"): NameObject("/Link"), NameObject("/Rect"): rect(),
            NameObject("/A"): DictionaryObject({NameObject("/S"): NameObject("/URI"), NameObject("/URI"): TextStringObject("https://example.invalid/")})}))
        uri[NameObject("/NM")] = TextStringObject("unsupported-uri")
    if catalog:
        writer.root_object[NameObject("/OpenAction")] = DictionaryObject({NameObject("/S"): NameObject("/JavaScript"),
            NameObject("/JS"): TextStringObject("Authored inert catalog sentinel, never executed")})
        writer.root_object[NameObject("/Lang")] = TextStringObject("en-US")
        writer.pages[0][NameObject("/B")] = ArrayObject([writer.pages[3].indirect_reference])
        square[NameObject("/IRT")] = writer.pages[3].indirect_reference
    if page_actions:
        for page, target in ((0, 1), (1, 3)):
            writer.pages[page][NameObject("/AA")] = DictionaryObject({NameObject("/O"): DictionaryObject({
                NameObject("/S"): NameObject("/GoTo"), NameObject("/D"): ArrayObject([
                    writer.pages[target].indirect_reference, NameObject("/Fit")])})})
    with path.open("xb") as stream:
        writer.write(stream)
    return path


def annotations(reader):
    return {str(reference.get_object().get("/NM")): (index, reference.get_object())
        for index, page in enumerate(reader.pages) for reference in page.get("/Annots", [])}


def serialized_pages(reader):
    found = []
    for generation, objects in reader.xref.items():
        if generation == 65535:
            continue
        for number in objects:
            if number == 0:
                continue
            obj = reader.get_object(IndirectObject(number, generation, reader))
            if isinstance(obj, dict) and obj.get("/Type") == "/Page":
                found.append(number)
    return found


class PdfFidelityTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory(prefix="wbs-fidelity-unit-")
        self.addCleanup(self.temporary.cleanup)
        self.work = Path(self.temporary.name).resolve()
        self.source = authored_pdf(self.work / "authored.pdf")
        self.original = self.source.read_bytes()
        self.output = self.work / "out"
        self.output.mkdir()
        self.neighbor = self.output / "neighbor.txt"
        self.neighbor.write_bytes(b"Preserve authored neighbor\n")

    def run_job(self, mode="1", starts=None, source=None):
        selected = source or self.source
        selected_before = Path(selected).read_bytes()
        with redirect_stdout(StringIO()):
            result = engine.run_split(selected, self.output, mode, starts)
        self.assertEqual(result["exit_code"], 0, result)
        self.assertEqual(self.source.read_bytes(), self.original)
        self.assertEqual(Path(selected).read_bytes(), selected_before)
        self.assertEqual(self.neighbor.read_bytes(), b"Preserve authored neighbor\n")
        final = Path(result["execution"]["final_directory"])
        return result, [(entry, PdfReader(final / entry["filename"])) for entry in result["execution"]["outputs"]]

    def test_safe_metadata_start_bookmark_and_original_page_geometry_all_modes(self):
        original = PdfReader(self.source)
        for mode, starts in (("manual", "1,3"), ("1", None), ("2", None)):
            with self.subTest(mode=mode):
                result, outputs = self.run_job(mode, starts)
                physical = []
                for entry, reader in outputs:
                    expected = re.sub(r"^[0-9]+ - ", "", entry["filename"][:-4])
                    self.assertEqual(reader.metadata.title, expected)
                    self.assertEqual(reader.metadata.author, "Authored author Å 日本")
                    self.assertEqual(reader.metadata.subject, "Source: Authored source Å 日本")
                    self.assertEqual(reader.metadata.creator, "WinBookSplit")
                    self.assertEqual(reader.metadata.producer, "WinBookSplit 1.0.0-dev (pypdf 6.19.0)")
                    self.assertEqual(len(reader.outline), 1)
                    self.assertEqual(reader.outline[0].title, expected)
                    self.assertEqual(reader.get_destination_page_number(reader.outline[0]), 0)
                    self.assertEqual(len(serialized_pages(reader)), len(reader.pages))
                    for index, page in enumerate(reader.pages, entry["start"]):
                        before = original.pages[index]
                        self.assertEqual(page.extract_text(), before.extract_text())
                        self.assertEqual(list(page.mediabox), list(before.mediabox))
                        self.assertEqual(list(page.cropbox), list(before.cropbox))
                        self.assertEqual(page.rotation, before.rotation)
                        physical.append(page.extract_text().strip())
                self.assertEqual(physical, [f"AUTHORED-PAGE-{index:02}" for index in range(1, 5)])
                self.assertEqual(result["warnings"], result["plan"]["warnings"])

    def test_local_direct_named_links_rebase_and_cross_links_cannot_embed_other_pages(self):
        result, outputs = self.run_job()
        first = outputs[0][1]
        records = annotations(first)
        self.assertEqual(set(records), {"direct-intra", "named-intra", "static-square"})
        for name, page, fit, args in (("direct-intra", 1, "/FitH", [123.0]), ("named-intra", 0, "/FitH", [826.0])):
            dest = records[name][1]["/Dest"]
            self.assertEqual(dest[0], first.pages[page].indirect_reference)
            self.assertEqual(dest[1], fit)
            self.assertEqual(list(dest[2:]), args)
            self.assertNotIn("/A", records[name][1])
        index, static = records["static-square"]
        self.assertEqual(static.raw_get("/P"), first.pages[index].indirect_reference)
        self.assertEqual(static["/Contents"], "Authored static square")
        self.assertEqual(len(serialized_pages(first)), 2)
        self.assertNotIn(b"AUTHORED-PAGE-03", first.stream.getvalue())
        self.assertEqual([warning["code"] for warning in result["warnings"]], ["cross_chapter_link_dropped"])
        self.assertIn("Count: 2.", result["warnings"][0]["message"])

    def test_preview_warns_before_reservation_or_writes_and_remains_same_execution_plan(self):
        with patch.object(engine, "OutputRun", side_effect=AssertionError("Preview cannot reserve")), \
                patch.object(engine, "write_slice", side_effect=AssertionError("Preview cannot write")), redirect_stdout(StringIO()) as displayed:
            preview = engine.run_split(self.source, self.output, "1", preview=True)
        self.assertEqual(preview["status"], "preview")
        self.assertIn("cross_chapter_link_dropped", displayed.getvalue())
        self.assertNotIn("outline None", displayed.getvalue())
        self.assertEqual({p.name for p in self.output.iterdir()}, {self.neighbor.name})
        result, _ = self.run_job()
        self.assertEqual(preview["plan"], result["plan"])

    def test_malformed_null_unresolved_numeric_and_external_destinations_are_warned_and_dropped(self):
        source = authored_pdf(self.work / "malformed.pdf", malformed=True)
        result, outputs = self.run_job(source=source)
        self.assertEqual(set(annotations(outputs[0][1])), {"direct-intra", "named-intra", "static-square"})
        warnings = {item["code"]: item["message"] for item in result["warnings"]}
        self.assertIn("Count: 7.", warnings["navigation_link_dropped"])
        self.assertIn("Count: 2.", warnings["cross_chapter_link_dropped"])
        self.assertTrue(all(warning["source_order"] is None and warning["depth"] is None for warning in result["warnings"]))

    def test_metadata_projection_omits_nontext_and_bounds_controls_without_false_source_producer(self):
        source = authored_pdf(self.work / "metadata.pdf", metadata_bad=True)
        result, outputs = self.run_job(source=source)
        for _, reader in outputs:
            self.assertNotIn("/Author", reader.metadata)
            subject = reader.metadata.subject
            self.assertEqual(subject, "Source: " + ("Å日本" + "X" * 300)[:256])
            self.assertNotIn("\u202e", subject)
        self.assertTrue({"metadata_omitted", "metadata_normalized"} <= {item["code"] for item in result["warnings"]})
        self.assertEqual(engine._metadata_text("Å日本\ud800\u200b\x00"), "Å日本")

    def test_catalog_articles_and_annotation_backlinks_are_not_wholesale_copied(self):
        source = authored_pdf(self.work / "catalog.pdf", catalog=True)
        result, outputs = self.run_job(source=source)
        first = outputs[0][1]
        self.assertNotIn("/OpenAction", first.root_object)
        self.assertNotIn("/Lang", first.root_object)
        self.assertNotIn("/B", first.pages[0])
        self.assertNotIn("/IRT", annotations(first)["static-square"][1])
        self.assertEqual(len(serialized_pages(first)), 2)
        self.assertTrue({"article_navigation_dropped", "annotation_relation_dropped"} <= {item["code"] for item in result["warnings"]})

    def test_annotation_inspection_limit_fails_before_output_and_preserves_input(self):
        with patch.object(engine, "MAX_PDF_ANNOTATIONS", 1), patch.object(engine, "OutputRun") as reserve, \
                patch.object(engine, "write_slice") as write, redirect_stdout(StringIO()):
            result = engine.run_split(self.source, self.output, "1")
        self.assertEqual(result["code"], "invalid_document")
        self.assertEqual(result["written_count"], 0)
        reserve.assert_not_called()
        write.assert_not_called()
        self.assertEqual({p.name for p in self.output.iterdir()}, {self.neighbor.name})
        self.assertEqual(self.source.read_bytes(), self.original)

    def test_annotation_projection_cannot_copy_type_less_nested_action_dictionary(self):
        source = authored_pdf(self.work / "nested-action.pdf", hidden_action=True)
        result, outputs = self.run_job(source=source)
        self.assertNotIn("static-square", annotations(outputs[0][1]))
        self.assertIn("annotation_dropped", {warning["code"] for warning in result["warnings"]})
        self.assertNotIn(b"Authored hidden action sentinel", outputs[0][1].stream.getvalue())
        self.assertEqual(len(serialized_pages(outputs[0][1])), 2)

    def test_malformed_rectangles_and_highlight_quadpoints_are_warned_before_publication(self):
        source = authored_pdf(self.work / "bad-annotation-geometry.pdf", bad_geometry=True)
        result, outputs = self.run_job(source=source)
        self.assertEqual(set(annotations(outputs[0][1])), {"direct-intra", "named-intra", "static-square"})
        warning = next(w for w in result["warnings"] if w["code"] == "annotation_dropped")
        self.assertIn("Count: 5.", warning["message"])
        self.assertEqual(len(serialized_pages(outputs[0][1])), 2)

    def test_page_actions_cannot_clone_local_or_outside_pages_and_source_is_immutable(self):
        source = authored_pdf(self.work / "page-actions.pdf", page_actions=True)
        original_bytes = source.read_bytes()
        prepared = engine.prepare_split(source, "1", output_base=self.output)
        captured_actions = [str(page.get("/AA")) for page in prepared._reader.pages]
        captured_bytes = prepared._reader.stream.getvalue()
        warning = next(w for w in prepared.plan["warnings"] if w["code"] == "navigation_link_dropped")
        self.assertIn("Count: 2.", warning["message"])
        with redirect_stdout(StringIO()):
            execution = engine.execute_split(prepared, self.output)
        for entry in execution["outputs"]:
            reader = PdfReader(Path(execution["final_directory"]) / entry["filename"])
            self.assertTrue(all("/AA" not in page for page in reader.pages))
            self.assertEqual(len(serialized_pages(reader)), len(reader.pages))
        first = PdfReader(Path(execution["final_directory"]) / execution["outputs"][0]["filename"])
        self.assertNotIn(b"AUTHORED-PAGE-03", first.stream.getvalue())
        self.assertNotIn(b"AUTHORED-PAGE-04", first.stream.getvalue())
        self.assertEqual([str(page.get("/AA")) for page in prepared._reader.pages], captured_actions)
        self.assertEqual(prepared._reader.stream.getvalue(), captured_bytes)
        self.assertEqual(source.read_bytes(), original_bytes)

    def test_captured_source_metadata_annotations_and_reader_are_immutable_during_publication(self):
        prepared = engine.prepare_split(self.source, "1", output_base=self.output)
        before = [str(reference.get_object()) for page in prepared._reader.pages for reference in page.get("/Annots", [])]
        metadata = dict(prepared._reader.metadata)
        snapshot = prepared._reader.stream.getvalue()
        replacement = self.work / "replacement.pdf"
        authored_pdf(replacement, metadata_bad=True)
        self.source.write_bytes(replacement.read_bytes())
        reopened = []
        def reopen_output(path):
            self.assertNotEqual(Path(path), self.source, "Never reopen the replaced source")
            reopened.append(Path(path))
            return PdfReader(path)
        with patch.object(engine, "PdfReader", side_effect=reopen_output), redirect_stdout(StringIO()):
            execution = engine.execute_split(prepared, self.output)
        self.assertEqual(len(reopened), 2)
        self.assertEqual(prepared._reader.stream.getvalue(), snapshot)
        self.assertEqual(dict(prepared._reader.metadata), metadata)
        self.assertEqual([str(reference.get_object()) for page in prepared._reader.pages for reference in page.get("/Annots", [])], before)
        first = PdfReader(Path(execution["final_directory"]) / execution["outputs"][0]["filename"])
        self.assertEqual(first.metadata.author, "Authored author Å 日本")
        self.assertEqual(execution["source_identity"]["sha256"], sha256(self.original).hexdigest())
        with self.assertRaises(TypeError):
            prepared.plan["warnings"][0]["message"] = "changed"


if __name__ == "__main__":
    unittest.main()
