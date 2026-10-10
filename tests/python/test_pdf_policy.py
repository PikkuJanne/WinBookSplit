"""Authored PDF policy units; no actions, attachments or viewers are invoked."""

from contextlib import redirect_stdout
from hashlib import sha256
import importlib.util
from io import BytesIO, StringIO
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

from pypdf import PdfReader, PdfWriter
from pypdf._encryption import Encryption
from pypdf.generic import (ArrayObject, BooleanObject, DecodedStreamObject,
    DictionaryObject, FloatObject, NameObject, NullObject, NumberObject, TextStringObject)


ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location("wbs_pdf_policy_units", ROOT / "engine/winbooksplit_engine.py")
engine = importlib.util.module_from_spec(spec)
spec.loader.exec_module(engine)
policy = engine._pdf_policy


def dictionary(**values):
    return DictionaryObject({NameObject("/" + key): value for key, value in values.items()})


def authored_writer(pages=4):
    writer = PdfWriter()
    for index in range(pages):
        page = writer.add_blank_page(200 + index, 300 + index)
        content = DecodedStreamObject()
        content.set_data(b"% ORIGINAL_POLICY_PAGE_" + str(index + 1).encode() + b"\n")
        page[NameObject("/Contents")] = writer._add_object(content)
    if pages >= 4:
        first = writer.add_outline_item("Original first", 0)
        writer.add_outline_item("Original first child", 1, parent=first)
        second = writer.add_outline_item("Original second", 2)
        writer.add_outline_item("Original second child", 3, parent=second)
    writer.add_metadata({"/Title": "Original policy source Å", "/Author": "Original fixture author"})
    return writer


def encoded(writer):
    stream = BytesIO()
    writer.write(stream)
    return stream.getvalue()


def inherited_writer():
    writer = authored_writer()
    pages = list(writer.pages)
    font = writer._add_object(dictionary(Type=NameObject("/Font"), Subtype=NameObject("/Type1"),
        BaseFont=NameObject("/Helvetica")))
    tree = writer._pages.get_object()
    tree[NameObject("/Resources")] = dictionary(Font=dictionary(F1=font))
    tree[NameObject("/MediaBox")] = ArrayObject([NumberObject(n) for n in (0, 0, 240, 340)])
    tree[NameObject("/CropBox")] = ArrayObject([NumberObject(n) for n in (10, 10, 180, 280)])
    tree[NameObject("/Rotate")] = NumberObject(90)
    for index, page in enumerate(pages, 1):
        for key in ("/Resources", "/MediaBox", "/CropBox", "/Rotate"):
            page.pop(key, None)
        content = DecodedStreamObject()
        content.set_data(f"BT /F1 12 Tf 20 250 Td (ORIGINAL_INHERITED_PAGE_{index:02}) Tj ET\n".encode())
        page[NameObject("/Contents")] = writer._add_object(content)
    return writer


def signature_structure():
    # This is a detection fixture, never a cryptographic-validity assertion.
    return dictionary(Type=NameObject("/Sig"), Contents=TextStringObject("Original inert signature marker"))


def embedded_structure():
    value = DecodedStreamObject()
    value[NameObject("/Type")] = NameObject("/EmbeddedFile")
    value.set_data(b"Original inert attachment marker\n")
    return value


class RecordingBytesIO(BytesIO):
    def __init__(self, data):
        super().__init__(data)
        self.read_sizes = []

    def read(self, size=-1):
        self.read_sizes.append(size)
        return super().read(size)


class PdfPolicyTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory(prefix="wbs-pdf-policy-unit-")
        self.addCleanup(temporary.cleanup)
        self.work = Path(temporary.name).resolve()
        self.output = self.work / "out"
        self.output.mkdir()
        self.neighbor = self.output / "authored-neighbor.txt"
        self.neighbor.write_bytes(b"Preserve authored earlier output\n")

    def source(self, writer, label):
        path = self.work / (label + ".pdf")
        path.write_bytes(encoded(writer))
        return path

    def assert_rejected_before_writer(self, source, code="unsupported_document", mode="manual", preview=False):
        before = sha256(source.read_bytes()).hexdigest()
        with patch.object(engine, "write_slice", side_effect=AssertionError("Policy failure must precede chapter writer")) as writer, \
                redirect_stdout(StringIO()):
            result = engine.run_split(source, self.output, mode, "1,3", preview=preview)
        writer.assert_not_called()
        self.assertEqual(result["code"], code)
        self.assertEqual(result["exit_code"], 7 if code == "unsupported_document" else 6)
        self.assertEqual(result["written_count"], 0)
        self.assertIsNone(result["execution"])
        self.assertEqual(result["fallback_modes"], ())
        self.assertIsNone(result.get("plan"))
        self.assertEqual(before, sha256(source.read_bytes()).hexdigest())
        self.assertEqual(list(self.output.iterdir()), [self.neighbor])
        return result

    def test_encrypted_inputs_never_read_verify_or_decrypt_even_empty_password(self):
        for empty_password in (False, True):
            writer = authored_writer()
            writer.encrypt("" if empty_password else "Original authored open credential",
                "Original authored owner credential", algorithm="RC4-128")
            source = self.source(writer, "encrypted-" + str(empty_password))
            with patch.object(Encryption, "read", side_effect=AssertionError("Encryption.read must not run")) as read, \
                    patch.object(Encryption, "verify", side_effect=AssertionError("Encryption.verify must not run")) as verify, \
                    patch.object(PdfReader, "decrypt", side_effect=AssertionError("PdfReader.decrypt must not run")) as decrypt:
                for mode in ("manual", "1", "2"):
                    with self.subTest(empty_password=empty_password, mode=mode):
                        result = self.assert_rejected_before_writer(source, mode=mode)
                        self.assertEqual(result["status"], "unsupported")
                self.assert_rejected_before_writer(source, preview=True)
                interaction = StringIO()
                with redirect_stdout(StringIO()):
                    result = engine.run_interactive_split(source, self.output, "manual", "1,3",
                        session="a" * 32, input_stream=BytesIO(), output_stream=interaction)
                self.assertEqual(result["code"], "unsupported_document")
                self.assertEqual(result["exit_code"], 7)
                self.assertEqual(result["written_count"], 0)
                self.assertIsNone(result["execution"])
                self.assertEqual(interaction.getvalue(), "")
                with self.assertRaises(policy.ConversionError) as conversion:
                    policy._validate_pdf(source.read_bytes(), str(source))
                self.assertEqual(conversion.exception.code, "unsupported_document")
                read.assert_not_called()
                verify.assert_not_called()
                decrypt.assert_not_called()

    def test_catalog_feature_presence_is_rejected_even_if_null_or_malformed(self):
        for key in ("/AcroForm", "/Perms", "/Collection", "/OpenAction", "/AA", "/AF"):
            with self.subTest(key=key):
                writer = authored_writer()
                writer.root_object[NameObject(key)] = NullObject()
                self.assert_rejected_before_writer(self.source(writer, "catalog-" + key[1:]))
        for key in ("/JavaScript", "/EmbeddedFiles"):
            with self.subTest(names_key=key):
                writer = authored_writer()
                writer.root_object[NameObject("/Names")] = DictionaryObject({NameObject(key): NullObject()})
                self.assert_rejected_before_writer(self.source(writer, "names-" + key[1:]))

    def test_widgets_signatures_attachments_and_multimedia_without_catalog_are_rejected(self):
        for subtype in ("/Widget", "/FileAttachment", "/RichMedia", "/Screen", "/Movie", "/Sound", "/3D"):
            with self.subTest(subtype=subtype):
                writer = authored_writer()
                annotation = dictionary(Type=NameObject("/Annot"), Subtype=NameObject(subtype),
                    Rect=ArrayObject([NumberObject(n) for n in (10, 10, 30, 30)]))
                writer.pages[0][NameObject("/Annots")] = ArrayObject([writer._add_object(annotation)])
                self.assert_rejected_before_writer(self.source(writer, "annotation-" + subtype[1:]))
        writer = authored_writer()
        writer.pages[0][NameObject("/AF")] = ArrayObject()
        self.assert_rejected_before_writer(self.source(writer, "page-associated-file"))

    def test_copied_static_appearance_rejects_nested_signature_and_attachment_shapes(self):
        features = {"signature": signature_structure(), "embedded": embedded_structure(),
            "ef": dictionary(EF=dictionary(F=TextStringObject("Original inert marker"))),
            "af": dictionary(AF=ArrayObject())}
        for label, feature in features.items():
            with self.subTest(feature=label):
                writer = authored_writer()
                appearance = DecodedStreamObject()
                appearance.set_data(b"q Q\n")
                appearance.update(dictionary(Type=NameObject("/XObject"), Subtype=NameObject("/Form"),
                    BBox=ArrayObject([NumberObject(n) for n in (0, 0, 20, 20)]),
                    Resources=dictionary(OriginalFeature=writer._add_object(feature))))
                annotation = dictionary(Type=NameObject("/Annot"), Subtype=NameObject("/Square"),
                    Rect=ArrayObject([NumberObject(n) for n in (10, 10, 30, 30)]),
                    AP=dictionary(N=writer._add_object(appearance)))
                writer.pages[0][NameObject("/Annots")] = ArrayObject([writer._add_object(annotation)])
                self.assert_rejected_before_writer(self.source(writer, "appearance-" + label))

    def test_copied_resource_keys_do_not_inherit_page_or_catalog_exclusions(self):
        for key in ("/Parent", "/P", "/AA", "/Outlines", "/Dests", "/B", "/Annots"):
            with self.subTest(resource_key=key):
                writer = authored_writer()
                resource = DictionaryObject({NameObject(key): writer._add_object(embedded_structure())})
                writer.pages[0][NameObject("/Resources")] = dictionary(OriginalResource=writer._add_object(resource))
                self.assert_rejected_before_writer(self.source(writer, "resource-" + key[1:]))

    def test_off_tree_page_typed_copied_resource_does_not_hide_actions(self):
        writer = authored_writer()
        resource = dictionary(Type=NameObject("/Page"), AA=dictionary(O=dictionary(
            S=NameObject("/JavaScript"), JS=TextStringObject("Original inert resource action, never executed"))))
        writer.pages[0][NameObject("/Resources")] = dictionary(OriginalResource=writer._add_object(resource))
        self.assert_rejected_before_writer(self.source(writer, "off-tree-page-resource"))

    def test_copied_resource_page_tree_or_catalog_backlinks_are_rejected_without_actions(self):
        for kind in ("physical-page", "page-tree", "catalog"):
            with self.subTest(kind=kind):
                writer = authored_writer()
                target = {"physical-page": writer.pages[3].indirect_reference,
                    "page-tree": writer._pages, "catalog": writer.root_object.indirect_reference}[kind]
                writer.pages[0][NameObject("/Resources")] = dictionary(OriginalAuthoredResource=target)
                source = self.source(writer, "copied-backlink-" + kind)
                result = self.assert_rejected_before_writer(source)
                self.assertEqual(result["status"], "unsupported")

    def test_copied_resource_actions_are_rejected_but_transparency_is_preserved(self):
        actions = {"typed-uri": dictionary(Type=NameObject("/Action"), S=NameObject("/URI"),
                URI=TextStringObject("https://example.invalid/original-inert-policy-fixture")),
            "untyped-uri": dictionary(S=NameObject("/URI"),
                URI=TextStringObject("https://example.invalid/original-inert-policy-fixture")),
            "untyped-remote": dictionary(S=NameObject("/GoToR"), F=TextStringObject("original-inert-never-open.pdf"),
                D=ArrayObject([NumberObject(0), NameObject("/Fit")]))}
        for kind, action in actions.items():
            with self.subTest(kind=kind):
                writer = authored_writer()
                writer.pages[0][NameObject("/Resources")] = dictionary(OriginalAuthoredResource=writer._add_object(action))
                source = self.source(writer, "copied-action-" + kind)
                with patch.object(engine, "OutputRun", side_effect=AssertionError("Policy must precede output reservation")) as reserve:
                    self.assert_rejected_before_writer(source)
                reserve.assert_not_called()
        writer = authored_writer()
        transparency = dictionary(Type=NameObject("/Group"), S=NameObject("/Transparency"))
        writer.pages[0][NameObject("/Resources")] = dictionary(OriginalAuthoredResource=writer._add_object(transparency))
        source = self.source(writer, "ordinary-transparency")
        before = sha256(source.read_bytes()).hexdigest()
        with redirect_stdout(StringIO()):
            result = engine.run_split(source, self.output, "manual", "1,3")
        self.assertEqual(result["exit_code"], 0)
        self.assertEqual(result["code"], "split_complete")
        first = Path(result["execution"]["final_directory"]) / result["execution"]["outputs"][0]["filename"]
        copied = PdfReader(first).pages[0]["/Resources"]["/OriginalAuthoredResource"]
        self.assertEqual(copied["/S"], "/Transparency")
        self.assertEqual(before, sha256(source.read_bytes()).hexdigest())

    def test_indirect_feature_names_are_rejected_before_reservation_and_conversion_acceptance(self):
        cases = [("resource", "/Type", value) for value in ("/Action", "/Sig", "/EmbeddedFile")]
        cases += [("resource", "/S", value) for value in ("/URI", "/GoToR", "/JavaScript")]
        cases += [("resource", "/FT", "/Sig"), ("annotation", "/FT", "/Sig"), ("annotation", "/Type", "/Sig")]
        cases += [("annotation", "/Subtype", value) for value in
            ("/Widget", "/FileAttachment", "/RichMedia", "/Screen", "/Movie", "/Sound", "/3D")]
        for index, (context, field, value) in enumerate(cases):
            with self.subTest(context=context, field=field, value=value):
                writer = authored_writer()
                feature = DictionaryObject({NameObject(field): writer._add_object(NameObject(value)),
                    NameObject("/URI"): TextStringObject("https://example.invalid/authored-inert-never-opened")})
                if context == "resource":
                    writer.pages[0][NameObject("/Resources")] = dictionary(OriginalResource=writer._add_object(feature))
                else:
                    feature[NameObject("/Rect")] = ArrayObject([NumberObject(n) for n in (10, 10, 30, 30)])
                    writer.pages[0][NameObject("/Annots")] = ArrayObject([writer._add_object(feature)])
                source = self.source(writer, "indirect-feature-" + str(index))
                with patch.object(engine, "OutputRun", side_effect=AssertionError("Policy must precede reservation")) as reserve:
                    result = self.assert_rejected_before_writer(source)
                self.assertEqual(result["status"], "unsupported")
                reserve.assert_not_called()
                with self.assertRaises(policy.ConversionError) as converted:
                    policy._validate_pdf(source.read_bytes(), str(source))
                self.assertEqual(converted.exception.code, "unsupported_document")

    def test_indirect_projected_action_page_and_catalog_names_are_warned_and_dropped(self):
        for field, value in (("/Type", "/Action"), ("/S", "/URI"), ("/S", "/GoToR"),
                ("/Type", "/Page"), ("/Type", "/Pages"), ("/Type", "/Catalog")):
            with self.subTest(field=field, value=value):
                writer = authored_writer()
                feature = DictionaryObject({NameObject(field): writer._add_object(NameObject(value)),
                    NameObject("/OriginalSentinel"): TextStringObject("AUTHORED_PROJECTED_FEATURE_NEVER_EXECUTED"),
                    NameObject("/Contents"): writer.pages[3]["/Contents"]})
                appearance = DecodedStreamObject()
                appearance.set_data(b"q 0 0 20 20 re S Q\n")
                appearance.update(dictionary(Type=NameObject("/XObject"), Subtype=NameObject("/Form"),
                    BBox=ArrayObject([NumberObject(n) for n in (0, 0, 20, 20)]),
                    Resources=dictionary(OriginalFeature=writer._add_object(feature))))
                annotation = dictionary(Type=NameObject("/Annot"), Subtype=NameObject("/Square"),
                    Rect=ArrayObject([NumberObject(n) for n in (10, 10, 30, 30)]), AP=dictionary(N=writer._add_object(appearance)))
                writer.pages[0][NameObject("/Annots")] = ArrayObject([writer._add_object(annotation)])
                source = self.source(writer, "indirect-projection-" + field[1:] + value[1:])
                original = source.read_bytes()
                output = self.work / ("projected-" + field[1:] + value[1:])
                output.mkdir()
                with redirect_stdout(StringIO()):
                    prepared = engine.prepare_split(source, "manual", "1,3", output_base=output)
                    self.assertIn("annotation_dropped", {warning["code"] for warning in prepared.plan["warnings"]})
                    execution = engine.execute_split(prepared, output)
                chapter_pages = []
                for item in execution["outputs"]:
                    reader = PdfReader(Path(execution["final_directory"]) / item["filename"])
                    chapter_pages.extend(reader.pages)
                    self.assertTrue(all(not page.get("/Annots") for page in reader.pages))
                    page_objects = 0
                    for generation, objects in reader.xref.items():
                        for identity in objects:
                            if identity == 0:
                                continue
                            obj = reader.get_object(identity)
                            if isinstance(obj, DictionaryObject):
                                self.assertNotIn("/OriginalSentinel", obj)
                                if obj.get("/Type") == "/Page":
                                    page_objects += 1
                    self.assertEqual(page_objects, len(reader.pages))
                self.assertEqual([page.get_contents().get_data() for page in chapter_pages],
                    [b"% ORIGINAL_POLICY_PAGE_" + str(n).encode() + b"\n" for n in range(1, 5)])
                self.assertEqual(source.read_bytes(), original)

    def test_ordinary_indirect_static_goto_and_transparency_names_preserve_fidelity(self):
        # Preserve existing physical-tree validity; this repair does not broaden it.
        for location in ("catalog", "tree", "leaf"):
            invalid = authored_writer()
            target, kind = {"catalog": (invalid.root_object, "/Catalog"),
                "tree": (invalid._pages.get_object(), "/Pages"), "leaf": (invalid.pages[0], "/Page")}[location]
            target[NameObject("/Type")] = invalid._add_object(NameObject(kind))
            with self.subTest(invalid_location=location), \
                    patch.object(engine, "OutputRun", side_effect=AssertionError("Invalid tree must precede reservation")) as reserve:
                self.assert_rejected_before_writer(self.source(invalid, "indirect-physical-type-" + location), "invalid_document")
            reserve.assert_not_called()
        writer = authored_writer()
        transparency = dictionary(Type=writer._add_object(NameObject("/Group")), S=writer._add_object(NameObject("/Transparency")))
        writer.pages[0][NameObject("/Resources")] = dictionary(OriginalGroup=writer._add_object(transparency))
        appearance = DecodedStreamObject()
        appearance.set_data(b"q 0 0 20 20 re S Q\n")
        appearance.update(dictionary(Type=writer._add_object(NameObject("/XObject")),
            Subtype=writer._add_object(NameObject("/Form")), BBox=ArrayObject([NumberObject(n) for n in (0, 0, 20, 20)])))
        rectangle = ArrayObject([NumberObject(n) for n in (10, 10, 30, 30)])
        square = dictionary(Type=writer._add_object(NameObject("/Annot")), Subtype=writer._add_object(NameObject("/Square")),
            Rect=rectangle, AP=dictionary(N=writer._add_object(appearance)))
        link = dictionary(Type=writer._add_object(NameObject("/Annot")), Subtype=writer._add_object(NameObject("/Link")),
            Rect=rectangle, A=dictionary(S=writer._add_object(NameObject("/GoTo")),
                D=ArrayObject([writer.pages[1].indirect_reference, NameObject("/FitH"), NumberObject(123)])))
        writer.pages[0][NameObject("/Annots")] = ArrayObject([writer._add_object(square), writer._add_object(link)])
        source = self.source(writer, "ordinary-indirect-names")
        original = source.read_bytes()
        with redirect_stdout(StringIO()):
            result = engine.run_split(source, self.output, "manual", "1,3")
        self.assertEqual(result["exit_code"], 0)
        self.assertEqual(result["written_count"], 2)
        self.assertNotIn("annotation_dropped", {warning["code"] for warning in result["warnings"]})
        all_pages = [page for item in result["execution"]["outputs"]
            for page in PdfReader(Path(result["execution"]["final_directory"]) / item["filename"]).pages]
        self.assertEqual(len(all_pages), 4)
        self.assertEqual([page.get_contents().get_data() for page in all_pages],
            [b"% ORIGINAL_POLICY_PAGE_" + str(n).encode() + b"\n" for n in range(1, 5)])
        first = PdfReader(Path(result["execution"]["final_directory"]) / result["execution"]["outputs"][0]["filename"])
        page = first.pages[0]
        self.assertEqual(page["/Resources"]["/OriginalGroup"]["/S"], "/Transparency")
        annotations = [reference.get_object() for reference in page["/Annots"]]
        self.assertEqual(annotations[0]["/Subtype"], "/Square")
        self.assertEqual(annotations[0]["/AP"]["/N"].get_data(), b"q 0 0 20 20 re S Q\n")
        self.assertEqual(annotations[1]["/Subtype"], "/Link")
        self.assertEqual(annotations[1]["/Dest"], ArrayObject([first.pages[1].indirect_reference, NameObject("/FitH"), FloatObject(123)]))
        self.assertEqual(source.read_bytes(), original)

    def test_inherited_font_text_boxes_rotation_and_captured_pages_are_preserved(self):
        source = self.source(inherited_writer(), "ordinary-inherited-fields")
        original = source.read_bytes()
        raw = PdfReader(source)
        self.assertTrue(all(not any(key in reference.get_object() for key in
            ("/Resources", "/MediaBox", "/CropBox", "/Rotate"))
            for reference in raw.root_object["/Pages"]["/Kids"]))
        def page_snapshot(page):
            return (page.get_contents().get_data(), tuple(page.mediabox), tuple(page.cropbox),
                page.get("/Rotate"), page["/Resources"]["/Font"]["/F1"].get_object())
        for mode in ("manual", "1", "2"):
            with self.subTest(mode=mode):
                output = self.work / ("inherited-" + mode)
                output.mkdir()
                neighbor = output / "original-neighbor.txt"
                neighbor.write_bytes(b"Preserve original inherited-field output neighbor\n")
                with redirect_stdout(StringIO()):
                    prepared = engine.prepare_split(source, mode, "1,3", output_base=output)
                    captured_before = [page_snapshot(page) for page in prepared._reader.pages]
                    execution = engine.execute_split(prepared, output)
                self.assertEqual(captured_before, [page_snapshot(page) for page in prepared._reader.pages])
                pages = [page for item in execution["outputs"]
                    for page in PdfReader(Path(execution["final_directory"]) / item["filename"]).pages]
                self.assertEqual(len(pages), 4)
                for index, page in enumerate(pages, 1):
                    self.assertEqual(page.extract_text().strip(), f"ORIGINAL_INHERITED_PAGE_{index:02}")
                    self.assertEqual(tuple(page.mediabox), (0, 0, 240, 340))
                    self.assertEqual(tuple(page.cropbox), (10, 10, 180, 280))
                    self.assertEqual(page["/Rotate"], 90)
                    self.assertEqual(page["/Resources"]["/Font"]["/F1"]["/BaseFont"], "/Helvetica")
                self.assertEqual(neighbor.read_bytes(), b"Preserve original inherited-field output neighbor\n")
        self.assertEqual(source.read_bytes(), original)

    def test_inherited_copied_resources_reject_document_backlinks_and_actions(self):
        for kind in ("physical-page", "page-tree", "catalog", "typed-action", "uri", "remote"):
            with self.subTest(kind=kind):
                writer = inherited_writer()
                target = {"physical-page": writer.pages[3].indirect_reference,
                    "page-tree": writer._pages, "catalog": writer.root_object.indirect_reference,
                    "typed-action": writer._add_object(dictionary(Type=NameObject("/Action"))),
                    "uri": writer._add_object(dictionary(S=NameObject("/URI"),
                        URI=TextStringObject("https://example.invalid/original-inert-inherited-fixture"))),
                    "remote": writer._add_object(dictionary(S=NameObject("/GoToR"),
                        F=TextStringObject("original-inert-never-open.pdf")))}[kind]
                writer._pages.get_object()[NameObject("/Resources")] = dictionary(OriginalAuthoredResource=target)
                source = self.source(writer, "inherited-resource-" + kind)
                with patch.object(engine, "OutputRun", side_effect=AssertionError("Policy must precede output reservation")) as reserve:
                    self.assert_rejected_before_writer(source)
                reserve.assert_not_called()

    def test_malformed_raw_children_counts_and_aliases_are_not_silently_skipped(self):
        def mutate(writer, kind):
            tree = writer._pages.get_object()
            kids = tree["/Kids"]
            if kind == "null-child":
                kids.append(NullObject())
                tree[NameObject("/Count")] = NumberObject(5)
            elif kind == "scalar-child":
                kids.append(NumberObject(7))
                tree[NameObject("/Count")] = NumberObject(5)
            elif kind == "wrong-child-type":
                kids.append(writer._add_object(dictionary(Type=NameObject("/Font"))))
                tree[NameObject("/Count")] = NumberObject(5)
            elif kind == "repeated-page":
                kids.append(kids[0])
                tree[NameObject("/Count")] = NumberObject(5)
            elif kind == "cycle":
                tree[NameObject("/Kids")] = ArrayObject([writer._pages])
                tree[NameObject("/Count")] = NumberObject(1)
            elif kind == "missing-kids":
                del tree["/Kids"]
            elif kind == "missing-pages":
                del writer.root_object["/Pages"]
            elif kind == "leaf-with-kids":
                writer.pages[0][NameObject("/Kids")] = ArrayObject()
            else:
                tree[NameObject("/Count")] = {"wrong-count": NumberObject(42), "float-count": FloatObject(4.5),
                    "boolean-count": BooleanObject(True), "negative-count": NumberObject(-1)}[kind]
        for kind in ("null-child", "scalar-child", "wrong-child-type", "repeated-page", "cycle", "missing-kids",
                "missing-pages", "leaf-with-kids", "wrong-count", "float-count", "boolean-count", "negative-count"):
            with self.subTest(kind=kind):
                writer = authored_writer()
                mutate(writer, kind)
                self.assert_rejected_before_writer(self.source(writer, kind), "invalid_document")
        self.assert_rejected_before_writer(self.source(authored_writer(0), "zero-pages"), "invalid_document")

    def test_indirect_integer_count_is_valid_and_preserves_complete_page_order(self):
        writer = authored_writer()
        writer._pages.get_object()[NameObject("/Count")] = writer._add_object(NumberObject(4))
        source = self.source(writer, "indirect-count")
        reader = policy.PolicyPdfReader(source)
        self.assertEqual(len(reader.pages), 4)
        self.assertEqual([page.get_contents().get_data() for page in reader.pages],
            [b"% ORIGINAL_POLICY_PAGE_" + str(n).encode() + b"\n" for n in range(1, 5)])

    def test_snapshot_reads_are_bounded_and_restore_callers_stream_position(self):
        data = encoded(authored_writer())
        source = RecordingBytesIO(data)
        source.seek(7)
        with patch.object(policy, "MAX_PDF_SNAPSHOT_BYTES", len(data) - 1):
            with self.assertRaises(policy.PdfPolicyError) as failure:
                policy.PolicyPdfReader(source)
        self.assertEqual(failure.exception.code, "invalid_document")
        self.assertEqual(source.read_sizes, [len(data)])
        self.assertEqual(source.tell(), 7)
        exact = RecordingBytesIO(data)
        exact.seek(11)
        with patch.object(policy, "MAX_PDF_SNAPSHOT_BYTES", len(data)):
            reader = policy.PolicyPdfReader(exact)
        self.assertEqual(exact.read_sizes, [len(data) + 1])
        self.assertEqual(exact.tell(), 11)
        self.assertEqual(reader.stream.getvalue(), data)
        self.assertEqual(len(reader.pages), 4)

    def test_converted_file_size_is_bounded_before_full_read_or_pdf_validation(self):
        source = self.work / "original.epub"
        source.write_bytes(b"Original inert ebook fixture for a controlled converter seam\n")
        runs, targets = [], []
        original_read_bytes = Path.read_bytes
        def new_run(base, input_path):
            run = engine.OutputRun(base, input_path)
            runs.append(run)
            return run
        def controlled_converter(argv, **_):
            target = Path(argv[2])
            targets.append(target)
            target.write_bytes(b"X" * 129)
            return {"exit_code": 0, "stdout_tail": "", "stderr_tail": "",
                "stdout_total_bytes": 0, "stderr_total_bytes": 0,
                "stdout_truncated": False, "stderr_truncated": False}
        def guarded_read_bytes(path):
            if path in targets:
                raise AssertionError("Generated PDF must use a bounded read, never read_bytes")
            return original_read_bytes(path)
        with patch.object(policy, "MAX_PDF_SNAPSHOT_BYTES", 128), \
                patch.object(policy, "_run_converter", side_effect=controlled_converter), \
                patch.object(policy, "_validate_pdf", side_effect=AssertionError("Oversize must fail before PDF parsing")) as validate, \
                patch.object(Path, "read_bytes", guarded_read_bytes):
            with self.assertRaises(policy.ConversionError) as failure:
                policy.convert_ebook(source, sys.executable, self.output, new_run=new_run)
        self.assertEqual(failure.exception.code, "conversion_output_invalid")
        validate.assert_not_called()
        self.assertEqual(len(targets), 1)
        self.assertTrue(all(not Path(run.stage).exists() for run in runs))
        self.assertEqual(list(self.output.iterdir()), [self.neighbor])
        self.assertEqual(source.read_bytes(), b"Original inert ebook fixture for a controlled converter seam\n")

    def test_unchecked_captured_reader_cannot_bypass_preflight_before_page_extraction(self):
        writer = authored_writer()
        writer.root_object[NameObject("/AcroForm")] = NullObject()
        source = self.source(writer, "unchecked-captured-reader")
        reader = PdfReader(source)
        with patch.object(PdfReader, "_flatten", side_effect=AssertionError("Policy must precede extraction")) as flatten:
            with self.assertRaises(policy.PdfPolicyError) as failure:
                engine.prepare_split(source, "manual", "1,3", _captured_reader=reader)
        self.assertEqual(failure.exception.code, "unsupported_document")
        flatten.assert_not_called()

    def test_ordinary_static_form_appearance_is_retained_without_form_catalog(self):
        writer = authored_writer()
        appearance = DecodedStreamObject()
        appearance.set_data(b"q 1 0 0 RG 0 0 20 20 re S Q\n")
        appearance.update(dictionary(Type=NameObject("/XObject"), Subtype=NameObject("/Form"),
            BBox=ArrayObject([NumberObject(n) for n in (0, 0, 20, 20)])))
        annotation = dictionary(Type=NameObject("/Annot"), Subtype=NameObject("/Square"),
            Rect=ArrayObject([NumberObject(n) for n in (10, 10, 30, 30)]),
            AP=dictionary(N=writer._add_object(appearance)))
        writer.pages[0][NameObject("/Annots")] = ArrayObject([writer._add_object(annotation)])
        source = self.source(writer, "ordinary-form-appearance")
        with redirect_stdout(StringIO()):
            result = engine.run_split(source, self.output, "manual", "1,3")
        self.assertEqual(result["exit_code"], 0)
        self.assertFalse({warning["code"] for warning in result["warnings"]} & {"annotation_dropped"})
        first = Path(result["execution"]["final_directory"]) / result["execution"]["outputs"][0]["filename"]
        output = PdfReader(first)
        self.assertNotIn("/AcroForm", output.root_object)
        copied = output.pages[0]["/Annots"][0].get_object()["/AP"]["/N"]
        self.assertEqual(copied.get_data(), b"q 1 0 0 RG 0 0 20 20 re S Q\n")

    def test_page_depth_node_page_and_feature_limits_fail_before_writer(self):
        source = self.source(authored_writer(), "bounded-ordinary")
        for constant, value in (("MAX_PDF_PAGE_TREE_NODES", 3), ("MAX_PDF_PAGES", 3),
                ("MAX_PDF_FEATURE_NODES", 1), ("MAX_PDF_FEATURE_DEPTH", 0)):
            with self.subTest(limit=constant), patch.object(policy, constant, value):
                self.assert_rejected_before_writer(source, "invalid_document")
        writer = authored_writer()
        child = writer.pages[0].indirect_reference
        for _ in range(policy.MAX_PDF_PAGE_TREE_DEPTH + 1):
            child = writer._add_object(dictionary(Type=NameObject("/Pages"), Kids=ArrayObject([child]), Count=NumberObject(1)))
        writer._pages.get_object()[NameObject("/Kids")] = ArrayObject([child])
        writer._pages.get_object()[NameObject("/Count")] = NumberObject(1)
        self.assert_rejected_before_writer(self.source(writer, "deep-pages"), "invalid_document")

    def test_ordinary_shared_resources_and_defined_page_action_exclusion_remain_allowed(self):
        writer = authored_writer()
        resource = dictionary(OriginalMarker=TextStringObject("Shared inert resource"))
        resource_ref = writer._add_object(resource)
        resource[NameObject("/Parent")] = resource_ref
        for page in writer.pages:
            page[NameObject("/Resources")] = dictionary(OriginalResource=resource_ref)
        writer.pages[0][NameObject("/AA")] = dictionary(O=dictionary(S=NameObject("/JavaScript"),
            JS=TextStringObject("Original inert page action, never executed")))
        source = self.source(writer, "ordinary-defined-exclusion")
        before = sha256(source.read_bytes()).hexdigest()
        for mode in ("manual", "1", "2"):
            output = self.work / ("ordinary-" + mode)
            output.mkdir()
            with self.subTest(mode=mode), redirect_stdout(StringIO()):
                result = engine.run_split(source, output, mode, "1,3")
            self.assertEqual(result["exit_code"], 0)
            self.assertEqual(result["code"], "split_complete")
            self.assertGreater(result["written_count"], 0)
            self.assertIn("navigation_link_dropped", {warning["code"] for warning in result["warnings"]})
            pages = [page for item in result["execution"]["outputs"]
                for page in PdfReader(Path(result["execution"]["final_directory"]) / item["filename"]).pages]
            self.assertEqual(len(pages), 4)
            self.assertEqual([page.get_contents().get_data() for page in pages],
                [b"% ORIGINAL_POLICY_PAGE_" + str(n).encode() + b"\n" for n in range(1, 5)])
            self.assertTrue(all("/AA" not in page for page in pages))
        self.assertEqual(before, sha256(source.read_bytes()).hexdigest())

    def test_malformed_outline_stays_a_planning_error_and_manual_remains_available(self):
        writer = authored_writer()
        root = writer._add_object(dictionary(Type=NameObject("/Outlines"), Count=NumberObject(1100)))
        previous = root
        for index in range(1100):
            current = writer._add_object(dictionary(Title=TextStringObject("Original deep outline " + str(index)),
                Parent=previous, Dest=ArrayObject([writer.pages[0].indirect_reference, NameObject("/Fit")]),
                Count=NumberObject(1)))
            previous.get_object()[NameObject("/First")] = current
            previous.get_object()[NameObject("/Last")] = current
            previous = current
        writer.root_object[NameObject("/Outlines")] = root
        source = self.source(writer, "deep-outline")
        self.assert_rejected_before_writer(source, "invalid_outline", mode="1")
        with redirect_stdout(StringIO()):
            result = engine.run_split(source, self.output, "manual", "1,3")
        self.assertEqual(result["exit_code"], 0)
        self.assertEqual(result["written_count"], 2)


if __name__ == "__main__":
    unittest.main()
