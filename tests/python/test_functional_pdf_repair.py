"""Application regressions for readable encrypted PDFs and local initial views.

Inputs are authored equivalents of the owner's repair-pack fixtures. They contain
only six synthetic page markers, never any part of the private regression book.
These tests assert the required behavior and intentionally fail before repair.
The ignored initial-view warning contract is one `opening_destination_ignored`
record per plan, independent of other outline or fidelity warnings.
"""

from contextlib import redirect_stdout
from hashlib import sha256
import importlib.util
from io import BytesIO, StringIO
import json
from pathlib import Path
import tempfile
import unittest

from pypdf import PdfReader, PdfWriter
from pypdf.generic import (ArrayObject, DecodedStreamObject, DictionaryObject,
                          FloatObject, NameObject, NumberObject, TextStringObject)


ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location("wbs_functional_repair_engine", ROOT / "engine/winbooksplit_engine.py")
engine = importlib.util.module_from_spec(spec)
spec.loader.exec_module(engine)
policy = engine._pdf_policy
CASES = ("plain", "empty_password_rc4", "local_goto", "destination_array",
         "empty_password_and_local_goto", "nonempty_password",
         "javascript_open_action", "chained_goto_javascript", "remote_goto",
         "malformed_local_goto")
RANGES = {"manual": ((0, 2), (2, 4), (4, 6)), "1": ((0, 4), (4, 6)),
          "2": ((0, 1), (1, 3), (3, 4), (4, 5), (5, 6))}


def build_fixture(case):
    """Public synthetic RC4-40/view-action inputs, equivalent to the supplied helper."""
    if case not in CASES:
        raise ValueError("Unknown authored fixture")
    writer = PdfWriter()
    for number in range(1, 7):
        page = writer.add_blank_page(width=612, height=792)
        font = DictionaryObject({NameObject("/Type"): NameObject("/Font"),
            NameObject("/Subtype"): NameObject("/Type1"), NameObject("/BaseFont"): NameObject("/Helvetica")})
        page[NameObject("/Resources")] = DictionaryObject({NameObject("/Font"):
            DictionaryObject({NameObject("/F1"): font})})
        stream = DecodedStreamObject()
        stream.set_data(f"BT /F1 16 Tf 72 720 Td (Synthetic test page {number}) Tj ET\n".encode("ascii"))
        page[NameObject("/Contents")] = stream
    parent_a = writer.add_outline_item("Part A", 0)
    writer.add_outline_item("Chapter A1", 1, parent=parent_a)
    writer.add_outline_item("Chapter A2", 3, parent=parent_a)
    parent_b = writer.add_outline_item("Part B", 4)
    writer.add_outline_item("Chapter B1", 5, parent=parent_b)
    destination = ArrayObject([writer.pages[0].indirect_reference, NameObject("/XYZ"),
                              NumberObject(-32768), NumberObject(-32768), FloatObject(0.75)])
    goto = DictionaryObject({NameObject("/S"): NameObject("/GoTo"), NameObject("/D"): destination})
    js = DictionaryObject({NameObject("/S"): NameObject("/JavaScript"),
        NameObject("/JS"): TextStringObject("/* inert negative control; never execute */")})
    if case in ("local_goto", "empty_password_and_local_goto"):
        writer.root_object[NameObject("/OpenAction")] = goto
    elif case == "destination_array":
        writer.root_object[NameObject("/OpenAction")] = destination
    elif case == "javascript_open_action":
        writer.root_object[NameObject("/OpenAction")] = js
    elif case == "chained_goto_javascript":
        goto[NameObject("/Next")] = js
        writer.root_object[NameObject("/OpenAction")] = goto
    elif case == "remote_goto":
        writer.root_object[NameObject("/OpenAction")] = DictionaryObject({
            NameObject("/S"): NameObject("/GoToR"), NameObject("/F"): TextStringObject("not-a-real-target.pdf"),
            NameObject("/D"): ArrayObject([NumberObject(0), NameObject("/Fit")])})
    elif case == "malformed_local_goto":
        writer.root_object[NameObject("/OpenAction")] = DictionaryObject({
            NameObject("/S"): NameObject("/GoTo"), NameObject("/D"): NumberObject(42)})
    if case in ("empty_password_rc4", "empty_password_and_local_goto"):
        writer.encrypt("", owner_password="synthetic-owner-only", algorithm="RC4-40", permissions_flag=4294967284)
    elif case == "nonempty_password":
        writer.encrypt("fixture-password", owner_password="fixture-owner", algorithm="RC4-40")
    data = BytesIO()
    writer.write(data)
    return data.getvalue()


class Dialogue:
    """Use the real engine's held-plan handshake with explicit confirmation."""
    def __init__(self):
        self.events, self.pending = [], []

    def write(self, line):
        if not line.startswith(engine.INTERACTION_PREFIX) or not line.endswith("\n"):
            raise AssertionError("Expected one engine interaction record")
        self.events.append(json.loads(line[len(engine.INTERACTION_PREFIX):]))

    def flush(self):
        event = self.events[-1]
        if event["stage"] == "input_ready":
            answer = {"action": "starts", "starts": "1,3,5"}
        elif event["stage"] == "plan_ready":
            answer = {"action": "execute", "plan_sha256": event["plan_sha256"]}
        else:
            raise AssertionError("Readable fixture must reach its ordinary split plan")
        self.pending.append((json.dumps({"protocol": "winbooksplit.interaction", "version": 1,
            "session": event["session"], "sequence": event["sequence"], **answer}) + "\n").encode("utf-8"))

    def readline(self, limit):
        return self.pending.pop(0)[:limit]


class FunctionalPdfRepairTests(unittest.TestCase):
    def setUp(self):
        owned = tempfile.TemporaryDirectory(prefix="wbs-functional-pdf-")
        self.addCleanup(owned.cleanup)
        self.work = Path(owned.name).resolve()

    def case(self, label):
        work = self.work / label
        work.mkdir()
        source = work / "Authored Å [1] & input.pdf"
        source.write_bytes(build_fixture(label.split("--")[0]))
        output = work / "out"
        output.mkdir()
        (output / "prior-output.txt").write_bytes(b"Preserve earlier output\n")
        return source, output, source.read_bytes()

    def assert_preserved(self, source, output, original, result=None):
        self.assertEqual(source.read_bytes(), original)
        self.assertEqual((output / "prior-output.txt").read_bytes(), b"Preserve earlier output\n")
        if result is not None and result["exit_code"] != 0:
            self.assertEqual({path.name for path in output.iterdir()}, {"prior-output.txt"})
            self.assertEqual(result["written_count"], 0)
            self.assertIsNone(result["execution"])

    def assert_plan(self, plan, mode, original):
        self.assertEqual(plan["ranges"], RANGES[mode])
        self.assertEqual(plan["total_pages"], 6)
        self.assertEqual([page for start, end in plan["ranges"] for page in range(start, end)], list(range(6)))
        self.assertTrue(plan["coverage"]["complete"])
        self.assertEqual(plan["source_identity"]["sha256"], sha256(original).hexdigest())

    def assert_outputs(self, result, source, output, original, mode):
        self.assert_preserved(source, output, original, result)
        self.assertEqual((result["code"], result["exit_code"]), ("split_complete", 0), result["message"])
        self.assert_plan(result["plan"], mode, original)
        self.assertEqual(result["written_count"], len(RANGES[mode]))
        reference = PdfReader(BytesIO(original))
        if reference.is_encrypted:
            self.assertTrue(reference.decrypt(""))
        expected = [page.get_contents().get_data() for page in reference.pages]
        observed = []
        for entry in result["execution"]["outputs"]:
            chapter = PdfReader(Path(result["execution"]["final_directory"]) / entry["filename"])
            self.assertFalse(chapter.is_encrypted)
            self.assertNotIn("/OpenAction", chapter.root_object)
            observed.extend(page.get_contents().get_data() for page in chapter.pages)
        self.assertEqual(observed, expected)

    def run_readable_case(self, case):
        for mode in ("manual", "1", "2"):
            for preview in (True, False):
                with self.subTest(case=case, mode=mode, preview=preview):
                    source, output, original = self.case(case + "--" + mode + "--" + str(preview))
                    with redirect_stdout(StringIO()):
                        result = engine.run_split(source, output, mode, "1,3,5", preview=preview)
                    self.assert_preserved(source, output, original, result)
                    if preview:
                        self.assertEqual((result["code"], result["exit_code"]), ("preview_complete", 0), result["message"])
                        self.assertEqual({path.name for path in output.iterdir()}, {"prior-output.txt"})
                        self.assert_plan(result["plan"], mode, original)
                    else:
                        self.assert_outputs(result, source, output, original, mode)
                    if case in ("local_goto", "destination_array", "empty_password_and_local_goto"):
                        ignored = [warning for warning in result["warnings"]
                                   if warning["code"] == "opening_destination_ignored"]
                        self.assertEqual(len(ignored), 1, "Ignored initial view requires one specific warning")

    def test_plain_control_preview_and_execution(self):
        self.run_readable_case("plain")

    def test_empty_password_rc4_preview_and_execution(self):
        self.run_readable_case("empty_password_rc4")

    def test_local_goto_preview_and_execution(self):
        self.run_readable_case("local_goto")

    def test_destination_array_preview_and_execution(self):
        self.run_readable_case("destination_array")

    def test_combined_empty_password_local_goto_preview_and_execution(self):
        self.run_readable_case("empty_password_and_local_goto")

    def test_authenticated_combined_reader_also_passes_shared_feature_preflight(self):
        source, output, original = self.case("empty_password_and_local_goto")
        reader = PdfReader(BytesIO(original))
        self.assertTrue(reader.decrypt(""))
        self.assertTrue(reader.is_encrypted)
        self.assertEqual(len(reader.pages), 6)
        try:
            policy._preflight_pdf_features(reader)
        finally:
            self.assert_preserved(source, output, original)
            self.assertEqual({path.name for path in output.iterdir()}, {"prior-output.txt"})

    def test_authenticated_still_encrypted_captured_reader_can_plan_preview_and_execute(self):
        for mode in ("manual", "1", "2"):
            with self.subTest(mode=mode):
                source, output, original = self.case("empty_password_rc4--captured--" + mode)
                reader = PdfReader(BytesIO(original))
                self.assertTrue(reader.decrypt(""))
                self.assertTrue(reader.is_encrypted)
                self.assertEqual(len(reader.pages), 6)
                try:
                    with redirect_stdout(StringIO()):
                        prepared = engine.prepare_split(source, mode, "1,3,5", output_base=output, _captured_reader=reader)
                        self.assertIs(prepared._reader, reader)
                        self.assert_plan(engine.preview_plan(prepared), mode, original)
                        result = engine._complete_prepared_split(prepared, output)
                    self.assert_outputs(result, source, output, original, mode)
                except Exception:
                    self.assertEqual({path.name for path in output.iterdir()}, {"prior-output.txt"})
                    raise
                finally:
                    self.assert_preserved(source, output, original)

    def test_interactive_readable_inputs_reuse_plan_and_execute_after_confirmation(self):
        for case in ("plain", "empty_password_rc4", "local_goto", "empty_password_and_local_goto"):
            with self.subTest(case=case):
                source, output, original = self.case(case + "--interactive")
                dialogue = Dialogue()
                with redirect_stdout(StringIO()):
                    result = engine.run_interactive_split(source, output, "manual", session="a" * 32,
                        input_stream=dialogue, output_stream=dialogue)
                self.assert_outputs(result, source, output, original, "manual")
                self.assertEqual([event["stage"] for event in dialogue.events], ["input_ready", "plan_ready"])

    def test_converted_pdf_validation_uses_same_readable_input_policy(self):
        for case in ("plain", "empty_password_rc4", "local_goto", "destination_array", "empty_password_and_local_goto"):
            with self.subTest(case=case):
                source, output, original = self.case(case + "--converted")
                try:
                    identity = policy._validate_pdf(original, str(source))
                    self.assertEqual(identity["page_count"], 6)
                    self.assertEqual(identity["sha256"], sha256(original).hexdigest())
                finally:
                    self.assert_preserved(source, output, original)
                    self.assertEqual({path.name for path in output.iterdir()}, {"prior-output.txt"})

    def test_active_or_malformed_opening_actions_remain_rejected(self):
        for case in ("javascript_open_action", "chained_goto_javascript", "remote_goto", "malformed_local_goto"):
            with self.subTest(case=case):
                source, output, original = self.case(case)
                with redirect_stdout(StringIO()):
                    result = engine.run_split(source, output, "manual", "1,3,5")
                self.assert_preserved(source, output, original, result)
                self.assertEqual((result["code"], result["exit_code"]), ("unsupported_document", 7))

    def test_nonempty_password_is_reported_as_password_required_without_output(self):
        source, output, original = self.case("nonempty_password")
        with redirect_stdout(StringIO()):
            result = engine.run_split(source, output, "manual", "1,3,5")
        self.assert_preserved(source, output, original, result)
        self.assertEqual(result["code"], "password_required", result["message"])
        self.assertNotEqual(result["exit_code"], 0)


if __name__ == "__main__":
    unittest.main(verbosity=2)
