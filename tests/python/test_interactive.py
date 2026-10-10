"""Held reader/plan, bounded interactive protocol and owned working-PDF lifetime."""

from hashlib import sha256
import importlib.util
from io import BytesIO
import json
from pathlib import Path
from types import SimpleNamespace
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, NumberObject


ROOT = Path(__file__).resolve().parents[2]
ENGINE = ROOT / "engine/winbooksplit_engine.py"
spec = importlib.util.spec_from_file_location("wbs_interactive_engine", ENGINE)
engine = importlib.util.module_from_spec(spec)
spec.loader.exec_module(engine)
SESSION = "a" * 32


def pdf_bytes(pages=6, bookmarks=False):
    writer = PdfWriter()
    for number in range(1, pages + 1):
        page = writer.add_blank_page(width=100, height=100)
        page[NameObject("/SyntheticPageID")] = NumberObject(number)
    if bookmarks:
        writer.add_outline_item("Authored Å title", 1)
        writer.add_outline_item("Authored second title", 4)
    stream = BytesIO()
    writer.write(stream)
    return stream.getvalue()


class Dialogue:
    def __init__(self, handler):
        self.events, self.pending, self.handler = [], [], handler

    def write(self, line):
        assert line.startswith(engine.INTERACTION_PREFIX) and line.endswith("\n")
        assert line.count("\n") == 1
        self.events.append(json.loads(line[len(engine.INTERACTION_PREFIX):]))

    def flush(self):
        event = self.events[-1]
        answer = self.handler(event)
        if answer is None:
            self.pending.append(b"")
        elif isinstance(answer, bytes):
            self.pending.append(answer)
        else:
            reply = {"protocol": "winbooksplit.interaction", "version": 1,
                     "session": event["session"], "sequence": event["sequence"], **answer}
            self.pending.append((json.dumps(reply, separators=(",", ":")) + "\n").encode("utf-8"))

    def readline(self, limit):
        return self.pending.pop(0)[:limit]


class InteractiveTests(unittest.TestCase):
    def setUp(self):
        self.owned = tempfile.TemporaryDirectory(prefix="wbs-interactive-unit-")
        self.addCleanup(self.owned.cleanup)
        self.work = Path(self.owned.name).resolve()
        self.base = self.work / "out"
        self.base.mkdir()
        self.neighbor = self.base / "preserved-neighbor.txt"
        self.neighbor.write_bytes(b"Authored neighbor must stay unchanged\n")
        self.source = self.work / "Authored Å [1] & %.pdf"
        self.source.write_bytes(pdf_bytes())

    def run_dialogue(self, handler, mode="manual", **kwargs):
        dialogue = Dialogue(handler)
        with patch.object(engine, "log"):
            result = engine.run_interactive_split(self.source, self.base, mode, session=SESSION,
                                                 input_stream=dialogue, output_stream=dialogue, **kwargs)
        self.assertEqual([event["sequence"] for event in dialogue.events], list(range(1, len(dialogue.events) + 1)))
        self.assertTrue(all(event["session"] == SESSION for event in dialogue.events))
        self.assertEqual(self.neighbor.read_bytes(), b"Authored neighbor must stay unchanged\n")
        return result, dialogue.events

    def assert_no_output(self):
        self.assertEqual({item.name for item in self.base.iterdir()}, {self.neighbor.name})

    def test_ac060_confirmation_holds_exact_prepared_reader_plan_and_original_page_identity(self):
        captured = []
        original_prepare, original_execute = engine.prepare_split, engine.execute_split

        def prepare(*args, **kwargs):
            job = original_prepare(*args, **kwargs)
            captured.append(job)
            return job

        def execute(job, output):
            self.assertIs(job, captured[0])
            self.assertIs(job.plan, captured[0].plan)
            def output_reader(path, *args, **kwargs):
                if Path(path) == self.source:
                    raise AssertionError("Cannot reopen source after consent")
                return PdfReader(path, *args, **kwargs)
            with patch.object(engine, "prepare_split", side_effect=AssertionError("Cannot reprepare after consent")), \
                    patch.object(engine, "plan_manual_starts", side_effect=AssertionError("Cannot replan after consent")), \
                    patch.object(engine, "PdfReader", side_effect=output_reader):
                # Output validation legitimately reopens every chapter PDF.
                return original_execute(job, output)

        def answer(event):
            self.assert_no_output()
            if event["stage"] == "input_ready":
                self.assertEqual(event["total_pages"], 6)
                return {"action": "starts", "starts": "3,5"}
            self.assertEqual(event["stage"], "plan_ready")
            self.assertEqual(json.loads(event["plan_json"]), event["plan"])
            self.assertEqual(sha256(event["plan_json"].encode("utf-8")).hexdigest(), event["plan_sha256"])
            self.source.write_bytes(pdf_bytes(1))
            return {"action": "execute", "plan_sha256": event["plan_sha256"]}

        with patch.object(engine, "prepare_split", side_effect=prepare), \
                patch.object(engine, "execute_split", side_effect=execute), \
                patch.object(engine, "PdfReader", wraps=PdfReader) as source_reader:
            result, events = self.run_dialogue(answer)
        self.assertEqual(source_reader.call_count, 1)
        self.assertEqual(result["status"], "success")
        plan = events[-1]["plan"]
        self.assertEqual(json.loads(engine.serialize_result(captured[0].plan))["entries"], plan["entries"])
        self.assertEqual([[entry["start"], entry["end"]] for entry in result["execution"]["outputs"]], plan["ranges"])
        observed = [int(page["/SyntheticPageID"]) for entry in result["execution"]["outputs"]
                    for page in PdfReader(Path(result["execution"]["final_directory"], entry["filename"])).pages]
        self.assertEqual(observed, list(range(1, 7)))
        self.assertEqual(result["written_count"], plan["coverage"]["section_count"])

    def test_ac060_cancel_and_eof_at_input_or_plan_never_call_writer_or_allocate_output(self):
        for stage, eof in (("input_ready", False), ("input_ready", True), ("plan_ready", False), ("plan_ready", True)):
            with self.subTest(stage=stage, eof=eof):
                def answer(event):
                    if event["stage"] == stage:
                        return None if eof else {"action": "cancel"}
                    return {"action": "starts", "starts": "1,3,5"}
                with patch.object(engine, "execute_split", side_effect=AssertionError("Cancel cannot execute")), \
                        patch.object(engine, "OutputRun", side_effect=AssertionError("Cancel cannot reserve")):
                    result, _ = self.run_dialogue(answer)
                self.assertEqual((result["status"], result["exit_code"]), ("cancelled", 130))
                self.assert_no_output()

    def test_ac060_no_plan_cancel_or_eof_preserves_original_failure_and_fallback_modes(self):
        for eof in (False, True):
            with self.subTest(eof=eof), patch.object(engine, "execute_split") as execute:
                result, events = self.run_dialogue(lambda event: None if eof else {"action": "cancel"}, mode="1")
            self.assertEqual([event["stage"] for event in events], ["no_plan"])
            self.assertEqual((result["status"], result["code"], result["exit_code"]), ("no_plan", "no_bookmarks", 5))
            self.assertEqual(json.loads(engine.serialize_result(result)), events[0]["result"])
            self.assertEqual(events[0]["fallback_modes"], ["manual"])
            execute.assert_not_called()
            self.assert_no_output()

    def test_ac060_explicit_fallback_reuses_same_reader_and_reports_final_manual_mode(self):
        def answer(event):
            if event["stage"] == "no_plan":
                return {"action": "retry", "mode": "manual"}
            if event["stage"] == "input_ready":
                return {"action": "starts", "starts": "1,3,5"}
            return {"action": "execute", "plan_sha256": event["plan_sha256"]}
        with patch.object(engine, "PdfReader", wraps=PdfReader) as reader, \
                patch.object(engine, "prepare_split", wraps=engine.prepare_split) as prepare:
            result, events = self.run_dialogue(answer, mode="2")
        self.assertEqual(sum(call.args[0] == self.source for call in reader.call_args_list), 1)
        self.assertEqual(prepare.call_count, 2)
        self.assertIs(prepare.call_args_list[0].kwargs["_captured_reader"], prepare.call_args_list[1].kwargs["_captured_reader"])
        self.assertEqual((result["mode"], result["written_count"]), ("manual", 3))
        self.assertEqual([event["stage"] for event in events], ["no_plan", "input_ready", "plan_ready"])

    def test_ac060_auto_plan_requires_consent_and_exact_declared_filename_range_and_title_parity(self):
        self.source.write_bytes(pdf_bytes(bookmarks=True))
        def answer(event):
            self.assertEqual(event["stage"], "plan_ready")
            self.assert_no_output()
            return {"action": "execute", "plan_sha256": event["plan_sha256"]}
        with patch.object(engine, "prepare_split", wraps=engine.prepare_split) as prepare, \
                patch.object(engine, "plan_level1", wraps=engine.plan_level1) as planner:
            result, events = self.run_dialogue(answer, mode="1")
        self.assertEqual((prepare.call_count, planner.call_count), (1, 1))
        self.assertEqual(result["status"], "success")
        plan = events[0]["plan"]
        self.assertEqual(plan["entries"][0]["reason"], "front_matter")
        for planned, written in zip(plan["entries"], result["execution"]["outputs"]):
            for key in ("title", "filename", "sequence", "start", "end", "reason", "parent_id"):
                self.assertEqual(planned[key], written[key])

    def test_ac060_level2_no_plan_retry_level1_retains_reader_and_only_offered_modes(self):
        self.source.write_bytes(pdf_bytes(bookmarks=True))
        def answer(event):
            if event["stage"] == "no_plan":
                self.assertEqual(event["result"]["code"], "no_bookmarks_at_level")
                self.assertEqual(event["fallback_modes"], ["1", "manual"])
                return {"action": "retry", "mode": "1"}
            return {"action": "execute", "plan_sha256": event["plan_sha256"]}
        with patch.object(engine, "prepare_split", wraps=engine.prepare_split) as prepare:
            result, events = self.run_dialogue(answer, mode="2")
        self.assertEqual((result["status"], result["mode"]), ("success", "1"))
        self.assertEqual(prepare.call_count, 2)
        self.assertIs(prepare.call_args_list[0].kwargs["_captured_reader"], prepare.call_args_list[1].kwargs["_captured_reader"])
        self.assertEqual([event["stage"] for event in events], ["no_plan", "plan_ready"])
        with patch.object(engine, "execute_split") as execute:
            result, _ = self.run_dialogue(lambda event: {"action": "retry", "mode": "2"}, mode="2")
        self.assertEqual((result["code"], result["exit_code"]), ("invalid_arguments", 2))
        execute.assert_not_called()

    def test_ac060_bad_session_sequence_json_action_or_digest_cannot_execute(self):
        base = {"protocol": "winbooksplit.interaction", "version": 1, "session": SESSION,
                "sequence": 1, "action": "starts", "starts": "1"}
        wrong_fields = [dict(base, session="b" * 32), dict(base, sequence=2), dict(base, sequence=True),
                        dict(base, version=True), dict(base, version=2), dict(base, protocol="other")]
        invalid = [(json.dumps(reply) + "\n").encode("utf-8") for reply in wrong_fields] + \
            [b'{}\n', b'{"session":"x","session":"y"}\n', b'not-json\n', b'\xff\n',
                   b'{}', b'x' * (engine.INTERACTION_REPLY_LIMIT_BYTES + 1)]
        for answer in invalid + [{"action": "execute", "plan_sha256": "0" * 64},
                                  {"action": "starts", "starts": "1", "mode": "manual"}]:
            with self.subTest(answer=str(answer)[:100]), patch.object(engine, "execute_split") as execute:
                result, _ = self.run_dialogue(lambda event: answer)
            self.assertEqual((result["code"], result["exit_code"]), ("invalid_arguments", 2))
            execute.assert_not_called()
            self.assert_no_output()
        def changed_digest(event):
            return {"action": "starts", "starts": "1"} if event["stage"] == "input_ready" else \
                {"action": "execute", "plan_sha256": "0" * 64}
        with patch.object(engine, "execute_split") as execute:
            result, _ = self.run_dialogue(changed_digest)
        self.assertEqual(result["code"], "invalid_arguments")
        execute.assert_not_called()

    def test_ac060_invalid_manual_tokens_are_not_filtered_after_physical_page_guidance(self):
        for starts in ("1,bad,3", "1,,3", "0", "1,999", ""):
            with self.subTest(starts=starts), patch.object(engine, "execute_split") as execute:
                result, events = self.run_dialogue(lambda event: {"action": "starts", "starts": starts})
            self.assertEqual([event["stage"] for event in events], ["input_ready"])
            self.assertEqual((result["code"], result["exit_code"]), ("invalid_start_pages", 2))
            execute.assert_not_called()
            self.assert_no_output()

    def converted_snapshot(self):
        self.source = self.work / "Authored book.epub"
        self.source.write_bytes(b"Authored conversion boundary, never claimed as a real Calibre fixture\n")
        data = pdf_bytes()
        path = self.work / "already-cleaned-conversion.pdf"
        original = {"path": str(self.source), "resolved_path": str(self.source), "binding": "ebook_snapshot",
                    "size_bytes": self.source.stat().st_size, "sha256": sha256(self.source.read_bytes()).hexdigest()}
        generated = {"path": str(path), "resolved_path": str(path), "binding": "reader_snapshot",
                     "size_bytes": len(data), "sha256": sha256(data).hexdigest(), "page_count": 6}
        return SimpleNamespace(pdf_bytes=data, original_source_identity=original, generated_pdf_identity=generated,
                               conversion={"generated_pdf_identity": generated, "original_source_identity": original,
                                           "workspace_cleanup": {"cleanup_complete": True, "retained_staging": None}})

    def test_ac061_requested_working_pdf_remains_readable_through_starts_and_consent_then_cleans(self):
        for confirm in (False, True):
            with self.subTest(confirm=confirm):
                converted = self.converted_snapshot()
                working_path = None
                def answer(event):
                    nonlocal working_path
                    if event["stage"] == "input_ready":
                        self.assertTrue(event["can_request_working_pdf"])
                        self.assertTrue(event["conversion"]["workspace_cleanup"]["cleanup_complete"])
                        self.assertFalse(Path(converted.generated_pdf_identity["path"]).exists())
                        return {"action": "request_working_pdf"}
                    if event["stage"] == "working_pdf_ready":
                        working_path = Path(event["working_pdf"]["path"])
                        self.assertEqual(working_path.read_bytes(), converted.pdf_bytes)
                        self.assertEqual(len(PdfReader(working_path).pages), 6)
                        self.assertEqual(event["working_pdf"]["binding"], "working_pdf_copy")
                        return {"action": "starts", "starts": "1,3,5"}
                    self.assertEqual(working_path.read_bytes(), converted.pdf_bytes)
                    return {"action": "execute", "plan_sha256": event["plan_sha256"]} if confirm else {"action": "cancel"}
                with patch.object(engine, "_convert_ebook_snapshot", return_value=converted) as convert:
                    result, _ = self.run_dialogue(answer)
                convert.assert_called_once()
                self.assertFalse(working_path.exists())
                self.assertFalse(working_path.parent.exists())
                self.assertTrue(result["diagnostic"]["working_pdf_cleanup"]["cleanup_complete"])
                self.assertEqual(result["status"], "success" if confirm else "cancelled")

    def test_ac061_working_pdf_cleanup_refuses_unexpected_member_and_retains_published_output_as_incomplete(self):
        converted = self.converted_snapshot()
        working_path = None
        def answer(event):
            nonlocal working_path
            if event["stage"] == "input_ready":
                return {"action": "request_working_pdf"}
            if event["stage"] == "working_pdf_ready":
                working_path = Path(event["working_pdf"]["path"])
                return {"action": "starts", "starts": "1,3,5"}
            (working_path.parent / "authored-unexpected.txt").write_bytes(b"Never adopt or delete this unregistered member")
            return {"action": "execute", "plan_sha256": event["plan_sha256"]}
        with patch.object(engine, "_convert_ebook_snapshot", return_value=converted):
            result, _ = self.run_dialogue(answer)
        self.assertEqual((result["status"], result["exit_code"]), ("incomplete", 6))
        self.assertEqual(result["written_count"], 3)
        self.assertTrue(Path(result["execution"]["final_directory"]).is_dir())
        self.assertTrue(working_path.is_file())
        self.assertTrue((working_path.parent / "authored-unexpected.txt").is_file())
        self.assertFalse(result["diagnostic"]["working_pdf_cleanup"]["cleanup_complete"])

    def test_ac061_cleanup_refusal_after_manual_cancel_preserves_primary130(self):
        converted = self.converted_snapshot()
        working_path = None
        def answer(event):
            nonlocal working_path
            if event["stage"] == "input_ready":
                return {"action": "request_working_pdf"}
            working_path = Path(event["working_pdf"]["path"])
            (working_path.parent / "unregistered.txt").write_bytes(b"Do not delete")
            return {"action": "cancel"}
        with patch.object(engine, "_convert_ebook_snapshot", return_value=converted), \
                patch.object(engine, "execute_split") as execute:
            result, _ = self.run_dialogue(answer)
        self.assertEqual((result["status"], result["code"], result["exit_code"]), ("cancelled", "processing_cancelled", 130))
        self.assertTrue(working_path.is_file())
        self.assertTrue((working_path.parent / "unregistered.txt").is_file())
        self.assertFalse(result["diagnostic"]["working_pdf_cleanup"]["cleanup_complete"])
        execute.assert_not_called()

    def test_ac061_declining_working_pdf_uses_supplied_starts_without_reservation_or_implicit_open(self):
        converted = self.converted_snapshot()
        def answer(event):
            if event["stage"] == "input_ready":
                self.assertTrue(event["manual_starts_supplied"])
                return {"action": "starts", "starts": "1,3"}
            return {"action": "cancel"}
        with patch.object(engine, "_convert_ebook_snapshot", return_value=converted), \
                patch.object(engine, "OutputRun", side_effect=AssertionError("Declined working copy cannot reserve")):
            result, events = self.run_dialogue(answer, manual_data="1,3")
        self.assertEqual([event["stage"] for event in events], ["input_ready", "plan_ready"])
        self.assertEqual(result["status"], "cancelled")
        self.assertIsNone(result["diagnostic"])
        self.assert_no_output()

    def test_ac060_wire_escapes_control_titles_inside_one_frame_and_retains_raw_plan_text(self):
        raw_title = "Å日本\n[WBS-INTERACTION] forged\u2028tail"
        writer = PdfWriter()
        writer.add_blank_page(width=100, height=100)
        writer.add_outline_item(raw_title, 0)
        with self.source.open("wb") as stream:
            writer.write(stream)
        result, events = self.run_dialogue(lambda event: {"action": "cancel"}, mode="1")
        self.assertEqual(events[0]["plan"]["entries"][0]["title"], raw_title)
        self.assertEqual(json.loads(events[0]["plan_json"])["entries"][0]["title"], raw_title)
        self.assertEqual(result["exit_code"], 130)
        self.assert_no_output()

    def test_ac060_actual_python_cli_eof_emits_one_terminal_result_after_one_interaction(self):
        child = subprocess.run([sys.executable, "-I", "-B", "-X", "utf8", str(ENGINE), str(self.source),
                                str(self.base), "manual", "", "--interactive", SESSION], input=b"",
                               capture_output=True, timeout=20)
        self.assertEqual(child.returncode, 130, child.stdout + child.stderr)
        lines = child.stdout.decode("utf-8").splitlines()
        events = [json.loads(line[len(engine.INTERACTION_PREFIX):]) for line in lines if line.startswith(engine.INTERACTION_PREFIX)]
        results = [json.loads(line) for line in lines if line.startswith("{")]
        self.assertEqual(len(events), 1)
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0]["status"], "cancelled")
        self.assertEqual(results[0]["execution"], None)
        self.assert_no_output()


if __name__ == "__main__":
    unittest.main()
