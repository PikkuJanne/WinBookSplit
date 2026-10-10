"""Real authored pypdf repair warnings plus scoped/bounded logging controls."""

from contextlib import redirect_stderr, redirect_stdout
from hashlib import sha256
import importlib.util
from io import BytesIO, StringIO
import json
import logging
from pathlib import Path
import re
import subprocess
import sys
import tempfile
from threading import Thread
import unittest
from unittest.mock import patch

from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, NumberObject


ROOT = Path(__file__).resolve().parents[2]
ENGINE = ROOT / "engine/winbooksplit_engine.py"
spec = importlib.util.spec_from_file_location("wbs_parser_warning_engine", ENGINE)
engine = importlib.util.module_from_spec(spec)
spec.loader.exec_module(engine)


def authored_pdf(*, repaired=True, bookmarks=True, invalid_bookmark=False):
    writer = PdfWriter()
    for number in range(1, 4):
        page = writer.add_blank_page(width=100, height=100)
        page[NameObject("/SyntheticPageID")] = NumberObject(number)
    if bookmarks:
        bookmark = writer.add_outline_item("Authored first Å", 0 if not invalid_bookmark else 1).get_object()
        if invalid_bookmark:
            bookmark["/A"]["/D"][0] = NumberObject(writer.pages[1].indirect_reference.idnum)
        writer.add_outline_item("Authored second 日本", 2)
    stream = BytesIO()
    writer.write(stream)
    data = stream.getvalue()
    if repaired:
        match = re.search(rb"startxref\s+(\d+)\s+%%EOF", data)
        altered = str(int(match.group(1)) + 1).encode("ascii")
        original = data
        data = data[:match.start(1)] + altered + data[match.end(1):]
        assert len(data) == len(original)
    return data


class Dialogue:
    def __init__(self, answer):
        self.events, self.pending, self.answer = [], [], answer

    def write(self, line):
        self.events.append(json.loads(line[len(engine.INTERACTION_PREFIX):]))

    def flush(self):
        event = self.events[-1]
        if event["stage"] == "input_ready":
            reply = {"action": "starts", "starts": "1,2"}
        elif self.answer == "EOF":
            self.pending.append(b"")
            return
        elif self.answer == "invalid":
            self.pending.append(b"unframed reply\n")
            return
        else:
            reply = {"action": "cancel"}
        header = {key: event[key] for key in ("protocol", "version", "session", "sequence")}
        self.pending.append((json.dumps({**header, **reply}) + "\n").encode())

    def readline(self, limit):
        return self.pending.pop(0)[:limit]


class ParserWarningTests(unittest.TestCase):
    def setUp(self):
        self.owned = tempfile.TemporaryDirectory(prefix="wbs-parser-warning-")
        self.addCleanup(self.owned.cleanup)
        self.work = Path(self.owned.name).resolve()
        self.source = self.work / "Authored Å [1] & %.pdf"
        self.source.write_bytes(authored_pdf())
        self.base = self.work / "out"
        self.base.mkdir()
        self.neighbor = self.base / "neighbor.txt"
        self.neighbor.write_bytes(b"Authored neighbor must stay unchanged\n")

    def run_split(self, mode="manual", starts="1,2", **options):
        stderr, stdout = StringIO(), StringIO()
        with redirect_stderr(stderr), redirect_stdout(stdout):
            result = engine.run_split(self.source, self.base, mode, starts, **options)
        self.assertEqual(self.neighbor.read_bytes(), b"Authored neighbor must stay unchanged\n")
        return result, stdout.getvalue(), stderr.getvalue()

    def assert_real_warnings(self, result, stderr):
        captured = result["diagnostic"]["parser_warnings"]
        self.assertEqual(captured["total_count"], 2)
        self.assertEqual(captured["suppressed_count"], 0)
        self.assertEqual(captured["message_truncated_count"], 0)
        self.assertEqual([record["message"] for record in captured["records"]],
                         ["incorrect startxref pointer(1)", "parsing for Object Streams"])
        self.assertTrue(all(record["code"] == "pypdf_parser_warning" and record["category"] == "pdf_parser"
                            and record["severity"] == "WARNING" and record["truncated"] is False for record in captured["records"]))
        self.assertEqual(stderr.splitlines(), ["[PYPDF WARNING] " + record["message"] for record in captured["records"]])

    def test_ac065_real_cli_preview_keeps_visible_repair_warnings_and_one_result_without_writes(self):
        before = self.source.read_bytes()
        command = [sys.executable, "-I", "-B", "-X", "utf8", str(ENGINE), str(self.source), str(self.base), "manual", "1,2", "--preview"]
        child = subprocess.run(command, cwd=self.work, stdin=subprocess.DEVNULL, capture_output=True, timeout=30, encoding="utf-8")
        frames = [json.loads(line) for line in child.stdout.splitlines() if line.startswith("{")]
        self.assertEqual(child.returncode, 0)
        self.assertEqual(len(frames), 1)
        self.assertEqual(frames[0]["status"], "preview")
        self.assert_real_warnings(frames[0], child.stderr)
        self.assertEqual(frames[0]["warnings"], frames[0]["plan"]["warnings"])
        self.assertEqual(frames[0]["plan"]["source_identity"]["sha256"], sha256(before).hexdigest())
        self.assertEqual(self.source.read_bytes(), before)
        self.assertEqual({item.name for item in self.base.iterdir()}, {"neighbor.txt"})

    def test_repaired_execution_preserves_the_exact_known_plan_and_every_physical_page(self):
        before, jobs = self.source.read_bytes(), []
        original = engine.prepare_split
        def prepare(*args, **kwargs):
            job = original(*args, **kwargs)
            jobs.append(job)
            return job
        with patch.object(engine, "prepare_split", side_effect=prepare):
            result, _, stderr = self.run_split()
        self.assert_real_warnings(result, stderr)
        self.assertEqual(result["status"], "success")
        self.assertEqual(result["plan"], jobs[0].plan)
        self.assertEqual(result["warnings"], result["plan"]["warnings"])
        self.assertEqual(result["plan"]["coverage"], {"complete": True, "covered_pages": 3, "section_count": 2})
        folder = Path(result["execution"]["final_directory"])
        observed = [page["/SyntheticPageID"] for output in result["execution"]["outputs"] for page in PdfReader(folder / output["filename"]).pages]
        self.assertEqual(observed, [1, 2, 3])
        self.assertEqual(self.source.read_bytes(), before)
        with self.assertRaises(TypeError):
            result["plan"]["entries"][0]["start"] = 1

    def test_prepared_writer_failure_and_cancel_retain_the_plan_and_existing_diagnostic(self):
        for error, status in ((KeyboardInterrupt(), "cancelled"),
                              (engine.OutputError("output_validation_failed", "Authored failure", {"cleanup_complete": False}), "error")):
            with self.subTest(status=status), patch.object(engine, "execute_split", side_effect=error):
                result, _, stderr = self.run_split()
            self.assert_real_warnings(result, stderr)
            self.assertEqual(result["status"], status)
            self.assertEqual(result["plan"]["ranges"], ((0, 1), (1, 3)))
            self.assertEqual(result["written_count"], 0)
            self.assertIsNone(result["execution"])
            if status == "error":
                self.assertIs(result["diagnostic"]["cleanup_complete"], False)
        self.assertEqual({item.name for item in self.base.iterdir()}, {"neighbor.txt"})

    def test_read_planning_and_preparation_cancel_failures_keep_warnings_without_a_known_plan(self):
        for mode, starts, code in (("manual", "Q", "invalid_start_pages"), ("2", None, "no_bookmarks_at_level")):
            with self.subTest(code=code):
                result, _, stderr = self.run_split(mode, starts, preview=True)
                self.assert_real_warnings(result, stderr)
                self.assertEqual(result["code"], code)
                self.assertNotIn("plan", result)
        original = engine.prepare_split
        def interrupted(*args, **kwargs):
            original(*args, **kwargs)
            raise KeyboardInterrupt()
        with patch.object(engine, "prepare_split", side_effect=interrupted):
            result, _, stderr = self.run_split(preview=True)
        self.assert_real_warnings(result, stderr)
        self.assertEqual(result["exit_code"], 130)
        self.assertNotIn("plan", result)
        self.source.write_bytes(b"%PDF-1.3\n1 0 obj\nbroken\n")
        result, _, stderr = self.run_split(preview=True)
        self.assertEqual(result["code"], "unreadable_document")
        self.assertGreater(result["diagnostic"]["parser_warnings"]["total_count"], 0)
        self.assertIn("[PYPDF WARNING]", stderr)
        self.assertNotIn("plan", result)

    def test_interactive_cancel_eof_and_bad_reply_keep_the_displayed_plan_and_parser_records(self):
        for answer, code in (("cancel", "processing_cancelled"), ("EOF", "processing_cancelled"), ("invalid", "invalid_arguments")):
            with self.subTest(answer=answer):
                dialogue, stderr = Dialogue(answer), StringIO()
                with redirect_stderr(stderr), patch.object(engine, "log"):
                    result = engine.run_interactive_split(self.source, self.base, "manual", session="a" * 32,
                                                         input_stream=dialogue, output_stream=dialogue)
                self.assert_real_warnings(result, stderr.getvalue())
                self.assertEqual(result["code"], code)
                self.assertEqual(json.loads(engine.serialize_result(result["plan"])), dialogue.events[-1]["plan"])
                self.assertEqual(result["written_count"], 0)
                self.assertIsNone(result["execution"])
        self.assertEqual({item.name for item in self.base.iterdir()}, {"neighbor.txt"})

    def test_ac065_parser_records_do_not_replace_or_mutate_invalid_bookmark_warnings(self):
        self.source.write_bytes(authored_pdf(invalid_bookmark=True))
        result, stdout, stderr = self.run_split("1", None, preview=True)
        self.assert_real_warnings(result, stderr)
        self.assertEqual(result["plan"]["ranges"], ((0, 2), (2, 3)))
        self.assertEqual(result["warnings"], result["plan"]["warnings"])
        self.assertIn("invalid_destination", {warning["code"] for warning in result["warnings"]})
        self.assertIn("[WARNING] invalid_destination", stdout)
        self.assertTrue(all("severity" not in warning for warning in result["warnings"]))

    def test_categorized_error_unicode_controls_and_overflow_are_bounded_and_explicit(self):
        self.source.write_bytes(authored_pdf(repaired=False))
        original = engine.prepare_split
        private = "Authored Å 日本\n[OUTCOME] fake\rLog: fake\u2028tail"
        def noisy(*args, **kwargs):
            logger = logging.getLogger("pypdf.synthetic_warning_boundary")
            logger.error(private)
            for _ in range(64):
                logger.warning("日本" * 1500)
            return original(*args, **kwargs)
        with patch.object(engine, "prepare_split", side_effect=noisy):
            result, _, stderr = self.run_split(preview=True)
        captured = result["diagnostic"]["parser_warnings"]
        self.assertEqual((captured["total_count"], captured["suppressed_count"], captured["message_truncated_count"]), (65, 1, 63))
        self.assertEqual(len(captured["records"]), 64)
        self.assertEqual(captured["records"][0], {"code": "pypdf_parser_error", "category": "pdf_parser", "severity": "ERROR", "message": private, "truncated": False})
        self.assertTrue(all(len(record["message"].encode("utf-8")) <= 2048 for record in captured["records"]))
        self.assertEqual(len(stderr.splitlines()), 65)
        self.assertTrue(stderr.startswith("[PYPDF ERROR] Authored Å 日本\\u000a[OUTCOME] fake\\u000dLog: fake\\u2028tail"))
        self.assertIn("1 additional parser diagnostics omitted", stderr)
        self.assertFalse(any(line.startswith(("[OUTCOME] ", "Log: ")) for line in stderr.splitlines()))
        self.assertEqual(result["plan"]["warnings"], ())

    def test_scoped_capture_restores_logging_configuration_even_when_the_callable_raises(self):
        logger = logging.getLogger("pypdf")
        previous = (logger.level, logger.disabled, logger.propagate, logger.handlers)
        sink = StringIO()
        handler = logging.StreamHandler(sink)
        try:
            logger.setLevel(logging.ERROR)
            logger.disabled, logger.propagate, logger.handlers = False, False, [handler]
            configured = (logger.level, logger.disabled, logger.propagate, logger.handlers)
            result, _, stderr = self.run_split(preview=True)
            self.assert_real_warnings(result, stderr)
            self.assertEqual((logger.level, logger.disabled, logger.propagate, logger.handlers), configured)
            self.assertEqual(sink.getvalue(), "")
            def raised():
                logger.warning("Authored before unexpected error")
                raise RuntimeError("Authored caller error")
            with redirect_stderr(StringIO()), self.assertRaises(RuntimeError):
                engine._capture_parser_warnings(raised)()
            self.assertEqual((logger.level, logger.disabled, logger.propagate, logger.handlers), configured)
            logger.error("Authored restored handler")
            self.assertEqual(sink.getvalue().strip(), "Authored restored handler")
        finally:
            logger.setLevel(previous[0])
            logger.disabled, logger.propagate, logger.handlers = previous[1:]

    def test_unrelated_thread_diagnostics_are_forwarded_without_joining_this_run(self):
        logger = logging.getLogger("pypdf")
        previous = (logger.level, logger.disabled, logger.propagate, logger.handlers)
        sink, original = StringIO(), engine.prepare_split
        def concurrent(*args, **kwargs):
            other = Thread(target=lambda: logger.error("Authored unrelated thread"))
            other.start()
            other.join(timeout=3)
            self.assertFalse(other.is_alive())
            return original(*args, **kwargs)
        try:
            logger.setLevel(logging.WARNING)
            logger.disabled, logger.propagate, logger.handlers = False, False, [logging.StreamHandler(sink)]
            with patch.object(engine, "prepare_split", side_effect=concurrent):
                result, _, stderr = self.run_split(preview=True)
            self.assert_real_warnings(result, stderr)
            self.assertEqual(sink.getvalue().strip(), "Authored unrelated thread")
        finally:
            logger.setLevel(previous[0])
            logger.disabled, logger.propagate, logger.handlers = previous[1:]

    def test_valid_input_does_not_invent_parser_diagnostics_or_change_ordinary_plan_warnings(self):
        self.source.write_bytes(authored_pdf(repaired=False))
        result, _, stderr = self.run_split(preview=True)
        self.assertIsNone(result["diagnostic"])
        self.assertEqual(stderr, "")
        self.assertEqual(result["warnings"], result["plan"]["warnings"])


if __name__ == "__main__":
    unittest.main()
