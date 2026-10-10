"""Local console authentication/cleanup units on authored files, not app runs."""

from copy import deepcopy
from hashlib import sha256
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[2]


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


console = load("wbs_console_unit_auth", ROOT / "tests/manual/current_launchers.py")
fixture = load("wbs_console_unit_fixture", ROOT / "tests/python/console_receipt_fixture.py")
conversion = load("wbs_console_unit_conversion", ROOT / "tests/conversion/characterize_conversion.py")


class ConsoleManifestTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory(prefix="wbs-console-unit-")
        self.addCleanup(self.temporary.cleanup)
        self.base = Path(self.temporary.name).resolve()
        self.source = self.base / "authored.pdf"
        self.source.write_bytes(b"Authored metadata-only input, never parsed as PDF\n")
        self.observed = {"size_bytes": self.source.stat().st_size, "sha256": sha256(self.source.read_bytes()).hexdigest()}
        self.outcome = {"protocol": "winbooksplit.outcome", "version": 1, "mode": "manual", "status": "read_error",
            "code": "unreadable_document", "exit_code": 6, "written_count": 0, "final_directory": None, "engine_result": None}
        self.text = "Authored Å 日本 console unit\r\n[OPERATION-OUTCOME] " + json.dumps(self.outcome) + "\r\n"
        self.receipt = fixture.make(self.outcome, sha256(self.text.encode()).hexdigest(), log_text=self.text,
            source_path=str(self.source), source_observation=self.observed)
        self.directory = self.base / (".WinBookSplit-console-" + self.receipt["owner"]["run_id"])
        self.directory.mkdir()
        self.log = self.directory / "console.log"
        self.log.write_bytes(self.text.encode())
        (self.directory / ".WinBookSplit-console-owner.json").write_text(json.dumps(self.receipt["owner"]), encoding="utf-8")
        self.publish(self.receipt)

    def publish(self, receipt):
        (self.directory / "WinBookSplit_Run.json").write_bytes(receipt["run_manifest_raw"].encode("utf-8"))

    def authenticate(self, **options):
        return console.authenticate_console_manifest(self.log, self.outcome,
            source_path=self.source, source_observation=self.observed, **options)

    def test_authenticates_exact_closed_bytes_and_removes_only_owned_members(self):
        observed = self.authenticate()
        self.assertEqual(observed["members"], sorted(console.CONSOLE_MEMBERS))
        self.assertEqual(observed["run_manifest"], self.receipt["run_manifest"])
        self.assertEqual(observed["log_size_bytes"], len(self.text.encode()))
        self.assertGreater(observed["file_identities"]["console.log"]["inode"], 0)
        console.remove_console_directory(self.log, self.base, self.outcome,
            source_path=self.source, source_observation=self.observed)
        self.assertFalse(self.directory.exists())
        self.assertEqual(self.source.read_bytes(), b"Authored metadata-only input, never parsed as PDF\n")

    def test_resealed_wrong_protocol_owner_outcome_log_source_and_warning_records_are_rejected(self):
        mutations = [lambda m: m.update(protocol="foreign.run"), lambda m: m.update(version=True),
            lambda m: m.update(run_id="f" * 32), lambda m: m.update(diagnostics_finalized=False),
            lambda m: m["outcome"].update(exit_code=0), lambda m: m["log"].update(bytes_written=1),
            lambda m: m["source_identity"].update(path="C:/foreign.pdf"), lambda m: m["source_identity"].update(size_bytes=1),
            lambda m: m["warnings"].update(parser=[{"code": "invented"}]),
            lambda m: m.update(finished_utc="2020-01-01T00:00:00Z"),
            lambda m: m["runtime_versions"].update(pypdf="9.99.0")]
        for number, change in enumerate(mutations):
            receipt = deepcopy(self.receipt)
            change(receipt["run_manifest"])
            fixture.reseal(receipt)
            self.publish(receipt)
            with self.subTest(mutation=number), self.assertRaises((ValueError, RuntimeError)):
                self.authenticate()
        self.publish(self.receipt)
        self.authenticate()

    def test_pending_or_foreign_members_never_authorize_cleanup(self):
        for name in ("WinBookSplit_Run.pending.json", "foreign.txt"):
            path = self.directory / name
            path.write_bytes(b"Preserve this extra member\n")
            with self.subTest(member=name), self.assertRaises((ValueError, RuntimeError)):
                console.remove_console_directory(self.log, self.base, self.outcome)
            self.assertEqual(path.read_bytes(), b"Preserve this extra member\n")
            self.assertTrue(self.log.exists())
            path.unlink()

    def test_two_members_only_allowed_for_explicit_nonzero_finalizer_fault(self):
        (self.directory / "WinBookSplit_Run.json").unlink()
        with self.assertRaises((ValueError, RuntimeError)):
            self.authenticate()
        failed = {**self.outcome, "status": "incomplete", "code": "console_finalize_failed"}
        receipt = console.authenticate_console_manifest(self.log, failed, allow_unfinalized=True)
        self.assertEqual(receipt["state"], "unfinalized")
        self.assertIsNone(receipt["run_manifest"])
        for outcome in ({**failed, "exit_code": 0}, {**failed, "status": "success"}):
            with self.assertRaises((ValueError, RuntimeError)):
                console.authenticate_console_manifest(self.log, outcome, allow_unfinalized=True)

    def test_changed_authenticated_file_is_preserved_by_held_cleanup(self):
        receipt = self.authenticate()
        self.log.write_bytes(self.log.read_bytes() + b"Changed after authentication\n")
        with patch.object(console, "authenticate_console_manifest", return_value=receipt):
            with self.assertRaises((ValueError, RuntimeError)):
                console.remove_console_directory(self.log, self.base, self.outcome)
        self.assertTrue(self.directory.exists())
        self.assertTrue(self.log.read_bytes().endswith(b"Changed after authentication\n"))

    def test_identical_bytes_replaced_authenticated_member_is_preserved(self):
        receipt = self.authenticate()
        raw, original = self.log.read_bytes(), self.log.lstat()
        self.log.unlink()
        self.log.write_bytes(raw)
        replacement = self.log.lstat()
        self.assertNotEqual((original.st_dev, original.st_ino), (replacement.st_dev, replacement.st_ino))
        with patch.object(console, "authenticate_console_manifest", return_value=receipt):
            with self.assertRaises((ValueError, RuntimeError)):
                console.remove_console_directory(self.log, self.base, self.outcome)
        self.assertEqual(self.log.read_bytes(), raw)
        self.assertTrue((self.directory / "WinBookSplit_Run.json").exists())

    def test_identical_bytes_replaced_authenticated_directory_is_preserved(self):
        receipt = self.authenticate()
        original = self.directory.lstat()
        preserved = self.base / "authored-original-console"
        self.directory.rename(preserved)
        self.directory.mkdir()
        for path in preserved.iterdir():
            (self.directory / path.name).write_bytes(path.read_bytes())
        replacement = self.directory.lstat()
        self.assertNotEqual((original.st_dev, original.st_ino), (replacement.st_dev, replacement.st_ino))
        with patch.object(console, "authenticate_console_manifest", return_value=receipt):
            with self.assertRaises((ValueError, RuntimeError)):
                console.remove_console_directory(self.log, self.base, self.outcome)
        self.assertTrue(preserved.exists())
        self.assertEqual({path.name: path.read_bytes() for path in self.directory.iterdir()},
                         {path.name: path.read_bytes() for path in preserved.iterdir()})

    def test_parser_only_diagnostic_has_no_cleanup_path_and_foreign_fields_are_not_adopted(self):
        parser = {"records": [{"code": "pypdf_parser_warning", "category": "pdf_parser", "severity": "WARNING",
            "message": "incorrect startxref pointer(1)", "truncated": False}],
            "total_count": 1, "suppressed_count": 0, "message_truncated_count": 0}
        with patch.object(conversion.launchers, "remove_known_directory") as remove:
            conversion.remove_diagnostic(self.base, {"diagnostic": {"parser_warnings": parser}})
            remove.assert_not_called()
            for diagnostic in ({"parser_warnings": parser, "foreign_path": str(self.source)},
                    {"parser_warnings": {**parser, "total_count": 0}}, {"parser_warnings": {}}):
                with self.subTest(diagnostic=diagnostic), self.assertRaises((KeyError, ValueError, RuntimeError)):
                    conversion.remove_diagnostic(self.base, {"diagnostic": diagnostic})
            remove.assert_not_called()
        self.assertTrue(self.source.exists())

    def test_parser_only_base_check_preserves_files_and_rejects_unowned_members(self):
        parser = {"records": [{"code": "pypdf_parser_warning", "category": "pdf_parser", "severity": "WARNING",
            "message": "incorrect startxref pointer(1)", "truncated": False}],
            "total_count": 1, "suppressed_count": 0, "message_truncated_count": 0}
        frame = {"protocol": "winbooksplit.result", "execution": None, "diagnostic": {"parser_warnings": parser}}
        preserved = {self.source.name, self.directory.name}
        before = {path: path.read_bytes() for path in (self.source, *self.directory.iterdir())}
        conversion.manual.check_base_members(self.base, frame, preserved)
        foreign = self.base / "foreign.txt"
        foreign.write_bytes(b"No parser warning owns or authorizes this file\n")
        with self.assertRaisesRegex(RuntimeError, "Unexpected output-base members"):
            conversion.manual.check_base_members(self.base, frame, preserved)
        for diagnostic in ({"parser_warnings": parser, "foreign_path": str(foreign)},
                {"parser_warnings": {**parser, "suppressed_count": 1}},
                {"parser_warnings": {**parser, "records": [{**parser["records"][0], "category": "foreign"}]}}):
            with self.subTest(diagnostic=diagnostic), self.assertRaises((KeyError, ValueError, RuntimeError)):
                conversion.manual.check_base_members(self.base, {**frame, "diagnostic": diagnostic}, preserved)
        self.assertEqual(foreign.read_bytes(), b"No parser warning owns or authorizes this file\n")
        self.assertEqual({path: path.read_bytes() for path in before}, before)

    def held_plan(self):
        plan = {"schema_version": 1, "mode": "2", "total_pages": 3, "section_count": 1,
            "source_identity": {"path": str(self.source), **self.observed, "binding": "reader_snapshot"},
            "normalized_inputs": {"mode": "2", "manual_starts": None}, "output_base": str(self.base),
            "entries": [{"start": 0, "end": 3, "title": "Authored Å 日本", "filename": "01 - Authored.pdf"}],
            "coverage": {"start": 0, "end": 3, "page_count": 3}, "warnings": []}
        canonical = json.dumps(plan, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
        event = {"protocol": "winbooksplit.interaction", "version": 1, "session": "a" * 32, "sequence": 1,
            "stage": "plan_ready", "plan": plan, "plan_json": canonical, "plan_sha256": sha256(canonical.encode()).hexdigest()}
        reply = {key: event[key] for key in ("protocol", "version", "session", "sequence", "plan_sha256")}
        reply["action"] = "execute"
        return plan, event, reply

    def publish_held_plan(self, event, reply=None):
        self.outcome = {**self.outcome, "mode": "2", "status": "timeout", "code": "processor_timeout", "exit_code": 130}
        self.text = "[PLAN] " + json.dumps(event, ensure_ascii=False) + "\r\n"
        if reply is not None:
            self.text += "[INTERACTION-REPLY] " + json.dumps(reply, ensure_ascii=False) + "\r\n"
        self.text += "[OPERATION-OUTCOME] " + json.dumps(self.outcome) + "\r\n"
        self.log.write_bytes(self.text.encode())
        self.receipt = fixture.make(self.outcome, sha256(self.text.encode()).hexdigest(), log_text=self.text,
            plan=event["plan"], source_path=str(self.source), source_observation=self.observed)
        self.publish(self.receipt)

    def test_displayed_unconfirmed_plan_keeps_metadata_only_record(self):
        plan, event, _ = self.held_plan()
        self.publish_held_plan(event)
        observed = self.authenticate()
        self.assertIsNone(observed["confirmed_plan"])
        self.assertIsNone(observed["run_manifest"]["plan"])
        self.assertEqual(observed["run_manifest"]["source_identity"]["binding"], "metadata_only")
        self.assertIsNone(observed["run_manifest"]["source_identity"]["sha256"])
        before = {path.name: path.read_bytes() for path in self.directory.iterdir()}
        forged = deepcopy(self.receipt)
        forged["run_manifest"].update(plan=plan, source_identity=plan["source_identity"])
        forged["run_manifest"]["settings"]["normalized_inputs"] = plan["normalized_inputs"]
        fixture.reseal(forged)
        self.publish(forged)
        with self.assertRaisesRegex(ValueError, "prepared/confirmed plan"):
            self.authenticate()
        self.publish(self.receipt)
        self.assertEqual({path.name: path.read_bytes() for path in self.directory.iterdir()}, before)

    def test_exact_execute_reply_confirms_plan_and_resealed_wrong_bindings_fail(self):
        plan, event, reply = self.held_plan()
        self.publish_held_plan(event, reply)
        self.assertEqual(self.authenticate()["confirmed_plan"], plan)
        for change in (lambda r: r.update(session="b" * 32), lambda r: r.update(sequence=2),
                lambda r: r.update(plan_sha256="0" * 64), lambda r: r.update(version=True)):
            wrong = deepcopy(reply)
            change(wrong)
            self.publish_held_plan(event, wrong)
            with self.subTest(reply=wrong), self.assertRaisesRegex(ValueError, "execute reply"):
                self.authenticate()
        duplicate = "[PLAN] " + json.dumps(event) + "\n" + ("[INTERACTION-REPLY] " + json.dumps(reply) + "\n") * 2
        with self.assertRaisesRegex(ValueError, "conflicting execute"):
            console.confirmed_console_plan(duplicate)
        self.publish_held_plan(event, reply)
        self.authenticate()


if __name__ == "__main__":
    unittest.main()
