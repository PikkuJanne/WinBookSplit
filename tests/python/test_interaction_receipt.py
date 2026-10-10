"""Synthetic handshake mutations only; no Windows, viewer or consent pass."""

from copy import deepcopy
import importlib.util
from pathlib import Path
import unittest

ROOT = Path(__file__).resolve().parents[2]

def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module

validator = load("wbs_interaction_receipt_unit", ROOT / "tests/manual/interaction_receipts.py")
fixture = load("wbs_interaction_fixture_unit", ROOT / "tests/python/session_receipt_fixture.py")

def complete():
    output = {"sequence": 1, "title": "Authored", "start": 0, "end": 2, "filename": "01 - Authored.pdf",
              "sha256": "b" * 64, "page_count": 2, "size_bytes": 100}
    execution = {"mode": "manual", "total_pages": 2, "source_identity": {"sha256": "a" * 64},
                 "coverage": {"complete": True, "covered_pages": 2, "section_count": 1}, "written_count": 1, "outputs": [output]}
    invocation = {"path": "C:/synthetic/python.exe", "arguments": ["-I", "-B", "-X", "utf8", "engine.py", "book.pdf", "output", "manual", "", "--interactive", "f" * 32]}
    record = fixture.make(execution, invocation, ["input_ready", "plan_ready"], ["starts", "execute"])
    summary = {"InteractionError": None, "InputError": None, "InputWriterStopped": True,
               "QueuedReplyCount": 2, "InteractionCount": 2, "ReplyCount": 2}
    return record, [summary], execution

class InteractionReceiptTests(unittest.TestCase):
    def test_one_captured_plan_and_literal_starts_are_bound_to_execution(self):
        record, summaries, execution = complete()
        validator.validate(record, summaries, ["input_ready", "plan_ready"], ["starts", "execute"], starts="2,3", execution=execution)

    def test_nonce_sequence_stage_action_and_totals_cannot_be_forged(self):
        original, totals, execution = complete()
        for mutation in (lambda r, s: r["replies"][0].update(session="e" * 32),
                         lambda r, s: r["requests"][1].update(sequence=True),
                         lambda r, s: r["requests"][0].update(stage="plan_ready"),
                         lambda r, s: r["replies"][0].update(starts="1"),
                         lambda r, s: s[0].update(ReplyCount=1),
                         lambda r, s: s.append(deepcopy(s[0])),
                         lambda r, s: r["engine_invocations"].append(deepcopy(r["engine_invocations"][0]))):
            record, summaries = deepcopy(original), deepcopy(totals)
            mutation(record, summaries)
            with self.assertRaises(ValueError):
                validator.validate(record, summaries, ["input_ready", "plan_ready"], ["starts", "execute"], starts="2,3", execution=execution)

    def test_changed_or_unconfirmed_plan_cannot_claim_written_sections(self):
        original, summaries, writer = complete()
        for mutation in (lambda r, e: r["replies"][1].update(plan_sha256="e" * 64),
                         lambda r, e: r["displayed_plans"].clear(),
                         lambda r, e: r["requests"][1].update(plan_json="{}"),
                         lambda r, e: e["outputs"][0].update(end=1),
                         lambda r, e: e.update(written_count=0)):
            record, execution = deepcopy(original), deepcopy(writer)
            mutation(record, execution)
            with self.assertRaises(ValueError):
                validator.validate(record, summaries, ["input_ready", "plan_ready"], ["starts", "execute"], starts="2,3", execution=execution)

    def test_failure_before_plan_requires_zero_requests_and_no_consent(self):
        record, summaries, _ = complete()
        record.update(requests=[], replies=[], displayed_plans=[])
        summaries[0].update(InteractionCount=0, ReplyCount=0, QueuedReplyCount=0)
        validator.validate(record, summaries, [], [])
        record["replies"] = [{"action": "execute"}]
        with self.assertRaises(ValueError):
            validator.validate(record, summaries, [], [])

    def test_noninteractive_frames_cannot_receive_fake_consent(self):
        record, summaries, _ = complete()
        with self.assertRaises(ValueError):
            validator.require_noninteractive(record, summaries)
        record.update(requests=[], replies=[], displayed_plans=[])
        record["engine_invocations"][0]["arguments"] = record["engine_invocations"][0]["arguments"][:9]
        summaries[0].update(InteractionCount=0, ReplyCount=0)
        validator.require_noninteractive(record, summaries)

if __name__ == "__main__":
    unittest.main(verbosity=2)
