"""STORE-006/028 contract regressions; synthetic receipts are NEVER native evidence.

The immutable v1.0.2 reports are read as historical evidence. Every newly written
report/receipt below is an explicitly labelled, temporary SYNTHETIC unit fixture.
The checker only establishes declaration/byte consistency, not execution truth.
"""
from copy import deepcopy
from hashlib import sha256
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[2]
HISTORY = ROOT / "docs/codex-v1.0.2/evidence"
TEMPLATE = ROOT / "docs/codex-v1.0.3/templates/native-store-report.json"
SYNTHETIC_COMMIT = "a" * 40
SYNTHETIC_ZIP_SHA256 = "b" * 64
SYNTHETIC_NOTE = "SYNTHETIC UNIT FIXTURE: no Windows/native/application execution occurred."


def load_helper(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


gate = load_helper("wbs_store_gate_under_test", ROOT / "tools/codex-v1.0.3/check_native_store_report.py")
legacy = load_helper("wbs_store_gate_legacy_consumer", ROOT / "tools/compatibility/native_smoke.py")


def synthetic_contract():
    """A complete declaration-only control, visibly labelled in report/receipts.

    native_execution/passed are the checker vocabulary being tested; they are
    fictional declarations here and must never be promoted to acceptance data.
    """
    report = json.loads(TEMPLATE.read_bytes())
    report.update({
        "evidence_kind": "native_execution", "status": "passed",
        "run_id": "SYNTHETIC_UNIT_ONLY_006", "source_commit": SYNTHETIC_COMMIT,
        "zip_sha256": SYNTHETIC_ZIP_SHA256, "notes": SYNTHETIC_NOTE,
        "synthetic_test_fixture": True,
    })
    report["host"] = {
        "execution_kind": "native", "owner_desktop": True, "os": "Windows 10",
        "build": 19045, "architecture": "x64", "powershell_version": "5.1.19041.1",
    }
    report["registration"] = {
        "family": "PythonSoftwareFoundation.PythonManager_qbz5n2kfra8p0",
        "publisher": "CN=4975D53F-AA7E-49A5-8B49-EA4FDC1BB66B",
        "signature_kind": "Store", "status": "Ok", "architecture": "X64",
        "version": "26.4.999.0",
    }
    report["forbidden_substitutions"] = {key: False for key in report["forbidden_substitutions"]}
    report["fallback_acquisition"] = {
        "policy": "disabled_before_discovery", "attempted": False, "used": False,
    }
    report["inventory"] = {"purpose": "verification_only", "supplied_to_application": False}
    runtime = {"runtime_ref": "SYNTHETIC_STORE_RUNTIME", "sha256": "c" * 64,
               "version": "3.15.0", "managed_by_store": True}
    report["before"] = {"application_venv_exists": False, "runtimes": [deepcopy(runtime)]}
    report["after"] = {"application_venv_exists": True, "runtimes": [deepcopy(runtime)]}
    report["discovery"] = {
        "mode": "automatic", "route": "registered_store_manager",
        "provenance": "pre_existing_owner_install", "runtime_ref": runtime["runtime_ref"],
        "version": "3.15.0", "interpreter_probe_exit": 0, "package_state": "setup_required",
    }
    report["setup"] = {
        "entry": "Setup-WinBookSplit.bat", "runtime_selection": "automatic", "exit_code": 0,
        "base_runtime_ref": runtime["runtime_ref"], "venv_base_runtime_ref": runtime["runtime_ref"],
    }
    report["checks"] = {key: True for key in report["checks"]}
    report['policy_component'].update(status='PASS', **dict.fromkeys(
        ('direct_remotesigned_denied','marked_setup_bat_noop','exact_formatter_action_noop',
         'source_bytes_and_marks_preserved','venv_unchanged','inbox_policy_unchanged'),True))
    return report


def write_synthetic_contract(directory, report, *, preservation_record=None):
    """Hash-bind labelled temporary receipts, including any negative mutation."""
    report = deepcopy(report)
    report["artifacts"] = []
    for kind, summary in gate.expected_summaries(report).items():
        record = {"fixture_kind": "SYNTHETIC", "note": SYNTHETIC_NOTE}
        if kind == "preservation" and preservation_record is not None:
            record.update(preservation_record)
        probe = {"Pid": 12345, "ExitCode": 0, "ElapsedSeconds": 0.1,
                 "Stdout": "SYNTHETIC", "Stderr": "", "StdoutTotalBytes": 9, "StderrTotalBytes": 0,
                 "redacted_stream_bytes": {"stdout": 9, "stderr": 0},
                 "raw_stream_sha256": {"stdout": sha256(b"SYNTHETIC").hexdigest(), "stderr": sha256(b"").hexdigest()}}
        probe.update(dict.fromkeys(("JobAssigned", "ParentStopped", "DescendantsStopped", "StreamsComplete", "InputWriterStopped"), True))
        probe.update(dict.fromkeys(("TimedOut", "Cancelled", "StdoutTruncated", "StderrTruncated"), False))
        probe.update(dict.fromkeys(("StartError", "StopError", "StreamError", "ResultError", "InteractionError", "InputError"), None))
        application = {"role": "application", "environment": {}, "synthetic_test_fixture": True,
                       "process": {"command": ["SYNTHETIC_APPLICATION.bat"], "cwd": "SYNTHETIC_CWD", "answers": [],
                                   "environment_before_after_restored": True, "failure": None, "probe": probe}}
        receipt = {
            "schema_version": gate.SCHEMA_VERSION, "kind": kind, "run_id": report["run_id"],
            "source_commit": SYNTHETIC_COMMIT, "zip_sha256": SYNTHETIC_ZIP_SHA256,
            "summary": summary, "records": [record] +
            [deepcopy(application) for _ in range(gate.APPLICATION_RECORD_MINIMUM.get(kind, 0))],
            "synthetic_test_fixture": True,
        }
        if kind == 'preservation':
            # Synthetic literal command controls, never native execution data.
            policy={'effective':'Bypass','scopes':[{'scope':scope,'policy':'Bypass' if scope=='Process' else 'Undefined'}
                for scope in ('MachinePolicy','UserPolicy','Process','CurrentUser','LocalMachine')]}
            marks=[{'path':'<APP>\\'+leaf,'sha256':'d'*64,'ads':'[ZoneTransfer]\r\nZoneId=3\r\n'}
                   for leaf in ('WinBookSplit.ps1','Setup-WinBookSplit.ps1')]
            denial=deepcopy(application);denial.update(role='expected_policy_denial',stage='direct_file_remotesigned')
            denial['process']['command']=['SYNTHETIC_INBOX','-NoProfile','-NonInteractive','-InputFormat','Text','-OutputFormat','Text',
                '-ExecutionPolicy','RemoteSigned','-File',marks[0]['path'],'-Version']
            denial['process']['probe'].update(ExitCode=1,Stdout='',Stderr='SYNTHETIC PSSecurityException',StdoutTotalBytes=0,StderrTotalBytes=29,
                redacted_stream_bytes={'stdout':0,'stderr':29},
                raw_stream_sha256={'stdout':sha256(b'').hexdigest(),'stderr':sha256(b'SYNTHETIC PSSecurityException').hexdigest()})
            bat=deepcopy(application);bat['stage']='marked_setup_bat_check'
            action=deepcopy(application);action['stage']='exact_automatic_formatter_action'
            action['action']="& 'SYNTHETIC_INBOX' -NoProfile -ExecutionPolicy Bypass -File '<APP>\\Setup-WinBookSplit.ps1' -NonInteractive"
            output='[EXACT_ACTION] '+action['action']
            action['process']['probe'].update(Stdout=output,StdoutTotalBytes=len(output),
                redacted_stream_bytes={'stdout':len(output),'stderr':0},
                raw_stream_sha256={'stdout':sha256(output.encode()).hexdigest(),'stderr':sha256(b'').hexdigest()})
            receipt['records'] += [{'role':'policy_observation','stage':'before','inbox_policy':policy},
                {'role':'policy_observation','stage':'authored_motw','marks':marks},denial,bat,action,
                {'role':'policy_observation','stage':'after','inbox_policy':policy}]
        raw = (json.dumps(receipt, ensure_ascii=True, indent=2) + "\n").encode("utf-8")
        name = "SYNTHETIC-" + kind + ".json"
        (directory / name).write_bytes(raw)
        report["artifacts"].append({"kind": kind, "path": name, "sha256": sha256(raw).hexdigest()})
    path = directory / "SYNTHETIC-report.json"
    path.write_text(json.dumps(report, ensure_ascii=True, indent=2) + "\n", encoding="utf-8")
    return path


class StoreExistingGateTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory(prefix="wbs-SYNTHETIC-store-gate-")
        self.addCleanup(temporary.cleanup)
        self.directory = Path(temporary.name)

    def validate_synthetic(self, report, **kwargs):
        path = write_synthetic_contract(self.directory, report, **kwargs)
        return gate.validate(path, SYNTHETIC_COMMIT, SYNTHETIC_ZIP_SHA256)

    def test_historical_private_runtime_pass_is_still_valid_under_its_old_contract(self):
        path = HISTORY / "V102-11-windows10.json"
        before = path.read_bytes()
        historical = json.loads(before)
        self.assertEqual(historical["status"], "PASS")
        raw = historical["returned_report"]
        self.assertEqual(raw["oracle"]["runtime_versions"]["pypdf"], "6.20.0")
        for route in raw["marked_setup_regressions"]["routes"]:
            for child in route["child_hosts"]:
                self.assertEqual(child["pypdf_version"], None if child["stage"] == "broken-default" else "6.20.0")
        with patch.object(legacy, "RUNTIME_POLICY", None):
            accepted = legacy.export_report(raw)
        self.assertTrue(accepted["oracle"]["success"])
        self.assertFalse(accepted["runtime_selection"]["owner_regular315_usable"])
        self.assertFalse(accepted["runtime_selection"]["owner_selected"])
        self.assertEqual(accepted["runtime_selection"]["baseline_kind"], "reviewed_private3150_archive")
        self.assertEqual(accepted["runtime_selection"]["selected_identity"]["source"], "explicit")
        self.assertEqual(path.read_bytes(), before)

    def test_real_historical_review_and_return_cannot_satisfy_store_gate(self):
        for name in ("V102-11-windows10-review.json", "V102-11-windows10.json"):
            with self.subTest(historical=name):
                path = HISTORY / name
                before = path.read_bytes()
                historical = json.loads(before)
                selection = historical.get("runtime_selection", historical.get("returned_report", {}).get("runtime_selection"))
                self.assertFalse(selection["owner_regular315_usable"])
                self.assertEqual(selection["baseline_kind"], "reviewed_private3150_archive")
                production = historical.get("production", historical)
                commit = production["source_commit"]
                zip_hash = (production["assets"]["WinBookSplit-v1.0.2.zip"]["sha256"]
                            if "assets" in production else historical["release_zip_sha256"])
                with self.assertRaisesRegex(gate.GateError, "unsupported report schema"):
                    gate.validate(path, commit, zip_hash)
                self.assertEqual(path.read_bytes(), before)

    def test_original_returned_report_cannot_satisfy_store_gate(self):
        historical = json.loads((HISTORY / "V102-11-windows10.json").read_bytes())
        # Copy exact old-schema data only to a test-owned path. Do not relabel it.
        path = self.directory / "historical-returned-report.json"
        path.write_text(json.dumps(historical["returned_report"]), encoding="utf-8")
        with self.assertRaisesRegex(gate.GateError, "unsupported report schema"):
            gate.validate(path, historical["production"]["source_commit"],
                          historical["production"]["assets"]["WinBookSplit-v1.0.2.zip"]["sha256"])

    def test_complete_synthetic_contract_is_consistent_positive_control_only(self):
        result = self.validate_synthetic(synthetic_contract())
        self.assertEqual(result["status"], "CONSISTENT_NATIVE_REPORT")
        self.assertEqual(result["receipt_count"], 6)
        self.assertIn("not physical execution truth", result["limitation"])

    def test_used_private_fallback_disqualifies_otherwise_complete_contract(self):
        report = synthetic_contract()
        report["forbidden_substitutions"]["private_runtime_fallback_used"] = True
        with self.assertRaisesRegex(gate.GateError, "disqualifying or unknown substitution: private_runtime_fallback_used"):
            self.validate_synthetic(report)

    def test_unknown_or_missing_fallback_usage_cannot_satisfy_gate(self):
        for state in (None, "false", "missing"):
            with self.subTest(state=state):
                report = synthetic_contract()
                if state == "missing":
                    del report["forbidden_substitutions"]["private_runtime_fallback_used"]
                else:
                    report["forbidden_substitutions"]["private_runtime_fallback_used"] = state
                with self.assertRaises(gate.GateError):
                    self.validate_synthetic(report)

    def test_each_hidden_runtime_substitution_disqualifies_contract(self):
        for key in ("runtime_supplied_by_kit", "runtime_copied_or_cloned", "runtime_environment_injected",
                    "explicit_python_path_supplied",
                    "hidden_runtime_selection_supplied", "child_path_injected",
                    "application_venv_prebuilt"):
            with self.subTest(substitution=key):
                report = synthetic_contract()
                report["forbidden_substitutions"][key] = True
                with self.assertRaisesRegex(gate.GateError, key):
                    self.validate_synthetic(report)

    def test_attempted_private_acquisition_disqualifies_even_when_not_used(self):
        # This was T01's reproduced expected failure. T05 requires rejecting any
        # attempted acquisition even if no downloaded/supplied runtime was used.
        report = synthetic_contract()
        attempt = {"policy": "disabled_before_discovery", "attempted": True, "used": False}
        report["fallback_acquisition"] = attempt
        with self.assertRaisesRegex(gate.GateError, "fallback|substitution|acquisition"):
            self.validate_synthetic(report, preservation_record=attempt)

    def test_attempted_fallback_flag_also_disqualifies_even_if_acquisition_block_claims_false(self):
        report = synthetic_contract()
        report["forbidden_substitutions"]["private_runtime_fallback_attempted"] = True
        with self.assertRaisesRegex(gate.GateError, "private_runtime_fallback_attempted"):
            self.validate_synthetic(report)

    def test_old_store_schema_cannot_be_promoted_by_adding_new_fields(self):
        report = synthetic_contract()
        report["schema_version"] = 1
        with self.assertRaisesRegex(gate.GateError, "unsupported report schema"):
            self.validate_synthetic(report)

    def test_independent_inventory_may_not_select_runtime_for_application(self):
        report = synthetic_contract()
        report["inventory"]["supplied_to_application"] = True
        with self.assertRaisesRegex(gate.GateError, "independent inventory"):
            self.validate_synthetic(report)

    def test_receipt_attempt_cannot_hide_behind_false_summary_flags(self):
        with self.assertRaisesRegex(gate.GateError, "private_runtime_fallback_attempted"):
            self.validate_synthetic(synthetic_contract(), preservation_record={"private_runtime_fallback_attempted": True})

    def test_receipt_private_baseline_cannot_hide_behind_false_summary_flags(self):
        with self.assertRaisesRegex(gate.GateError, "baseline"):
            self.validate_synthetic(synthetic_contract(), preservation_record={
                "runtime_selection": {"baseline_kind": "reviewed_private3150_archive", "owner_regular315_usable": False},
            })


if __name__ == "__main__":
    unittest.main(verbosity=2)
