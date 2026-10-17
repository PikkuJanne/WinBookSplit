"""STORE-006 contract regressions; synthetic receipts are NEVER native evidence.

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
    return report


def write_synthetic_contract(directory, report, *, preservation_record=None):
    """Hash-bind labelled temporary receipts, including any negative mutation."""
    report = deepcopy(report)
    report["artifacts"] = []
    for kind, summary in gate.expected_summaries(report).items():
        record = {"fixture_kind": "SYNTHETIC", "note": SYNTHETIC_NOTE}
        if kind == "preservation" and preservation_record is not None:
            record.update(preservation_record)
        receipt = {
            "schema_version": 1, "kind": kind, "run_id": report["run_id"],
            "source_commit": SYNTHETIC_COMMIT, "zip_sha256": SYNTHETIC_ZIP_SHA256,
            "summary": summary, "records": [record], "synthetic_test_fixture": True,
        }
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
                with self.assertRaisesRegex(gate.GateError, "not the existing-owner Store scenario"):
                    gate.validate(path, commit, zip_hash)
                self.assertEqual(path.read_bytes(), before)

    def test_original_returned_report_cannot_satisfy_store_gate(self):
        historical = json.loads((HISTORY / "V102-11-windows10.json").read_bytes())
        # Copy exact old-schema data only to a test-owned path. Do not relabel it.
        path = self.directory / "historical-returned-report.json"
        path.write_text(json.dumps(historical["returned_report"]), encoding="utf-8")
        with self.assertRaisesRegex(gate.GateError, "not the existing-owner Store scenario"):
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
        for key in ("runtime_copied_or_cloned", "explicit_python_path_supplied",
                    "hidden_runtime_selection_supplied", "child_path_injected",
                    "application_venv_prebuilt"):
            with self.subTest(substitution=key):
                report = synthetic_contract()
                report["forbidden_substitutions"][key] = True
                with self.assertRaisesRegex(gate.GateError, key):
                    self.validate_synthetic(report)

    @unittest.expectedFailure
    def test_attempted_private_acquisition_disqualifies_even_when_not_used(self):
        # T01 preserved RED requirement: the supplied checker only checks USED.
        # An attempted acquisition must also block this scenario. The hashes and
        # all other declarations remain valid, so malformed evidence is no excuse.
        report = synthetic_contract()
        attempt = {"private_runtime_fallback_attempted": True,
                   "private_runtime_fallback_used": False,
                   "outcome": "SYNTHETIC acquisition refused before download"}
        report["fallback_acquisition"] = attempt
        with self.assertRaisesRegex(gate.GateError, "fallback|substitution|acquisition"):
            self.validate_synthetic(report, preservation_record=attempt)


if __name__ == "__main__":
    unittest.main(verbosity=2)
