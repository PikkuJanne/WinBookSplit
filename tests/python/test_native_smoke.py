"""Targeted oracle controls; synthetic receipts never count as native acceptance."""
import ast
import copy
import importlib.util
import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from pypdf import PdfReader, PdfWriter

ROOT = Path(__file__).resolve().parents[2]
HELPER_ROOT = ROOT / "tools/compatibility"
spec = importlib.util.spec_from_file_location("wbs_native_smoke_under_test", HELPER_ROOT / "native_smoke.py")
smoke = importlib.util.module_from_spec(spec)
spec.loader.exec_module(smoke)


def synthetic_probe():
    return {"Pid": 91, "ExitCode": 0, "JobAssigned": True, "ParentStopped": True, "DescendantsStopped": True,
            "StreamsComplete": True, "InputWriterStopped": True, "TimedOut": False, "Cancelled": False,
            "StdoutTruncated": False, "StderrTruncated": False, "StartError": None, "StopError": None,
            "StreamError": None, "ResultError": None, "InteractionError": None, "InputError": None,
            "Stdout": "", "Stderr": "", "StdoutTotalBytes": 0, "StderrTotalBytes": 0, "ElapsedSeconds": 0.2}


def synthetic_export(pypdf_version="6.21.1"):
    oracle = {"protocol": "winbooksplit.native-oracle", "schema_version": 1, "success": True,
              "plan_sha256": "a" * 64, "application_version": "1.0.1", "policy_sha256": "b" * 64,
              "runtime_versions": {"python": "3.15.0", "pypdf": pypdf_version, "calibre": "9.15.0"},
              "source_unchanged": True, "cases": [{"id": name, "status": "PASS", "exit_code": 0, "error_code": None} for name in smoke.CASE_IDS],
              "limits": ["No Explorer gesture was performed.", "All native artifacts are preserved locally."]}
    policies = [{"scope": scope, "policy": "Bypass" if scope == "Process" else "Undefined"}
                for scope in ("MachinePolicy", "UserPolicy", "Process", "CurrentUser", "LocalMachine")]
    identity = {"source": "explicit", "executable_sha256": "f"*64, "version": "3.15.0", "implementation": "cpython",
                "platform": "win32", "machine": "AMD64", "bits": 64, "gil_disabled": False, "releaselevel": "final",
                "serial": 0, "isolated": True, "dont_write_bytecode": True, "policy_sha256": "b"*64}
    routes = []
    for identifier in ("PS51-fresh", "PS51-stale", "PS51-marked-bat", "CMD-marked-bat", "PS51-organization-simulation"):
        mock, bat = identifier.endswith("-organization-simulation"), identifier.endswith("-marked-bat")
        stages = ("simulation",) if mock else ("fresh", "check", "noop", "zeroargs-redirected", "broken-default", "repair", "pdf") if bat else ("direct", "before", "action", "retry")
        child_stages = () if mock else stages[:-1] if bat else ("action",)
        route = {"id": identifier, "status": "PASS", "kind": "MOCKED_POLICY_QUERY" if mock else "BAT" if bat else "ACTION",
                 "fixture": "MOCKED_POLICY_QUERY" if mock else "AUTHORED_BROKEN_VENV" if bat else "AUTHORED_BROKEN_VENV; no legacy Python execution" if identifier.endswith("-stale") else "ACTUAL_BARE315",
                 "exits": [{"stage": stage, "exit_code": 1 if mock or stage == "direct" else 3 if stage in {"before", "broken-default"} else 0} for stage in stages],
                 "child_hosts": [{"stage": stage, "executable_status": "observed", "version": "5.1.19041.6000", "edition": "Desktop",
                                  "execution_policy": {"status": "observed", "effective": "Bypass", "scopes": copy.deepcopy(policies), "reason_code": None},
                                  "selected_version": None if stage == "broken-default" else "3.15.0",
                                  "pypdf_version": None if stage == "broken-default" else pypdf_version} for stage in child_stages],
                 "pdf": None if mock else {"success": True, "runtime_version": "3.15.0", "written_count": 3, "physical_pages": 10, "source_sha256": "c"*64}}
        if mock:
            route["simulated_positive_launch_refused"] = True
        routes.append(route)
    return {"oracle": oracle, "execution": "USER_LAUNCHED_AUTOMATED",
            "runtime_selection": {"owner_version": "3.15.0", "owner_regular315_usable": True, "owner_selected": True,
                "selected_version": "3.15.0", "required_baseline_version": "3.15.0", "baseline_version": "3.15.0",
                "baseline_kind": "existing3150_without_pypdf", "reason": "usable_regular315_discovered",
                "selected_identity": copy.deepcopy(identity), "baseline_identity": copy.deepcopy(identity)},
            "marked_setup_regressions": {"status": "PASS", "required_runtime_version": "3.15.0", "exact_asset_bound": True,
                "Action_roundtrips": 2, "marked_BAT_routes": 2, "mocked_policy_refusals": 1, "raw_receipt_sha256": "d"*64,
                "probe_sha256": "e"*64, "actual_inbox_policy_snapshots": {"before": policies, "after": policies},
                "fixture": "Authored broken venv, not an owner legacy runtime; marks assigned by automation; all source/ADS/PDF checks verified", "routes": routes},
            "os": {"Caption": "Microsoft Windows 10 Pro", "Version": "10.0.19045",
            "BuildNumber": "19045", "OSArchitecture": "64-bit", "ProductType": 1, "UBR": 6000,
            "powershell_version": "5.1.19041.6000", "process_bits": 64, "os_bits": 64, "DisplayVersion": "22H2",
            "processor_architectures": [9], "platform_scope": "NATIVE_WINDOWS10"}, "private": "C:\\Users\\secret-owner\\file.pdf"}


class PureOracleControls(unittest.TestCase):
    def test_store_positive_recipes_validate_then_reject_command_substitution(self):
        # Synthetic receipts exercise the complete new oracle function; actual
        # native PDF recipes are also run as explicitly labelled developer work.
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder).resolve()
            app = root / "app"
            app.mkdir()
            python = app / ".venv/Scripts/python.exe"
            runtime = {"version": "3.14.8", "ok": True}
            cases, rows, immutable = [], [], []
            for identifier in ("manual", "auto1", "auto2", "bat"):
                base = root / identifier
                base.mkdir()
                published, console = base / "publication", base / "console"
                published.mkdir()
                console.mkdir()
                log = console / "console.log"
                log.write_text('[DEPENDENCY] ' + json.dumps({"Runtime": {"Path": str(python), "Source": "application_venv", "Details": runtime}}), encoding="utf-8")
                source = root / (identifier + ".pdf")
                source.write_bytes(b"Authored immutable source")
                immutable.append(smoke.observation(source))
                arguments = ["/d", "/c", "authored.cmd"] if identifier == "bat" else ["-File", str(app / "WinBookSplit.ps1")]
                selector = "cmd" if identifier == "bat" else "inbox_ps"
                cases.append({"id": identifier, "exe_selector": selector, "arguments": arguments, "cwd": str(root), "answers": [], "environment": {},
                              "source_path": str(source), "output_base": str(base), "ranges": [[0, 1]], "prior_members": []})
                probe = synthetic_probe()
                final = {"protocol": "winbooksplit.outcome", "version": 1, "status": "success", "code": "split_complete", "message": "Authored synthetic receipt", "mode": "Manual",
                         "exit_code": 0, "written_count": 1, "final_directory": str(published), "engine_result": {"execution": {}}}
                probe["Stdout"] = '[OUTCOME] ' + json.dumps(final) + '\nLog: ' + str(log) + '\n'
                probe["StdoutTotalBytes"] = len(probe["Stdout"].encode("utf-8"))
                executable = str(root / ("cmd.exe" if identifier == "bat" else "WindowsPowerShell/v1.0/powershell.exe"))
                rows.append({"id": identifier, "command": [executable, *arguments], "cwd": str(root), "answers": [], "environment_before_after_restored": True, "failure": None, "probe": probe})
            plan = {"protocol": "winbooksplit.native-store-plan", "schema_version": 2, "application_root": str(app), "policy_sha256": "b"*64,
                    "python_path": str(python), "runtime": runtime, "immutable_files": immutable, "cases": cases}
            raw = {"protocol": "winbooksplit.native-store-process", "schema_version": 2, "plan_sha256": "a"*64, "cases": rows}
            with patch.object(smoke, "bind_application", return_value=app), patch.object(smoke, "windows_system_directory", return_value=str(root)), \
                 patch.object(smoke, "pdf_snapshot", return_value={}), patch.object(smoke, "validate_publication", side_effect=lambda base, *unused: base / "publication"), \
                 patch.object(smoke, "authenticate_console_manifest", return_value={}):
                self.assertTrue(smoke.validate_store(plan, raw, "a"*64)["success"])
                rows[0]["command"][0] = "C:\\authored-decoy\\powershell.exe"
                with self.assertRaisesRegex(smoke.OracleError, "store_process_recipe_differ"):
                    smoke.validate_store(plan, raw, "a"*64)

    def test_oracle_never_launches_or_deletes(self):
        source = (HELPER_ROOT / "native_smoke.py").read_text(encoding="utf-8")
        tree = ast.parse(source)
        imports = {node.module for node in ast.walk(tree) if isinstance(node, ast.ImportFrom)}
        self.assertNotIn("subprocess", imports)
        self.assertNotIn("reportlab", imports)
        self.assertNotIn("PIL", imports)
        self.assertFalse(any(isinstance(node, ast.Attribute) and node.attr in {"unlink", "rmdir", "rmtree", "Popen", "run", "system"} for node in ast.walk(tree)))

    def test_duplicate_and_nonfinite_json_fail_closed(self):
        for raw in ('{"x":1,"x":2}', '{"x":NaN}', '{"x":Infinity}'):
            with self.subTest(raw=raw), self.assertRaises(ValueError):
                smoke.unique_json(raw)

    def test_native_plan_and_process_schema_reject_boolean_alias_and_unknown_fields(self):
        cases = [{"id": name} for name in smoke.CASE_IDS]
        with self.assertRaisesRegex(smoke.OracleError, "plan_schema_invalid"):
            smoke.validate({"protocol": "winbooksplit.native-plan", "schema_version": True, "cases": cases}, {}, "a"*64)
        plan = {"protocol": "winbooksplit.native-plan", "schema_version": 1, "cases": cases,
                "application_root": "C:\\synthetic\\app", "policy_sha256": "b"*64, "application_version": "1.0.1"}
        raw = {"protocol": "winbooksplit.native-process", "schema_version": 1, "plan_sha256": "a"*64, "cases": cases}
        with patch.object(smoke, "bind_application"), patch.object(smoke, "APP_VERSION", "1.0.1"):
            for bad in ({**raw, "schema_version": True}, {**raw, "extra": "PRIVATE"}):
                with self.assertRaisesRegex(smoke.OracleError, "process_report_binding_invalid"):
                    smoke.validate(plan, bad, "a"*64)

    def test_fixture_manifest_schema_boolean_alias_rejected(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder).resolve()
            smoke.write_json(root / "native-fixtures.json", {"protocol": "winbooksplit.native-fixtures", "schema_version": True,
                             "license": "MIT", "authored_original": True, "remote_resources": False})
            with self.assertRaisesRegex(smoke.OracleError, "fixture_provenance_invalid"):
                smoke.fixture_catalog(root)

    def test_complete_synthetic_probe_shape_accepted(self):
        self.assertIs(smoke.native_probe(probe := synthetic_probe()), probe)

    def test_each_unsafe_synthetic_transport_flag_rejected(self):
        for key in ("JobAssigned", "ParentStopped", "DescendantsStopped", "StreamsComplete", "InputWriterStopped"):
            for value in (False, 1, None):
                row = synthetic_probe()
                row[key] = value
                with self.subTest(key=key, value=value), self.assertRaises(smoke.OracleError):
                    smoke.native_probe(row)
        for key in ("TimedOut", "Cancelled", "StdoutTruncated", "StderrTruncated"):
            row = synthetic_probe()
            row[key] = True
            with self.subTest(key=key), self.assertRaises(smoke.OracleError):
                smoke.native_probe(row)
        for key in ("StartError", "StopError", "StreamError", "ResultError", "InteractionError", "InputError"):
            row = synthetic_probe()
            row[key] = "synthetic error"
            with self.subTest(key=key), self.assertRaises(smoke.OracleError):
                smoke.native_probe(row)

    def test_missing_synthetic_transport_error_field_rejected(self):
        for key in ("StartError", "StopError", "StreamError", "ResultError", "InteractionError", "InputError"):
            row = synthetic_probe()
            del row[key]
            with self.subTest(key=key), self.assertRaisesRegex(smoke.OracleError, "native_transport_error"):
                smoke.native_probe(row)

    def test_boolean_alias_native_exit_rejected(self):
        row = synthetic_probe()
        row["ExitCode"] = False
        with self.assertRaises(smoke.OracleError):
            smoke.native_probe(row)

    def test_typed_exact_utf8_stream_lengths_required(self):
        row = synthetic_probe()
        row.update(Stdout="Å日本", StdoutTotalBytes=len("Å日本".encode("utf-8")))
        smoke.native_probe(row)
        for key in ("StdoutTotalBytes", "StderrTotalBytes"):
            for value in (True, -1, 999):
                bad = copy.deepcopy(row)
                bad[key] = value
                with self.subTest(key=key, value=value), self.assertRaises(smoke.OracleError):
                    smoke.native_probe(bad)

    def test_terminal_frame_duplicates_unknown_keys_and_bool_rejected(self):
        row = {"protocol": "winbooksplit.outcome", "version": 1, "status": "success", "code": "split_complete", "message": "done",
               "exit_code": 0, "mode": "manual", "written_count": 1, "final_directory": None, "engine_result": None}
        self.assertEqual(smoke.outcome("[OUTCOME] " + json.dumps(row)), row)
        for bad in ({**row, "exit_code": False}, {**row, "unknown": True}, {**row, "version": True}):
            with self.assertRaises(smoke.OracleError):
                smoke.outcome("[OUTCOME] " + json.dumps(bad))
        with self.assertRaises(smoke.OracleError):
            smoke.outcome("[OUTCOME] " + json.dumps(row) + "\n[OUTCOME] " + json.dumps(row))

    def test_export_omits_planted_private_raw_fields(self):
        raw = synthetic_export()
        raw["os"]["CSName"] = "SECRET_HOST"
        raw["probe"] = {"Stdout": "SECRET_USERNAME PRIVATE_DOCUMENT"}
        result = smoke.export_report(raw)
        text = json.dumps(result)
        for marker in ("secret-owner", "SECRET_HOST", "SECRET_USERNAME", "PRIVATE_DOCUMENT"):
            self.assertNotIn(marker, text)
        self.assertEqual(result["execution"], "USER_LAUNCHED_AUTOMATED")

    def test_standalone_export_accepts_coherent_historical_and_current_package_versions(self):
        for version in ("6.20.0", "6.21.1"):
            raw = synthetic_export(pypdf_version=version)
            with self.subTest(version=version), patch.object(smoke, "RUNTIME_POLICY", None):
                result = smoke.export_report(raw)
            self.assertEqual(result["oracle"]["runtime_versions"]["pypdf"], version)
            self.assertTrue(result["oracle"]["success"])
            for route in result["marked_setup_regressions"]["routes"]:
                for child in route["child_hosts"]:
                    self.assertEqual(child["pypdf_version"], None if child["stage"] == "broken-default" else version)

    def test_standalone_export_rejects_header_and_child_package_version_mismatches(self):
        for version, other in (("6.20.0", "6.21.1"), ("6.21.1", "6.20.0")):
            for target in ("header", "child", "broken-child"):
                raw = synthetic_export(pypdf_version=version)
                if target == "header":
                    raw["oracle"]["runtime_versions"]["pypdf"] = other
                else:
                    children = [child for route in raw["marked_setup_regressions"]["routes"]
                                for child in route["child_hosts"]]
                    child = next(child for child in children
                                 if (child["stage"] == "broken-default") == (target == "broken-child"))
                    child["pypdf_version"] = other
                with self.subTest(version=version, target=target), patch.object(smoke, "RUNTIME_POLICY", None), \
                        self.assertRaisesRegex(smoke.OracleError, "redacted_child_runtime_invalid"):
                    smoke.export_report(raw)

    def test_export_refuses_case_leakage_and_false_success(self):
        mutations = [lambda row: row["oracle"]["cases"][0].update(path="C:\\Users\\PRIVATE"),
                     lambda row: row["oracle"]["cases"][0].update(error_code="C:\\Users\\PRIVATE"),
                     lambda row: row["oracle"]["cases"][0].update(status="FAIL"),
                     lambda row: row["oracle"].update(source_unchanged=False),
                     lambda row: row["oracle"].update(limits=["PRIVATE_PATH"]),
                     lambda row: row["oracle"]["runtime_versions"].update(python="PRIVATE_USERNAME"),
                     lambda row: row["os"].update(ProductType=3),
                     lambda row: row["os"].update(Caption="PRIVATE_HOST")]
        for mutate in mutations:
            row = synthetic_export()
            mutate(row)
            with self.subTest(mutate=mutate), self.assertRaises(smoke.OracleError):
                smoke.export_report(row)

    def test_owner_later_patch_and_required_baseline_remain_distinct_in_export(self):
        row = synthetic_export()
        row["oracle"]["runtime_versions"]["python"] = "3.15.1"
        row["runtime_selection"].update(owner_version="3.15.1", selected_version="3.15.1",
                                        baseline_kind="reviewed_private3150_archive")
        row["runtime_selection"]["selected_identity"]["version"] = "3.15.1"
        result = smoke.export_report(row)
        self.assertEqual(result["runtime_selection"]["selected_version"], "3.15.1")
        self.assertEqual(result["runtime_selection"]["baseline_version"], "3.15.0")
        self.assertEqual(result["marked_setup_regressions"]["required_runtime_version"], "3.15.0")
        for section, key, value in (("runtime_selection", "baseline_version", "3.15.1"),
                                   ("runtime_selection", "selected_version", "3.15.0"),
                                   ("marked_setup_regressions", "fixture", "PRIVATE_USER_PATH"),
                                   ("marked_setup_regressions", "Action_roundtrips", True)):
            bad = copy.deepcopy(row); bad[section][key] = value
            with self.subTest(section=section, key=key), self.assertRaises(smoke.OracleError):
                smoke.export_report(bad)

    def test_export_preserves_fixed_native_runtime_and_marked_route_proofs(self):
        raw = synthetic_export()
        result = smoke.export_report(raw)
        self.assertEqual(result["runtime_selection"]["selected_identity"], raw["runtime_selection"]["selected_identity"])
        self.assertEqual(result["marked_setup_regressions"]["routes"], raw["marked_setup_regressions"]["routes"])
        self.assertEqual(result["marked_setup_regressions"]["routes"][0]["child_hosts"][0]["execution_policy"]["effective"], "Bypass")

    def test_fixed_identity_and_route_proof_reject_private_or_false_fields(self):
        mutations = [lambda r: r["runtime_selection"]["selected_identity"].update(executable="C:\\Users\\PRIVATE"),
                     lambda r: r["runtime_selection"]["baseline_identity"].update(source="PRIVATE_USERNAME"),
                     lambda r: r["runtime_selection"]["selected_identity"].update(bits=True),
                     lambda r: r["runtime_selection"]["selected_identity"].update(gil_disabled=0),
                     lambda r: r["runtime_selection"]["baseline_identity"].update(policy_sha256="c"*64),
                     lambda r: r["marked_setup_regressions"]["routes"][0].update(id="PRIVATE_PATH"),
                     lambda r: r["marked_setup_regressions"]["routes"][0].update(status="FAIL"),
                     lambda r: r["marked_setup_regressions"]["routes"][0]["exits"][0].update(exit_code=False),
                     lambda r: r["marked_setup_regressions"]["routes"][0]["exits"].__setitem__(0, "PRIVATE_RAW"),
                     lambda r: r["marked_setup_regressions"]["routes"][0]["child_hosts"].__setitem__(0, "PRIVATE_RAW"),
                     lambda r: r["marked_setup_regressions"]["routes"][0]["child_hosts"][0].update(executable="C:\\PRIVATE"),
                     lambda r: r["marked_setup_regressions"]["routes"][0]["child_hosts"][0]["execution_policy"].update(effective="RemoteSigned"),
                     lambda r: r["marked_setup_regressions"]["routes"][0]["child_hosts"][0]["execution_policy"]["scopes"][0].update(policy="AllSigned"),
                     lambda r: r["marked_setup_regressions"]["routes"][0]["pdf"].update(written_count=True),
                     lambda r: r["marked_setup_regressions"]["routes"][0]["pdf"].update(runtime_version="3.15.1"),
                     lambda r: r["marked_setup_regressions"]["routes"][-1].update(simulated_positive_launch_refused=1)]
        for mutation in mutations:
            raw = synthetic_export(); mutation(raw)
            with self.subTest(mutation=mutation), self.assertRaises(smoke.OracleError):
                smoke.export_report(raw)

    def test_nonpass_cases_remain_nonpass(self):
        row = synthetic_export()
        row["oracle"]["success"] = False
        row["oracle"]["cases"][2].update(status="BLOCKED", exit_code=None, error_code="missing_prerequisite")
        self.assertFalse(smoke.export_report(row)["oracle"]["success"])

    def test_native_architecture_and_exact_win10_build_display_gate(self):
        for delta in ({"DisplayVersion": "21H2"}, {"BuildNumber": "22631"}, {"processor_architectures": [12]},
                      {"processor_architectures": [True]}, {"os_bits": 32}, {"process_bits": True}, {"platform_scope": "BLOCKED"}, {"Version": "10.0.99999"}):
            row = synthetic_export()
            row["os"].update(delta)
            with self.subTest(delta=delta), self.assertRaises(smoke.OracleError):
                smoke.export_report(row)

    def test_finite_workstation_caption_variants_and_foreign_suffix_refusal(self):
        editions = ("Home", "Home N", "Home Single Language", "Home Country Specific", "Home China", "Pro", "Pro N",
                    "Pro for Workstations", "Pro for Workstations N", "Pro N for Workstations", "Pro Education",
                    "Pro Education N", "Enterprise", "Enterprise N", "Enterprise G", "Enterprise G N", "Education", "Education N", "IoT Enterprise")
        for edition in editions:
            row = synthetic_export()
            row["os"]["Caption"] = "Microsoft Windows 10 " + edition
            with self.subTest(edition=edition):
                self.assertEqual(smoke.export_report(row)["os"]["Caption"], row["os"]["Caption"])
        row = synthetic_export()
        row["os"].update(Caption="Microsoft Windows 11 SE", Version="10.0.26300", BuildNumber="26300", DisplayVersion="25H2", platform_scope="NATIVE_WINDOWS11_DRY_RUN")
        self.assertEqual(smoke.export_report(row)["os"]["Caption"], "Microsoft Windows 11 SE")
        for caption in ("Microsoft Windows Server 2025 Datacenter", "Microsoft Windows 10 Pro PRIVATE_HOST",
                        "Microsoft Windows 10 IoT Core", "Microsoft Windows 10 Enterprise LTSC 2021"):
            row = synthetic_export()
            row["os"]["Caption"] = caption
            with self.subTest(caption=caption), self.assertRaises(smoke.OracleError):
                smoke.export_report(row)

    def test_systemroot_environment_cannot_substitute_native_known_system_path(self):
        expected = smoke.windows_system_directory()
        with patch.dict(os.environ, {"SystemRoot": "C:\\PRIVATE_HOST\\decoy"}):
            self.assertEqual(smoke.windows_system_directory(), expected)
        self.assertTrue(Path(expected).is_dir())
        self.assertTrue((Path(expected) / "WindowsPowerShell/v1.0/powershell.exe").is_file())

    def test_no_bookmark_expected_exit_matches_shipped_outcome_contract(self):
        root = ROOT
        contract = json.loads((root / "engine/WinBookSplit.Outcomes.json").read_text(encoding="utf-8"))
        self.assertEqual(smoke.NO_BOOKMARK_EXIT, contract["codes"]["no_bookmarks"])

    def test_negative_outcomes_reject_wrong_stage_codes_and_publication(self):
        for identifier, expected in (("invalid-manual", "invalid_arguments"), ("cancel-plan", "processing_cancelled"),
                                     ("no-bookmarks", "no_bookmarks"), ("preview", "preview_complete")):
            final = {"code":expected, "written_count":0, "final_directory":None}
            if identifier in {"cancel-plan", "no-bookmarks"}:
                status = "no_plan" if identifier == "no-bookmarks" else "cancelled"
                final.update(status=status, engine_result={"status":status, "code":expected, "execution":None})
            smoke.validate_negative_outcome(identifier, final)
            for delta in ({"code":"invalid_start_pages"}, {"code":"cancelled"}, {"written_count":1}, {"final_directory":"C:\\foreign"}):
                with self.subTest(identifier=identifier, delta=delta), self.assertRaises(smoke.OracleError):
                    smoke.validate_negative_outcome(identifier, {**final, **delta})

    def test_negative_engine_result_requires_no_execution_and_correct_stage(self):
        for identifier, status, code in (("no-bookmarks", "no_plan", "no_bookmarks"),
                                         ("cancel-plan", "cancelled", "processing_cancelled")):
            final = {"code":code, "status":status, "written_count":0, "final_directory":None,
                     "engine_result":{"code":code, "status":status, "execution":None}}
            smoke.validate_negative_outcome(identifier, final)
            for frame in (None, {"code":code, "status":status},
                          {"code":code, "status":status, "execution":{"written_count":0}},
                          {"code":"split_complete", "status":status, "execution":None},
                          {"code":code, "status":"success", "execution":None}):
                with self.subTest(identifier=identifier, frame=frame), self.assertRaisesRegex(smoke.OracleError, "negative_engine_execution_differ"):
                    smoke.validate_negative_outcome(identifier, {**final, "engine_result":frame})
            with self.assertRaisesRegex(smoke.OracleError, "negative_engine_execution_differ"):
                smoke.validate_negative_outcome(identifier, {**final, "status":"success"})

    def test_windows11_dry_run_is_never_win10_scope(self):
        row = synthetic_export()
        row["os"].update(Caption="Microsoft Windows 11 Pro", Version="10.0.26300", BuildNumber="26300", DisplayVersion="25H2", platform_scope="NATIVE_WINDOWS11_DRY_RUN")
        self.assertEqual(smoke.export_report(row)["os"]["platform_scope"], "NATIVE_WINDOWS11_DRY_RUN")

    def test_exact_json_write_never_overwrites(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder).resolve() / "report.json"
            smoke.write_json(path, {"a": 1})
            original = path.read_bytes()
            with self.assertRaises(smoke.OracleError):
                smoke.write_json(path, {"a": 2})
            self.assertEqual(path.read_bytes(), original)

    def test_immutable_byte_and_identity_observation(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder).resolve() / "input.bin"
            path.write_bytes(b"authored")
            observed = smoke.observation(path)
            smoke.unchanged(observed)
            path.write_bytes(b"altered")
            with self.assertRaises(smoke.OracleError):
                smoke.unchanged(observed)


class PublicationOracleControls(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.base = Path(self.temporary.name).resolve()
        self.run_id = "a" * 32
        self.final = self.base / ("Original_" + self.run_id)
        self.final.mkdir()
        smoke.RUNTIME_POLICY = {"pypdf": {"version": "6.21.1"}}
        source = self.base / "source.pdf"
        writer = PdfWriter()
        for width in (101, 102, 103):
            writer.add_blank_page(width=width, height=117)
        with source.open("xb") as stream:
            writer.write(stream)
        self.source = smoke.pdf_snapshot(source)
        self.ranges = [[0, 1], [1, 3]]
        entries = []
        for index, (start, end) in enumerate(self.ranges):
            target = self.final / (str(index) + ".pdf")
            title = "Chapter " + str(index)
            writer = PdfWriter()
            reader = PdfReader(source)
            for page in reader.pages[start:end]:
                writer.add_page(page)
            writer.add_metadata({"/Title": title, "/Creator": "WinBookSplit", "/Producer": "WinBookSplit 1.0.1 (pypdf 6.21.1)"})
            writer.add_outline_item(title, 0)
            with target.open("xb") as stream:
                writer.write(stream)
            observed = smoke.observation(target)
            entries.append({"filename": target.name, "title": title, "start": start, "end": end, "page_count": end-start,
                            "sha256": observed["sha256"], "size_bytes": observed["size_bytes"]})
        self.execution = {"run_id": self.run_id, "final_directory": str(self.final), "mode": "manual", "total_pages": 3,
                          "source_identity": {"path": str(source)}, "coverage": {"complete": True}, "written_count": 2, "outputs": entries}
        self.manifest = {"schema_version": 1, "status": "complete", **copy.deepcopy(self.execution)}
        self.execution["manifest"] = self.manifest
        smoke.write_json(self.final / ".WinBookSplit-owner.json", {"schema_version": 1, "kind": "run", "run_id": self.run_id})
        smoke.write_json(self.final / "WinBookSplit_Manifest.json", self.manifest)

    def tearDown(self):
        self.temporary.cleanup()

    def check(self):
        return smoke.validate_publication(self.base, self.execution, self.source, self.ranges, False, "1.0.1")

    def test_reopened_full_order_geometry_hash_manifest_and_bookmarks(self):
        self.assertEqual(self.check(), self.final)

    def test_repeat_cannot_reuse_previous_ordinary_directory(self):
        identities = set()
        smoke.publication_once(self.final, identities)
        with self.assertRaises(smoke.OracleError):
            smoke.publication_once(self.final, identities)
        second = self.base / ("New_" + "b" * 32)
        second.mkdir()
        smoke.publication_once(second, identities)
        self.assertEqual(len(identities), 2)

    def test_missing_reordered_or_repeated_physical_page_rejected(self):
        for bad in ([[0, 1], [2, 3]], [[1, 3], [0, 1]], [[0, 1], [0, 3]]):
            with self.subTest(ranges=bad), self.assertRaises(smoke.OracleError):
                smoke.validate_publication(self.base, self.execution, self.source, bad, False, "1.0.1")

    def test_foreign_member_owner_and_manifest_rejected(self):
        foreign = self.final / "foreign.txt"
        foreign.write_bytes(b"preserve")
        with self.assertRaises(smoke.OracleError):
            self.check()
        foreign.unlink()  # Exact authored unit fixture only.
        self.execution["manifest"] = {**self.manifest, "written_count": 1}
        with self.assertRaises(smoke.OracleError):
            self.check()

    def test_changed_pdf_bytes_and_geometry_rejected(self):
        target = self.final / self.execution["outputs"][0]["filename"]
        with target.open("ab") as stream:
            stream.write(b"changed")
        with self.assertRaises(smoke.OracleError):
            self.check()

    def test_claimed_retention_without_exact_full_pdf_rejected(self):
        self.execution["retained_intermediate"] = {"filename": "WinBookSplit_Converted.pdf"}
        with self.assertRaises(smoke.OracleError):
            smoke.validate_publication(self.base, self.execution, self.source, self.ranges, True, "1.0.1")


if __name__ == "__main__":
    unittest.main(verbosity=2)
