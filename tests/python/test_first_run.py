"""Unchanged native BAT cold start; never provision or select a test runtime.

The desired missing-package offer requires a naturally discoverable supported
Python without locked pypdf. A missing prerequisite is an explicit skip, after
recording what the real BAT actually did. Component tests remain separate.
"""

from hashlib import sha256
import importlib.util
import json
import os
from pathlib import Path
import shutil
import tempfile
import unittest

from pypdf import PdfWriter

ROOT = Path(__file__).resolve().parents[2]
SPEC = importlib.util.spec_from_file_location(
    "wbs_first_run_launcher", ROOT / "tests/launcher/characterize_launcher.py")
launcher = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(launcher)


@unittest.skipUnless(os.name == "nt", "Ordinary BAT cold start requires native Windows")
class OrdinaryFirstRunTests(unittest.TestCase):
    def test_ordinary_bat_offers_local_setup_once_before_decline(self):
        work = Path(tempfile.mkdtemp(prefix="WBS-first-run-"))
        app, books, cwd = (work / name for name in ("a", "b", "c"))
        copied = launcher.process_tests.copy_application(app)
        # copy_application predates setup BAT; include the complete setup entry.
        shutil.copyfile(ROOT / "Setup-WinBookSplit.bat", app / "Setup-WinBookSplit.bat")
        copied["Setup-WinBookSplit.bat"] = sha256((app / "Setup-WinBookSplit.bat").read_bytes()).hexdigest()
        for directory in (books, cwd):
            directory.mkdir()
        source = books / "authored-cold-start.pdf"
        writer = PdfWriter()
        writer.add_blank_page(width=300, height=400)
        with source.open("wb") as stream:
            writer.write(stream)
        before = sha256(source.read_bytes()).hexdigest()
        self.assertFalse((app / ".venv").exists())
        inbox = Path(os.environ["SystemRoot"]) / "System32/WindowsPowerShell/v1.0/powershell.exe"
        probe = cwd / "preflight.ps1"
        probe.write_text(
            "[Console]::OutputEncoding = New-Object Text.UTF8Encoding($false)\n"
            "$ErrorActionPreference='Stop'\n"
            ". '..\\a\\engine\\WinBookSplit.Runtime.ps1'\n"
            "function Get-ProbeFailure($failure) {\n"
            " return @{code=$failure.Data['Code'];message=$failure.Message;attempts=@($failure.Data['Attempts'])}\n"
            "}\n"
            "$r=[ordered]@{base_found=$false;processing_ready=$false;package_missing=$false;base_failure=$null;processing_failure=$null}\n"
            "try {\n"
            " $base=Resolve-WinBookSplitRuntime -ApplicationRoot (Resolve-Path '..\\a').Path -InterpreterOnly\n"
            " $r.base_found=$true; $r.base_version=$base.Version; $r.base_source=$base.Source\n"
            " try { $null=Resolve-WinBookSplitRuntime -ApplicationRoot (Resolve-Path '..\\a').Path; $r.processing_ready=$true }\n"
            " catch { $r.processing_failure=Get-ProbeFailure $_.Exception;"
            " $r.package_missing=@($r.processing_failure.attempts | Where-Object { $_.IdentityVerified -and $_.Reason -ceq 'missing_pypdf' }).Count -gt 0 }\n"
            "} catch { $r.base_failure=Get-ProbeFailure $_.Exception }\n"
            "Write-Output ('[COLD-START-PREFLIGHT] '+($r | ConvertTo-Json -Depth 15 -Compress))\n",
            encoding="utf-8-sig")
        evidence = {"scope": "NATIVE_UNCHANGED_BAT_NATURAL_DISCOVERY_NO_PROVISIONING",
                    "work": str(work), "application_sha256": copied,
                    "source_sha256_before": before, "environment_inherited_unchanged": True,
                    "native_missing_package_status": "NOT_RUN"}
        try:
            inherited = os.environ.copy()
            preflight = launcher.run_prompted(
                [str(inbox), "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "preflight.ps1"],
                cwd, inherited, "", timeout=90)
            evidence["preflight_process"] = preflight
            records = [json.loads(line.removeprefix("[COLD-START-PREFLIGHT] "))
                       for line in preflight["stdout"].splitlines()
                       if line.startswith("[COLD-START-PREFLIGHT] ")]
            self.assertEqual(preflight["exit_code"], 0, str(work))
            self.assertEqual(len(records), 1, str(work))
            observed = records[0]
            evidence["preflight"] = observed
            # Literal relative paths keep the wrapper independent of account-name
            # encoding. The BAT, PS and engine bytes and parent environment stay
            # unchanged; CWD is deliberately unrelated to the application.
            wrapper = cwd / "invoke.cmd"
            wrapper.write_bytes(b'@echo off\r\nsetlocal DisableDelayedExpansion\r\n"..\\a\\WinBookSplit.bat" "..\\b\\authored-cold-start.pdf"\r\n')
            command = [str(Path(os.environ["SystemRoot"]) / "System32/cmd.exe"),
                       "/d", "/v:off", "/c", "invoke.cmd"]
            actual = launcher.run_prompted(command, cwd, inherited, "N\n\n", timeout=90)
            evidence["bat_process"] = actual
            evidence["source_sha256_after"] = sha256(source.read_bytes()).hexdigest()
            evidence["app_venv_exists_after"] = (app / ".venv").exists()
            outcomes = [json.loads(line.removeprefix("[OUTCOME] "))
                        for line in actual["stdout"].splitlines() if line.startswith("[OUTCOME] ")]
            evidence["outcomes"] = outcomes
            self.assertEqual(evidence["source_sha256_after"], before)
            self.assertFalse(evidence["app_venv_exists_after"])
            self.assertEqual(sorted(path.name for path in books.iterdir()), [source.name])
            self.assertEqual(len(outcomes), 1, str(work))
            self.assertEqual(outcomes[0]["written_count"], 0, str(work))
            self.assertEqual(outcomes[0]["exit_code"], actual["exit_code"], str(work))
            self.assertEqual(copied, {name: sha256((app / name).read_bytes()).hexdigest() for name in copied})
            for flag in ("base_found", "processing_ready", "package_missing"):
                self.assertIs(type(observed[flag]), bool, str(work))
            for failure in (observed["base_failure"], observed["processing_failure"]):
                if failure is None:
                    continue
                self.assertIsInstance(failure, dict, str(work))
                self.assertIn(failure["code"], {"runtime_not_found", "runtime_invalid", "setup_required"}, failure)
                self.assertIsInstance(failure["message"], str, str(work))
                self.assertTrue(failure["message"].strip(), str(work))
                self.assertIsInstance(failure["attempts"], list, str(work))
                self.assertTrue(failure["attempts"], str(work))
                for row in failure["attempts"]:
                    self.assertIsInstance(row, dict, str(work))
                    for field in ("Reason", "Stage", "Source", "Message"):
                        self.assertIsInstance(row[field], str, str(work))
                        self.assertTrue(row[field].strip(), str(work))
                    self.assertIs(type(row["Accepted"]), bool, str(work))
            if not observed["base_found"]:
                failure = observed["base_failure"]
                self.assertIsInstance(failure, dict, str(work))
                self.assertEqual(failure["code"], "runtime_not_found", failure)
                reasons = {row["Reason"] for row in failure["attempts"]}
                # A known unavailable prerequisite may skip. Malformed policy,
                # probe failures, unexpected exceptions and package errors fail.
                unavailable = {"no_candidates", "missing_python", "manager_not_registered", "manager_no_runtimes",
                               "python_security_floor", "python_compatibility_floor", "unsupported_python",
                               "unsupported_architecture", "unsupported_flavor", "unstable_python"}
                incidental = {"rejected_entry_point", "manager_entry_rejected", "manager_registered",
                              "installed_runtimes_listed"}
                self.assertTrue(reasons & unavailable, failure)
                self.assertFalse(reasons - unavailable - incidental, failure)
                evidence["prerequisite_reason"] = "No naturally discoverable supported base Python; observed=" + ",".join(sorted(reasons))
                self.skipTest(evidence["prerequisite_reason"] + "; actual BAT captured at " + str(work))
            self.assertIsNone(observed["base_failure"], str(work))
            if observed["processing_ready"]:
                self.assertIsNone(observed["processing_failure"], str(work))
                self.assertFalse(observed["package_missing"], str(work))
                evidence["prerequisite_reason"] = "Naturally discovered processing runtime is already healthy"
                self.skipTest(evidence["prerequisite_reason"] + "; actual BAT captured at " + str(work))
            self.assertIsInstance(observed["processing_failure"], dict, str(work))
            self.assertTrue(observed["package_missing"], observed["processing_failure"])
            evidence["native_missing_package_status"] = "FAIL"
            prompt = "First-time setup is required. Download and install the PDF component into this copy of WinBookSplit now? [Y/N]"
            self.assertEqual(actual["stdout"].count(prompt), 1, str(work))
            evidence["native_missing_package_status"] = "PASS_DECLINE_ONLY"
        except unittest.SkipTest:
            raise
        except Exception as error:
            evidence["driver_error"] = str(error)
            raise
        finally:
            (work / "evidence.json").write_text(json.dumps(evidence, ensure_ascii=False, indent=2), encoding="utf-8")
            print("FIRST_RUN_EVIDENCE=" + str(work / "evidence.json"), flush=True)


if __name__ == "__main__":
    unittest.main()
