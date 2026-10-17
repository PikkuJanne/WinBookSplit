"""Store no-acquisition tooling controls; supplied developer Python stays separate."""
import ctypes
from contextlib import redirect_stderr
from hashlib import sha256
import io
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[2]
PROBE = ROOT / "tests/compatibility/Probe-StoreExisting.ps1"
COMPONENTS = (PROBE, ROOT / "tools/compatibility/Invoke-StoreExistingAcceptance.ps1",
              ROOT / "tools/compatibility/Invoke-NativeAcceptance.ps1",
              ROOT / "tools/compatibility/Invoke-NativeDiscoveryCollector.ps1")
IDS = {"store-inputs-refuse-explicit-baseline-and-provisioning", "store-inputs-refuse-runtime-environment",
       "redaction-parses-nested-json-case-and-key-paths", "actual-redacted-stream-lengths-pass-rejecting-validator",
       "no-store-registration-before-any-bootstrap-or-fallback", "inventory-positive-control-real-child-is-actually-used",
       "inventory-nonzero-listing-refuses-runtime-probe", "existing-inventory-positive-reaches-bootstrap-no-acquisition",
       "rejected-runtime-probe-survives-returned-failure", "finite-return-zip-positive-and-foreign-member-refusal"}
IDS.update({"policy-action-refuses-runtime-expression-and-extra-command",
            "policy-actual-authored-motw-positive-and-controller-domain"})


@unittest.skipUnless(os.name == "nt", "Actual Windows Store-scenario tooling controls")
class StoreScenarioControls(unittest.TestCase):
    def run_probe(self, controller):
        if not controller:
            self.skipTest("PowerShell executable unavailable")
        before = {str(p.relative_to(ROOT)): sha256(p.read_bytes()).hexdigest() for p in COMPONENTS}
        temporary = tempfile.TemporaryDirectory(prefix="wbs-store-scenario-controls-")
        work = Path(temporary.name).resolve()
        command = [str(controller), "-NoProfile", "-NonInteractive", "-InputFormat", "Text", "-OutputFormat", "Text",
                   "-ExecutionPolicy", "RemoteSigned", "-File", str(PROBE), "-RepositoryRoot", str(ROOT),
                   "-TestPython", sys.executable, "-WorkRoot", str(work / "owned"), "-ReportPath", str(work / "report.json")]
        safe = False
        try:
            result = subprocess.run(command, cwd=ROOT, stdin=subprocess.DEVNULL, stdout=subprocess.PIPE,
                                    stderr=subprocess.PIPE, timeout=120, check=False)
            (work / "stdout.bin").write_bytes(result.stdout)
            (work / "stderr.bin").write_bytes(result.stderr)
            driver = {"command": command, "exit_code": result.returncode,
                      "stdout": result.stdout.decode("utf-8", errors="strict"),
                      "stderr": result.stderr.decode("utf-8", errors="strict"),
                      "source_before": before}
            (work / "driver.json").write_text(json.dumps(driver, ensure_ascii=False), encoding="utf-8")
            self.assertEqual(result.returncode, 0, f"Store tooling probe failed; retained {work}")
            body = json.loads((work / "report.json").read_text(encoding="utf-8"))
            self.assertTrue(body["success"], f"{body}; retained {work}")
            self.assertEqual(body["kind"], "T05_STORE_SCENARIO_CONTROLLED_NATIVE_TOOLING")
            self.assertEqual({r["id"] for r in body["cases"]}, IDS)
            self.assertEqual(len(body["cases"]), len(IDS))
            self.assertTrue(all(r["status"] == "PASS" for r in body["cases"]))
            self.assertEqual(before, {str(p.relative_to(ROOT)): sha256(p.read_bytes()).hexdigest() for p in COMPONENTS})
            safe = True
        except Exception as error:
            timed_out = isinstance(error, subprocess.TimeoutExpired)
            if timed_out:
                # A finite outer timeout may still carry partial native bytes.
                # Preserve those bytes before re-raising without a stop proof.
                for name, data in (("stdout", error.output), ("stderr", error.stderr)):
                    if isinstance(data, str):
                        data = data.encode("utf-8")
                    (work / f"{name}.bin").write_bytes(data or b"")
            failure = {"protocol": "winbooksplit.native-store-driver-failure", "schema_version": 1,
                       "kind": "outer_timeout" if timed_out else "driver_failure",
                       "command": command, "timeout_seconds": error.timeout if timed_out else None,
                       "error_type": type(error).__name__, "message": str(error),
                       "source_before": before, "cleanup_safe": False}
            (work / "driver-failure.json").write_text(json.dumps(failure, ensure_ascii=False), encoding="utf-8")
            raise
        finally:
            if safe:
                temporary.cleanup()
            else:
                temporary._finalizer.detach()
                print(f"Unproved cleanup safety; retained native control evidence: {work}", file=sys.stderr, flush=True)

    def test_inbox_no_fallback_and_positive_controls(self):
        buffer = ctypes.create_unicode_buffer(32768)
        size = ctypes.windll.kernel32.GetSystemDirectoryW(buffer, len(buffer))
        self.assertTrue(0 < size < len(buffer))
        self.run_probe(Path(buffer.value) / "WindowsPowerShell/v1.0/powershell.exe")

    def test_ps7_no_fallback_and_positive_controls(self):
        self.run_probe(shutil.which("pwsh"))


class StoreDriverFailureRetentionControls(unittest.TestCase):
    def run_authored_failure(self, *, timeout):
        real_temporary = tempfile.TemporaryDirectory
        # No native process starts; cleanup owns only these authored bytes.
        with real_temporary(prefix="wbs-store-retention-controls-") as outer:
            retained = []

            def owned_temporary(**options):
                directory = real_temporary(dir=outer, **options)
                retained.append(directory)
                return directory

            stdout, stderr = b"AUTHORED partial stdout \xe6\x97\xa5", b"AUTHORED partial stderr"
            error = subprocess.TimeoutExpired(["AUTHORED controller"], 120, output=stdout, stderr=stderr)
            completed = subprocess.CompletedProcess(["AUTHORED controller"], 23, stdout=stdout, stderr=stderr)
            captured = io.StringIO()
            driver = StoreScenarioControls("test_inbox_no_fallback_and_positive_controls")
            with patch.object(tempfile, "TemporaryDirectory", side_effect=owned_temporary), \
                    patch.object(subprocess, "run", side_effect=error if timeout else None,
                                 return_value=completed) as launched, redirect_stderr(captured):
                with self.assertRaises(subprocess.TimeoutExpired if timeout else AssertionError):
                    driver.run_probe(Path("AUTHORED-controller.exe"))
            self.assertEqual(launched.call_count, 1)
            self.assertEqual(launched.call_args.kwargs["timeout"], 120)
            self.assertEqual(len(retained), 1)
            work = Path(retained[0].name)
            self.assertTrue(work.is_dir(), "Store failure evidence was removed")
            self.assertFalse(retained[0]._finalizer.alive, "Unproved Store failure cleanup remained registered")
            self.assertIn(f"Unproved cleanup safety; retained native control evidence: {work}", captured.getvalue())
            self.assertEqual((work / "stdout.bin").read_bytes(), stdout)
            self.assertEqual((work / "stderr.bin").read_bytes(), stderr)
            receipt = json.loads((work / "driver-failure.json").read_bytes())
            self.assertEqual(receipt["kind"], "outer_timeout" if timeout else "driver_failure")
            self.assertFalse(receipt["cleanup_safe"])
            self.assertEqual(receipt["timeout_seconds"], 120 if timeout else None)
            self.assertEqual(receipt["error_type"], "TimeoutExpired" if timeout else "AssertionError")
            self.assertTrue(receipt["source_before"])
            if not timeout:
                self.assertEqual(json.loads((work / "driver.json").read_bytes())["exit_code"], 23)

    def test_nonzero_driver_retains_raw_bytes_and_outer_marker(self):
        self.run_authored_failure(timeout=False)

    def test_timeout_driver_retains_partial_bytes_and_outer_marker(self):
        self.run_authored_failure(timeout=True)


if __name__ == "__main__":
    unittest.main(verbosity=2)
