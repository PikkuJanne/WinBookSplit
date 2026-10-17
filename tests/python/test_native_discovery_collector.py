"""Actual native collector transport controls; no application Python bootstrap."""
from hashlib import sha256
from contextlib import redirect_stderr
import ctypes
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
PROBE = ROOT / "tests/compatibility/Probe-NativeDiscoveryCollector.ps1"
COMPONENTS = (PROBE, ROOT / "tools/compatibility/Invoke-NativeDiscoveryCollector.ps1",
              ROOT / "engine/WinBookSplit.Process.ps1")
CONTROL_IDS = {
    "dual-unicode-streams-and-plain-json-positive-control",
    "nonzero-listing-retains-both-streams", "missing-launch-is-explicit-failure",
    "timeout-stops-child-after-positive-reach-marker",
    "descendant-inherited-stream-positive-control",
    "hung-descendant-and-inherited-streams-are-bounded",
    "bounded-output-cannot-become-success", "existing-receipt-is-preserved",
    "application-environment-is-unchanged",
}


@unittest.skipUnless(os.name == "nt", "Actual Windows PowerShell collector controls")
class NativeDiscoveryCollectorControls(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        buffer = ctypes.create_unicode_buffer(32768)
        length = ctypes.windll.kernel32.GetSystemDirectoryW(buffer, len(buffer))
        if not 0 < length < len(buffer):
            raise RuntimeError("Cannot locate the genuine Windows system directory")
        cls.inbox = Path(buffer.value) / "WindowsPowerShell/v1.0/powershell.exe"
        if not cls.inbox.is_file():
            raise unittest.SkipTest("Actual inbox Windows PowerShell 5.1 is unavailable")
        located = shutil.which("pwsh")
        cls.ps7 = Path(located).resolve() if located else None

    def run_control_matrix(self, controller, child):
        if controller is None or child is None:
            self.skipTest("Actual PowerShell 7 executable is unavailable")
        before = {str(path.relative_to(ROOT)): sha256(path.read_bytes()).hexdigest()
                  for path in COMPONENTS}
        # Keep failed or interrupted evidence: do not clean a directory after an
        # unproved native stop. A passing probe proves stop/drain in every case.
        temporary = tempfile.TemporaryDirectory(prefix="wbs-native-collector-controls-")
        work = Path(temporary.name).resolve()
        report = work / "report.json"
        command = [str(controller), "-NoLogo", "-NoProfile", "-NonInteractive",
                   "-InputFormat", "Text", "-OutputFormat", "Text",
                   "-ExecutionPolicy", "RemoteSigned", "-File", str(PROBE),
                   "-RepositoryRoot", str(ROOT), "-WorkRoot", str(work / "owned"),
                   "-ReportPath", str(report), "-ChildPowerShell", str(child)]
        safe = False
        try:
            result = subprocess.run(command, cwd=ROOT, stdin=subprocess.DEVNULL,
                                    stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                    timeout=180, check=False)
            (work / "outer.stdout.bin").write_bytes(result.stdout)
            (work / "outer.stderr.bin").write_bytes(result.stderr)
            after = {str(path.relative_to(ROOT)): sha256(path.read_bytes()).hexdigest()
                     for path in COMPONENTS}
            receipt = {"command": command, "exit_code": result.returncode,
                       "stdout": result.stdout.decode("utf-8", errors="strict"),
                       "stderr": result.stderr.decode("utf-8", errors="strict"),
                       "component_hashes_before": before, "component_hashes_after": after}
            (work / "driver.json").write_text(json.dumps(receipt, ensure_ascii=False), encoding="utf-8")
            self.assertEqual(before, after, f"Collector component bytes changed: {work}")
            self.assertEqual(result.returncode, 0, f"Collector probe failed; retained evidence: {work}")
            raw = report.read_bytes()
            self.assertTrue(raw.startswith(b"{"), "Collector report must be BOMless plain JSON")
            body = json.loads(raw.decode("utf-8", errors="strict"))
            self.assertEqual(body["protocol"], "winbooksplit.native-discovery-collector-tests")
            self.assertEqual(body["schema_version"], 1)
            self.assertEqual(body["scope"], "CONTROLLED_NATIVE_COLLECTOR_ONLY")
            self.assertFalse(body["app_python_required"])
            self.assertEqual(body["status"], "PASS")
            self.assertEqual({case["id"] for case in body["cases"]}, CONTROL_IDS)
            self.assertEqual(len(body["cases"]), len(CONTROL_IDS))
            for case in body["cases"]:
                self.assertEqual(case["status"], "PASS", f"{case}; retained evidence: {work}")
            safe = True
        except Exception as error:
            timed_out = isinstance(error, subprocess.TimeoutExpired)
            if timed_out:
                # communicate() supplies partial bytes on a finite outer timeout.
                # Retain them before re-raising; an absent stop proof forbids cleanup.
                for name, data in (("stdout", error.output), ("stderr", error.stderr)):
                    if isinstance(data, str):
                        data = data.encode("utf-8")
                    (work / f"outer.{name}.bin").write_bytes(data or b"")
            failure = {"protocol": "winbooksplit.native-collector-driver-failure", "schema_version": 1,
                       "kind": "outer_timeout" if timed_out else "driver_failure",
                       "command": command, "timeout_seconds": error.timeout if timed_out else None,
                       "error_type": type(error).__name__, "message": str(error),
                       "component_hashes_before": before, "cleanup_safe": False}
            (work / "driver-failure.json").write_text(json.dumps(failure, ensure_ascii=False), encoding="utf-8")
            raise
        finally:
            if safe:
                temporary.cleanup()
            else:
                temporary._finalizer.detach()
                # The outer maintained runner recognizes Unproved and retains
                # its enclosing TEMP tree as well as this detached directory.
                print(f"Unproved cleanup safety; retained native control evidence: {work}", file=sys.stderr, flush=True)

    def test_inbox_controller_inbox_child(self):
        self.run_control_matrix(self.inbox, self.inbox)

    def test_inbox_controller_ps7_child(self):
        self.run_control_matrix(self.inbox, self.ps7)

    def test_ps7_controller_inbox_child(self):
        self.run_control_matrix(self.ps7, self.inbox)

    def test_ps7_controller_ps7_child(self):
        self.run_control_matrix(self.ps7, self.ps7)


class NativeCollectorDriverFailureRetentionControls(unittest.TestCase):
    def run_authored_failure(self, *, timeout):
        real_temporary = tempfile.TemporaryDirectory
        # No native process starts in these controls. Only authored failure bytes
        # enter the test-owned outer directory, which this test may safely clean.
        with real_temporary(prefix="wbs-collector-retention-controls-") as outer:
            retained = []

            def owned_temporary(**options):
                directory = real_temporary(dir=outer, **options)
                retained.append(directory)
                return directory

            stdout, stderr = b"AUTHORED partial stdout \xe6\x97\xa5", b"AUTHORED partial stderr"
            error = subprocess.TimeoutExpired(["AUTHORED controller"], 180, output=stdout, stderr=stderr)
            completed = subprocess.CompletedProcess(["AUTHORED controller"], 23, stdout=stdout, stderr=stderr)
            captured = io.StringIO()
            driver = NativeDiscoveryCollectorControls("test_inbox_controller_inbox_child")
            with patch.object(tempfile, "TemporaryDirectory", side_effect=owned_temporary), \
                    patch.object(subprocess, "run", side_effect=error if timeout else None,
                                 return_value=completed) as launched, redirect_stderr(captured):
                with self.assertRaises(subprocess.TimeoutExpired if timeout else AssertionError):
                    driver.run_control_matrix(Path("AUTHORED-controller.exe"), Path("AUTHORED-child.exe"))
            self.assertEqual(launched.call_count, 1)
            self.assertEqual(launched.call_args.kwargs["timeout"], 180)
            self.assertEqual(len(retained), 1)
            work = Path(retained[0].name)
            self.assertTrue(work.is_dir(), "Failure evidence was removed")
            self.assertFalse(retained[0]._finalizer.alive, "Failure evidence cleanup was still registered")
            self.assertIn(f"Unproved cleanup safety; retained native control evidence: {work}", captured.getvalue())
            self.assertEqual((work / "outer.stdout.bin").read_bytes(), stdout)
            self.assertEqual((work / "outer.stderr.bin").read_bytes(), stderr)
            receipt = json.loads((work / "driver-failure.json").read_bytes())
            self.assertEqual(receipt["kind"], "outer_timeout" if timeout else "driver_failure")
            self.assertFalse(receipt["cleanup_safe"])
            self.assertEqual(receipt["timeout_seconds"], 180 if timeout else None)
            if not timeout:
                self.assertEqual(json.loads((work / "driver.json").read_bytes())["exit_code"], 23)

    def test_failed_driver_retains_raw_bytes_and_outer_runner_marker(self):
        self.run_authored_failure(timeout=False)

    def test_timeout_retains_partial_bytes_and_outer_runner_marker(self):
        self.run_authored_failure(timeout=True)


if __name__ == "__main__":
    unittest.main(verbosity=2)
