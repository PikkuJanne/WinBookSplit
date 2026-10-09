"""Mechanical-extraction and owned-output safety checks, without private inputs."""

import importlib.util
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import Mock, patch
import uuid

ROOT = Path(__file__).resolve().parents[2]
SPEC = importlib.util.spec_from_file_location("wbs_extraction_test", ROOT / "tests/extraction/characterize_extraction.py")
extraction = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(extraction)


class ExtractionTests(unittest.TestCase):
    @unittest.skipUnless(sys.platform == "win32", "Tracked taskkill tree requires Windows")
    def test_entrypoint_timeout_terminates_actual_parent_and_descendant(self):
        import ctypes
        from ctypes import wintypes

        with tempfile.TemporaryDirectory(prefix="wbs-owned-tree-test-") as directory:
            work = Path(directory).resolve()
            ready = work / "owned-pids.json"
            descendant = "import time; time.sleep(120)"
            parent = (
                "import json,os,pathlib,subprocess,sys,time; "
                "child=subprocess.Popen([sys.executable,'-I','-B','-c'," + repr(descendant) + "]); "
                "pathlib.Path(os.environ['WBS_OWNED_PIDS']).write_text("
                "json.dumps([os.getpid(),child.pid]),encoding='utf-8'); time.sleep(120)"
            )
            environment = extraction.clean_environment(work)
            environment["WBS_OWNED_PIDS"] = str(ready)
            with self.assertRaises(extraction.EntryPointFailure) as raised:
                extraction.run_entrypoint([sys.executable, "-I", "-B", "-c", parent], work,
                                          environment=environment, stdin="", timeout=2)
            self.assertTrue(raised.exception.cleanup_safe)
            pids = json.loads(ready.read_text(encoding="utf-8"))
            # A Windows venv redirector can be the tracked parent of the
            # interpreter executing this code; verify that PID as well.
            pids.append(raised.exception.pid)
            kernel = ctypes.WinDLL("kernel32", use_last_error=True)
            kernel.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
            kernel.OpenProcess.restype = wintypes.HANDLE
            kernel.WaitForSingleObject.argtypes = [wintypes.HANDLE, wintypes.DWORD]
            kernel.WaitForSingleObject.restype = wintypes.DWORD
            kernel.CloseHandle.argtypes = [wintypes.HANDLE]
            kernel.CloseHandle.restype = wintypes.BOOL
            for pid in pids:
                handle = kernel.OpenProcess(0x00100000, False, pid)  # SYNCHRONIZE; no terminate right.
                if handle:
                    try:
                        self.assertEqual(kernel.WaitForSingleObject(handle, 0), 0,
                                         "Owned process is still running after timeout cleanup permission")
                    finally:
                        kernel.CloseHandle(handle)
                else:
                    self.assertEqual(ctypes.get_last_error(), 87, "Cannot verify owned process state")

    def test_failed_tree_termination_cannot_mask_unsafe_cleanup_by_closing_pipes(self):
        child = Mock()
        child.pid = 1234
        child.poll.return_value = None
        child.communicate.side_effect = subprocess.TimeoutExpired("owned fixture", 1)
        for name in ("stdin", "stdout", "stderr"):
            stream = Mock()
            stream.close.side_effect = BrokenPipeError("Simulated uncertain pipe close")
            setattr(child, name, stream)
        with patch.object(extraction.subprocess, "Popen", return_value=child), \
                patch.object(extraction.subprocess, "run", return_value=Mock(returncode=1)) as terminate, \
                patch.dict(os.environ, {"SystemRoot": "C:/Windows"}):
            with self.assertRaises(extraction.EntryPointFailure) as raised:
                extraction.run_entrypoint(["owned-test-child"], Path.cwd(), environment={}, stdin="", timeout=1)
        self.assertFalse(raised.exception.cleanup_safe)
        self.assertEqual(terminate.call_args.args[0][-4:], ["/PID", "1234", "/T", "/F"])
        child.kill.assert_not_called()
        for stream in (child.stdin, child.stdout, child.stderr):
            stream.close.assert_not_called()

    @unittest.skipUnless(sys.platform == "win32", "Actual Windows host path preflight")
    def test_uncertain_tree_preserves_marked_owned_output(self):
        with tempfile.TemporaryDirectory(prefix="wbs-unsafe-tree-test-") as directory:
            work = Path(directory).resolve()
            fake_documents = work / "synthetic-documents"
            fake_documents.mkdir()
            fixture = work / "fixture.pdf"
            fixture.write_bytes(b"Synthetic test-only bytes; process call is mocked")
            system_ps51 = Path(os.environ["SystemRoot"]) / "System32/WindowsPowerShell/v1.0/powershell.exe"
            fake_observation = {"exit_code": 0, "stdout": str(fake_documents) + "\n", "stderr": ""}
            cases = [{"oracle_id": name, "original": {"outputs": []}} for name in ("MAN-03", "BM-03")]
            with patch.object(extraction, "run", return_value=fake_observation), \
                    patch.object(extraction, "run_entrypoint", side_effect=extraction.EntryPointFailure(
                        "Simulated uncertain tracked tree", cleanup_safe=False, pid=1234)):
                with self.assertRaisesRegex(RuntimeError, "preserve owned marked output"):
                    extraction.launchers(work, {"simple10": fixture}, None,
                                         [system_ps51, Path(sys.executable)], cases)
            # This is a temp-only synthetic Documents root, not a private listing.
            reserved = list(fake_documents.iterdir())
            self.assertEqual(len(reserved), 1)
            self.assertTrue((reserved[0] / ".wbs-m1-owner.json").is_file())
            self.assertEqual((reserved[0] / ".wbs-m1-neighbor.txt").read_bytes(),
                             b"Original synthetic Documents neighbor\n")

    def test_historical_ac011_extracted_commit_preserved_immutable_original_ast(self):
        # AC-011 records M1-T01's mechanical extraction. Later scoped behavior
        # fixes must not rewrite that evidence or be required to retain defects.
        sources = extraction.original_sources()
        prior = extraction.function_asts(extraction.embedded_body(sources["WinBookSplit.ps1"]))
        git = shutil.which("git")
        self.assertIsNotNone(git, "Git is required to read the historical extraction")
        reference = subprocess.run(
            [git, "cat-file", "blob", "88c2149b3b3034fbd0d7ef23c2382f4b01648e4c:engine/winbooksplit_engine.py"],
            cwd=ROOT, env=extraction.clean_environment(), capture_output=True, timeout=30, check=False)
        self.assertEqual(reference.returncode, 0, reference.stderr.decode("utf-8", errors="replace"))
        extracted = extraction.function_asts(reference.stdout.decode("utf-8"))
        self.assertEqual(set(prior), {"log", "split_pdf", "write_slice"})
        self.assertTrue(all(extracted.get(name) == tree for name, tree in prior.items()))

    def test_ac011_import_has_no_processing_or_configuration_side_effects(self):
        with tempfile.TemporaryDirectory(prefix="wbs-import-test-") as directory:
            observation = extraction.probe_import(Path(directory).resolve())
            self.assertTrue(observation["import_safe"])
            self.assertTrue(observation["logger_stdout_configuration_unchanged"])
            self.assertTrue(observation["cwd_unchanged"])

    def test_cleanup_wrong_owner_cannot_delete_any_files(self):
        with tempfile.TemporaryDirectory(prefix="wbs-docs-owner-test-") as directory:
            documents = Path(directory).resolve()
            token = uuid.uuid4().hex
            output = extraction.owned_document_directory(documents, token)
            marker = output / ".wbs-m1-owner.json"
            marker.write_text(json.dumps({"token": "another-owner"}), encoding="utf-8")
            before = {path.name: path.read_bytes() for path in output.iterdir()}
            with self.assertRaisesRegex(RuntimeError, "ownership"):
                extraction.cleanup_document_directory(output, documents, token, set())
            self.assertEqual(before, {path.name: path.read_bytes() for path in output.iterdir()})

    def test_cleanup_unknown_neighbor_cannot_delete_any_files(self):
        with tempfile.TemporaryDirectory(prefix="wbs-docs-neighbor-test-") as directory:
            documents = Path(directory).resolve()
            token = uuid.uuid4().hex
            output = extraction.owned_document_directory(documents, token)
            (output / "another-owner.txt").write_bytes(b"Must not be deleted")
            before = {path.name: path.read_bytes() for path in output.iterdir()}
            with self.assertRaisesRegex(RuntimeError, "unexpected member"):
                extraction.cleanup_document_directory(output, documents, token, set())
            self.assertEqual(before, {path.name: path.read_bytes() for path in output.iterdir()})

    def test_cleanup_wrong_containment_cannot_delete_owned_directory(self):
        with tempfile.TemporaryDirectory(prefix="wbs-docs-containment-test-") as directory:
            documents = Path(directory).resolve()
            token = uuid.uuid4().hex
            output = extraction.owned_document_directory(documents, token)
            wrong_root = documents / "different-root"
            wrong_root.mkdir()
            with self.assertRaisesRegex(RuntimeError, "containment"):
                extraction.cleanup_document_directory(output, wrong_root, token, set())
            self.assertTrue((output / ".wbs-m1-owner.json").is_file())


if __name__ == "__main__":
    unittest.main(verbosity=2)
