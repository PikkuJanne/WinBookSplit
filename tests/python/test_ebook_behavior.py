"""Meaningful local observer/configuration controls; no PDF/Calibre claim."""
from contextlib import ExitStack
import importlib.util
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
from urllib.error import HTTPError
from urllib.request import ProxyHandler, Request, build_opener

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location("wbs_ebook_behavior_units", ROOT / "tests/ebooks/behavior.py")
behavior = importlib.util.module_from_spec(spec)
spec.loader.exec_module(behavior)


class EbookBehaviorTests(unittest.TestCase):
    def test_owned_configuration_restores_existing_environment_on_failure(self):
        with tempfile.TemporaryDirectory() as directory, ExitStack() as stack:
            work = Path(directory)
            protected = work / "existing-config.txt"
            protected.write_bytes(b"Unchanged authored prior configuration")
            stack.enter_context(patch.dict(os.environ, {name: str(protected) for name in behavior.VARIABLES}))
            selected = behavior.CalibreEnvironment(work / "owned")
            with self.assertRaisesRegex(RuntimeError, "authored failure"):
                with selected.inherited():
                    for name, path in selected.values.items():
                        self.assertEqual(os.environ[name], path)
                    (Path(selected.values[behavior.VARIABLES[0]]) / "observed.txt").write_bytes(b"Owned generated setting")
                    raise RuntimeError("authored failure")
            self.assertTrue(selected.receipt()["inherited_environment_restored"])
            self.assertTrue(selected.receipt()["new_empty_process_directories"])
            self.assertEqual(set(selected.receipt()["after"][behavior.VARIABLES[0]]["files"]), {"observed.txt"})
            self.assertEqual(protected.read_bytes(), b"Unchanged authored prior configuration")
            self.assertTrue(all(os.environ[name] == str(protected) for name in behavior.VARIABLES))

    def test_positive_controls_reach_both_same_resource_endpoints(self):
        observer = behavior.LoopbackCanary()
        try:
            observer.positive("before")
            with observer.conversion("authored-no-fetch"):
                pass
            observer.positive("after")
        finally:
            observer.stop()
        result = observer.receipt()
        self.assertEqual(len(result["positive_controls"]), 4)
        self.assertEqual([row["phase"] for row in result["events"]], ["control-before"] * 2 + ["control-after"] * 2)
        self.assertEqual(result["conversion_requests"], [])
        self.assertTrue(result["listener_stopped"])
        self.assertFalse(result["system_network_settings_changed"])

    def test_conversion_request_is_recorded_even_for_unknown_path_or_body(self):
        observer = behavior.LoopbackCanary()
        try:
            opener = build_opener(ProxyHandler({}))
            with observer.conversion("authored-unexpected-request"):
                with self.assertRaises(HTTPError) as failure:
                    opener.open(Request(observer.base_url + "unexpected", data=b"authored-body", method="POST"), timeout=5)
                self.assertEqual(failure.exception.code, 404)
                failure.exception.close()
        finally:
            observer.stop()
        requests = observer.receipt()["conversion_requests"]
        self.assertEqual(len(requests), 1)
        self.assertEqual(requests[0]["method"], "POST")
        self.assertEqual(requests[0]["path"], "/" + observer.nonce + "/unexpected")
        self.assertEqual(requests[0]["body_length_declared"], len(b"authored-body"))


if __name__ == "__main__":
    unittest.main()
