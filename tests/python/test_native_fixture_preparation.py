"""Dependency-light controls for the developer fixture preparation boundary."""
import ast
import importlib.util
from pathlib import Path
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[2]
HELPER_ROOT = ROOT / "tools/compatibility"
spec = importlib.util.spec_from_file_location("wbs_native_preparer_under_test", HELPER_ROOT / "prepare_fixtures.py")
prepare = importlib.util.module_from_spec(spec)
spec.loader.exec_module(prepare)


class FixturePreparationControls(unittest.TestCase):
    def test_existing_supervised_converter_and_original_generators_reused(self):
        tree = ast.parse((HELPER_ROOT / "prepare_fixtures.py").read_text(encoding="utf-8"))
        calls = [node for node in ast.walk(tree) if isinstance(node, ast.Call)]
        attributes = {node.func.attr for node in calls if isinstance(node.func, ast.Attribute)}
        self.assertIn("_run_converter", attributes)
        self.assertIn("_default_process_factory", attributes)
        self.assertIn("_pdf_bytes", attributes)
        self.assertIn("authored_epub", attributes)
        self.assertFalse(attributes & {"Popen", "communicate", "run", "unlink", "rmtree"})

    def test_nonempty_owned_output_rejected_before_tools_or_writes(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder).resolve()
            repository, output = root / "r", root / "o"
            repository.mkdir()
            output.mkdir()
            foreign = output / "foreign.txt"
            foreign.write_bytes(b"keep")
            with self.assertRaises(ValueError):
                prepare.generate(repository, output, root / "missing.exe", root / "report.json")
            self.assertEqual(foreign.read_bytes(), b"keep")
            self.assertEqual(list(output.iterdir()), [foreign])
            self.assertFalse((root / "report.json").exists())

    def test_checkout_target_rejected_before_tools_or_writes(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder).resolve()
            repository = root / "r"
            repository.mkdir()
            output = repository / "o"
            output.mkdir()
            with self.assertRaises(ValueError):
                prepare.generate(repository, output, root / "missing.exe", root / "report.json")
            self.assertEqual(list(output.iterdir()), [])
            self.assertFalse((root / "report.json").exists())


if __name__ == "__main__":
    unittest.main(verbosity=2)
