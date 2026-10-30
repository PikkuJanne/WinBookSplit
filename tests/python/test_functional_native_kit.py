"""Independent native-fixture oracle failures; these units are not Store acceptance."""
from contextlib import redirect_stdout
from hashlib import sha256
import importlib.util
import json
from io import StringIO
from pathlib import Path
import tempfile
import unittest
from pypdf import PdfReader, PdfWriter

ROOT = Path(__file__).resolve().parents[2]


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


fixture = load("native_kit_fixture", ROOT / "tests/python/test_functional_pdf_repair.py")
oracle = load("native_kit_oracle", ROOT / "tools/compatibility/functional_oracle.py")


class FunctionalNativeOracleTests(unittest.TestCase):
    def setUp(self):
        work = tempfile.TemporaryDirectory(prefix="WBS-functional-native-oracle-")
        self.addCleanup(work.cleanup)
        self.root = Path(work.name)
        self.source = self.root / "combined.pdf"
        data = fixture.build_fixture("empty_password_and_local_goto")
        self.source.write_bytes(data)
        self.expected = {"source": {"sha256": sha256(data).hexdigest(), "bytes": len(data)},
                         "ranges": [[0, 2], [2, 4], [4, 6]],
                         "pages": [sha256(f"BT /F1 16 Tf 72 720 Td (Synthetic test page {n}) Tj ET\n".encode()).hexdigest()
                                   for n in range(1, 7)]}
        with redirect_stdout(StringIO()):
            result = fixture.engine.run_split(self.source, self.root, "manual", "1,3,5")
        self.assertEqual(result["exit_code"], 0)
        result = json.loads(fixture.engine.serialize_result(result))
        self.outcome = {"code": result["code"], "exit_code": result["exit_code"],
                        "written_count": result["written_count"], "engine_result": result,
                        "final_directory": result["execution"]["final_directory"]}

    def test_actual_combined_engine_output_matches_authored_page_oracle(self):
        self.assertEqual(oracle.validate(self.expected, self.source, self.outcome)["status"], "PASS")

    def test_changed_source_fails_even_when_outputs_are_valid(self):
        self.source.write_bytes(self.source.read_bytes() + b"\n")
        with self.assertRaisesRegex(ValueError, "source_changed"):
            oracle.validate(self.expected, self.source, self.outcome)

    def test_wrong_ranges_fail_independently_of_page_hashes(self):
        self.outcome["engine_result"]["plan"]["ranges"][0] = [0, 1]
        with self.assertRaisesRegex(ValueError, "independent_ranges_differ"):
            oracle.validate(self.expected, self.source, self.outcome)

    def test_missing_ignored_view_warning_fails(self):
        self.outcome["engine_result"]["warnings"] = []
        with self.assertRaisesRegex(ValueError, "opening_warning_count_invalid"):
            oracle.validate(self.expected, self.source, self.outcome)

    def rewrite_first(self, *, encrypted=False, reverse=False):
        result = self.outcome["engine_result"]
        path = Path(self.outcome["final_directory"]) / result["execution"]["outputs"][0]["filename"]
        reader = PdfReader(path)
        writer = PdfWriter()
        for page in list(reader.pages)[::-1] if reverse else reader.pages:
            writer.add_page(page)
        if encrypted:
            writer.encrypt("", algorithm="RC4-40")
        with path.open("wb") as stream:
            writer.write(stream)

    def test_reordered_chapter_pages_fail(self):
        self.rewrite_first(reverse=True)
        with self.assertRaisesRegex(ValueError, "physical_page_content_or_order_differ"):
            oracle.validate(self.expected, self.source, self.outcome)

    def test_encrypted_chapter_fails(self):
        self.rewrite_first(encrypted=True)
        with self.assertRaisesRegex(ValueError, "chapter_structure_invalid"):
            oracle.validate(self.expected, self.source, self.outcome)


if __name__ == "__main__":
    unittest.main(verbosity=2)
