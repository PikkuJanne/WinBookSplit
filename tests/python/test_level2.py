"""Parent-aware Level 2 boundaries using generated synthetic inputs only."""

from hashlib import sha256
import importlib.util
from io import BytesIO
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, NullObject, TextStringObject


ROOT = Path(__file__).resolve().parents[2]


def load_module(name, path):
    spec = importlib.util.spec_from_file_location(name, path.resolve())
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


engine = load_module("wbs_level2_test_engine", ROOT / "engine/winbooksplit_engine.py")
fixtures = load_module("wbs_level2_test_fixtures", ROOT / "tests/fixtures/generate_pdf_fixtures.py")
ORACLES = [case for case in json.loads(
    (ROOT / "docs/codex-v1.0.0/PLAN_ORACLES.json").read_text(encoding="utf-8")
)["bookmarks"] if case["level"] == 2]


class Destination(dict):
    def __init__(self, title, page):
        super().__init__({"/Title": title})
        self.title, self.page, self.node = title, page, None


class SyntheticReader:
    def __init__(self, outline, pages=10):
        self.outline, self.pages = outline, [None] * pages

    def get_destination_page_number(self, node):
        if isinstance(node.page, Exception):
            raise node.page
        return node.page


def oracle_outline(nodes):
    result = []
    for node in nodes:
        result.append(Destination(node["title"], node["start_page"] - 1))
        children = oracle_outline(node.get("children", []))
        if children:
            result.append(children)
    return result


def fixture_outline(nodes):
    return [{"title": node["title"], "start_page": node["start_page"],
             "children": fixture_outline(node.get("children", []))} for node in nodes]


class Level2PlanTests(unittest.TestCase):
    def assert_plan_error(self, outline, code, pages=10):
        with self.assertRaises(engine.BookmarkPlanError) as raised:
            engine.plan_level2(SyntheticReader(outline, pages), pages)
        self.assertEqual(raised.exception.code, code)
        self.assertTrue(str(raised.exception))
        self.assertIsInstance(raised.exception.warnings, list)
        return raised.exception

    def test_four_unchanged_level2_oracles(self):
        self.assertEqual([case["id"] for case in ORACLES], ["BM-03", "BM-04", "BM-06", "BM-07"])
        for case in ORACLES:
            with self.subTest(oracle=case["id"]):
                outline = oracle_outline(case["parents"])
                if "expected_error" in case:
                    self.assert_plan_error(outline, case["expected_error"], case["pages"])
                else:
                    plan = engine.plan_level2(SyntheticReader(outline, case["pages"]), case["pages"])
                    self.assertEqual(plan["ranges"], [tuple(item) for item in case["expected_ranges"]])
                    self.assertEqual([(entry["start"], entry["end"]) for entry in plan["entries"]], plan["ranges"])
                    self.assertTrue(set(case.get("expected_warnings", [])) <= {warning["code"] for warning in plan["warnings"]})
                    self.assertEqual(plan["mode"], "2")
                    self.assertEqual(plan["total_pages"], case["pages"])

    def test_ac023_openings_and_children_have_parent_identity_and_end_at_parent_boundary(self):
        outline = oracle_outline(ORACLES[0]["parents"])
        plan = engine.plan_level2(SyntheticReader(outline, 12), 12)
        self.assertEqual([entry["title"] for entry in plan["entries"]],
                         ["Front matter", "A - Opening pages", "A1", "A2", "B - Opening pages", "B1"])
        self.assertEqual([entry["reason"] for entry in plan["entries"]],
                         ["front_matter", "parent_opening", "bookmark", "bookmark", "parent_opening", "bookmark"])
        nodes = {node["title"]: node for node in plan["bookmarks"]}
        self.assertIsNone(plan["entries"][0]["parent_id"])
        self.assertEqual([entry["parent_id"] for entry in plan["entries"][1:4]], [nodes["A"]["id"]] * 3)
        self.assertEqual([entry["parent_id"] for entry in plan["entries"][4:]], [nodes["B"]["id"]] * 2)
        self.assertEqual(plan["entries"][2]["bookmark_id"], nodes["A1"]["id"])
        self.assertEqual(plan["entries"][3]["bookmark_id"], nodes["A2"]["id"])
        self.assertEqual(plan["entries"][3]["end"], 8)

    def test_ac024_parent_without_children_is_an_intact_fallback(self):
        plan = engine.plan_level2(SyntheticReader(oracle_outline(ORACLES[1]["parents"]), 12), 12)
        self.assertEqual(plan["ranges"], [(0, 3), (3, 6), (6, 12)])
        fallback = plan["entries"][-1]
        parent = next(node for node in plan["bookmarks"] if node["title"] == "B")
        self.assertEqual(fallback["title"], "B")
        self.assertEqual(fallback["reason"], "parent_fallback")
        self.assertEqual(fallback["parent_id"], parent["id"])

    def test_fallback_parents_before_or_between_children_cannot_be_absorbed_by_a_neighbor(self):
        cases = [([Destination("A", 0), Destination("B", 5), [Destination("B1", 6)]],
                  [(0, 5), (5, 6), (6, 10)], ["A", "B - Opening pages", "B1"]),
                 ([Destination("A", 0), [Destination("A1", 1)], Destination("B", 3),
                   Destination("C", 6), [Destination("C1", 8)]],
                  [(0, 1), (1, 3), (3, 6), (6, 8), (8, 10)],
                  ["A - Opening pages", "A1", "B", "C - Opening pages", "C1"])]
        for number, (outline, ranges, titles) in enumerate(cases):
            with self.subTest(case=number):
                plan = engine.plan_level2(SyntheticReader(outline), 10)
                self.assertEqual(plan["ranges"], ranges)
                self.assertEqual([entry["title"] for entry in plan["entries"]], titles)
                self.assertTrue(any(entry["reason"] == "parent_fallback" for entry in plan["entries"]))

    def test_ac025_equal_start_creates_no_empty_opening_and_first_child_alias_wins(self):
        outline = [Destination("A", 2), [Destination("A2", 4), Destination("A1", 2),
                   Destination("Alias of A1", 2)], Destination("B", 6)]
        plan = engine.plan_level2(SyntheticReader(outline), 10)
        self.assertEqual(plan["ranges"], [(0, 2), (2, 4), (4, 6), (6, 10)])
        self.assertEqual([entry["title"] for entry in plan["entries"]], ["Front matter", "A1", "A2", "B"])
        self.assertTrue(all(entry["start"] < entry["end"] for entry in plan["entries"]))
        self.assertTrue({"duplicate_destination", "outline_reordered"} <= {warning["code"] for warning in plan["warnings"]})

    def test_outside_children_before_parent_at_parent_end_and_after_parent_end_are_ignored(self):
        outline = [Destination("A", 2), [Destination("Before A", 1), Destination("A1", 3),
                   Destination("At B start", 6), Destination("After B start", 8)], Destination("B", 6)]
        plan = engine.plan_level2(SyntheticReader(outline), 10)
        self.assertEqual(plan["ranges"], [(0, 2), (2, 3), (3, 6), (6, 10)])
        records = {node["title"]: node for node in plan["bookmarks"]}
        outside_orders = {records[title]["source_order"] for title in ("Before A", "At B start", "After B start")}
        self.assertEqual({warning["source_order"] for warning in plan["warnings"]
                          if warning["code"] == "child_outside_parent"}, outside_orders)
        self.assertEqual(plan["entries"][-1]["reason"], "parent_fallback")

    def test_child_at_next_parent_start_can_only_split_its_actual_parent(self):
        outline = [Destination("A", 0), [Destination("A1", 1), Destination("A outside", 5)],
                   Destination("B", 5), [Destination("B1", 5)]]
        plan = engine.plan_level2(SyntheticReader(outline), 10)
        self.assertEqual(plan["ranges"], [(0, 1), (1, 5), (5, 10)])
        self.assertEqual([entry["title"] for entry in plan["entries"]], ["A - Opening pages", "A1", "B1"])
        self.assertTrue(any(warning["code"] == "child_outside_parent" for warning in plan["warnings"]))

    def test_first_parent_and_its_subtree_win_over_parent_aliases(self):
        outline = [Destination("A", 0), [Destination("A1", 2)],
                   Destination("Alias A", 0), [Destination("Alias child", 1)], Destination("B", 5)]
        plan = engine.plan_level2(SyntheticReader(outline), 10)
        self.assertEqual(plan["ranges"], [(0, 2), (2, 5), (5, 10)])
        self.assertEqual([entry["title"] for entry in plan["entries"]], ["A - Opening pages", "A1", "B"])
        self.assertTrue({"duplicate_destination", "duplicate_parent_subtree"} <= {warning["code"] for warning in plan["warnings"]})

    def test_alias_subtree_cannot_satisfy_the_selected_level_requirement(self):
        outline = [Destination("A", 0), Destination("Alias A", 0), [Destination("Alias child", 1)], Destination("B", 5)]
        error = self.assert_plan_error(outline, "no_bookmarks_at_level")
        self.assertTrue(any(warning["code"] == "duplicate_parent_subtree" for warning in error.warnings))

    def test_only_outside_children_cannot_satisfy_the_selected_level_requirement(self):
        outline = [Destination("A", 2), [Destination("Before A", 1), Destination("At B start", 6),
                   Destination("After B start", 8)], Destination("B", 6)]
        error = self.assert_plan_error(outline, "no_bookmarks_at_level")
        self.assertEqual(len([warning for warning in error.warnings if warning["code"] == "child_outside_parent"]), 3)
        with patch.object(engine, "PdfReader", return_value=SyntheticReader(outline)), \
                patch.object(engine, "write_slice") as writer, patch.object(engine, "log"):
            with self.assertRaises(SystemExit) as raised:
                engine.split_pdf("synthetic-input.pdf", "synthetic-unused-output", "2")
        self.assertEqual(raised.exception.code, 55)
        writer.assert_not_called()

    def test_unusable_parent_descendants_are_warned_and_not_orphaned_into_other_parents(self):
        outline = [Destination("Invalid parent", None), [Destination("Orphan child", 1)],
                   Destination("A", 2), [Destination("A1", 3)], Destination("B", 6)]
        plan = engine.plan_level2(SyntheticReader(outline), 10)
        self.assertEqual(plan["ranges"], [(0, 2), (2, 3), (3, 6), (6, 10)])
        self.assertEqual([entry["title"] for entry in plan["entries"]], ["Front matter", "A - Opening pages", "A1", "B"])
        self.assertTrue(any(warning["code"] == "invalid_parent_subtree" for warning in plan["warnings"]))
        orphan = next(node for node in plan["bookmarks"] if node["title"] == "Orphan child")
        self.assertTrue(all(entry.get("bookmark_id") != orphan["id"] for entry in plan["entries"]))

    def test_unusable_parent_children_and_depth3_children_cannot_satisfy_level2(self):
        cases = [[Destination("Invalid parent", None), [Destination("Orphan child", 1)], Destination("B", 5)],
                 [Destination("A", 0), [Destination("Invalid child", None), [Destination("Grandchild", 2)]]]]
        for number, outline in enumerate(cases):
            with self.subTest(case=number):
                self.assert_plan_error(outline, "no_bookmarks_at_level")

    def test_depth3_metadata_is_retained_without_creating_boundaries(self):
        outline = [Destination("A", 0), [Destination("A1", 1), [Destination("Grandchild", 2)]]]
        plan = engine.plan_level2(SyntheticReader(outline), 10)
        self.assertEqual(plan["ranges"], [(0, 1), (1, 10)])
        grandchild = next(node for node in plan["bookmarks"] if node["title"] == "Grandchild")
        self.assertEqual(grandchild["depth"], 3)
        self.assertEqual(len(grandchild["lineage"]), 2)
        self.assertTrue(all(entry.get("bookmark_id") != grandchild["id"] for entry in plan["entries"]))

    def test_recoverable_invalid_and_external_children_are_diagnosed_with_a_complete_valid_plan(self):
        invalid = [None, -1, 10, False, 3.5, ValueError("Synthetic child destination failure")]
        children = [Destination(f"Invalid {number}", page) for number, page in enumerate(invalid)]
        external = Destination("External child", 5)
        external.node = {"/A": {"/S": "/URI"}}
        children.extend([external, Destination("A1", 3)])
        plan = engine.plan_level2(SyntheticReader([Destination("A", 0), children]), 10)
        self.assertEqual(plan["ranges"], [(0, 3), (3, 10)])
        rejected = plan["bookmarks"][1:-1]
        self.assertEqual([node["page"] for node in rejected], [None] * len(rejected))
        for node in rejected:
            self.assertTrue(any(warning["source_order"] == node["source_order"] and warning["message"]
                                for warning in plan["warnings"]))

    def test_physically_out_of_order_parents_and_children_keep_their_own_intervals(self):
        outline = [Destination("B", 6), [Destination("B1", 7)], Destination("A", 0),
                   [Destination("A2", 4), Destination("A1", 1)]]
        plan = engine.plan_level2(SyntheticReader(outline), 10)
        self.assertEqual(plan["ranges"], [(0, 1), (1, 4), (4, 6), (6, 7), (7, 10)])
        self.assertEqual([entry["title"] for entry in plan["entries"]],
                         ["A - Opening pages", "A1", "A2", "B - Opening pages", "B1"])
        self.assertTrue(any(warning["code"] == "outline_reordered" for warning in plan["warnings"]))
        self.assertEqual([entry["sequence"] for entry in plan["entries"]], list(range(1, 6)))
        self.assertEqual([entry["filename"].split(" - ", 1)[0] for entry in plan["entries"]],
                         ["01", "02", "03", "04", "05"])

    def test_ac026_absent_outlines_invalid_parents_and_missing_selected_level_are_distinct(self):
        self.assert_plan_error([], "no_bookmarks")
        self.assert_plan_error([Destination("Invalid parent", None), [Destination("Child", 1)]], "no_usable_bookmarks")
        self.assert_plan_error([Destination("A", 0), Destination("B", 5)], "no_bookmarks_at_level")

    def test_zero_pages_and_malformed_cyclic_outlines_reject_before_execution(self):
        self.assert_plan_error([Destination("A", 0), [Destination("A1", 0)]], "invalid_document", pages=0)
        cycle = [Destination("A", 0)]
        cycle.append(cycle)
        shared = [Destination("Shared child", 1)]
        for number, outline in enumerate((cycle, [[Destination("Orphan", 1)]],
                                          [Destination("A", 0), shared, Destination("B", 5), shared])):
            with self.subTest(case=number):
                error = self.assert_plan_error(outline, "invalid_outline")
                self.assertTrue(error.warnings)

    def test_no_plan_and_malformed_requests_write_nothing_and_keep_meaningful_exit_codes(self):
        cycle = [Destination("A", 0)]
        cycle.append(cycle)
        cases = [([], 10, 55, "no_bookmarks"),
                 ([Destination("Invalid parent", None)], 10, 55, "no_usable_bookmarks"),
                 ([Destination("A", 0)], 10, 55, "no_bookmarks_at_level"),
                 (cycle, 10, 1, "invalid_outline"),
                 ([Destination("A", 0), [Destination("A1", 0)]], 0, 1, "invalid_document")]
        for outline, pages, exit_code, code in cases:
            with self.subTest(code=code):
                with patch.object(engine, "PdfReader", return_value=SyntheticReader(outline, pages)), \
                        patch.object(engine, "write_slice") as writer, patch.object(engine, "log") as log:
                    with self.assertRaises(SystemExit) as raised:
                        engine.split_pdf("synthetic-input.pdf", "synthetic-unused-output", "2")
                self.assertEqual(raised.exception.code, exit_code)
                writer.assert_not_called()
                messages = " ".join(str(call.args[0]) for call in log.call_args_list)
                self.assertIn(code, messages)
                self.assertNotIn("[NO_BOOKMARKS_FOUND]", messages)
                payloads = [json.loads(str(call.args[0])) for call in log.call_args_list
                            if str(call.args[0]).startswith("{")]
                self.assertEqual(len(payloads), 1)
                self.assertEqual(payloads[0]["protocol"], "winbooksplit.result")
                self.assertEqual(payloads[0]["code"], code)
                self.assertEqual(payloads[0]["exit_code"], exit_code)


class Level2SyntheticPdfTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.owned = tempfile.TemporaryDirectory(prefix="wbs-level2-unit-")
        cls.addClassCleanup(cls.owned.cleanup)
        cls.work = Path(cls.owned.name).resolve()

    def synthetic_page_writer(self):
        definition = {"pages": 10, "title": "Synthetic Level 2 lineage regression", "outline": []}
        writer = PdfWriter()
        for page in PdfReader(BytesIO(fixtures._pdf_bytes(definition))).pages:
            writer.add_page(page)
        return writer

    def test_real_unusable_parents_retain_children_without_orphaning_them_into_valid_intervals(self):
        for kind in ("null", "unknown-named", "external"):
            with self.subTest(destination=kind):
                writer = self.synthetic_page_writer()
                invalid_parent = writer.add_outline_item("Unusable parent", 0)
                writer.add_outline_item("Orphan candidate", 1, parent=invalid_parent)
                action = invalid_parent.get_object()["/A"]
                if kind == "null":
                    action[NameObject("/D")] = NullObject()
                elif kind == "unknown-named":
                    action[NameObject("/D")] = TextStringObject("Missing synthetic named destination")
                else:
                    action[NameObject("/S")] = NameObject("/URI")
                    del action["/D"]
                    action[NameObject("/URI")] = TextStringObject("https://example.invalid/synthetic")
                valid_parent = writer.add_outline_item("B", 5)
                writer.add_outline_item("B1", 6, parent=valid_parent)
                source = self.work / ("unusable-parent-" + kind + ".pdf")
                with source.open("xb") as stream:
                    writer.write(stream)
                before = source.read_bytes()
                reader = PdfReader(source)
                raw = reader.outline
                self.assertEqual(raw[0].title, "Unusable parent")
                self.assertIn(reader.get_destination_page_number(raw[0]), (None, -1))
                self.assertEqual(raw[1][0].title, "Orphan candidate")
                self.assertEqual(reader.get_destination_page_number(raw[1][0]), 1)
                plan = engine.plan_level2(reader, 10)
                nodes = {node["title"]: node for node in plan["bookmarks"]}
                self.assertEqual(nodes["Orphan candidate"]["parent_id"], nodes["Unusable parent"]["id"])
                self.assertEqual(nodes["Orphan candidate"]["lineage"], [nodes["Unusable parent"]["id"]])
                self.assertEqual(plan["ranges"], [(0, 5), (5, 6), (6, 10)])
                self.assertEqual([entry["title"] for entry in plan["entries"]], ["Front matter", "B - Opening pages", "B1"])
                self.assertTrue(any(warning["code"] == "invalid_parent_subtree" for warning in plan["warnings"]))
                output = self.work / ("unusable-parent-" + kind + "-output")
                output.mkdir()
                with patch.object(engine, "log"):
                    execution = engine.split_pdf(str(source), str(output), "2")
                self.assertEqual([fixtures.page_ids(path) for path in sorted(Path(execution["final_directory"]).glob("*.pdf"))],
                                 [list(range(1, 6)), [6], list(range(7, 11))])
                self.assertEqual(source.read_bytes(), before)

    def test_real_alias_children_and_grandchildren_below_unusable_direct_child_cannot_create_level2(self):
        for kind in ("alias-only", "grandchild-only"):
            with self.subTest(lineage=kind):
                writer = self.synthetic_page_writer()
                parent = writer.add_outline_item("A", 0)
                if kind == "alias-only":
                    alias = writer.add_outline_item("Alias A", 0)
                    writer.add_outline_item("Alias child", 1, parent=alias)
                else:
                    unusable_child = writer.add_outline_item("Unusable direct child", 1, parent=parent)
                    writer.add_outline_item("Grandchild", 2, parent=unusable_child)
                    unusable_child.get_object()["/A"][NameObject("/D")] = NullObject()
                source = self.work / (kind + ".pdf")
                with source.open("xb") as stream:
                    writer.write(stream)
                before = source.read_bytes()
                reader = PdfReader(source)
                raw = reader.outline
                if kind == "alias-only":
                    self.assertEqual([raw[0].title, raw[1].title, raw[2][0].title], ["A", "Alias A", "Alias child"])
                    self.assertEqual(reader.get_destination_page_number(raw[2][0]), 1)
                else:
                    self.assertEqual([raw[0].title, raw[1][0].title, raw[1][1][0].title],
                                     ["A", "Unusable direct child", "Grandchild"])
                    self.assertIn(reader.get_destination_page_number(raw[1][0]), (None, -1))
                    self.assertEqual(reader.get_destination_page_number(raw[1][1][0]), 2)
                normalized = engine.normalize_bookmarks(reader, 10)["bookmarks"]
                self.assertEqual(normalized[-1]["depth"], 2 if kind == "alias-only" else 3)
                self.assertEqual(normalized[-1]["parent_id"], normalized[-2]["id"])
                with self.assertRaises(engine.BookmarkPlanError) as raised:
                    engine.plan_level2(reader, 10)
                self.assertEqual(raised.exception.code, "no_bookmarks_at_level")
                expected_warning = "duplicate_parent_subtree" if kind == "alias-only" else "invalid_destination"
                self.assertTrue(any(warning["code"] == expected_warning for warning in raised.exception.warnings))
                output = self.work / (kind + "-output")
                output.mkdir()
                with patch.object(engine, "write_slice") as writer_spy, patch.object(engine, "log"):
                    with self.assertRaises(SystemExit) as exited:
                        engine.split_pdf(str(source), str(output), "2")
                self.assertEqual(exited.exception.code, 55)
                writer_spy.assert_not_called()
                self.assertEqual(list(output.iterdir()), [])
                self.assertEqual(source.read_bytes(), before)

    def test_real_oracle_outputs_preserve_every_page_and_the_expected_parent_boundaries(self):
        for case in ORACLES:
            with self.subTest(oracle=case["id"]):
                definition = {"pages": case["pages"], "title": "Synthetic Level 2 " + case["id"],
                              "outline": fixture_outline(case["parents"])}
                source = self.work / (case["id"] + ".pdf")
                with source.open("xb") as stream:
                    stream.write(fixtures._pdf_bytes(definition))
                before = sha256(source.read_bytes()).hexdigest()
                output = self.work / (case["id"] + "-output")
                output.mkdir()
                neighbor = output / "synthetic-neighbor.txt"
                neighbor.write_bytes(b"Original synthetic neighbor preserved\n")
                with patch.object(engine, "log"):
                    if "expected_error" in case:
                        with self.assertRaises(SystemExit) as raised:
                            engine.split_pdf(str(source), str(output), "2")
                        self.assertEqual(raised.exception.code, 55)
                        self.assertEqual({path.name for path in output.iterdir()}, {neighbor.name})
                    else:
                        plan = engine.plan_level2(PdfReader(source), case["pages"])
                        execution = engine.split_pdf(str(source), str(output), "2")
                        published = Path(execution["final_directory"])
                        observed = [fixtures.page_ids(path) for path in sorted(published.glob("*.pdf"))]
                        self.assertEqual(observed, [list(range(start + 1, end + 1)) for start, end in case["expected_ranges"]])
                        self.assertEqual([page for pages in observed for page in pages], list(range(1, case["pages"] + 1)))
                        self.assertEqual({path.name for path in published.iterdir()},
                                         {entry["filename"] for entry in plan["entries"]} |
                                         {".WinBookSplit-owner.json", execution["manifest_filename"]})
                        self.assertEqual({path.name for path in output.iterdir()}, {published.name, neighbor.name})
                self.assertEqual(neighbor.read_bytes(), b"Original synthetic neighbor preserved\n")
                self.assertEqual(sha256(source.read_bytes()).hexdigest(), before)


if __name__ == "__main__":
    unittest.main(verbosity=2)
