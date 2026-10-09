"""Level 1 coverage and bounded-outline checks using synthetic content only."""

from hashlib import sha256
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, NumberObject, TextStringObject


ROOT = Path(__file__).resolve().parents[2]


def load_module(name, path):
    spec = importlib.util.spec_from_file_location(name, path.resolve())
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


engine = load_module("wbs_level1_test_engine", ROOT / "engine/winbooksplit_engine.py")
fixtures = load_module("wbs_level1_test_fixtures", ROOT / "tests/fixtures/generate_pdf_fixtures.py")
ORACLES = json.loads((ROOT / "docs/codex-v1.0.0/PLAN_ORACLES.json").read_text(encoding="utf-8"))["bookmarks"]


class Destination(dict):
    """Synthetic dictionary destination, including its original raw action."""
    def __init__(self, title, page, action=None):
        super().__init__({"/Title": title})
        self.title, self.page = title, page
        self.node = None if action is None else {"/A": action}


class SyntheticReader:
    def __init__(self, outline, pages=10):
        self.outline, self.pages = outline, [None] * pages
        self.resolved_nodes = []

    def get_destination_page_number(self, node):
        self.resolved_nodes.append(node)
        page = node.page if hasattr(node, "page") else node["test-page"]
        if isinstance(page, Exception):
            raise page
        return page


def oracle_outline(parents):
    result = []
    for parent in parents:
        result.append(Destination(parent["title"], parent["start_page"] - 1))
        children = oracle_outline(parent.get("children", []))
        if children:
            result.append(children)
    return result


def nested_outline(depth):
    result = [Destination(f"Depth {depth}", 0)]
    for current in range(depth - 1, 0, -1):
        result = [Destination(f"Depth {current}", 0), result]
    return result


class Level1PlanTests(unittest.TestCase):
    def assert_plan_error(self, reader, pages, code):
        with self.assertRaises(engine.BookmarkPlanError) as raised:
            engine.plan_level1(reader, pages)
        self.assertEqual(raised.exception.code, code)
        self.assertTrue(str(raised.exception), "Rejected outlines need a diagnostic")
        self.assertIsInstance(raised.exception.warnings, list)
        if code == "invalid_outline":
            self.assertTrue(raised.exception.warnings, "Malformed outlines need diagnostic warnings")
        return raised.exception

    def test_four_predeclared_level1_oracles(self):
        cases = [case for case in ORACLES if case["level"] == 1]
        self.assertEqual([case["id"] for case in cases], ["BM-01", "BM-02", "BM-05", "BM-08"])
        for case in cases:
            with self.subTest(oracle=case["id"]):
                reader = SyntheticReader(oracle_outline(case["parents"]), case["pages"])
                if "expected_error" in case:
                    self.assert_plan_error(reader, case["pages"], case["expected_error"])
                else:
                    plan = engine.plan_level1(reader, case["pages"])
                    self.assertEqual(plan["ranges"], [tuple(item) for item in case["expected_ranges"]])
                    self.assertEqual([(item["start"], item["end"]) for item in plan["entries"]], plan["ranges"])
                    self.assertEqual([item["sequence"] for item in plan["entries"]], list(range(1, len(plan["entries"]) + 1)))
                    self.assertEqual(plan["total_pages"], case["pages"])
                    self.assertTrue(set(case.get("expected_warnings", [])) <= {item["code"] for item in plan["warnings"]})

    def test_ac019_front_matter_has_its_own_title_and_complete_range(self):
        plan = engine.plan_level1(SyntheticReader([Destination("A", 3), Destination("B", 6)]), 10)
        self.assertEqual(plan["ranges"], [(0, 3), (3, 6), (6, 10)])
        front = plan["entries"][0]
        self.assertEqual(front["title"], "Front matter")
        self.assertEqual(front["reason"], "front_matter")
        self.assertIsNone(front["parent_id"])
        self.assertEqual([item["title"] for item in plan["entries"][1:]], ["A", "B"])

    def test_ac020_first_outline_alias_wins_and_sort_duplicates_are_warned(self):
        outline = [Destination("B", 6), Destination("A", 3), [Destination("A child", 4)],
                   Destination("Alias", 3), [Destination("Alias child", 5)]]
        plan = engine.plan_level1(SyntheticReader(outline), 10)
        self.assertEqual(plan["ranges"], [(0, 3), (3, 6), (6, 10)])
        self.assertEqual([item["title"] for item in plan["entries"]], ["Front matter", "A", "B"])
        self.assertTrue({"duplicate_destination", "outline_reordered"} <= {item["code"] for item in plan["warnings"]})
        # Deeper aliases remain diagnostic records, without becoming boundaries.
        self.assertEqual([item["title"] for item in plan["bookmarks"]], ["B", "A", "A child", "Alias", "Alias child"])

    def test_ac021_invalid_resolution_types_bounds_and_exceptions_are_diagnosed(self):
        invalid = [None, -1, -2, 10, True, False, 1.0, 1.5, float("nan"), float("inf"), "3",
                   ValueError("Synthetic destination resolution failure")]
        outline = [Destination("A", 3)] + [Destination(f"Invalid {number}", page) for number, page in enumerate(invalid)]
        outline.append(Destination("B", 6))
        plan = engine.plan_level1(SyntheticReader(outline), 10)
        self.assertEqual(plan["ranges"], [(0, 3), (3, 6), (6, 10)])
        rejected = plan["bookmarks"][1:-1]
        self.assertEqual([item["page"] for item in rejected], [None] * len(invalid))
        for number, item in enumerate(rejected):
            relevant = [warning for warning in plan["warnings"] if warning["source_order"] == item["source_order"]]
            self.assertTrue(relevant, f"Missing diagnostic for {item['title']}")
            self.assertTrue(all(warning["depth"] == 1 and warning["code"] and warning["message"] for warning in relevant))
            expected = "destination_error" if number == len(invalid) - 1 else "invalid_destination"
            self.assertIn(expected, {warning["code"] for warning in relevant})

    def test_external_actions_are_not_interpreted_as_internal_page_boundaries(self):
        for kind in ("/URI", "/GoToR", "/Launch", "/JavaScript"):
            with self.subTest(action=kind):
                external = Destination("External", 5, {"/S": kind})
                reader = SyntheticReader([external, Destination("A", 3)])
                plan = engine.plan_level1(reader, 10)
                self.assertEqual(plan["ranges"], [(0, 3), (3, 10)])
                self.assertIsNone(plan["bookmarks"][0]["page"])
                self.assertTrue(any(item["source_order"] == 1 and item["code"] == "external_destination" and item["message"]
                                    for item in plan["warnings"]))
                self.assertTrue(all(node is not external for node in reader.resolved_nodes))

    def test_dictionary_titles_and_empty_titles_have_deterministic_fallback(self):
        outline = [{"/Title": "Dictionary title", "test-page": 0}, Destination("", 5)]
        plan = engine.plan_level1(SyntheticReader(outline), 10)
        self.assertEqual([item["title"] for item in plan["entries"]], ["Dictionary title", "Untitled"])
        self.assertEqual(plan["ranges"], [(0, 5), (5, 10)])

    def test_raw_invalid_destination_arrays_cannot_use_a_synthesized_page(self):
        for raw in ({}, {"/Dest": []}, {"/Dest": [3]}, {"/Dest": [3, "/UnknownFit"]},
                    {"/A": {"/S": "/GoTo", "/D": [3, "/UnknownFit"]}}, {"/A": "invalid action"}):
            with self.subTest(raw=raw):
                invalid = Destination("Invalid raw destination", 3)
                invalid.node = raw
                reader = SyntheticReader([invalid, Destination("A", 6)])
                plan = engine.plan_level1(reader, 10)
                self.assertEqual(plan["ranges"], [(0, 6), (6, 10)])
                self.assertIsNone(plan["bookmarks"][0]["page"])
                self.assertTrue(any(warning["source_order"] == 1 and warning["code"] == "invalid_destination"
                                    for warning in plan["warnings"]))
                self.assertTrue(all(node is not invalid for node in reader.resolved_nodes))

    def test_ac022_source_order_depth_and_lineage_are_retained(self):
        outline = [Destination("Parent", 3), [Destination("Child", 4), [Destination("Grandchild", 5)]],
                   Destination("Second parent", 7)]
        reader = SyntheticReader(outline)
        normalized = engine.normalize_outline(reader, outline, 10)
        nodes = normalized["bookmarks"]
        self.assertEqual([item["source_order"] for item in nodes], [1, 2, 3, 4])
        self.assertEqual([item["depth"] for item in nodes], [1, 2, 3, 1])
        self.assertEqual(len({item["id"] for item in nodes}), 4)
        self.assertTrue(all(isinstance(item["id"], str) and item["id"] for item in nodes))
        self.assertEqual(nodes[0]["lineage"], [])
        self.assertIsNone(nodes[0]["parent_id"])
        self.assertEqual(nodes[1]["parent_id"], nodes[0]["id"])
        self.assertEqual(nodes[1]["lineage"], [nodes[0]["id"]])
        self.assertEqual(nodes[2]["parent_id"], nodes[1]["id"])
        self.assertEqual(nodes[2]["lineage"], [nodes[0]["id"], nodes[1]["id"]])
        self.assertEqual(nodes[3]["lineage"], [])
        plan = engine.plan_level1(SyntheticReader(outline), 10)
        self.assertEqual(plan["ranges"], [(0, 3), (3, 7), (7, 10)])

    def test_valid_depth_at_64_is_iterative_and_deeper_boundaries_are_ignored(self):
        self.assertEqual(engine.MAX_OUTLINE_DEPTH, 64)
        outline = nested_outline(64)
        normalized = engine.normalize_outline(SyntheticReader(outline), outline, 10)
        self.assertEqual([item["depth"] for item in normalized["bookmarks"]], list(range(1, 65)))
        self.assertEqual(len(normalized["bookmarks"][-1]["lineage"]), 63)
        plan = engine.plan_level1(SyntheticReader(nested_outline(64)), 10)
        self.assertEqual(plan["ranges"], [(0, 10)])

    def test_empty_outline_and_no_usable_level1_are_distinct_errors(self):
        self.assert_plan_error(SyntheticReader([]), 10, "no_bookmarks")
        error = self.assert_plan_error(SyntheticReader([Destination("Invalid", None)]), 10, "no_usable_bookmarks")
        self.assertTrue(error.warnings)
        # A valid child cannot establish a Level 1 interval for an unusable parent.
        self.assert_plan_error(SyntheticReader([Destination("Invalid", None), [Destination("Child", 0)]]), 10,
                               "no_usable_bookmarks")

    def test_zero_page_document_is_rejected(self):
        self.assert_plan_error(SyntheticReader([Destination("A", 0)], 0), 0, "invalid_document")

    def test_cycles_reused_containers_orphan_children_and_malformed_nodes_are_rejected(self):
        cycle = [Destination("A", 0)]
        cycle.append(cycle)
        shared = [Destination("Shared child", 1)]
        cases = [cycle, [Destination("A", 0), shared, Destination("B", 5), shared],
                 [[Destination("Orphan", 1)], Destination("A", 5)],
                 {"unexpected": "outline container"}]
        for number, outline in enumerate(cases):
            with self.subTest(case=number):
                self.assert_plan_error(SyntheticReader(outline), 10, "invalid_outline")
        # Unsupported scalar entries are unusable destinations with warnings;
        # unlike cyclic containers, they cannot make traversal ambiguous.
        for scalar in (None, "not a destination"):
            with self.subTest(invalid_entry=scalar):
                plan = engine.plan_level1(SyntheticReader([Destination("A", 0), scalar]), 10)
                self.assertEqual(plan["ranges"], [(0, 10)])
                self.assertIsNone(plan["bookmarks"][1]["page"])
                self.assertTrue(any(warning["code"] == "invalid_destination" and warning["source_order"] == 2
                                    for warning in plan["warnings"]))

    def test_depth_and_node_limits_are_enforced_before_writing(self):
        self.assertEqual(engine.MAX_OUTLINE_NODES, 10000)
        self.assert_plan_error(SyntheticReader(nested_outline(65)), 10, "invalid_outline")
        outline = [Destination(f"Node {number}", 0) for number in range(10001)]
        self.assert_plan_error(SyntheticReader(outline), 10, "invalid_outline")
        normalized = engine.normalize_outline(SyntheticReader(outline[:10000]), outline[:10000], 10)
        self.assertEqual(len(normalized["bookmarks"]), 10000)

    def test_reused_destination_cannot_acquire_contradictory_source_lineage(self):
        shared = Destination("Repeated destination object", 0)
        for outline in ([shared, shared], [shared, [shared]]):
            with self.subTest(nested=isinstance(outline[-1], list)):
                self.assert_plan_error(SyntheticReader(outline), 10, "invalid_outline")

    def test_outline_property_exception_is_a_structured_error(self):
        class BrokenReader:
            @property
            def outline(self):
                raise ValueError("Synthetic malformed PDF outline")
        self.assert_plan_error(BrokenReader(), 10, "invalid_outline")

    def test_rejected_outline_requests_never_reach_slice_writer(self):
        cycle = [Destination("A", 0)]
        cycle.append(cycle)
        cases = [([], 10), ([Destination("Invalid", None)], 10), (cycle, 10),
                 (nested_outline(65), 10), ([[Destination("Orphan", 0)]], 10),
                 ([Destination("A", 0)], 0)]
        for number, (outline, pages) in enumerate(cases):
            with self.subTest(case=number):
                with patch.object(engine, "PdfReader", return_value=SyntheticReader(outline, pages)), \
                        patch.object(engine, "write_slice") as writer, patch.object(engine, "log") as log:
                    with self.assertRaises(SystemExit) as raised:
                        engine.split_pdf("synthetic-test-input.pdf", "synthetic-unused-output", "1")
                self.assertNotEqual(raised.exception.code, 0)
                writer.assert_not_called()
                self.assertTrue(log.call_args_list, "Rejected run needs an error diagnostic")


class Level1SyntheticPdfTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.owned = tempfile.TemporaryDirectory(prefix="wbs-level1-unit-")
        cls.addClassCleanup(cls.owned.cleanup)
        cls.work = Path(cls.owned.name).resolve()
        cls.inputs = fixtures.generate_fixtures(cls.work / "fixtures")

    def test_real_simple_and_nested_pdf_plans_and_outputs_preserve_page_identity(self):
        for name, expected in (("simple10", [(0, 3), (3, 6), (6, 10)]),
                               ("nested12", [(0, 2), (2, 8), (8, 12)])):
            with self.subTest(fixture=name):
                source = self.inputs[name]
                before = sha256(source.read_bytes()).hexdigest()
                reader = PdfReader(source)
                plan = engine.plan_level1(reader, len(reader.pages))
                self.assertEqual(plan["ranges"], expected)
                output = self.work / name
                output.mkdir()
                with patch.object(engine, "log"):
                    execution = engine.split_pdf(str(source), str(output), "1")
                files = sorted(Path(execution["final_directory"]).glob("*.pdf"))
                self.assertEqual(len(files), len(expected))
                observed = [fixtures.page_ids(path) for path in files]
                self.assertEqual(observed, [list(range(start + 1, end + 1)) for start, end in expected])
                self.assertEqual([item for pages in observed for item in pages], list(range(1, len(reader.pages) + 1)))
                self.assertEqual(sha256(source.read_bytes()).hexdigest(), before)

    def test_real_duplicate_out_of_order_pdf_retains_first_title_and_warnings(self):
        path = self.work / "duplicate-outline.pdf"
        writer = PdfWriter()
        for page in PdfReader(self.inputs["simple10"]).pages:
            writer.add_page(page)
        writer.add_outline_item("B", 6)
        writer.add_outline_item("A", 3)
        writer.add_outline_item("Alias", 3)
        with path.open("xb") as stream:
            writer.write(stream)
        before = path.read_bytes()
        plan = engine.plan_level1(PdfReader(path), 10)
        self.assertEqual(plan["ranges"], [(0, 3), (3, 6), (6, 10)])
        self.assertEqual([item["title"] for item in plan["entries"]], ["Front matter", "A", "B"])
        self.assertTrue({"duplicate_destination", "outline_reordered"} <= {item["code"] for item in plan["warnings"]})
        self.assertEqual(path.read_bytes(), before)

    def test_real_named_destination_and_indirect_goto_action_keep_the_requested_page(self):
        path = self.work / "named-indirect-action.pdf"
        writer = PdfWriter()
        for page in PdfReader(self.inputs["simple10"]).pages:
            writer.add_page(page)
        writer.add_named_destination("Chapter A", 3)
        bookmark = writer.add_outline_item("Named A", 0).get_object()
        action = bookmark["/A"]
        action[NameObject("/D")] = TextStringObject("Chapter A")
        action[NameObject("/S")] = writer._add_object(NameObject("/GoTo"))
        with path.open("xb") as stream:
            writer.write(stream)
        before = path.read_bytes()
        plan = engine.plan_level1(PdfReader(path), 10)
        self.assertEqual(plan["ranges"], [(0, 3), (3, 10)])
        self.assertEqual(plan["bookmarks"][0]["page"], 3)
        self.assertEqual(plan["warnings"], [])
        self.assertEqual(path.read_bytes(), before)

    def test_real_indirect_fit_cannot_silently_move_the_bookmark_to_page_zero(self):
        path = self.work / "indirect-fit.pdf"
        writer = PdfWriter()
        for page in PdfReader(self.inputs["simple10"]).pages:
            writer.add_page(page)
        bookmark = writer.add_outline_item("A at physical page 4", 3).get_object()
        bookmark["/A"]["/D"][1] = writer._add_object(NameObject("/Fit"))
        with path.open("xb") as stream:
            writer.write(stream)
        plan = engine.plan_level1(PdfReader(path), 10)
        self.assertEqual(plan["ranges"], [(0, 3), (3, 10)])
        self.assertEqual(plan["bookmarks"][0]["page"], 3)

    def test_real_unknown_fit_is_diagnosed_and_never_creates_a_fallback_page_zero_boundary(self):
        path = self.work / "unknown-fit.pdf"
        writer = PdfWriter()
        for page in PdfReader(self.inputs["simple10"]).pages:
            writer.add_page(page)
        invalid = writer.add_outline_item("Invalid fit", 3).get_object()
        invalid["/A"]["/D"][1] = NameObject("/UnknownFit")
        writer.add_outline_item("Valid B", 6)
        with path.open("xb") as stream:
            writer.write(stream)
        plan = engine.plan_level1(PdfReader(path), 10)
        self.assertEqual(plan["ranges"], [(0, 6), (6, 10)])
        self.assertIsNone(plan["bookmarks"][0]["page"])
        self.assertEqual(plan["bookmarks"][1]["page"], 6)
        self.assertTrue(any(warning["code"] == "invalid_destination" and warning["source_order"] == 1
                            for warning in plan["warnings"]))

    def test_raw_numeric_page_object_id_is_rejected_despite_pypdf_resolving_it(self):
        path = self.work / "numeric-page-object-id.pdf"
        writer = PdfWriter()
        for page in PdfReader(self.inputs["simple10"]).pages:
            writer.add_page(page)
        invalid = writer.add_outline_item("Numeric object ID masquerading as a page reference", 3).get_object()
        invalid["/A"]["/D"][0] = NumberObject(writer.pages[3].indirect_reference.idnum)
        writer.add_outline_item("Valid B", 6)
        with path.open("xb") as stream:
            writer.write(stream)
        before = path.read_bytes()
        reader = PdfReader(path)
        # Reproduce the dependency's permissive interpretation: a bare number
        # must not acquire the meaning of an internal page-object reference.
        self.assertEqual(reader.get_destination_page_number(reader.outline[0]), 3)
        plan = engine.plan_level1(reader, 10)
        self.assertEqual(plan["ranges"], [(0, 6), (6, 10)])
        self.assertIsNone(plan["bookmarks"][0]["page"])
        self.assertEqual(plan["bookmarks"][1]["page"], 6)
        self.assertTrue(any(warning["code"] == "invalid_destination" and warning["source_order"] == 1
                            and warning["message"] for warning in plan["warnings"]))
        output = self.work / "numeric-page-output"
        output.mkdir()
        with patch.object(engine, "log"):
            execution = engine.split_pdf(str(path), str(output), "1")
        observed = [fixtures.page_ids(file) for file in sorted(Path(execution["final_directory"]).glob("*.pdf"))]
        self.assertEqual(observed, [list(range(1, 7)), list(range(7, 11))])
        self.assertEqual(path.read_bytes(), before)


class RawTreePreflightTests(unittest.TestCase):
    class Reader:
        def __init__(self, catalog):
            self.root_object, self.pages = catalog, [None] * 10
            self.outline_reads = 0

        @property
        def outline(self):
            self.outline_reads += 1
            raise AssertionError("Unbounded pypdf outline parser must not be entered")

    def assert_rejected_before_outline_or_output(self, catalog):
        reader = self.Reader(catalog)
        with self.assertRaises(engine.BookmarkPlanError) as raised:
            engine.plan_level1(reader, 10)
        self.assertEqual(raised.exception.code, "invalid_outline")
        self.assertTrue(raised.exception.warnings)
        self.assertEqual(reader.outline_reads, 0)
        with patch.object(engine, "PdfReader", return_value=reader), \
                patch.object(engine, "write_slice") as writer, patch.object(engine, "log"):
            with self.assertRaises(SystemExit) as exited:
                engine.split_pdf("synthetic-raw-tree.pdf", "synthetic-unused-output", "1")
        self.assertNotEqual(exited.exception.code, 0)
        writer.assert_not_called()
        self.assertEqual(reader.outline_reads, 0)

    def test_cyclic_raw_outline_links_are_rejected_before_the_recursive_pdf_parser(self):
        for link in ("/Next", "/First"):
            with self.subTest(link=link):
                node = {"/Title": "Synthetic cyclic bookmark"}
                node[link] = node
                self.assert_rejected_before_outline_or_output({"/Outlines": {"/First": node}})

    def test_malformed_cyclic_or_reused_named_destination_trees_are_rejected_before_outline(self):
        cycle = {"/Kids": []}
        cycle["/Kids"].append(cycle)
        shared = {"/Names": ["A", [3, "/Fit"]]}
        cases = [{"/Kids": "invalid child list"}, {"/Kids": ["invalid child node"]},
                 {"/Names": ["A"]}, cycle, {"/Kids": [shared, shared]},
                 {"/Names": ["A", [3, "/Fit"], "A", [5, "/Fit"]]}]
        for number, tree in enumerate(cases):
            with self.subTest(case=number):
                self.assert_rejected_before_outline_or_output({"/Names": {"/Dests": tree}})

    def test_deep_or_over_budget_raw_name_trees_are_rejected_before_outline(self):
        deep = {"/Names": ["A", [3, "/Fit"]]}
        for unused in range(64):
            deep = {"/Kids": [deep]}
        self.assert_rejected_before_outline_or_output({"/Names": {"/Dests": deep}})
        # A reduced budget gives a quick regression for breadth, independently
        # of the production 10000-node normalization boundary exercised above.
        wide = {"/Kids": [{"/Names": [f"Name {number}", [3, "/Fit"]]} for number in range(5)]}
        with patch.object(engine, "MAX_OUTLINE_NODES", 4):
            self.assert_rejected_before_outline_or_output({"/Names": {"/Dests": wide}})

    def test_unreadable_raw_catalog_or_outline_root_is_rejected_before_outline(self):
        for catalog in ("invalid catalog", {"/Outlines": "invalid outline root"},
                        {"/Outlines": {"/First": "invalid outline child"}}, {"/Dests": "invalid names tree"}):
            with self.subTest(catalog=catalog):
                self.assert_rejected_before_outline_or_output(catalog)


if __name__ == "__main__":
    unittest.main(verbosity=2)
