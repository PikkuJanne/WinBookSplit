"""Corrected Level 1 evidence using original synthetic PDFs and bounded mock trees.

No shell probes or private input documents are used by this route.
"""

from __future__ import annotations

import argparse
from contextlib import redirect_stdout
from datetime import datetime, timezone
from io import StringIO
import json
import os
from pathlib import Path
import platform
import random
import subprocess
import sys
import tempfile
from unittest.mock import patch

from pypdf import PdfReader, PdfWriter
from pypdf.generic import ArrayObject, DictionaryObject, NameObject, NumberObject, TextStringObject

ROOT = Path(__file__).resolve().parents[2]
ENGINE = ROOT / "engine/winbooksplit_engine.py"
ACCEPTANCE_IDS = [f"AC-{number:03}" for number in range(19, 23)]
ORACLE_IDS = {"BM-01", "BM-02", "BM-05", "BM-08"}

# The existing helper uses trusted absolute imports, argv children, immutable
# Git guards and page identity checks. Keep those historical sources intact.
import importlib.util

spec = importlib.util.spec_from_file_location("wbs_level1_manual_helpers", ROOT / "tests/manual/characterize_manual.py")
manual = importlib.util.module_from_spec(spec)
spec.loader.exec_module(manual)
history, require = manual.history, manual.require


class Destination(dict):
    def __init__(self, title, page, raw=None):
        super().__init__({"/Title": title})
        self.title, self.page, self.node = title, page, raw


class SyntheticReader:
    def __init__(self, outline):
        self.outline = outline
        self.resolved = []

    def get_destination_page_number(self, node):
        self.resolved.append(node)
        if isinstance(node.page, Exception):
            raise node.page
        return node.page


class PreflightReader:
    """Verify raw outline/name trees are bounded before pypdf outline retrieval."""
    def __init__(self, catalog):
        self.root_object, self.outline_read = catalog, False

    @property
    def outline(self):
        self.outline_read = True
        raise AssertionError("Malformed raw tree reached recursive outline retrieval")


def partition(plan, pages, expected=None):
    ranges = plan["ranges"]
    require(bool(ranges) and ranges[0][0] == 0 and ranges[-1][1] == pages
            and all(0 <= start < end <= pages for start, end in ranges)
            and all(left[1] == right[0] for left, right in zip(ranges, ranges[1:])), "Invalid Level 1 partition")
    require([page for start, end in ranges for page in range(start, end)] == list(range(pages)),
            "Level 1 plan omitted/repeated/reordered a physical page")
    require([(entry["start"], entry["end"]) for entry in plan["entries"]] == ranges,
            "Writer entries differ from the validated ranges")
    require([entry["sequence"] for entry in plan["entries"]] == list(range(1, len(ranges) + 1)),
            "Invalid entry sequence")
    if expected is not None:
        require([list(pair) for pair in ranges] == expected, "Unexpected Level 1 ranges")


def rejected(engine, reader, pages, code="invalid_outline", warning=None):
    reader.pages = [None] * pages  # The split entry point reads this before planning.
    try:
        engine.plan_level1(reader, pages)
    except engine.BookmarkPlanError as error:
        require(error.code == code and str(error), "Rejected outline lost its structured diagnostic")
        if warning is not None:
            require(any(item["code"] == warning and item["message"] for item in error.warnings),
                    "Rejected outline lost its warning")
        with patch.object(engine, "PdfReader", return_value=reader), patch.object(engine, "write_slice") as writer:
            with redirect_stdout(StringIO()) as captured:
                try:
                    engine.split_pdf("synthetic-unused-input.pdf", "synthetic-unused-output", "1")
                except SystemExit as exit_error:
                    require(exit_error.code == (55 if code in {"no_bookmarks", "no_usable_bookmarks"} else 1),
                            "Rejected outline returned the wrong process status")
                else:
                    raise RuntimeError("Rejected outline unexpectedly succeeded")
            require(not writer.called and code in captured.getvalue(), "Rejected outline reached the writer or lost error")
        if isinstance(reader, PreflightReader):
            require(not reader.outline_read, "Raw preflight did not precede outline retrieval")
        return {"code": error.code, "warnings": error.warnings, "writer_not_reached": True}
    raise RuntimeError("Invalid outline was accepted")


def nested_outline(depth):
    result = [Destination(f"Depth {depth}", 0)]
    for number in range(depth - 1, 0, -1):
        result = [Destination(f"Depth {number}", 0), result]
    return result


def fixture(path, generator, pages, parents):
    def fixture_nodes(nodes):
        return [{"title": node["title"], "start_page": node["start_page"],
                 "children": fixture_nodes(node.get("children", []))} for node in nodes]
    with path.open("xb") as stream:
        stream.write(generator._pdf_bytes({"title": "Original Level 1 synthetic fixture",
                                           "pages": pages, "outline": fixture_nodes(parents)}))
    require(generator.page_ids(path) == list(range(1, pages + 1)), "Generated input page identities changed")
    return path


def cli_case(source, cwd, generator, *, expected=None, error=None, warnings=()):
    cwd.mkdir()
    output = cwd / "output"
    output.mkdir()
    neighbor = output / "synthetic-neighbor.txt"
    neighbor.write_bytes(b"Preserve this Level 1 synthetic neighbor\n")
    before = history.file_digest(source)
    result = manual.observe(ENGINE, source, output, cwd, "1")
    records = manual.published_outputs(output, generator, result)
    if error is not None:
        require(result["exit_code"] == (55 if error in {"no_bookmarks", "no_usable_bookmarks"} else 1)
                and error in result["stdout"] and not records and "[Writing]" not in result["stdout"],
                "Rejected CLI outline wrote output or returned an incorrect error")
    else:
        require(result["exit_code"] == 0 and not result["stderr"], "Valid Level 1 CLI run failed: " + repr(result))
        manual.check_outputs(records, expected, len(PdfReader(source).pages))
    require(all(code in result["stdout"] for code in warnings), "Missing CLI bookmark warnings")
    require(history.file_digest(source) == before, "CLI modified its input")
    require(neighbor.read_bytes() == b"Preserve this Level 1 synthetic neighbor\n", "CLI modified its neighbor")
    manual.check_base_members(output, result, [neighbor.name])
    return {"outputs": records, "input_unchanged": True, "neighbor_unchanged": True, **result}


def raw_destination_pdf(source, path):
    """Original marked pages with named/raw/external destination variants."""
    writer = PdfWriter()
    for page in PdfReader(source).pages:
        writer.add_page(page)
    page3, page6 = writer.pages[3].indirect_reference, writer.pages[6].indirect_reference
    valid_named = DictionaryObject({NameObject("/D"): ArrayObject([page3, NameObject("/Fit")])})
    invalid_named = DictionaryObject({NameObject("/D"): ArrayObject([page3, NameObject("/BadFit")])})
    writer._root_object[NameObject("/Names")] = DictionaryObject({NameObject("/Dests"): DictionaryObject({
        NameObject("/Names"): ArrayObject([TextStringObject("ValidNamed"), valid_named,
                                           TextStringObject("InvalidNamed"), invalid_named])})})
    variants = [
        ("Valid named", {"/Dest": TextStringObject("ValidNamed")}),
        ("External", {"/A": DictionaryObject({NameObject("/S"): NameObject("/URI"),
                                                 NameObject("/URI"): TextStringObject("https://example.invalid/")})}),
        ("Missing named", {"/Dest": TextStringObject("MissingNamed")}),
        ("Invalid named", {"/Dest": TextStringObject("InvalidNamed")}),
        ("Invalid raw fit", {"/Dest": ArrayObject([page3, NameObject("/BadFit")])}),
        ("Invalid raw integer", {"/Dest": ArrayObject([NumberObject(page3.idnum), NameObject("/Fit")])}),
        ("Valid raw", {"/Dest": ArrayObject([page6, NameObject("/Fit")])}),
    ]
    for title, replacement in variants:
        node = writer.add_outline_item(title, 0).get_object()
        del node["/A"]
        for key, value in replacement.items():
            node[NameObject(key)] = value
    with path.open("xb") as stream:
        writer.write(stream)
    return path


def cyclic_pdf(source, path):
    writer = PdfWriter()
    for page in PdfReader(source).pages:
        writer.add_page(page)
    root = DictionaryObject({NameObject("/Type"): NameObject("/Outlines")})
    root_ref = writer._add_object(root)
    node = DictionaryObject({NameObject("/Title"): TextStringObject("Original cycle test"),
                             NameObject("/Parent"): root_ref,
                             NameObject("/Dest"): ArrayObject([writer.pages[0].indirect_reference, NameObject("/Fit")])})
    node_ref = writer._add_object(node)
    node[NameObject("/Next")] = node_ref
    root[NameObject("/First")] = node_ref
    root[NameObject("/Last")] = node_ref
    writer._root_object[NameObject("/Outlines")] = root_ref
    with path.open("xb") as stream:
        writer.write(stream)
    return path


def normalization_cases(engine, work, fixtures, generator):
    results = []
    bad_values = [None, -1, -2, 10, True, False, 1.0, 1.5, float("nan"), float("inf"), "3",
                  ValueError("Original synthetic resolution failure")]
    invalid_nodes = [Destination(f"Invalid {number}", value) for number, value in enumerate(bad_values)]
    external_nodes = [Destination(kind, 5, {"/A": {"/S": kind}})
                      for kind in ("/URI", "/GoToR", "/Launch", "/JavaScript")]
    raw_nodes = [Destination(f"Invalid raw {number}", 3, raw) for number, raw in enumerate(
        ({}, {"/Dest": []}, {"/Dest": [3]}, {"/Dest": [3, "/BadFit"]},
         {"/A": {"/S": "/GoTo", "/D": [3, "/BadFit"]}}, {"/A": "malformed"}))]
    reader = SyntheticReader([Destination("A", 3), *invalid_nodes, *external_nodes, *raw_nodes, Destination("B", 6)])
    plan = engine.plan_level1(reader, 10)
    partition(plan, 10, [[0, 3], [3, 6], [6, 10]])
    unusable = plan["bookmarks"][1:-1]
    require(all(item["page"] is None for item in unusable), "Invalid destination became a split boundary")
    require(all(any(warning["source_order"] == item["source_order"] and warning["message"]
                    for warning in plan["warnings"]) for item in unusable), "Invalid destination lacked a diagnostic")
    require(all(all(resolved is not node for resolved in reader.resolved) for node in [*external_nodes, *raw_nodes]),
            "External or malformed raw destination reached internal-page resolution")
    no_usable = rejected(engine, SyntheticReader([Destination("Invalid", None), [Destination("Child", 0)]]),
                         10, "no_usable_bookmarks")
    zero = rejected(engine, SyntheticReader([Destination("A", 0)]), 0, "invalid_document")
    raw_path = raw_destination_pdf(fixtures["simple10"], work / "raw-named-destinations.pdf")
    raw_cli = cli_case(raw_path, work / "raw-named-cli", generator, expected=[[0, 3], [3, 6], [6, 10]],
                       warnings=("external_destination", "invalid_destination"))
    actual_plan = engine.plan_level1(PdfReader(raw_path), 10)
    require([entry["title"] for entry in actual_plan["entries"]] == ["Front matter", "Valid named", "Valid raw"],
            "Valid raw/named destinations or malformed aliases changed the selected titles")
    raw_integer = next(node for node in actual_plan["bookmarks"] if node["title"] == "Invalid raw integer")
    require(raw_integer["page"] is None and any(warning["source_order"] == raw_integer["source_order"]
            and warning["code"] == "invalid_destination" for warning in actual_plan["warnings"]),
            "Raw integer colliding with a real page object ID was accepted")
    results.append({"id": "invalid-destinations", "passed": True, "warnings": plan["warnings"],
                    "invalid_resolution_count": len(bad_values), "external_action_count": len(external_nodes),
                    "invalid_raw_count": len(raw_nodes), "no_usable": no_usable, "zero_pages": zero, "raw_named_cli": raw_cli})

    outline = [Destination("Parent", 3), [Destination("Child", 4), [Destination("Grandchild", 5)]],
               Destination("Second", 7)]
    deep = engine.plan_level1(SyntheticReader(outline), 10)
    partition(deep, 10, [[0, 3], [3, 7], [7, 10]])
    nodes = deep["bookmarks"]
    require([node["depth"] for node in nodes] == [1, 2, 3, 1]
            and [node["source_order"] for node in nodes] == [1, 2, 3, 4]
            and nodes[1]["lineage"] == [nodes[0]["id"]]
            and nodes[2]["lineage"] == [nodes[0]["id"], nodes[1]["id"]]
            and nodes[2]["parent_id"] == nodes[1]["id"]
            and not nodes[3]["lineage"], "Depth/lineage/source order was lost")
    at_limit = engine.plan_level1(SyntheticReader(nested_outline(64)), 10)
    partition(at_limit, 10, [[0, 10]])
    require([node["depth"] for node in at_limit["bookmarks"]] == list(range(1, 65)), "Depth 64 was not retained")
    results.append({"id": "deep-lineage", "passed": True, "bookmarks": nodes, "valid_depth_limit": 64})

    cycle = [Destination("A", 0)]
    cycle.append(cycle)
    shared_child = [Destination("Shared", 1)]
    shared_node = Destination("Shared object", 0)
    cycles = [cycle, [Destination("A", 0), shared_child, Destination("B", 5), shared_child],
              [shared_node, shared_node], [shared_node, [shared_node]]]
    cycle_errors = [rejected(engine, SyntheticReader(item), 10, warning="outline_cycle") for item in cycles]
    raw_cycle = {}
    raw_cycle["/Next"] = raw_cycle
    cycle_errors.append(rejected(engine, PreflightReader({"/Outlines": {"/First": raw_cycle}}), 10, warning="outline_cycle"))
    names_cycle = {}
    names_cycle["/Kids"] = [names_cycle]
    cycle_errors.append(rejected(engine, PreflightReader({"/Names": {"/Dests": names_cycle}}), 10, warning="outline_cycle"))
    cycle_path = cyclic_pdf(fixtures["simple10"], work / "cyclic-outline.pdf")
    cycle_cli = cli_case(cycle_path, work / "cyclic-outline-cli", generator, error="invalid_outline", warnings=("outline_cycle",))
    results.append({"id": "cyclic-outline", "passed": True, "rejections": cycle_errors, "serialized_pdf_cli": cycle_cli})

    malformed = [rejected(engine, SyntheticReader({"unexpected": "container"}), 10, warning="malformed_outline"),
                 rejected(engine, SyntheticReader([[Destination("Orphan", 0)]]), 10, warning="malformed_outline")]
    for catalog in ([], {"/Outlines": "bad root"}, {"/Outlines": {"/First": "bad node"}},
                    {"/Names": {"/Dests": {"/Names": ["unpaired"]}}},
                    {"/Names": {"/Dests": {"/Kids": "bad children"}}},
                    {"/Names": {"/Dests": {"/Names": ["duplicate", [0, "/Fit"], "duplicate", [1, "/Fit"]]}}}):
        malformed.append(rejected(engine, PreflightReader(catalog), 10, warning="malformed_outline"))
    scalars = engine.plan_level1(SyntheticReader([Destination("A", 0), None, "unsupported scalar"]), 10)
    partition(scalars, 10, [[0, 10]])
    require(len(scalars["warnings"]) == 2, "Malformed scalar entries were silently ignored")
    results.append({"id": "malformed-outline", "passed": True, "rejections": malformed, "scalar_warnings": scalars["warnings"]})

    require(engine.MAX_OUTLINE_DEPTH == 64 and engine.MAX_OUTLINE_NODES == 10000, "Unexpected traversal limits")
    limits = [rejected(engine, SyntheticReader(nested_outline(65)), 10, warning="outline_limit"),
              rejected(engine, SyntheticReader([Destination(f"Node {number}", 0) for number in range(10001)]),
                       10, warning="outline_limit")]
    normalized = engine.normalize_outline(SyntheticReader([]), [Destination(f"Node {number}", 0) for number in range(10000)], 10)
    require(len(normalized["bookmarks"]) == 10000, "Valid 10000-node bound was rejected")
    first = {}
    cursor = first
    for _ in range(64):
        cursor["/First"] = {}
        cursor = cursor["/First"]
    limits.append(rejected(engine, PreflightReader({"/Outlines": {"/First": first}}), 10, warning="outline_limit"))
    named = {}
    cursor = named
    for _ in range(64):
        child = {}
        cursor["/Kids"] = [child]
        cursor = child
    limits.append(rejected(engine, PreflightReader({"/Names": {"/Dests": named}}), 10, warning="outline_limit"))
    excessive_pairs = {"/Names": [value for number in range(10001) for value in (f"Name {number}", [0, "/Fit"])]}
    limits.append(rejected(engine, PreflightReader({"/Names": {"/Dests": excessive_pairs}}), 10, warning="outline_limit"))
    results.append({"id": "traversal-limits", "passed": True, "rejections": limits, "valid_node_limit": 10000})
    return results, [raw_path, cycle_path]


def characterize(work):
    sources = history.original_sources()
    paths = (*history.BASELINE_PATHS, "engine/winbooksplit_engine.py", "tests/extraction/characterize_extraction.py",
             "tests/manual/characterize_manual.py", "tests/bookmarks/characterize_level1.py", "tests/bookmarks/README.md",
             "tests/run_tests.py")
    before = {path: history.file_digest(ROOT / path) for path in paths}
    engine = manual.load_module("wbs_level1_current_engine", ENGINE)
    generator = history.load_generator()
    fixtures = generator.generate_fixtures(work / "fixtures")
    input_paths = list(fixtures.values())
    cases = []
    for oracle in json.loads(sources["docs/codex-v1.0.0/PLAN_ORACLES.json"])["bookmarks"]:
        if oracle["id"] not in ORACLE_IDS:
            continue
        source = fixture(work / (oracle["id"] + ".pdf"), generator, oracle["pages"], oracle["parents"])
        input_paths.append(source)
        result = cli_case(source, work / (oracle["id"] + "-cli"), generator,
                          expected=oracle.get("expected_ranges"), error=oracle.get("expected_error"),
                          warnings=oracle.get("expected_warnings", []))
        if "expected_ranges" in oracle:
            plan = engine.plan_level1(PdfReader(source), oracle["pages"])
            partition(plan, oracle["pages"], oracle["expected_ranges"])
            if oracle["id"] == "BM-08":
                require([entry["title"] for entry in plan["entries"]] == ["Front matter", "A", "B"],
                        "Duplicate destination did not preserve the first source title")
            require([record["filename"] for record in result["outputs"]] == [entry["filename"] for entry in plan["entries"]],
                    "CLI writer did not use the planned filenames")
        cases.append({"oracle_id": oracle["id"], "passed": True, "expected_ranges": oracle.get("expected_ranges"),
                      "expected_error": oracle.get("expected_error"), **result})
    input_hashes = {str(path): history.file_digest(path) for path in input_paths}
    normalization, extra_inputs = normalization_cases(engine, work, fixtures, generator)
    input_paths.extend(extra_inputs)
    input_hashes.update({str(path): history.file_digest(path) for path in extra_inputs})

    seed, count = 20261009, 150
    rng = random.Random(seed)
    samples = []
    for index in range(count):
        pages = rng.randint(1, 120)
        starts = [rng.randrange(pages) for _ in range(rng.randint(1, 25))]
        names = [f"Node {number}" for number in range(len(starts))]
        plan = engine.plan_level1(SyntheticReader([Destination(name, page) for name, page in zip(names, starts)]), pages)
        boundaries = sorted({0, *starts}) + [pages]
        expected = [list(pair) for pair in zip(boundaries, boundaries[1:])]
        partition(plan, pages, expected)
        expected_titles = (["Front matter"] if min(starts) else []) + [names[starts.index(page)] for page in sorted(set(starts))]
        require([entry["title"] for entry in plan["entries"]] == expected_titles, "Seeded duplicate-title selection changed")
        if index < 10:
            starts10 = [page % 10 for page in starts]
            parents = [{"title": name, "start_page": page + 1, "children": []} for name, page in zip(names, starts10)]
            source = fixture(work / f"writer-{index}.pdf", generator, 10, parents)
            source_hash = history.file_digest(source)
            plan10 = engine.plan_level1(PdfReader(source), 10)
            output = work / f"writer-output-{index}"
            output.mkdir()
            neighbor = output / "synthetic-neighbor.txt"
            neighbor.write_bytes(b"Original writer neighbor\n")
            with redirect_stdout(StringIO()):
                result = engine.run_split(str(source), str(output), "1")
            records = manual.published_outputs(output, generator, result)
            manual.check_base_members(output, result, [neighbor.name])
            manual.check_outputs(records, [list(pair) for pair in plan10["ranges"]], 10)
            require(history.file_digest(source) == source_hash and neighbor.read_bytes() == b"Original writer neighbor\n",
                    "Sampled writer modified source or neighbor")
            input_paths.append(source)
            input_hashes[str(source)] = source_hash
            samples.append({"index": index, "starts": starts10, "outputs": records})

    import_observation = history.probe_import(work)
    guard_path = work / "baseline-must-refuse.json"
    guard = history.run([sys.executable, "-I", "-B", str(ROOT / "tests/baseline/characterize_original.py"),
                         "--report", str(guard_path)], work, environment=history.clean_environment(work))
    require(guard["exit_code"] == 1 and "Original launchers changed" in guard["stderr"] and not guard_path.exists(),
            "Immutable historical baseline guard was weakened")
    require(input_hashes == {str(path): history.file_digest(path) for path in input_paths}, "Synthetic inputs changed")
    require(before == {path: history.file_digest(ROOT / path) for path in paths}, "Tested source changed during Level 1 run")
    return {"schema_version": 1, "task_id": "M1-T03", "observed_at": datetime.now(timezone.utc).isoformat(),
            "result": "LEVEL1_REGRESSION_PASSED", "success": True, "exit_code": 0, "acceptance_ids": ACCEPTANCE_IDS,
            "engine_case_count": len(cases), "engine_cases": cases, "normalization_cases": normalization,
            "seeded_level1_cases": {"seed": seed, "count": count, "passed": True}, "sampled_writer_cases": samples,
            "source_unchanged": True, "baseline_guards_preserved": True, "input_and_neighbor_unchanged": True,
            "immutable_original_commit": history.ORIGINAL_COMMIT, "historical_original_guard": guard,
            "tested_path_sha256": before, "input_sha256": input_hashes, "import_observation": import_observation,
            "environment": {"python": sys.version, "python_executable": sys.executable, "platform": platform.platform(),
                            "pypdf": history.pypdf.__version__},
            "not_run": ["corrected Level 2", "actual Level 1 launcher paths", "Explorer", "Calibre conversion",
                        "rendered fidelity", "output transactions", "release-package acceptance"]}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--report", type=Path, required=True)
    args = parser.parse_args()
    try:
        require(sys.flags.isolated and sys.dont_write_bytecode, "Use explicit developer Python -I -B")
        require(args.report.is_absolute() and not os.path.lexists(args.report), "New absolute report required")
        report_path = args.report.resolve()
        require(not report_path.is_relative_to(ROOT) and report_path.parent.is_dir(), "External report required")
        with tempfile.TemporaryDirectory(prefix="WinBookSplit-M1-T03-tests-") as directory:
            report = characterize(Path(directory).resolve())
        report["owned_temp_removed"] = not Path(directory).exists()
        require(report["owned_temp_removed"], "Owned temporary cleanup failed")
        with report_path.open("x", encoding="utf-8", newline="\n") as stream:
            json.dump(report, stream, indent=2, ensure_ascii=True)
            stream.write("\n")
        print("Level 1 acceptance passed: four CLI oracles, five normalization aggregates, 150 seeded plans, ten writer samples")
        return 0
    except (OSError, RuntimeError, ValueError, subprocess.TimeoutExpired) as error:
        print("Level 1 acceptance failed: " + str(error), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
