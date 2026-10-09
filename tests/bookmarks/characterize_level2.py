"""Corrected parent-aware Level 2 evidence using original synthetic PDFs.

Actual CLI/writer observations are distinct from callable hierarchy mock cases.
This route has no shell probes and reads no private input document.
"""

from __future__ import annotations

import argparse
from contextlib import redirect_stdout
from datetime import datetime, timezone
import importlib.util
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
ACCEPTANCE_IDS = [f"AC-{number:03}" for number in range(23, 27)]
ORACLE_IDS = {"BM-03", "BM-04", "BM-06", "BM-07"}
spec = importlib.util.spec_from_file_location("wbs_level2_level1_helpers", ROOT / "tests/bookmarks/characterize_level1.py")
level1 = importlib.util.module_from_spec(spec)
spec.loader.exec_module(level1)
manual, history, require = level1.manual, level1.history, level1.require
Destination, SyntheticReader = level1.Destination, level1.SyntheticReader


def oracle_outline(parents):
    result = []
    for parent in parents:
        result.append(Destination(parent["title"], parent["start_page"] - 1))
        children = oracle_outline(parent.get("children", []))
        if children:
            result.append(children)
    return result


def check_plan(plan, pages, expected=None, titles=None):
    level1.partition(plan, pages, expected)
    records = {record["id"]: record for record in plan["bookmarks"]}
    parents = sorted((record for record in records.values()
                      if record["depth"] == 1 and record["page"] is not None and "alias_of" not in record),
                     key=lambda record: record["page"])
    intervals = {parent["id"]: (parent["page"], parents[index + 1]["page"] if index + 1 < len(parents) else pages)
                 for index, parent in enumerate(parents)}
    for entry in plan["entries"]:
        if entry["reason"] == "front_matter":
            require(entry["parent_id"] is None and entry["start"] == 0, "Front matter acquired a parent")
            continue
        require(entry["parent_id"] in intervals, "Entry belongs to an unusable or duplicate parent")
        start, end = intervals[entry["parent_id"]]
        require(start <= entry["start"] < entry["end"] <= end, "Child/opening/fallback crossed its parent interval")
        if entry["reason"] == "bookmark":
            child = records[entry["bookmark_id"]]
            require(child["depth"] == 2 and child["parent_id"] == entry["parent_id"]
                    and child["page"] == entry["start"], "Selected child lost its direct parent/identity")
        else:
            require(entry["reason"] in {"parent_opening", "parent_fallback"}, "Unknown hierarchy segment reason")
    if titles is not None:
        require([entry["title"] for entry in plan["entries"]] == titles, "Incorrect opening/fallback/child titles")


def rejected(engine, reader, pages, code, warning=None):
    reader.pages = [None] * pages
    try:
        engine.plan_level2(reader, pages)
    except engine.BookmarkPlanError as error:
        require(error.code == code and str(error), "Rejected Level 2 outline lost its diagnostic")
        if warning is not None:
            require(any(item["code"] == warning and item["message"] for item in error.warnings), "Missing rejection warning")
        with patch.object(engine, "PdfReader", return_value=reader), patch.object(engine, "write_slice") as writer:
            with redirect_stdout(StringIO()) as captured:
                try:
                    engine.split_pdf("synthetic-unused-input.pdf", "synthetic-unused-output", "2")
                except SystemExit as status:
                    expected = 55 if code in {"no_bookmarks", "no_usable_bookmarks", "no_bookmarks_at_level"} else 1
                    require(status.code == expected, "Rejected Level 2 call returned the wrong status")
                    if expected == 55:
                        require("[NO_BOOKMARKS_FOUND]" in captured.getvalue(), "No-plan compatibility sentinel missing")
                else:
                    raise RuntimeError("Rejected Level 2 call succeeded")
            require(not writer.called and code in captured.getvalue(), "Rejected Level 2 call wrote or lost its error")
        if isinstance(reader, level1.PreflightReader):
            require(not reader.outline_read, "Malformed raw tree reached recursive pypdf retrieval")
        return {"code": error.code, "warnings": error.warnings, "writer_not_reached": True}
    raise RuntimeError("Rejected Level 2 outline was accepted")


def cli_case(source, cwd, generator, *, expected=None, error=None, warnings=()):
    cwd.mkdir()
    output = cwd / "output"
    output.mkdir()
    neighbor = output / "synthetic-neighbor.txt"
    neighbor.write_bytes(b"Original Level 2 synthetic neighbor\n")
    before = history.file_digest(source)
    result = manual.observe(ENGINE, source, output, cwd, "2")
    records = history.outputs(output, generator)
    if error is not None:
        status = 55 if error in {"no_bookmarks", "no_usable_bookmarks", "no_bookmarks_at_level"} else 1
        require(result["exit_code"] == status and error in result["stdout"] and not records
                and "[Writing]" not in result["stdout"], "Rejected Level 2 CLI call wrote or returned the wrong error")
        if status == 55:
            require("[NO_BOOKMARKS_FOUND]" in result["stdout"], "No-plan CLI sentinel missing")
    else:
        require(result["exit_code"] == 0 and not result["stderr"], "Valid Level 2 CLI failed: " + repr(result))
        manual.check_outputs(records, expected, len(PdfReader(source).pages))
    require(all(code in result["stdout"] for code in warnings), "Missing CLI hierarchy warning")
    require(history.file_digest(source) == before and neighbor.read_bytes() == b"Original Level 2 synthetic neighbor\n",
            "Level 2 CLI modified input/neighbor")
    require({path.name for path in output.iterdir()} == {neighbor.name, *(record["filename"] for record in records)},
            "Level 2 CLI wrote unexpected files")
    return {"outputs": records, "input_unchanged": True, "neighbor_unchanged": True, **result}


def raw_child_pdf(source, path):
    writer = PdfWriter()
    for page in PdfReader(source).pages:
        writer.add_page(page)
    parent = writer.add_outline_item("Parent", 0)
    writer.add_outline_item("Valid child", 2, parent=parent)
    page3, page8 = writer.pages[3].indirect_reference, writer.pages[8].indirect_reference
    writer._root_object[NameObject("/Names")] = DictionaryObject({NameObject("/Dests"): DictionaryObject({
        NameObject("/Names"): ArrayObject([TextStringObject("ValidNamed"), DictionaryObject({
            NameObject("/D"): ArrayObject([page3, NameObject("/Fit")])})])})})
    variants = [
        ("Valid named child", {"/Dest": TextStringObject("ValidNamed")}),
        ("Raw integer child", {"/Dest": ArrayObject([NumberObject(page8.idnum), NameObject("/Fit")])}),
        ("Invalid fit child", {"/Dest": ArrayObject([page8, NameObject("/BadFit")])}),
        ("Unresolved named child", {"/Dest": TextStringObject("MissingNamed")}),
        ("External child", {"/A": DictionaryObject({NameObject("/S"): NameObject("/URI"),
                NameObject("/URI"): TextStringObject("https://example.invalid/")})}),
    ]
    for title, replacement in variants:
        node = writer.add_outline_item(title, 0, parent=parent).get_object()
        del node["/A"]
        for key, value in replacement.items():
            node[NameObject(key)] = value
    with path.open("xb") as stream:
        writer.write(stream)
    return path


def hierarchy_cases(engine, work, fixtures, generator):
    cases = []
    duplicate_outline = [Destination("Canonical", 0), [Destination("Own", 2)],
                         Destination("Alias", 0), [Destination("Alias child", 5)],
                         Destination("Next", 6), [Destination("Next child", 7)]]
    duplicate = engine.plan_level2(SyntheticReader(duplicate_outline), 10)
    check_plan(duplicate, 10, [[0, 2], [2, 6], [6, 7], [7, 10]],
               ["Canonical - Opening pages", "Own", "Next - Opening pages", "Next child"])
    require(any(item["code"] == "duplicate_parent_subtree" for item in duplicate["warnings"]), "Alias subtree warning missing")
    alias_only = rejected(engine, SyntheticReader([Destination("Canonical", 0), Destination("Alias", 0),
                           [Destination("Alias-only child", 2)]]), 10, "no_bookmarks_at_level", "duplicate_parent_subtree")
    cases.append({"id": "duplicate-parent-subtrees", "passed": True, "plan": duplicate, "alias_only_rejection": alias_only})

    invalid_outline = [Destination("Invalid", None), [Destination("Orphaned child", 0)],
                       Destination("Valid", 0), [Destination("Own", 2)]]
    invalid = engine.plan_level2(SyntheticReader(invalid_outline), 10)
    check_plan(invalid, 10, [[0, 2], [2, 10]], ["Valid - Opening pages", "Own"])
    require(any(item["code"] == "invalid_parent_subtree" for item in invalid["warnings"]), "Invalid subtree warning missing")
    invalid_only = rejected(engine, SyntheticReader([Destination("Invalid", None), [Destination("Child", 0)],
                                           Destination("Valid without children", 0)]),
                            10, "no_bookmarks_at_level", "invalid_parent_subtree")
    cases.append({"id": "invalid-parent-subtrees", "passed": True, "plan": invalid, "invalid_only_rejection": invalid_only})

    child_outline = [Destination("A", 0), [Destination("Last", 5), Destination("At parent", 0),
                     Destination("First duplicate", 3), Destination("Alias child", 3), Destination("Outside", 6)],
                     Destination("B", 6)]
    ordered = engine.plan_level2(SyntheticReader(child_outline), 10)
    check_plan(ordered, 10, [[0, 3], [3, 5], [5, 6], [6, 10]], ["At parent", "First duplicate", "Last", "B"])
    require({"duplicate_destination", "outline_reordered", "child_outside_parent"}
            <= {item["code"] for item in ordered["warnings"]}, "Child alias/order/outside warning missing")
    require(ordered["entries"][0]["reason"] == "bookmark" and ordered["entries"][-1]["reason"] == "parent_fallback",
            "Child at parent start created an empty opening or fallback disappeared")
    cases.append({"id": "child-order-and-aliases", "passed": True, "plan": ordered})

    bad = [Destination(f"Invalid {number}", page) for number, page in enumerate(
        (None, -1, 10, True, False, 1.0, 1.5, float("nan"), "3", ValueError("Original child resolution failure")))]
    bad.append(Destination("External", 4, {"/A": {"/S": "/URI"}}))
    reader = SyntheticReader([Destination("Parent", 0), [*bad, Destination("Usable", 2)], Destination("Next", 6)])
    child_invalid = engine.plan_level2(reader, 10)
    check_plan(child_invalid, 10, [[0, 2], [2, 6], [6, 10]], ["Parent - Opening pages", "Usable", "Next"])
    rejected_children = [node for node in child_invalid["bookmarks"] if node["depth"] == 2 and node["title"] != "Usable"]
    require(all(node["page"] is None and any(warning["source_order"] == node["source_order"] and warning["message"]
                                            for warning in child_invalid["warnings"]) for node in rejected_children),
            "Invalid child became usable or lost its warning")
    raw_path = raw_child_pdf(fixtures["simple10"], work / "raw-child-destinations.pdf")
    raw_cli = cli_case(raw_path, work / "raw-child-cli", generator, expected=[[0, 2], [2, 3], [3, 10]],
                       warnings=("invalid_destination", "external_destination"))
    raw_plan = engine.plan_level2(PdfReader(raw_path), 10)
    check_plan(raw_plan, 10, [[0, 2], [2, 3], [3, 10]], ["Parent - Opening pages", "Valid child", "Valid named child"])
    cases.append({"id": "invalid-child-destinations", "passed": True, "plan": child_invalid, "actual_raw_child_cli": raw_cli})

    deep_outline = [Destination("A", 3), [Destination("A1", 4), [Destination("Ignored grandchild", 5)]],
                    Destination("B", 7), [Destination("B1", 8)]]
    deep = engine.plan_level2(SyntheticReader(deep_outline), 10)
    check_plan(deep, 10, [[0, 3], [3, 4], [4, 7], [7, 8], [8, 10]],
               ["Front matter", "A - Opening pages", "A1", "B - Opening pages", "B1"])
    require(any(node["depth"] == 3 and len(node["lineage"]) == 2 for node in deep["bookmarks"]), "Deep lineage was dropped")
    at_limit = engine.plan_level2(SyntheticReader(level1.nested_outline(64)), 10)
    check_plan(at_limit, 10, [[0, 10]])
    cycle = [Destination("A", 0)]
    cycle.append(cycle)
    failures = [rejected(engine, SyntheticReader(cycle), 10, "invalid_outline", "outline_cycle"),
                rejected(engine, SyntheticReader([[Destination("Orphan", 0)]]), 10, "invalid_outline", "malformed_outline"),
                rejected(engine, SyntheticReader(level1.nested_outline(65)), 10, "invalid_outline", "outline_limit")]
    raw_cycle_path = level1.cyclic_pdf(fixtures["simple10"], work / "level2-cyclic.pdf")
    cycle_cli = cli_case(raw_cycle_path, work / "level2-cyclic-cli", generator, error="invalid_outline", warnings=("outline_cycle",))
    cases.append({"id": "deep-and-malformed-outlines", "passed": True, "plan": deep,
                  "accepted_depth": 64, "rejections": failures, "serialized_cycle_cli": cycle_cli})

    no_level2 = [rejected(engine, SyntheticReader([]), 10, "no_bookmarks"),
                rejected(engine, SyntheticReader([Destination("Invalid parent", None), [Destination("Child", 0)]]),
                         10, "no_usable_bookmarks"),
                rejected(engine, SyntheticReader([Destination("Parent", 0)]), 10, "no_bookmarks_at_level"),
                rejected(engine, SyntheticReader([Destination("Parent", 0), [Destination("Invalid child", None),
                                  [Destination("Grandchild cannot replace child", 0)]]]), 10, "no_bookmarks_at_level"),
                rejected(engine, SyntheticReader([Destination("Parent", 0), [Destination("Outside", 6)],
                                                    Destination("Next", 6)]), 10, "no_bookmarks_at_level", "child_outside_parent"),
                rejected(engine, SyntheticReader([Destination("Parent", 0), [Destination("Child", 0)]]), 0, "invalid_document")]
    cases.append({"id": "no-usable-level2", "passed": True, "rejections": no_level2})
    return cases, [raw_path, raw_cycle_path]


def seeded_definition(rng, pages):
    parent_starts = sorted(rng.sample(range(pages), rng.randint(1, min(6, pages))))
    parents, expected = [], []
    if parent_starts[0]:
        expected.append((0, parent_starts[0], "Front matter"))
    for index, start in enumerate(parent_starts):
        end = parent_starts[index + 1] if index + 1 < len(parent_starts) else pages
        child_pages = [rng.randrange(start, end) for _ in range(rng.randint(0, 5))]
        if index == 0 and not child_pages:  # Every accepted seeded case has a real selected-level boundary.
            child_pages = [rng.randrange(start, end)]
        rng.shuffle(child_pages)
        children = [{"title": f"Parent {index} child {number}", "start_page": page + 1, "children": []}
                    for number, page in enumerate(child_pages)]
        title = f"Parent {index}"
        parents.append({"title": title, "start_page": start + 1, "children": children})
        if not children:
            expected.append((start, end, title))
            continue
        sorted_starts = sorted(set(child_pages))
        if sorted_starts[0] > start:
            expected.append((start, sorted_starts[0], title + " - Opening pages"))
        for number, page in enumerate(sorted_starts):
            child_end = sorted_starts[number + 1] if number + 1 < len(sorted_starts) else end
            expected.append((page, child_end, children[child_pages.index(page)]["title"]))
    rng.shuffle(parents)
    # These alias descendants must not add cuts or replace the first source subtree.
    first = parents[0]
    parents.append({"title": "Alias parent", "start_page": first["start_page"],
                    "children": [{"title": "Ignored alias child", "start_page": first["start_page"], "children": []}]})
    return parents, expected


def characterize(work):
    sources = history.original_sources()
    paths = (*history.BASELINE_PATHS, "engine/winbooksplit_engine.py", "tests/extraction/characterize_extraction.py",
             "tests/manual/characterize_manual.py", "tests/bookmarks/characterize_level1.py",
             "tests/bookmarks/characterize_level2.py", "tests/bookmarks/README.md", "tests/run_tests.py")
    before = {path: history.file_digest(ROOT / path) for path in paths}
    engine = manual.load_module("wbs_level2_current_engine", ENGINE)
    generator = history.load_generator()
    fixtures = generator.generate_fixtures(work / "fixtures")
    input_hashes = {str(path): history.file_digest(path) for path in fixtures.values()}
    cases = []
    for oracle in json.loads(sources["docs/codex-v1.0.0/PLAN_ORACLES.json"])["bookmarks"]:
        if oracle["id"] not in ORACLE_IDS:
            continue
        source = level1.fixture(work / (oracle["id"] + ".pdf"), generator, oracle["pages"], oracle["parents"])
        input_hashes[str(source)] = history.file_digest(source)
        result = cli_case(source, work / (oracle["id"] + "-cli"), generator, expected=oracle.get("expected_ranges"),
                          error=oracle.get("expected_error"), warnings=oracle.get("expected_warnings", []))
        if "expected_ranges" in oracle:
            plan = engine.plan_level2(PdfReader(source), oracle["pages"])
            check_plan(plan, oracle["pages"], oracle["expected_ranges"])
            require([record["filename"] for record in result["outputs"]] == [entry["filename"] for entry in plan["entries"]],
                    "Actual CLI filenames differ from the planned titles")
            if oracle["id"] == "BM-03":
                check_plan(plan, 12, oracle["expected_ranges"], ["Front matter", "A - Opening pages", "A1", "A2", "B - Opening pages", "B1"])
        cases.append({"oracle_id": oracle["id"], "passed": True, "expected_ranges": oracle.get("expected_ranges"),
                      "expected_error": oracle.get("expected_error"), **result})
    hierarchy, extra_inputs = hierarchy_cases(engine, work, fixtures, generator)
    input_hashes.update({str(path): history.file_digest(path) for path in extra_inputs})
    seed, count = 20261009, 150
    rng, samples = random.Random(seed), []
    for index in range(count):
        pages = rng.randint(1, 120)
        parents, expected = seeded_definition(rng, pages)
        plan = engine.plan_level2(SyntheticReader(oracle_outline(parents)), pages)
        check_plan(plan, pages, [[start, end] for start, end, _ in expected], [title for _, _, title in expected])
        if index < 10:
            parents10, expected10 = seeded_definition(rng, 10)
            source = level1.fixture(work / f"writer-{index}.pdf", generator, 10, parents10)
            input_hashes[str(source)] = history.file_digest(source)
            plan10 = engine.plan_level2(PdfReader(source), 10)
            check_plan(plan10, 10, [[start, end] for start, end, _ in expected10], [title for _, _, title in expected10])
            result = cli_case(source, work / f"writer-output-{index}", generator,
                              expected=[[start, end] for start, end, _ in expected10])
            samples.append({"index": index, "parents": parents10, **result})
    import_observation = history.probe_import(work)
    guard_path = work / "baseline-must-refuse.json"
    guard = history.run([sys.executable, "-I", "-B", str(ROOT / "tests/baseline/characterize_original.py"),
                         "--report", str(guard_path)], work, environment=history.clean_environment(work))
    require(guard["exit_code"] == 1 and "Original launchers changed" in guard["stderr"] and not guard_path.exists(),
            "Immutable historical baseline guard was weakened")
    require(input_hashes == {path: history.file_digest(Path(path)) for path in input_hashes}, "Level 2 inputs changed")
    require(before == {path: history.file_digest(ROOT / path) for path in paths}, "Level 2 source changed during testing")
    return {"schema_version": 1, "task_id": "M1-T04", "observed_at": datetime.now(timezone.utc).isoformat(),
            "result": "LEVEL2_REGRESSION_PASSED", "success": True, "exit_code": 0, "acceptance_ids": ACCEPTANCE_IDS,
            "engine_case_count": len(cases), "engine_cases": cases, "hierarchy_cases": hierarchy,
            "seeded_level2_cases": {"seed": seed, "count": count, "passed": True}, "sampled_writer_cases": samples,
            "source_unchanged": True, "baseline_guards_preserved": True, "input_and_neighbor_unchanged": True,
            "immutable_original_commit": history.ORIGINAL_COMMIT, "historical_original_guard": guard,
            "tested_path_sha256": before, "input_sha256": input_hashes, "import_observation": import_observation,
            "environment": {"python": sys.version, "python_executable": sys.executable,
                            "platform": platform.platform(), "pypdf": history.pypdf.__version__},
            "not_run": ["interactive no-plan fallback offer", "all launcher/error paths", "Explorer", "Calibre conversion",
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
        with tempfile.TemporaryDirectory(prefix="WinBookSplit-M1-T04-tests-") as directory:
            report = characterize(Path(directory).resolve())
        report["owned_temp_removed"] = not Path(directory).exists()
        require(report["owned_temp_removed"], "Owned temporary cleanup failed")
        with report_path.open("x", encoding="utf-8", newline="\n") as stream:
            json.dump(report, stream, indent=2, ensure_ascii=True)
            stream.write("\n")
        print("Level 2 acceptance passed: four CLI oracles, six hierarchy aggregates, 150 seeded plans, ten writer samples")
        return 0
    except (OSError, RuntimeError, ValueError, subprocess.TimeoutExpired) as error:
        print("Level 2 acceptance failed: " + str(error), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
