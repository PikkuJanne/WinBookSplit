"""Current manual acceptance with a corrected Level 2 launcher reference."""

from __future__ import annotations

import argparse
from contextlib import redirect_stdout
from datetime import datetime, timezone
import importlib.util
from io import StringIO
import json
import os
from pathlib import Path
import random
import subprocess
import sys
import tempfile

from pypdf import PdfReader, PdfWriter

ROOT = Path(__file__).resolve().parents[2]
ENGINE = ROOT / "engine/winbooksplit_engine.py"
ACCEPTANCE_IDS = [f"AC-{number:03d}" for number in range(13, 19)]


def load_module(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


history = load_module("wbs_manual_history", ROOT / "tests/extraction/characterize_extraction.py")
require = history.require


def check_outputs(records, expected_ranges, pages):
    require([record["range"] for record in records] == expected_ranges, "Unexpected manual ranges")
    require([record["page_ids"] for record in records]
            == [list(range(start + 1, end + 1)) for start, end in expected_ranges],
            "Incorrect page identity/order within a slice")
    require([page for record in records for page in record["page_ids"]] == list(range(1, pages + 1)),
            "Missing, repeated or reordered physical pages")


def observe(engine, source, output, cwd, mode, manual_data=None):
    command = [sys.executable, "-I", "-B", str(engine), str(source), str(output), mode]
    if manual_data is not None:
        command.append(manual_data)
    return history.run(command, cwd, environment=history.clean_environment(cwd))


def characterize(work, shells):
    sources = history.original_sources()  # Keep immutable baseline raw-byte guards.
    paths = (*history.BASELINE_PATHS, "engine/winbooksplit_engine.py",
             "tests/extraction/characterize_extraction.py", "tests/manual/characterize_manual.py")
    before = {path: history.file_digest(ROOT / path) for path in paths}
    engine = load_module("wbs_manual_engine", ENGINE)
    generator = history.load_generator()
    fixtures = generator.generate_fixtures(work / "fixtures")
    single = work / "single-page.pdf"
    single.write_bytes(generator._pdf_bytes({"title": "Synthetic one-page fixture", "pages": 1, "outline": []}))
    zero = work / "zero-page.pdf"
    with zero.open("xb") as stream:
        PdfWriter().write(stream)
    by_pages = {10: fixtures["simple10"], 1: single, 0: zero}
    inputs_before = {str(path): history.file_digest(path) for path in (*fixtures.values(), single, zero)}
    oracles = json.loads(sources["docs/codex-v1.0.0/PLAN_ORACLES.json"])
    cases = []
    for oracle in oracles["manual"]:
        cwd = work / oracle["id"]
        cwd.mkdir()
        output = cwd / "output"
        output.mkdir()
        neighbor = output / "synthetic-neighbor.txt"
        neighbor.write_bytes(b"Preserve this synthetic neighbor\n")
        result = observe(ENGINE, by_pages[oracle["pages"]], output, cwd, "manual", oracle["input"])
        records = history.outputs(output, generator)
        if "expected_error" in oracle:
            require(result["exit_code"] == 1 and oracle["expected_error"] in result["stdout"]
                    and not records and "[Writing]" not in result["stdout"], "Invalid request wrote or succeeded")
        else:
            require(result["exit_code"] == 0 and not result["stderr"], "Valid manual request failed")
            check_outputs(records, oracle["expected_ranges"], oracle["pages"])
            if oracle["id"] in {"MAN-01", "MAN-06"}:
                require("One section; no internal split." in result["stdout"], "Missing whole-document notice")
            if oracle["id"] in {"MAN-03", "MAN-05"}:
                require("Physical page 1 added" in result["stdout"], "Missing first-page notice")
            if oracle["id"] == "MAN-04":
                require("sorted into physical page order" in result["stdout"]
                        and "Duplicate start pages removed" in result["stdout"], "Missing normalization notices")
        require(neighbor.read_bytes() == b"Preserve this synthetic neighbor\n", "Neighbor changed")
        require({path.name for path in output.iterdir()} == {"synthetic-neighbor.txt", *(r["filename"] for r in records)},
                "Unexpected files written")
        cases.append({"oracle_id": oracle["id"], "passed": True, "input": oracle["input"],
                      "expected_ranges": oracle.get("expected_ranges"),
                      "expected_error": oracle.get("expected_error"), "outputs": records, **result})

    # Extra grammar/digit-limit probes use the actual CLI and preserve inputs.
    extras = [("comma", ",", None), ("zero", "0", None), ("missing", None, None),
              ("fullwidth", "１,４", None), ("embedded-space", "1,4 7", None),
              ("late-invalid", "7,04,4,1,abc", None), ("huge", "9" * 5000, None),
              ("long-leading-zeros", "0" * 5000 + "1", [[0, 10]])]
    extra_cases = []
    for label, text, expected_ranges in extras:
        cwd = work / ("extra-" + label)
        cwd.mkdir()
        output = cwd / "output"
        output.mkdir()
        result = observe(ENGINE, fixtures["simple10"], output, cwd, "manual", text)
        records = history.outputs(output, generator)
        if expected_ranges is None:
            require(result["exit_code"] == 1 and "invalid_start_pages" in result["stdout"]
                    and not list(output.iterdir()), "Extra invalid syntax was accepted")
        else:
            require(result["exit_code"] == 0 and not result["stderr"], "Valid long leading-zero token failed")
            check_outputs(records, expected_ranges, 10)
        extra_cases.append({"id": label, "input_length": len(text) if text is not None else None,
                            "expected_ranges": expected_ranges,
                            "expected_error": "invalid_start_pages" if expected_ranges is None else None,
                            "passed": True, "outputs": records, **result})

    seed = 20261009
    rng = random.Random(seed)
    sampled_writes = []
    for index in range(250):
        pages = rng.randint(1, 120)
        requested = [rng.randint(1, pages) for _ in range(rng.randint(1, 25))]
        rng.shuffle(requested)
        text = ", ".join("0" * rng.randint(0, 4) + str(page) for page in requested)
        plan = engine.plan_manual_starts(text, pages)
        ranges = plan["ranges"]
        require(ranges[0][0] == 0 and ranges[-1][1] == pages
                and all(0 <= start < end <= pages for start, end in ranges)
                and all(left[1] == right[0] for left, right in zip(ranges, ranges[1:])), "Invalid seeded partition")
        require([page for start, end in ranges for page in range(start, end)] == list(range(pages)),
                "Seeded partition omitted/repeated/reordered a page")
        require(plan["starts"] == sorted(set([1, *requested])), "Seeded normalization changed valid starts")
        if index < 25:
            # Exercise the real writer on bounded ten-page marker fixtures.
            starts10 = [1 + ((page - 1) % 10) for page in requested]
            text10 = ",".join(map(str, starts10))
            plan10 = engine.plan_manual_starts(text10, 10)
            output = work / f"seeded-output-{index}"
            output.mkdir()
            with redirect_stdout(StringIO()):
                engine.split_pdf(str(fixtures["simple10"]), str(output), "manual", text10)
            records = history.outputs(output, generator)
            check_outputs(records, [list(pair) for pair in plan10["ranges"]], 10)
            sampled_writes.append({"index": index, "starts": starts10, "outputs": records})

    oracle_map = {case["id"]: case for case in oracles["bookmarks"]}
    # The last known-bad current comparison is replaced deliberately in M1-T04.
    # Immutable original evidence remains guarded; actual current launchers use
    # independently checked corrected BM-03 output bytes as their reference.
    reference_cwd = work / "corrected-level2-reference"
    reference_cwd.mkdir()
    reference_output = reference_cwd / "output"
    reference_output.mkdir()
    reference_neighbor = reference_output / "synthetic-neighbor.txt"
    reference_neighbor.write_bytes(b"Original corrected-reference neighbor\n")
    reference_result = observe(ENGINE, fixtures["nested12"], reference_output, reference_cwd, "2")
    reference_records = history.outputs(reference_output, generator)
    require(reference_result["exit_code"] == 0 and not reference_result["stderr"], "Corrected BM-03 reference failed")
    reference_ranges = oracle_map["BM-03"]["expected_ranges"]
    check_outputs(reference_records, reference_ranges, 12)
    reference_plan = engine.plan_level2(PdfReader(fixtures["nested12"]), 12)
    reference_titles = [entry["title"] for entry in reference_plan["entries"]]
    require(reference_titles == ["Front matter", "A - Opening pages", "A1", "A2", "B - Opening pages", "B1"],
            "Incorrect corrected Level 2 titles")
    require([record["filename"] for record in reference_records] == [entry["filename"] for entry in reference_plan["entries"]],
            "Corrected Level 2 writer did not use planned titles")
    require(reference_neighbor.read_bytes() == b"Original corrected-reference neighbor\n"
            and {path.name for path in reference_output.iterdir()} == {reference_neighbor.name, *(r["filename"] for r in reference_records)},
            "Corrected reference modified its neighbor or wrote unexpected files")
    level2_reference = {"oracle_id": "BM-03", "passed": True, "expected_ranges": reference_ranges,
                        "titles": reference_titles, "outputs": reference_records, **reference_result}

    import_observation = history.probe_import(work)
    guard_path = work / "baseline-must-refuse.json"
    guard = history.run([sys.executable, "-I", "-B", str(ROOT / "tests/baseline/characterize_original.py"),
                         "--report", str(guard_path)], work, environment=history.clean_environment(work))
    require(guard["exit_code"] == 1 and "Original launchers changed" in guard["stderr"]
            and not guard_path.exists(), "Historical baseline guard was weakened")
    # Existing owned launcher probes compare exact output bytes, independent of
    # notices. Both manual 4,7 and Level 2 BM-03 now use corrected references.
    references = [{"oracle_id": "MAN-03", "original": {"outputs": next(
        case["outputs"] for case in cases if case["oracle_id"] == "MAN-03")}},
        {"oracle_id": "BM-03", "original": {"outputs": reference_records}}]
    entrypoints = history.launchers(work, fixtures, generator, shells, references) if shells else None
    require(inputs_before == {str(path): history.file_digest(path) for path in (*fixtures.values(), single, zero)},
            "Synthetic source input changed")
    require(before == {path: history.file_digest(ROOT / path) for path in paths}, "Tested source changed during run")
    return {"schema_version": 1, "task_id": "M1-T02", "observed_at": datetime.now(timezone.utc).isoformat(),
            "result": "MANUAL_REGRESSION_PASSED", "success": True, "exit_code": 0,
            "acceptance_ids": ACCEPTANCE_IDS, "engine_cases": cases, "engine_case_count": len(cases),
            "extra_cli_cases": extra_cases, "seeded_property_cases": {"seed": seed, "count": 250, "passed": True},
            "sampled_writer_cases": sampled_writes, "historical_bookmark_cases": [],
            "level2_launcher_reference": level2_reference,
            "baseline_guards_preserved": True, "historical_original_guard": guard,
            "immutable_original_commit": history.ORIGINAL_COMMIT, "source_unchanged": True,
            "tested_path_sha256": before, "input_and_neighbor_unchanged": True,
            "fixtures": {name: {"sha256": history.file_digest(path), "page_ids": generator.page_ids(path)}
                         for name, path in fixtures.items()},
            "import_observation": import_observation, "entrypoints": entrypoints,
            "not_run": ["full corrected Level 1/2 acceptance (separate layers)", "Explorer",
                        "Calibre conversion", "release-package acceptance"]
                       + ([] if shells else ["actual entrypoints"])}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--report", type=Path, required=True)
    parser.add_argument("--shell-path", type=Path, action="append", default=[])
    args = parser.parse_args()
    try:
        require(sys.flags.isolated and sys.dont_write_bytecode, "Use explicit developer Python -I -B")
        require(args.report.is_absolute() and not os.path.lexists(args.report), "New absolute report required")
        report_path = args.report.resolve()
        require(not report_path.is_relative_to(ROOT) and report_path.parent.is_dir(), "External report required")
        with tempfile.TemporaryDirectory(prefix="WinBookSplit-M1-T02-tests-") as directory:
            report = characterize(Path(directory).resolve(), args.shell_path)
        report["owned_temp_removed"] = not Path(directory).exists()
        require(report["owned_temp_removed"], "Owned temp cleanup failed")
        with report_path.open("x", encoding="utf-8", newline="\n") as stream:
            json.dump(report, stream, indent=2, ensure_ascii=True)
            stream.write("\n")
        print("Manual acceptance passed: 22 oracles, 8 extra CLI cases, 250 seeded plans, 25 real writer samples")
        return 0
    except (OSError, RuntimeError, subprocess.TimeoutExpired) as error:
        print("Manual acceptance failed: " + str(error), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
