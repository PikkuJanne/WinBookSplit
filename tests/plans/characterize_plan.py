"""Shared immutable plan, real writer parity and captured-source acceptance.

Only authored synthetic PDFs are read. Deliberately changed source copies are
owned by this run; the immutable fixtures and historical guards stay intact.
"""

from __future__ import annotations

import argparse
from collections.abc import Mapping
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

ROOT = Path(__file__).resolve().parents[2]
ENGINE = ROOT / "engine/winbooksplit_engine.py"
ACCEPTANCE_IDS = ["AC-027", "AC-028", "AC-029"]
spec = importlib.util.spec_from_file_location("wbs_shared_plan_helpers", ROOT / "tests/bookmarks/characterize_level2.py")
level2 = importlib.util.module_from_spec(spec)
spec.loader.exec_module(level2)
level1, manual, history, require = level2.level1, level2.manual, level2.history, level2.require


def plain(value):
    if isinstance(value, Mapping):
        return {key: plain(item) for key, item in value.items()}
    if isinstance(value, (tuple, list)):
        return [plain(item) for item in value]
    return value


def model(ranges=((0, 10),), *, mode="manual", pages=10):
    return {"mode": mode, "total_pages": pages,
            "entries": [{"sequence": index, "title": f"Section {index}", "start": start, "end": end,
                         "parent_id": None, "reason": "manual" if mode == "manual" else "bookmark",
                         "filename": f"{index:02d} - Section {index}.pdf", "warnings": []}
                        for index, (start, end) in enumerate(ranges, 1)]}


def check_complete(plan, pages, expected=None):
    ranges = tuple((entry["start"], entry["end"]) for entry in plan["entries"])
    require(plan["total_pages"] == pages and isinstance(plan["ranges"], tuple) and plan["ranges"] == ranges,
            "Shared validated ranges differ from entries or remain mutable")
    require([page for start, end in ranges for page in range(start, end)] == list(range(pages)),
            "Validated plan omitted, repeated or reordered a physical page")
    require(all(0 <= start < end <= pages for start, end in ranges), "Validated plan contains an empty/out-of-range section")
    require(plain(plan["coverage"]) == {"complete": True, "covered_pages": pages, "section_count": len(ranges)},
            "Coverage metadata differs from the complete partition")
    if expected is not None:
        require(ranges == tuple(map(tuple, expected)), "Validated boundaries differ from independent expected cuts")


def reject_model(engine, definition):
    try:
        engine.validate_plan(definition)
    except engine.PlanError as error:
        require(isinstance(error.code, str) and error.code and str(error), "Plan rejection lost its diagnostic")
        return {"error_code": error.code, "message": str(error)}
    raise RuntimeError("Malformed shared plan was accepted: " + repr(definition))


def structural_cases(engine):
    malformed = {
        "gap": [model(((0, 3), (4, 10))), model(((1, 10),)), model(((0, 9),))],
        "overlap": [model(((0, 6), (5, 10)))],
        "empty": [model(()), model(((0, 4), (4, 4), (4, 10)))],
        "negative": [model(((-1, 10),))],
        "reversed": [model(((0, 7), (7, 3), (3, 10)))],
        "overflow": [model(((0, 11),))],
        "zero-pages": [model(((0, 0),), pages=0)],
        "noninteger": [model(pages=True), model(pages=10.0), model(((False, 10),)), model(((0, 10.0),))],
    }
    # The same aggregate also checks entry/range contracts, not just sums.
    for field in ("sequence", "title", "start", "end", "parent_id", "reason", "filename", "warnings"):
        missing = model()
        del missing["entries"][0][field]
        malformed["noninteger"].append(missing)
    for changes in ({"sequence": 2}, {"sequence": True}, {"filename": "../escaped.pdf"},
                    {"filename": "unsafe\\name.pdf"}, {"filename": "unsafe\x00name.pdf"}, {"parent_id": []}):
        changed = model()
        changed["entries"][0].update(changes)
        malformed["noninteger"].append(changed)
    duplicate = model(((0, 5), (5, 10)))
    duplicate["entries"][1]["filename"] = duplicate["entries"][0]["filename"].upper().replace(".PDF", ".pdf")
    malformed["noninteger"].extend([duplicate, {**model(), "ranges": [[0, 9]]},
                                   {**model(), "mode": "3"}])
    records = []
    for name, definitions in malformed.items():
        rejections = [reject_model(engine, definition) for definition in definitions]
        records.append({"id": name, "passed": True, "rejected": True,
                        "error_code": rejections[0]["error_code"], "rejections": rejections})
    accepted = []
    for mode in ("manual", "1", "2"):
        frozen = engine.validate_plan({**model(mode=mode), "coverage": {"complete": False, "covered_pages": 0}})
        check_complete(frozen, 10, [[0, 10]])
        accepted.append({"mode": mode, "coverage": plain(frozen["coverage"]), "ranges": plain(frozen["ranges"])})
    records.append({"id": "valid-whole-document", "passed": True, "accepted": True, "modes": accepted})
    return records


def mutation_refused(action):
    try:
        action()
    except (TypeError, AttributeError):
        return
    raise RuntimeError("Read-only shared plan/result accepted a mutation")


def assign(container, key, value):
    container[key] = value


def preview_cases(engine, source, work):
    before_files = {path.relative_to(work).as_posix(): history.file_digest(path)
                    for path in work.rglob("*") if path.is_file()}
    prepared = engine.prepare_split(source, "manual", "7,4,4")
    preview = engine.preview_plan(prepared)
    require(preview is prepared.plan and engine.preview_plan(prepared) is preview, "Preview is not the prepared read-only plan")
    before = plain(preview)
    require(before["normalized_inputs"]["starts"] == [1, 4, 7] and len(before["notices"]) == 3,
            "Shared preview lost normalized inputs or notices")
    mutations = [lambda: assign(preview, "total_pages", 1),
                 lambda: assign(preview["entries"][0], "start", 8),
                 lambda: assign(preview["entries"], 0, {}),
                 lambda: assign(preview["normalized_inputs"]["starts"], 0, 8),
                 lambda: assign(preview["source_identity"], "sha256", "changed"),
                 lambda: setattr(prepared, "plan", model())]
    for action in mutations:
        mutation_refused(action)
    require(plain(preview) == before and before_files == {path.relative_to(work).as_posix(): history.file_digest(path)
                                                         for path in work.rglob("*") if path.is_file()},
            "Preparation/preview changed data or wrote chapters")
    records = [{"id": "read-only-preview", "passed": True, "mutation_rejected": True,
                "chapter_files_written": 0, "mutation_attempts": len(mutations)}]
    mutable = model(((0, 4), (4, 10)))
    mutable["entries"][0]["warnings"] = [{"code": "synthetic-warning", "details": ["original"]}]
    mutable["normalized_inputs"] = {"starts": [1, 5]}
    frozen = engine.validate_plan(mutable)
    snapshot = plain(frozen)
    mutable["entries"][0]["start"] = 9
    mutable["entries"][0]["warnings"][0]["details"][0] = "changed"
    mutable["normalized_inputs"]["starts"].append(9)
    require(plain(frozen) == snapshot, "Validated metadata still aliases caller-owned input containers")
    mutation_refused(lambda: assign(frozen["entries"][0]["warnings"][0]["details"], 0, "changed"))
    output = work / "unprepared-output"
    output.mkdir()
    try:
        engine.execute_split(frozen, output)
    except engine.PlanError as error:
        require(error.code == "invalid_prepared_split", "Plain mapping execution lost its diagnostic")
    else:
        raise RuntimeError("Execution accepted an unbound plain plan")
    require(not list(output.iterdir()), "Rejected unbound plan wrote files")
    records.append({"id": "isolated-preview-data", "passed": True, "detached_from_input": True,
                    "unbound_execution_rejected": True, "chapter_files_written": 0})
    return records


def page_content(path):
    return [history.digest(page.extract_text().encode("utf-8")) for page in history.pypdf.PdfReader(path).pages]


def write_prepared(engine, prepared, output, generator, expected, *, original_content=None):
    output.mkdir()
    neighbor = output / "synthetic-neighbor.txt"
    neighbor.write_bytes(b"Preserve this authored synthetic neighbor\n")
    neighbor_hash = history.file_digest(neighbor)
    preview = engine.preview_plan(prepared)
    check_complete(preview, preview["total_pages"], expected)
    preview_before = plain(preview)
    # A prepared job must use its captured plan and reader, not rerun a mode.
    with patch.object(engine, "prepare_split", side_effect=RuntimeError("Execution replanned")), \
            patch.object(engine, "plan_manual_starts", side_effect=RuntimeError("Execution replanned manual")), \
            patch.object(engine, "plan_level1", side_effect=RuntimeError("Execution replanned Level 1")), \
            patch.object(engine, "plan_level2", side_effect=RuntimeError("Execution replanned Level 2")), \
            patch.object(engine, "PdfReader", side_effect=RuntimeError("Execution reopened source")), \
            redirect_stdout(StringIO()):
        result = engine.execute_split(prepared, output)
    actual = history.outputs(output, generator)
    manual.check_outputs(actual, expected, preview["total_pages"])
    require([record["filename"] for record in actual] == [entry["filename"] for entry in preview["entries"]],
            "Real writer filenames differ from the preview")
    require(result["written_count"] == len(preview["entries"]) > 0
            and result["mode"] == preview["mode"] and result["total_pages"] == preview["total_pages"]
            and result["source_identity"] == preview["source_identity"] and result["coverage"] == preview["coverage"],
            "Execution result disagrees with its prepared plan/source/coverage")
    require(plain(result["outputs"]) == [{**entry, "page_count": entry["end"] - entry["start"]}
                                         for entry in preview_before["entries"]],
            "Execution results differ from preview entries")
    mutation_refused(lambda: assign(result, "written_count", 0))
    mutation_refused(lambda: assign(result["outputs"][0], "start", 9))
    require(plain(engine.preview_plan(prepared)) == preview_before, "Writing changed the shared preview")
    require(history.file_digest(neighbor) == neighbor_hash, "Shared writer changed a neighbor")
    require({path.name for path in output.iterdir()} == {"synthetic-neighbor.txt", *(record["filename"] for record in actual)},
            "Shared writer created unexpected files")
    content = [digest for record in actual for digest in page_content(output / record["filename"])]
    if original_content is not None:
        require(content == original_content, "Shared writer used replacement source content instead of the captured original")
    return {"total_pages": preview["total_pages"], "preview_entries": preview_before["entries"],
            "outputs": actual, "coverage": plain(result["coverage"]), "written_count": result["written_count"],
            "source_identity": plain(result["source_identity"]), "writer_result": plain(result),
            "neighbor_unchanged": True, "no_replanning_or_reopening": True,
            "page_content_sha256": content, "original_page_content_sha256": original_content}


def parity_cases(engine, work, fixtures, generator):
    definitions = [("manual", fixtures["simple10"], "4,7", [[0, 3], [3, 6], [6, 10]]),
                   ("1", fixtures["simple10"], None, [[0, 3], [3, 6], [6, 10]]),
                   ("2", fixtures["nested12"], None, [[0, 2], [2, 3], [3, 6], [6, 8], [8, 10], [10, 12]])]
    records = []
    for mode, source, data, expected in definitions:
        prepared = engine.prepare_split(source, mode, data)
        before = history.file_digest(source)
        record = write_prepared(engine, prepared, work / ("parity-" + mode), generator, expected,
                                original_content=page_content(source))
        require(history.file_digest(source) == before, "Shared writer changed its original input")
        records.append({"mode": mode, "passed": True, **record})
    return records


def source_cases(engine, work, fixtures, generator):
    records = []
    original_bytes = fixtures["simple10"].read_bytes()
    replacement_bytes = fixtures["nested12"].read_bytes()
    original_content = page_content(fixtures["simple10"])
    replacement_content = page_content(fixtures["nested12"])
    require(original_bytes != replacement_bytes, "Source mutation fixture must differ")
    require(original_content != replacement_content[:len(original_content)], "Replacement must change visible synthetic page content")
    for name in ("changed-same-path", "repointed-source"):
        source = work / (name + ".pdf")
        source.write_bytes(original_bytes)
        prepared = engine.prepare_split(source, "manual", "4,7")
        captured_sha = history.file_digest(source)
        if name == "repointed-source":
            moved = work / (name + "-original-moved.pdf")
            source.rename(moved)
        source.write_bytes(replacement_bytes)
        replacement_sha = history.file_digest(source)
        record = write_prepared(engine, prepared, work / (name + "-output"), generator, [[0, 3], [3, 6], [6, 10]],
                                original_content=original_content)
        require(record["source_identity"]["sha256"] == captured_sha
                and record["source_identity"]["binding"] == "reader_snapshot"
                and record["source_identity"]["size_bytes"] == len(original_bytes),
                "Prepared source identity changed after replacing its path")
        require(history.file_digest(source) == replacement_sha, "Execution changed the deliberate replacement")
        if name == "repointed-source":
            require(moved.read_bytes() == original_bytes, "Execution changed the moved original")
            # Reusing the same binding after removal must still use its snapshot.
            source.unlink()
            deleted = write_prepared(engine, prepared, work / "deleted-path-output", generator, [[0, 3], [3, 6], [6, 10]],
                                     original_content=original_content)
            require(not source.exists() and deleted["source_identity"] == record["source_identity"],
                    "Deleted source path invalidated/recreated the safe captured source")
            source.write_bytes(replacement_bytes)
            record["deleted_path_snapshot"] = deleted
        records.append({"id": name, "passed": True, "outcome": "bound_original_preserved",
                        "captured_sha256": captured_sha, "replacement_sha256": replacement_sha,
                        "replacement_unchanged": True, "replacement_page_content_sha256": replacement_content, **record})
    return records


def seeded_cases(engine):
    seed, rng = 20261009, random.Random(20261009)
    counts = {"manual": 0, "1": 0, "2": 0}
    observations = []
    for mode in counts:
        for index in range(100):
            pages = rng.randint(1, 300)
            if mode == "manual":
                starts = [rng.randint(1, pages) for _ in range(rng.randint(1, 12))]
                raw = engine.plan_manual_starts(",".join(map(str, starts)), pages)
                ordered = sorted({0, *(start - 1 for start in starts)})
                expected = list(zip(ordered, ordered[1:] + [pages]))
                require(raw["ranges"] == expected, "Seeded manual mode changed independent physical cuts")
                definition = model(raw["ranges"], mode=mode, pages=pages)
                definition["normalized_inputs"] = {"starts": raw["starts"]}
            elif mode == "1":
                starts = sorted(rng.sample(range(pages), rng.randint(1, min(8, pages))))
                shuffled = starts[:]
                rng.shuffle(shuffled)
                reader = level1.SyntheticReader([level1.Destination(f"Parent {start}", start) for start in shuffled])
                definition = engine.plan_level1(reader, pages)
                cuts = sorted({0, *starts})
                expected = list(zip(cuts, cuts[1:] + [pages]))
            else:
                parents, expected_sections = level2.seeded_definition(rng, pages)
                definition = engine.plan_level2(level1.SyntheticReader(level2.oracle_outline(parents)), pages)
                expected = [(start, end) for start, end, _ in expected_sections]
            frozen = engine.validate_plan(definition)
            check_complete(frozen, pages, expected)
            counts[mode] += 1
            observations.append({"mode": mode, "index": index, "pages": pages, "ranges": plain(frozen["ranges"])})
    return {"seed": seed, "count": sum(counts.values()), "per_mode": counts, "passed": True,
            "observations_sha256": history.digest(json.dumps(observations, sort_keys=True).encode())}


def characterize(work):
    history.original_sources()
    paths = (*history.BASELINE_PATHS, "engine/winbooksplit_engine.py", "tests/run_tests.py", "tests/README.md",
             "tests/extraction/characterize_extraction.py", "tests/manual/characterize_manual.py",
             "tests/bookmarks/characterize_level1.py", "tests/bookmarks/characterize_level2.py",
             "tests/plans/characterize_plan.py", "tests/plans/README.md")
    before = {path: history.file_digest(ROOT / path) for path in paths}
    engine = manual.load_module("wbs_current_shared_plan_engine", ENGINE)
    generator = history.load_generator()
    fixtures = generator.generate_fixtures(work / "fixtures")
    inputs_before = {str(path): history.file_digest(path) for path in fixtures.values()}
    structural = structural_cases(engine)
    previews = preview_cases(engine, fixtures["simple10"], work)
    parity = parity_cases(engine, work, fixtures, generator)
    bindings = source_cases(engine, work, fixtures, generator)
    properties = seeded_cases(engine)
    import_observation = history.probe_import(work)
    guard_path = work / "baseline-must-refuse.json"
    guard = history.run([sys.executable, "-I", "-B", str(ROOT / "tests/baseline/characterize_original.py"),
                         "--report", str(guard_path)], work, environment=history.clean_environment(work))
    require(guard["exit_code"] == 1 and "Original launchers changed" in guard["stderr"] and not guard_path.exists(),
            "Immutable historical baseline guard was weakened")
    require(inputs_before == {path: history.file_digest(Path(path)) for path in inputs_before},
            "Shared-plan immutable synthetic fixtures changed")
    require(before == {path: history.file_digest(ROOT / path) for path in paths}, "Shared-plan source changed during testing")
    return {"schema_version": 1, "task_id": "M1-T05", "observed_at": datetime.now(timezone.utc).isoformat(),
            "result": "SHARED_PLAN_REGRESSION_PASSED", "success": True, "exit_code": 0,
            "acceptance_ids": ACCEPTANCE_IDS, "structural_cases": structural, "preview_cases": previews,
            "parity_cases": parity, "source_cases": bindings, "seeded_plan_cases": properties,
            "source_unchanged": True, "baseline_guards_preserved": True, "input_and_neighbor_unchanged": True,
            "immutable_original_commit": history.ORIGINAL_COMMIT, "historical_original_guard": guard,
            "tested_path_sha256": before, "input_sha256": inputs_before, "import_observation": import_observation,
            "environment": {"python": sys.version, "python_executable": sys.executable,
                            "platform": platform.platform(), "pypdf": history.pypdf.__version__},
            "not_run": ["interactive preview UI", "additional actual launcher probes", "Explorer", "Calibre conversion",
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
        with tempfile.TemporaryDirectory(prefix="WinBookSplit-M1-T05-tests-") as directory:
            report = characterize(Path(directory).resolve())
        report["owned_temp_removed"] = not Path(directory).exists()
        require(report["owned_temp_removed"], "Owned temporary cleanup failed")
        with report_path.open("x", encoding="utf-8", newline="\n") as stream:
            json.dump(report, stream, indent=2, ensure_ascii=True)
            stream.write("\n")
        print("Shared-plan acceptance passed: nine structural aggregates, three real writer modes, immutable preview, source binding, 300 seeded plans")
        return 0
    except (OSError, RuntimeError, ValueError, subprocess.TimeoutExpired) as error:
        print("Shared-plan acceptance failed: " + str(error), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
