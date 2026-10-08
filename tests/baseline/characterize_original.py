"""Execute the unchanged embedded engine; verify explicitly known ORIGINAL results.

This is deliberately separate from future fixed-engine acceptance tests. A zero
exit status means the baseline was reproduced, including its known defects.
"""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
import hashlib
import importlib.util
from importlib.metadata import version
import json
import os
from pathlib import Path
import platform
import re
import shutil
import subprocess
import sys
import tempfile

import pypdf


ROOT = Path(__file__).resolve().parents[2]
EXPECTED_PATH = Path(__file__).with_name("expected_original.json")


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def run(argv: list[str], cwd: Path, stdin: str | None = None) -> dict:
    # communicate() drains stdout/stderr concurrently; never evaluate book paths.
    result = subprocess.run(
        argv, cwd=cwd, input=stdin, capture_output=True, text=True,
        encoding="utf-8", errors="replace", timeout=30, check=False,
    )
    return {"exit_code": result.returncode, "stdout": result.stdout,
            "stderr": result.stderr}


def require(condition: bool, message: str) -> None:
    if not condition:
        raise RuntimeError(message)


def load_generator():
    sys.dont_write_bytecode = True
    path = ROOT / "tests/fixtures/generate_pdf_fixtures.py"
    spec = importlib.util.spec_from_file_location("wbs_fixture_generator", path)
    require(spec is not None and spec.loader is not None, "Cannot load generator")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def outline_rows(reader: pypdf.PdfReader, nodes=None, depth=1) -> list[list]:
    rows = []
    for node in reader.outline if nodes is None else nodes:
        if isinstance(node, list):
            rows.extend(outline_rows(reader, node, depth + 1))
        else:
            rows.append([depth, node.title,
                         reader.get_destination_page_number(node) + 1])
    return rows


def characterize(temp: Path, include_launcher: bool) -> dict:
    expected = json.loads(EXPECTED_PATH.read_text(encoding="utf-8"))
    source_hashes = {path: sha256(ROOT / path) for path in expected["source_sha256"]}
    require(source_hashes == expected["source_sha256"],
            "Original launchers changed. Do not run known-bad expectations against fixed code.")
    source = (ROOT / "WinBookSplit.ps1").read_text(encoding="utf-8-sig")
    matches = re.findall(r'^\$pyScriptContent = @"\n(.*?)\n"@$', source,
                         flags=re.MULTILINE | re.DOTALL)
    require(len(matches) == 1, "Expected exactly one original embedded engine")
    engine_source = matches[0] + "\n"
    require("$" not in engine_source and "`" not in engine_source,
            "Here-string contains interpolation; direct extraction is not equivalent")
    engine = temp / "original_embedded_engine.py"
    engine.write_text(engine_source, encoding="utf-8")
    compile(engine_source, str(engine), "exec")

    generator = load_generator()
    fixtures = generator.generate_fixtures(temp / "fixtures")
    repeated = generator.generate_fixtures(temp / "fixtures-repeat")
    provenance = json.loads((temp / "fixtures/provenance.json").read_text(encoding="utf-8"))
    require((temp / "fixtures/provenance.json").read_bytes()
            == (temp / "fixtures-repeat/provenance.json").read_bytes(),
            "Non-repeatable fixture provenance")
    fixtures_report = {}
    expected_outlines = {
        "simple10": [[1, "A", 4], [1, "B", 7]],
        "nested12": [[1, "A", 3], [2, "A1", 4], [2, "A2", 7],
                     [1, "B", 9], [2, "B1", 11]],
    }
    for name, path in fixtures.items():
        pages = 10 if name == "simple10" else 12
        ids = generator.page_ids(path)
        require(ids == list(range(1, pages + 1)), f"Bad fixture page IDs: {name}")
        require(sha256(path) == sha256(repeated[name]), f"Non-repeatable fixture: {name}")
        reader = pypdf.PdfReader(path)
        rows = outline_rows(reader)
        require(rows == expected_outlines[name], f"Bad outline: {name}: {rows}")
        fixtures_report[name] = {"sha256": sha256(path), "page_ids": ids,
                                 "outlines": rows, "repeat_bytes_equal": True,
                                 "metadata": dict(reader.metadata)}

    oracles = json.loads((ROOT / "docs/codex-v1.0.0/PLAN_ORACLES.json").read_text(
        encoding="utf-8"))
    oracle_map = {case["id"]: case for kind in ("manual", "bookmarks")
                  for case in oracles[kind]}
    cases = []
    for original in expected["cases"]:
        oracle = oracle_map[original["oracle_id"]]
        case_dir = temp / original["oracle_id"]
        case_dir.mkdir()
        output = case_dir / "output"
        output.mkdir()
        neighbor = case_dir / "prior-output.pdf"
        neighbor.write_bytes(b"Original synthetic neighbor sentinel\n")
        neighbor_hash = sha256(neighbor)
        input_path = fixtures[original["fixture"]]
        input_hash = sha256(input_path)
        mode = "manual" if "input" in oracle else str(oracle["level"])
        argv = [sys.executable, "-I", str(engine), str(input_path), str(output), mode]
        if mode == "manual":
            argv.append(oracle["input"])
        result = run(argv, cwd=case_dir)
        outputs = []
        for path in sorted(output.glob("*.pdf")):
            ids = generator.page_ids(path)
            require(bool(ids), f"Empty slice: {path.name}")
            require(ids == list(range(ids[0], ids[-1] + 1)),
                    f"Disordered or duplicated pages: {path.name}")
            outputs.append({"filename": path.name, "page_ids": ids,
                            "range": [ids[0] - 1, ids[-1]], "sha256": sha256(path)})
        ranges = [item["range"] for item in outputs]
        require(len(outputs) == len(list(output.iterdir())),
                f"Unexpected non-PDF output: {oracle['id']}")
        require(result["exit_code"] == original["exit_code"],
                f"Unexpected original exit: {oracle['id']}: {result}")
        require(ranges == original["ranges"],
                f"Unexpected original ranges: {oracle['id']}: {ranges}")
        require(input_hash == sha256(input_path), f"Input mutated: {oracle['id']}")
        require(neighbor_hash == sha256(neighbor), f"Neighbor mutated: {oracle['id']}")
        require(source_hashes == {path: sha256(ROOT / path) for path in source_hashes},
                "Original source mutated")
        all_ids = [page for item in outputs for page in item["page_ids"]]
        target = {key: oracle[key] for key in ("expected_ranges", "expected_error")
                  if key in oracle}
        fixed_match = (result["exit_code"] == 0 and ranges == target["expected_ranges"]
                       if "expected_ranges" in target else
                       result["exit_code"] != 0 and not outputs)
        cases.append({"oracle_id": oracle["id"], "acceptance_id":
                      "AC-004" if mode == "manual" else "AC-005",
                      "mode": mode, "input": oracle.get("input"),
                      "original_expectation_matched": True,
                      "known_defect": original["defect"], **result,
                      "actual_ranges": ranges, "outputs": outputs,
                      "missing_page_ids": sorted(set(range(1, oracle["pages"] + 1))
                                                 - set(all_ids)),
                      "complete_ordered_coverage": all_ids == list(range(1, oracle["pages"] + 1)),
                      "fixed_target": target, "fixed_target_matched": fixed_match,
                      "source_and_neighbor_unchanged": True})

    launcher = []
    if include_launcher:
        require(os.name == "nt", "Launcher probes require actual Windows")
        powershell = shutil.which("powershell.exe")
        cmd = shutil.which("cmd.exe")
        require(bool(powershell and cmd), "Windows launcher hosts missing")
        missing = temp / "never-created.pdf"
        require(not missing.exists(), "Missing-input probe must have no input file")
        result = run([powershell, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
                      str(ROOT / "WinBookSplit.ps1"), str(missing)], temp, "\n")
        require(result["exit_code"] == 0 and "Error: Invalid file." in result["stdout"],
                f"Unexpected original missing-input result: {result}")
        launcher.append({"id": "PS51-missing-input", "scope": "unmodified early rejection path",
                         "known_defect": "invalid_input_exit_zero", **result})
        result = run([cmd, "/d", "/c", str(ROOT / "WinBookSplit.bat")], temp, "\n")
        require(result["exit_code"] == 0 and "No file dropped." in result["stdout"],
                f"Unexpected original no-drop result: {result}")
        launcher.append({"id": "BAT-no-input", "scope": "unmodified no-argument path",
                         "known_defect": "no_input_exit_zero", **result})
        version_probe = run([powershell, "-NoProfile", "-Command",
                             "$PSVersionTable.PSVersion.ToString()"], temp)
        require(version_probe["exit_code"] == 0, "Cannot record launcher shell version")
        launcher.append({"id": "PS51-version", **version_probe})

    require(source_hashes == {path: sha256(ROOT / path) for path in source_hashes},
            "Launcher probes mutated original source")

    digest_paths = ["WinBookSplit.ps1", "WinBookSplit.bat",
                    "tests/baseline/characterize_original.py",
                    "tests/baseline/expected_original.json",
                    "tests/fixtures/generate_pdf_fixtures.py",
                    "docs/codex-v1.0.0/PLAN_ORACLES.json"]
    digests = {path: sha256(ROOT / path) for path in digest_paths}
    source_digest = hashlib.sha256(json.dumps(digests, sort_keys=True).encode()).hexdigest()
    base = run(["git", "rev-parse", "HEAD"], ROOT)
    tree = run(["git", "rev-parse", "HEAD^{tree}"], ROOT)
    require(base["exit_code"] == tree["exit_code"] == 0, "Cannot record Git source identity")
    return {
        "schema_version": 1,
        "task_id": "M0-T02", "observed_at": datetime.now(timezone.utc).isoformat(),
        "result": "ORIGINAL_BEHAVIOR_REPRODUCED",
        "meaning": "Known-bad characterization matched; no fixed-engine acceptance pass is claimed.",
        "base_commit": base["stdout"].strip(), "base_tree": tree["stdout"].strip(),
        "tested_path_sha256": digests, "tested_paths_digest": source_digest,
        "embedded_engine_sha256": sha256(engine),
        "environment": {"python": sys.version, "pypdf": pypdf.__version__,
                        "platform": platform.platform(), "reportlab": version("reportlab"),
                        "python_isolated_children": True},
        "fixture_provenance": "tests/fixtures/README.md; generator-authored MIT fixtures only",
        "generated_provenance": provenance,
        "fixtures": fixtures_report, "engine_case_count": len(cases),
        "engine_cases": cases, "launcher_probes": launcher,
        "not_run": ["full PowerShell splitting workflow", "Explorer drag/drop",
                    "Calibre conversion", "release-package matrix", "fixed-engine acceptance"],
        "cleanup": "All generated fixtures/engine/slices are inside one TemporaryDirectory owned by this run.",
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--report", type=Path, required=True,
                        help="New JSON evidence file; existing files are never overwritten")
    parser.add_argument("--launcher-probes", action="store_true")
    args = parser.parse_args()
    try:
        require(not args.report.exists(), "Report path already exists")
        with tempfile.TemporaryDirectory(prefix="WinBookSplit-M0-T02-") as directory:
            report = characterize(Path(directory), args.launcher_probes)
        # No success record is written until the complete characterization passes.
        with args.report.open("x", encoding="utf-8", newline="\n") as stream:
            json.dump(report, stream, indent=2, ensure_ascii=False)
            stream.write("\n")
        print(f"Original behavior reproduced: {report['engine_case_count']} engine cases; "
              f"{len(report['launcher_probes']) - 1 if args.launcher_probes else 0} launcher boundary probes.")
        print("This is characterization, not a fixed-version acceptance pass.")
        return 0
    except (OSError, RuntimeError, subprocess.TimeoutExpired) as error:
        print(f"Characterization failed: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
