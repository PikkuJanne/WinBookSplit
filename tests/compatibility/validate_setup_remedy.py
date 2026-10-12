"""Reopen actual remedy-roundtrip PDF output using the maintained native oracle.

This finite authored fixture never launches a process or changes an input.
Private paths stay in the local receipt; the task evidence exporter redacts it.
"""
from __future__ import annotations

import argparse
from hashlib import sha256
import importlib.util
from pathlib import Path
import sys


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    for name in ("application-root", "input-file", "output-base", "process-report", "output"):
        parser.add_argument("--" + name, type=Path, required=True)
    parser.add_argument("--input-sha256", required=True)
    args = parser.parse_args()
    oracle_path = Path(__file__).resolve().parents[2] / "tools/compatibility/native_smoke.py"
    spec = importlib.util.spec_from_file_location("wbs_remedy_native_oracle", oracle_path)
    oracle = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(oracle)
    oracle.need(sys.flags.isolated and sys.dont_write_bytecode, "isolated_python_required")
    policy_hash = sha256((args.application_root / "engine/WinBookSplit.Compatibility.json").read_bytes()).hexdigest()
    oracle.bind_application(args.application_root, policy_hash)
    source = oracle.observation(args.input_file)
    oracle.need(source["sha256"] == args.input_sha256, "authored_source_changed")
    probe = oracle.native_probe(oracle.read_json(args.process_report))
    final = oracle.outcome(probe["Stdout"])
    oracle.need(probe["ExitCode"] == final["exit_code"] == 0 and final["written_count"] == 3,
                "remedy_retry_failed")
    engine = final["engine_result"]
    physical = oracle.pdf_snapshot(args.input_file)
    oracle.need(len(physical["pages"]) == 10 and
                [page["markers"] for page in physical["pages"]] ==
                [[f"{number:03}"] for number in range(1, 11)], "authored_fixture_invalid")
    directory = oracle.validate_publication(oracle.ordinary(args.output_base, directory=True),
                                            engine["execution"], physical, [[0, 3], [3, 6], [6, 10]],
                                            False, oracle.APP_VERSION)
    logs = [line[len("Log: "):] for line in probe["Stdout"].splitlines() if line.startswith("Log: ")]
    oracle.need(len(logs) == 1, "console_log_count")
    console = oracle.authenticate_console_manifest(Path(logs[0]), final, source_path=args.input_file,
                                                    source_observation=source,
                                                    expected_python_version="3.15.0")
    oracle.unchanged(source)
    oracle.write_json(args.output, {"protocol": "winbooksplit.setup-remedy-oracle", "schema_version": 1,
                                   "success": True, "physical_pages": 10, "written_count": 3,
                                   "ranges": [[0, 3], [3, 6], [6, 10]], "source": source,
                                   "final_directory": str(directory), "console": console})
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
