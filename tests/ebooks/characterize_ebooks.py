"""M4-T04 real authored ebook retention, physical pages and local behavior.

The inherited conversion layer remains a separate runner stage. This companion
adds ten actual application calls, four independent profile references and two
selected-format help calls; none are human/GUI or system-network-denial tests.
"""
from __future__ import annotations

import argparse
from contextlib import nullcontext
from datetime import datetime, timezone
from hashlib import sha256
import importlib.util
import json
import os
from pathlib import Path
import re
import shutil
import sys
import tempfile

ROOT = Path(__file__).resolve().parents[2]


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


support = load("wbs_ebooks_support", ROOT / "tests/support/characterize_support.py")
assets = load("wbs_ebooks_assets", ROOT / "tests/ebooks/assets.py")
behavior = load("wbs_ebooks_behavior", ROOT / "tests/ebooks/behavior.py")
validator = None  # Load after module definitions so pure helper imports remain usable.
cli, history, require = support.cli, support.history, support.require
paths, launchers, conversion, runner = support.paths, support.launchers, support.conversion, support.runner
RENDER_ROOT = None
COMPLETED_CASES = []
LAST_CASE = None
BEHAVIOR = {}
SOURCE_BEFORE = None


def tree_identity(directory):
    directory = Path(directory)
    require(directory.is_dir() and not history.is_reparse(directory), "Prior directory must be an ordinary owned directory")
    details = directory.lstat()
    members = {}
    for path in sorted(directory.iterdir()):
        require(path.is_file() and not history.is_reparse(path), "Owned publication has an unexpected nonregular member")
        members[path.name] = cli.identity(path)
    return {"path": str(directory), "device": details.st_dev, "inode": details.st_ino, "members": members}


def prior_directories(base, excluded=()):
    return {path.name: tree_identity(path) for path in sorted(base.iterdir())
            if path.is_dir() and path.name not in excluded}


def remove_publication(base, execution, authenticated):
    directory = Path(execution["final_directory"])
    expected = {".WinBookSplit-owner.json", "WinBookSplit_Manifest.json", "WinBookSplit_Converted.pdf",
                *(item["filename"] for item in execution["outputs"])}
    require(directory.parent == base and set(authenticated["members"]) == expected
            and tree_identity(directory) == authenticated,
            "Cleanup refuses an unexpected, replaced or changed publication member")
    launchers.remove_known_directory(directory, base, expected, ".WinBookSplit-owner.json",
        {"schema_version": 1, "kind": "run", "run_id": execution["run_id"]},
        expected_sha256={name: item["sha256"] for name, item in authenticated["members"].items()},
        expected_file_identities=authenticated["members"],
        expected_directory_identity={"device": authenticated["device"], "inode": authenticated["inode"]})


def protected(source, neighbor, prior):
    return {name: cli.identity(path) for name, path in (("source", source), ("neighbor", neighbor), ("prior", prior))}


def failure_record(result, base):
    diagnostic = result["diagnostic"]
    require(diagnostic["cleanup_complete"] is True and diagnostic["retained_staging"] is None
            and diagnostic["cleanup_error"] is None, "Real converter failure did not prove owned workspace cleanup")
    path = Path(diagnostic["record_path"])
    require(path.name == "failure.json" and path.parent.parent == base
            and path.parent.name == ".WinBookSplit-failed-" + diagnostic["run_id"], "Failure record escaped explicit base")
    tree = tree_identity(path.parent)
    require(set(tree["members"]) == {".WinBookSplit-owner.json", "failure.json"}, "Failure record membership differs")
    owner_raw = (path.parent / ".WinBookSplit-owner.json").read_bytes().decode("utf-8", "strict")
    require(json.loads(owner_raw) == {"schema_version": 1, "kind": "failed", "run_id": diagnostic["run_id"]},
            "Real converter failure owner marker differs")
    return {"path": str(path), "raw": path.read_bytes().decode("utf-8", "strict"), "tree": tree, "owner_raw": owner_raw}


def application_case(work, host, fmt, kind, source_template, base, prior, reference, calibre, renderer, *, canary=None):
    global LAST_CASE
    suffix = "auto1-keep" if kind == "retained" else kind
    identifier = host["id"] + "-" + fmt + "-" + suffix
    directory = work / identifier
    directory.mkdir()
    app, books, cwd = [directory / name for name in ("a", "b", "c")]
    copied = support.copy_application(app)
    books.mkdir()
    cwd.mkdir()
    source = books / ("Book Å 日本 [1] & %! $(literal)." + fmt)
    neighbor = source.with_suffix(".pdf")
    if kind == "malformed":
        source.write_bytes(("Authored malformed " + fmt.upper() + " bytes, not a valid ebook.\n").encode("ascii"))
    else:
        shutil.copyfile(source_template, source)
    shutil.copyfile(reference["observation"]["path"], neighbor)
    initial = protected(source, neighbor, prior)
    parameters = ["-InputFile", str(source), "-OutputDirectory", str(base), "-PythonPath", sys.executable,
                  "-CalibrePath", str(calibre), "-Mode", "Auto", "-BookmarkLevel", "1",
                  "-KeepConvertedPdf", "-NonInteractive"]
    command = [host["shell_executable"], "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "RemoteSigned",
               "-File", str(app / "WinBookSplit.ps1"), *parameters]
    config = behavior.CalibreEnvironment(directory / "settings")
    prior_before = prior_directories(base)
    LAST_CASE = {"id": identifier, "kind": kind, "input_path": str(source), "output_base": str(base),
                 "command": command, "cwd": str(cwd), "workspace": str(directory),
                 "source_observations_initial": initial, "prior_publications_before": prior_before}
    with paths.read_only_source(source):
        before = protected(source, neighbor, prior)
        with config.inherited():
            environment = config.environment(support.controlled_environment(cwd, host["shell_executable"]))
            try:
                with canary.conversion(identifier) if canary is not None else nullcontext():
                    observed = behavior.native(command, cwd, environment, timeout=180)
            except behavior.NativeFailure as error:
                LAST_CASE.update(error.process)
                raise
        LAST_CASE.update(observed, calibre_environment=config.receipt())
        final = cli.outcome(observed["stdout"])
        LAST_CASE["outcome"] = final
        expected_exit = 4 if kind == "malformed" else 0
        require(observed["exit_code"] == final["exit_code"] == expected_exit,
                "Actual native ebook outcome differs: " + identifier)
        log, records, processes, log_hash = cli.console_records(observed["stdout"], base)
        require(log is not None and len(records) == 1, "Actual ebook run did not retain one engine/log record")
        text = log.read_bytes().decode("utf-8", "strict")
        console = launchers.authenticate_console_manifest(log, final, source_path=source, source_observation=before["source"])
        engine = final["engine_result"]
        row = {**LAST_CASE, **observed, "host_id": host["id"], "host_version": host["host_version"],
            "shell_executable": host["shell_executable"], "actual_process": True, "input_kind": fmt,
            "parameters": parameters, "application_sha256": copied, "source_read_only_observed": True,
            "source_observations_before": before, "source_observations_after": protected(source, neighbor, prior),
            "outcome": final, "engine_records": records, "process_summaries": processes,
            "console_log": text, "console_log_sha256": log_hash, "console_path": str(log), "console_evidence": console,
            "engine_record": engine, "retained_pdf": None, "chapters": [], "publication": None,
            "failure_record": None, "prior_publications_after": None,
            "calibre_environment": config.receipt(), "canary_urls": [] if canary is None else canary.urls}
        LAST_CASE = row
        if kind == "malformed":
            require(engine["code"] == "conversion_failed" and engine["execution"] is None
                    and engine["written_count"] == 0 and engine.get("plan") is None,
                    "Real malformed ebook was not rejected before chapter planning/writing")
            row["failure_record"] = failure_record(engine, base)
            conversion_observation = engine["diagnostic"]["conversion"]
            require(type(conversion_observation["exit_code"]) is int and conversion_observation["exit_code"] != 0,
                    "Malformed ebook did not reach real Calibre nonzero failure")
            row["converter_process"] = conversion_observation
            new_names = {Path(row["failure_record"]["path"]).parent.name, log.parent.name}
        else:
            execution = engine["execution"]
            row["converter_process"] = execution["conversion"]
            row["loopback_diagnostics"] = {} if canary is None else {
                url: [{"stream": name, "line": line} for name in ("stdout_tail", "stderr_tail")
                      for line in execution["conversion"][name].splitlines() if url in line and "block" in line.lower()]
                for url in canary.urls}
            expected_vector = [page["content_sha256"] for page in reference["observation"]["pages"]] if kind == "retained" else None
            final_directory = Path(execution["final_directory"])
            require(final_directory.parent == base and execution["retained_intermediate"] is not None,
                    "Actual retained ebook run omitted its complete generated PDF")
            artifact = RENDER_ROOT / identifier
            artifact.mkdir()
            retained = final_directory / "WinBookSplit_Converted.pdf"
            row["retained_pdf"] = assets.observe_pdf(retained, renderer, artifact / "retained")
            actual_vector = [page["content_sha256"] for page in row["retained_pdf"]["pages"]]
            checked, members = conversion.check_publication(base, execution, expected_vector or actual_vector)
            row["checked_publication"] = checked
            row["chapters"] = [{"entry": entry, "observation": assets.observe_pdf(final_directory / entry["filename"], renderer,
                                  artifact / ("chapter-" + str(index)))} for index, entry in enumerate(execution["outputs"], 1)]
            tree = tree_identity(final_directory)
            require(set(tree["members"]) == members, "Actual retained publication authentication differs")
            row["publication"] = {"tree": tree,
                "owner_raw": (final_directory / ".WinBookSplit-owner.json").read_bytes().decode("utf-8", "strict"),
                "manifest_raw": (final_directory / "WinBookSplit_Manifest.json").read_bytes().decode("utf-8", "strict")}
            new_names = {final_directory.name, log.parent.name}
        row["prior_publications_after"] = prior_directories(base, excluded=new_names)
        require(row["prior_publications_after"] == prior_before and row["source_observations_before"] == row["source_observations_after"],
                "Ebook processing changed source, neighbor, prior file or earlier publication")
        row["application_unchanged"] = copied == {name: cli.digest(app / name) for name in copied}
        require(row["application_unchanged"], "Native ebook run changed copied shipped application")
        if validator is not None:
            validator.validate_case(row, cleanup_complete=False)
    row["source_observations_restored"] = protected(source, neighbor, prior)
    row["source_attributes_restored"] = row["source_observations_restored"] == initial
    require(row["source_attributes_restored"], "Owned ebook read-only attributes were not restored")
    COMPLETED_CASES.append(row)
    return row


def cleanup_case(row):
    base, source = Path(row["output_base"]), Path(row["input_path"])
    neighbor, prior = source.with_suffix(".pdf"), base / "prior-output.pdf"
    require(protected(source, neighbor, prior) == row["source_observations_initial"],
            "Cleanup refuses a changed protected ebook/neighbor/prior file")
    if row["publication"] is not None:
        remove_publication(base, row["engine_record"]["execution"], row["publication"]["tree"])
    if row["failure_record"] is not None:
        directory = Path(row["failure_record"]["path"]).parent
        actual = tree_identity(directory)
        require(actual == row["failure_record"]["tree"], "Failure record changed before authenticated cleanup")
        launchers.remove_known_directory(directory, base, set(actual["members"]), ".WinBookSplit-owner.json",
            json.loads(row["failure_record"]["owner_raw"]),
            expected_sha256={name: item["sha256"] for name, item in actual["members"].items()},
            expected_file_identities=actual["members"],
            expected_directory_identity={"device": actual["device"], "inode": actual["inode"]})
    console = row["console_evidence"]
    fresh = launchers.authenticate_console_manifest(Path(row["console_path"]), row["outcome"], source_path=source,
                                                   source_observation=row["source_observations_before"]["source"])
    require(fresh == console, "Console bytes or identity changed after original authenticated observation")
    launchers.remove_known_directory(Path(row["console_path"]).parent, base, set(console["members"]),
        ".WinBookSplit-console-owner.json", console["owner"],
        expected_sha256={name: item["sha256"] for name, item in console["file_identities"].items()},
        expected_file_identities=console["file_identities"], expected_directory_identity=console["directory_identity"])
    removed = not Path(row["console_path"]).parent.exists()
    if row["publication"] is not None:
        removed = removed and not Path(row["publication"]["tree"]["path"]).exists()
    if row["failure_record"] is not None:
        removed = removed and not Path(row["failure_record"]["path"]).parent.exists()
    require(removed, "Authenticated current-run output/console directory remains after cleanup")
    row.update(owned_outputs_removed=removed, source_observations_after_cleanup=protected(source, neighbor, prior), passed=True)


def characterize(work, shells, calibre, renderer):
    global BEHAVIOR, SOURCE_BEFORE, validator
    if validator is None:
        validator = load("wbs_ebooks_validator", ROOT / "tests/ebooks/validate_ebooks_report.py")
    require(os.name == "nt" and len(shells) == 2, "Actual Windows and both explicit supported hosts required")
    require(RENDER_ROOT is not None and RENDER_ROOT.is_absolute() and not RENDER_ROOT.exists()
            and not RENDER_ROOT.is_relative_to(ROOT) and not RENDER_ROOT.is_relative_to(work),
            "Persistent render root must be new and outside checkout/owned temporary workspace")
    COMPLETED_CASES.clear()
    behavior.NATIVE_OPERATIONS.clear()
    before = runner.source_manifest()
    SOURCE_BEFORE = before
    RENDER_ROOT.mkdir()
    tools_before = {"calibre": cli.identity(calibre), "renderer": cli.identity(renderer)}
    setup = behavior.CalibreEnvironment(work / "reference-settings")
    BEHAVIOR = {"stage": {"name": "fixture_generation", "state": "attempted", "directory": str(work / "fixtures")},
                "native_operations": behavior.NATIVE_OPERATIONS, "reference_config": setup.receipt()}
    try:
        with setup.inherited():
            provenance = assets.generate(work / "fixtures", calibre)
            BEHAVIOR["fixture_provenance"] = provenance
            sources = {item["format"]: work / "fixtures" / item["name"] for item in provenance["files"]}
            originals_before = {fmt: cli.identity(path) for fmt, path in sources.items()}
            BEHAVIOR["stage"] = {"name": "independent_profile_references", "state": "attempted",
                                 "inputs": {fmt: str(path) for fmt, path in sources.items()}, "render_root": str(RENDER_ROOT / "references")}
            references = assets.reference_conversions(sources, work, calibre, renderer, render_root=RENDER_ROOT / "references")
            BEHAVIOR["references"] = references
            BEHAVIOR["stage"] = {"name": "selected_format_help", "state": "attempted"}
            local_help = behavior.local_help(calibre, sources, work, setup.environment(history.clean_environment(work)))
            BEHAVIOR["local_help"] = local_help
    finally:
        BEHAVIOR["reference_config"] = setup.receipt()
        BEHAVIOR["asset_native_operations"] = getattr(assets, "NATIVE_OPERATIONS", [])
    host_probe = cli.process_tests.host_probe(work)
    hosts = [paths.host_observation(shell, work, host_probe) for shell in shells]
    require({host["id"] for host in hosts} == {"PS51", "PS7"}, "Two distinct supported PowerShell hosts required")
    for host in hosts:
        BEHAVIOR["stage"] = {"name": "application_matrix", "host_id": host["id"], "state": "attempted"}
        base = work / ("o-" + host["id"])
        base.mkdir()
        prior = base / "prior-output.pdf"
        shutil.copyfile(references["epub"]["tablet"]["observation"]["path"], prior)
        rows = []
        for fmt in ("epub", "azw3"):
            rows.append(application_case(work, host, fmt, "retained", sources[fmt], base, prior,
                                         references[fmt]["tablet"], calibre, renderer))
        # The second real publication must coexist with the first unchanged run.
        require(rows[1]["prior_publications_before"] and rows[1]["prior_publications_before"] == rows[1]["prior_publications_after"],
                "Same-base independent format run lacked immutable prior publication evidence")
        for row in rows:
            cleanup_case(row)
        require(sorted(path.name for path in base.iterdir()) == ["prior-output.pdf"], "Success cleanup left unexplained base members")
        for row in rows:
            row["output_members_after_cleanup"] = ["prior-output.pdf"]
            validator.validate_case(row)
        for fmt in ("epub", "azw3"):
            failure_base = work / ("f-" + host["id"] + "-" + fmt)
            failure_base.mkdir()
            failure_prior = failure_base / "prior-output.pdf"
            shutil.copyfile(prior, failure_prior)
            row = application_case(work, host, fmt, "malformed", sources[fmt], failure_base, failure_prior,
                                   references[fmt]["tablet"], calibre, renderer)
            cleanup_case(row)
            row["output_members_after_cleanup"] = sorted(path.name for path in failure_base.iterdir())
            validator.validate_case(row)
    observer = behavior.LoopbackCanary()
    BEHAVIOR["canary"] = observer.receipt()
    BEHAVIOR["stage"] = {"name": "loopback_canary", "state": "attempted"}
    try:
        observer.positive("before")
        canary_source = work / "authored-loopback-canary.epub"
        canary_asset = assets.authored_canary_epub(canary_source, observer.base_url, sources["epub"])
        for host in hosts:
            base = work / ("n-" + host["id"])
            base.mkdir()
            prior = base / "prior-output.pdf"
            shutil.copyfile(references["epub"]["tablet"]["observation"]["path"], prior)
            row = application_case(work, host, "epub", "canary", canary_source, base, prior,
                                   references["epub"]["tablet"], calibre, renderer, canary=observer)
            row["canary_asset"] = canary_asset
            cleanup_case(row)
            row["output_members_after_cleanup"] = sorted(path.name for path in base.iterdir())
            validator.validate_case(row)
        observer.positive("after")
    finally:
        observer.stop()
        BEHAVIOR["canary"] = observer.receipt()
    BEHAVIOR["canary_asset"] = canary_asset
    require(not BEHAVIOR["canary"]["conversion_requests"], "Real conversion requested an authored loopback resource")
    for host in hosts:
        host["policies_after"] = paths.host_observation(Path(host["shell_executable"]), work, host_probe)["stored_policies"]
    render_probe = behavior.native([renderer, "-v"], work, history.clean_environment(work), timeout=30)
    require(render_probe["exit_code"] == 0 and "pdftoppm version 26.07.0" in render_probe["stderr"], "Pinned native Poppler version differs")
    tools_after = {"calibre": cli.identity(calibre), "renderer": cli.identity(renderer)}
    after = runner.source_manifest()
    require(before == after and originals_before == {fmt: cli.identity(path) for fmt, path in sources.items()}
            and tools_before == tools_after, "Task source/original ebook/runtime executable changed during acceptance")
    BEHAVIOR["stage"] = {"name": "completed", "state": "passed"}
    return {"schema_version": 1, "task_id": "M4-T04", "result": "EBOOK_REGRESSION_PASSED", "success": True, "exit_code": 0,
            "observed_at": datetime.now(timezone.utc).isoformat(), "acceptance_ids": ["AC-076", "AC-077", "AC-078"],
            "tested_path_sha256": before, "source_unchanged": True, "host_cases": hosts, "cases": COMPLETED_CASES,
            "case_count": len(COMPLETED_CASES), "application_cases": 4, "failure_cases": 4, "canary_cases": 2,
            "fixture_provenance": provenance, "format_checks": provenance["format_checks"], "resource_closure": provenance["resource_closure"],
            "references": references, "behavior": BEHAVIOR, "original_inputs_before": originals_before,
            "original_inputs_after": {fmt: cli.identity(path) for fmt, path in sources.items()},
            "tools_before": tools_before, "tools_after": tools_after,
            "calibre": provenance["calibre"], "renderer": {"path": str(renderer), "sha256": cli.digest(renderer), "version": "26.07.0", "probe": render_probe},
            "render_directory": str(RENDER_ROOT), "render_artifacts_retained": True,
            "retained_render_files": {str(path): cli.identity(path) for path in sorted(RENDER_ROOT.rglob("*")) if path.is_file()},
            "input_neighbor_prior_unchanged": True, "machine_settings_unchanged": True,
            "human_or_GUI_tested": False, "environment": {"python": sys.version, "python_executable": sys.executable},
            "limits": ["Only authored self-contained EPUB and genuine generated unencrypted AZW3 are included.",
                       "Default-profile references are diagnostics; the shipped application always uses tablet.",
                       "Loopback observation is not OS network denial, a sandbox, or proof about arbitrary external converters/plugins.",
                       "No private documents, DRM removal, native ebook chapter output, installation, GUI, human or CI claims."]}


def main():
    global RENDER_ROOT, validator
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--report", required=True, type=Path)
    parser.add_argument("--shell-path", required=True, action="append", type=Path)
    parser.add_argument("--calibre-path", required=True, type=Path)
    parser.add_argument("--renderer-path", required=True, type=Path)
    parser.add_argument("--render-directory", required=True, type=Path)
    args = parser.parse_args()
    target = runner.new_external_path(args.report)
    temporary = tempfile.TemporaryDirectory(prefix="E-")
    work = Path(temporary.name).resolve()
    RENDER_ROOT = args.render_directory.resolve()
    try:
        require(sys.flags.isolated and sys.dont_write_bytecode, "Use pinned isolated developer Python -I -B")
        validator = load("wbs_ebooks_validator", ROOT / "tests/ebooks/validate_ebooks_report.py")
        shells, calibre, renderer = [path.resolve() for path in args.shell_path], args.calibre_path.resolve(), args.renderer_path.resolve()
        report = characterize(work, shells, calibre, renderer)
        report["owned_temp_removed"] = False
        validator.validate_ebooks_report(report, list(map(str, shells)), str(calibre), str(renderer), cleanup_complete=False,
                                        expected_source_sha256=report["tested_path_sha256"])
        # No reparse node may expand cleanup outside this exclusively authored tree.
        require(not history.is_reparse(work) and all(not history.is_reparse(path) for path in work.rglob("*")),
                "Authored temporary tree contains a reparse object; preserve it")
        temporary.cleanup()
        report["owned_temp_removed"] = not work.exists()
        validator.validate_ebooks_report(report, list(map(str, shells)), str(calibre), str(renderer),
                                        expected_source_sha256=report["tested_path_sha256"])
        with target.open("x", encoding="utf-8", newline="\n") as stream:
            json.dump(report, stream, ensure_ascii=True, indent=2)
            stream.write("\n")
        print("Ebook acceptance passed: " + str(report["case_count"]) + " actual application controls")
        return 0
    except BaseException as error:
        temporary._finalizer.detach()
        code = 130 if isinstance(error, KeyboardInterrupt) else error.code if isinstance(error, SystemExit) and type(error.code) is int and error.code != 0 else 1
        if isinstance(getattr(error, "process", None), dict):
            last = {"phase": "owned-native-timeout", "process": error.process, "attempted_case": LAST_CASE}
        else:
            last = LAST_CASE
        source_after = runner.source_manifest()
        failed = {"schema_version": 1, "task_id": "M4-T04", "result": "EBOOK_REGRESSION_FAILED", "success": False,
                  "exit_code": code, "error": str(error), "error_type": type(error).__name__, "cleanup_safe": False,
                  "workspace_retained": str(work) if work.exists() else None, "owned_temp_removed": not work.exists(),
                  "render_directory": str(RENDER_ROOT), "completed_cases": COMPLETED_CASES, "last_case": last,
                  "behavior": BEHAVIOR, "tested_path_sha256": SOURCE_BEFORE,
                  "tested_path_sha256_after_failure": source_after,
                  "source_unchanged": SOURCE_BEFORE is not None and SOURCE_BEFORE == source_after}
        with target.open("x", encoding="utf-8", newline="\n") as stream:
            json.dump(failed, stream, ensure_ascii=True, indent=2)
            stream.write("\n")
        print("Ebook acceptance failed; authored workspace retained: " + str(error), file=sys.stderr)
        return code


if __name__ == "__main__":
    raise SystemExit(main())
