"""M2-T03 real offline ebook conversion and owned intermediate acceptance."""

from __future__ import annotations

import argparse
from contextlib import ExitStack
from datetime import datetime, timezone
from hashlib import sha256
import json
import os
from pathlib import Path
import platform
import re
import shutil
import subprocess
import sys
import tempfile
from unittest.mock import patch

from pypdf import PdfReader, PdfWriter

ROOT = Path(__file__).resolve().parents[2]


def load(name, path):
    import importlib.util
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


manual = load("wbs_conversion_manual", ROOT / "tests/manual/characterize_manual.py")
launchers = load("wbs_conversion_launchers", ROOT / "tests/manual/current_launchers.py")
paths = load("wbs_conversion_paths", ROOT / "tests/paths/characterize_paths.py")
runner = load("wbs_conversion_runner", ROOT / "tests/run_tests.py")
fixtures = load("wbs_conversion_fixtures", ROOT / "tests/conversion/generate_ebook_fixtures.py")
history, require = manual.history, manual.require


def page_content(reader):
    result = []
    for page in reader.pages:
        stream = page.get_contents()
        result.append(sha256(stream.get_data() if stream is not None else b"").hexdigest())
    return result


def markers(reader):
    return ["WBS-PAGE-" + number for page in reader.pages
            for number in re.findall(r"(?<![A-Z0-9-])WBS-PAGE-([0-9]{3})(?![0-9])", page.extract_text() or "")]


def reference_pdf(source, target, calibre, work):
    observation = fixtures.run([calibre, source, target, "--output-profile", "tablet"], work)
    require(observation["exit_code"] == 0 and target.is_file(), "Independent actual ebook conversion failed: " + repr(observation))
    reader = PdfReader(target)
    require(markers(reader) == fixtures.CHAPTER_MARKERS, "Real converter lost or duplicated an authored chapter marker")
    require(len(reader.pages) == (3 if source.suffix == ".epub" else 4), "Original fixture's actual physical PDF pages changed")
    return {"process": observation, "path": str(target), "sha256": history.file_digest(target),
            "size_bytes": target.stat().st_size, "page_count": len(reader.pages), "page_content_sha256": page_content(reader)}


def original_neighbor(source, inputs):
    neighbor = source.with_suffix(".pdf")
    writer = PdfWriter()
    writer.add_blank_page(width=117, height=119)
    with neighbor.open("xb") as stream:
        writer.write(stream)
    inputs[str(source)] = history.file_digest(source)
    inputs[str(neighbor)] = history.file_digest(neighbor)
    return neighbor


def check_publication(base, execution, expected_content):
    """Conversion-only exact member adapter; existing PDF helper stays strict."""
    final = manual.published_directory(base, execution)
    retained = execution.get("retained_intermediate", "missing")
    require(retained != "missing", "Ebook result omitted explicit retention record")
    chapter_names = [entry["filename"] for entry in execution["outputs"]]
    names = {".WinBookSplit-owner.json", "WinBookSplit_Manifest.json", *chapter_names}
    if retained is not None:
        require(retained["filename"] == "WinBookSplit_Converted.pdf", "Retained intermediate used an unexpected name")
        names.add(retained["filename"])
    require(len(chapter_names) == len(set(name.casefold() for name in chapter_names)) and not set(chapter_names) & {"WinBookSplit_Converted.pdf"},
            "Retained intermediate collided with a chapter")
    require({path.name for path in final.iterdir()} == names
            and all(path.is_file() and not history.is_reparse(path) for path in final.iterdir()), "Publication has unexpected or unsafe members")
    owner = json.loads((final / ".WinBookSplit-owner.json").read_text(encoding="utf-8"))
    require(owner == {"schema_version": 1, "kind": "run", "run_id": execution["run_id"]}, "Conversion output owner differs")
    manifest = json.loads((final / "WinBookSplit_Manifest.json").read_text(encoding="utf-8"))
    require(manifest == execution["manifest"], "Actual complete manifest differs from execution")
    for key in ("schema_version", "status", "run_id", "final_directory", "mode", "total_pages", "source_identity", "coverage", "written_count", "outputs",
                "original_ebook_identity", "conversion", "retained_intermediate"):
        require(manifest[key] == execution[key], "Conversion manifest field differs: " + key)
    outputs, actual_content, chapter_markers = [], [], []
    for entry in execution["outputs"]:
        target = final / entry["filename"]
        reader = PdfReader(target)
        content = page_content(reader)
        start, end = entry["start"], entry["end"]
        require(content == expected_content[start:end] and len(reader.pages) == end - start == entry["page_count"],
                "A real converted slice lost, reordered or duplicated physical source content")
        require(history.file_digest(target) == entry["sha256"] and target.stat().st_size == entry["size_bytes"], "Chapter bytes differ from manifest")
        outputs.append({"filename": entry["filename"], "range": [start, end], "page_ids": list(range(start + 1, end + 1)),
                        "sha256": entry["sha256"], "size_bytes": entry["size_bytes"], "page_count": len(reader.pages)})
        actual_content.extend(content)
        chapter_markers.extend(markers(reader))
    require(actual_content == expected_content and chapter_markers == fixtures.CHAPTER_MARKERS, "Converted physical coverage or chapter markers differ")
    retained_sha = retained_content = False
    if retained is not None:
        target = final / retained["filename"]
        reader = PdfReader(target)
        require(history.file_digest(target) == retained["sha256"] and target.stat().st_size == retained["size_bytes"]
                and len(reader.pages) == retained["page_count"] == execution["total_pages"], "Retained full PDF differs from captured conversion bytes")
        require(page_content(reader) == expected_content and markers(reader) == fixtures.CHAPTER_MARKERS, "Retained full PDF content differs")
        retained_sha = retained_content = True
    return {"outputs": outputs, "page_content_sha256": actual_content, "chapter_markers": chapter_markers,
            "retained_sha256_verified": retained_sha, "retained_page_content_verified": retained_content,
            "retained_absent_verified": retained is None, "manifest_validated": True}, names


def remove_publication(base, execution, members):
    launchers.remove_known_directory(Path(execution["final_directory"]), base, members, ".WinBookSplit-owner.json",
                                    {"schema_version": 1, "kind": "run", "run_id": execution["run_id"]})


def remove_diagnostic(base, frame):
    diagnostic = frame.get("diagnostic")
    if diagnostic is None:
        return
    require(diagnostic["cleanup_complete"] is True and diagnostic["retained_staging"] is None
            and diagnostic["cleanup_error"] is None, "Ordinary converter failure retained a workspace")
    target = Path(diagnostic["record_path"])
    require(target.name == "failure.json" and target.parent.parent == base
            and target.parent.name == ".WinBookSplit-failed-" + diagnostic["run_id"], "Converter diagnostic escaped its explicit base")
    owner = {"schema_version": 1, "kind": "failed", "run_id": diagnostic["run_id"]}
    require(json.loads((target.parent / ".WinBookSplit-owner.json").read_text(encoding="utf-8")) == owner, "Converter failure marker differs")
    launchers.remove_known_directory(target.parent, base, {".WinBookSplit-owner.json", "failure.json"}, ".WinBookSplit-owner.json", owner)


def parity_record(source, base, execution, preview, checked, reference):
    conversion = execution["conversion"]
    generated = conversion["generated_pdf_identity"]
    require(generated["page_content_sha256"] == reference["page_content_sha256"], "Actual conversion metadata differs from independent actual conversion content")
    require(not Path(generated["path"]).exists(), "Completed conversion left its owned intermediate workspace file")
    record = {"passed": True, "format": source.suffix[1:], "mode": execution["mode"], "total_pages": execution["total_pages"],
              "preview_entries": preview, "coverage": execution["coverage"], "source_identity": execution["source_identity"],
              "original_ebook_identity": execution["original_ebook_identity"], "conversion": conversion,
              "writer_result": execution, "written_count": execution["written_count"],
              "keep_converted_pdf": execution["retained_intermediate"] is not None, "retained_intermediate": execution["retained_intermediate"],
              "original_page_content_sha256": generated["page_content_sha256"],
              "converted_reference_page_content_sha256": reference["page_content_sha256"],
              "input_sha256": history.file_digest(source), "output_base": str(base), "input_unchanged": True,
              "neighbor_unchanged": True, "no_replanning_or_reopening": True, "workspace_absent_verified": True,
              **checked}
    return record


def direct_case(fmt, engine, work, source, base, calibre, reference, inputs):
    keep = fmt == "azw3"
    neighbor = original_neighbor(source, inputs)
    before = {path: history.file_digest(Path(path)) for path in (str(source), str(neighbor))}
    with paths.read_only_source(source):
        prepared = engine.prepare_ebook(source, "manual", "2,3", output_base=base, calibre_path=calibre, keep_converted_pdf=keep)
        preview = manual.plain(engine.preview_plan(prepared))
        captured = page_content(prepared._reader)
        captured_bytes = prepared._reader.stream.getvalue()
        require(sha256(captured_bytes).hexdigest() == preview["conversion"]["generated_pdf_identity"]["sha256"]
                and len(captured_bytes) == preview["conversion"]["generated_pdf_identity"]["size_bytes"], "Captured actual conversion bytes differ from their independent reader snapshot")
        require(captured == reference["page_content_sha256"] == preview["conversion"]["generated_pdf_identity"]["page_content_sha256"],
                "Independent captured reader does not match actual converter metadata and reference")
        counters = {"planner_calls_during_execute": 0, "source_reader_calls_during_execute": 0}
        captured_reader = engine.PdfReader
        source_paths = {source.resolve(), Path(preview["source_identity"]["path"]).resolve()}

        def no_planner(*args, **kwargs):
            counters["planner_calls_during_execute"] += 1
            raise RuntimeError("Ebook execution tried to prepare or replan")

        def guarded_reader(target, *args, **kwargs):
            if isinstance(target, (str, os.PathLike)) and Path(target).resolve() in source_paths:
                counters["source_reader_calls_during_execute"] += 1
                raise RuntimeError("Ebook execution tried to reopen a captured source")
            return captured_reader(target, *args, **kwargs)

        with ExitStack() as stack:
            for name in ("prepare_ebook", "prepare_split", "plan_manual_starts", "plan_level1", "plan_level2"):
                stack.enter_context(patch.object(engine, name, side_effect=no_planner))
            stack.enter_context(patch.object(engine, "PdfReader", side_effect=guarded_reader))
            execution = manual.plain(engine.execute_split(prepared, base))
        require(manual.plain(engine.preview_plan(prepared)) == preview, "Ebook execution changed the frozen preview")
    checked, members = check_publication(base, execution, captured)
    record = parity_record(source, base, execution, preview["entries"], checked, reference)
    require(before == {path: history.file_digest(Path(path)) for path in before}, "Direct conversion modified ebook or same-name neighbor")
    remove_publication(base, execution, members)
    require(not any(base.iterdir()), "Direct conversion left unexplained output-base members")
    record.update(id=fmt, **counters, owned_outputs_removed=True,
                  captured_pdf_bytes_verified=True,
                  page_content_observation="independent-captured-reader-and-real-conversion-and-slices",
                  no_replanning_observation="direct-api-planner-and-source-reopen-traps")
    return record


def make_fake_converter(work, ps51):
    source, executable = work / "FakeConverter.cs", work / "FakeConverter.exe"
    zero = work / "zero.pdf"
    with zero.open("xb") as stream:
        PdfWriter().write(stream)
    source.write_text('''using System;
using System.IO;
public class FakeConverter {
    public static int Main(string[] args) {
        if (args.Length == 1 && args[0] == "--version") { Console.WriteLine("ebook-convert.exe (calibre 9.15.0)"); return 0; }
        if (args.Length < 2) return 3;
        string mode = Environment.GetEnvironmentVariable("WBS_FAKE_MODE");
        if (mode == "empty") File.WriteAllBytes(args[1], new byte[0]);
        if (mode == "corrupt") File.WriteAllText(args[1], "Authored corrupt PDF bytes\\n");
        if (mode == "zero-page") File.WriteAllBytes(args[1], File.ReadAllBytes(Environment.GetEnvironmentVariable("WBS_ZERO_PDF")));
        File.WriteAllText(Environment.GetEnvironmentVariable("WBS_FAKE_RECEIPT"), mode + "\\n" + args[1] + "\\n0\\n");
        Console.WriteLine("WBS-FAKE-NATIVE-EXIT-0");
        return 0;
    }
}''', encoding="utf-8")
    environment = history.clean_environment(work)
    environment.update(WBS_FAKE_SOURCE=str(source), WBS_FAKE_EXE=str(executable))
    command = [str(ps51), "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "RemoteSigned", "-Command",
               "Add-Type -Path $env:WBS_FAKE_SOURCE -OutputAssembly $env:WBS_FAKE_EXE -OutputType ConsoleApplication"]
    compilation = history.run(command, work, environment=environment)
    require(compilation["exit_code"] == 0 and executable.is_file(), "Authored native converter control did not compile: " + repr(compilation))
    return executable, zero, {"command": command, "cwd": str(work), **compilation,
                              "source_sha256": history.file_digest(source), "executable_sha256": history.file_digest(executable)}


def launcher_case(identifier, shell, source, base, cwd, wrapper, calibre, reference, inputs, *, keep=False, fake_kind=None, batch=False, zero=None):
    system = Path(os.environ["SystemRoot"])
    ps51 = system / "System32/WindowsPowerShell/v1.0/powershell.exe"
    environment = history.clean_environment(cwd)
    environment.update(PATH=os.pathsep.join((str(Path(sys.executable).parent), str(ps51.parent), str(system / "System32"), str(system))),
                       PSModulePath=str(shell.parent / "Modules"), WBS_CONVERSION_SCRIPT=str(ROOT / "WinBookSplit.ps1"),
                       WBS_CONVERSION_INPUT=str(source), WBS_CONVERSION_BASE=str(base), WBS_CONVERSION_CALIBRE=str(calibre),
                       WBS_CONVERSION_KEEP="1" if keep else "0", WBS_CONVERSION_BAT=str(ROOT / "WinBookSplit.bat"))
    fake_receipt = cwd / "fake-receipt.txt"
    if fake_kind is not None:
        environment.update(WBS_FAKE_MODE=fake_kind, WBS_ZERO_PDF=str(zero), WBS_FAKE_RECEIPT=str(fake_receipt))
    command = [str(system / "System32/cmd.exe"), "/d", "/v:off", "/c", str(wrapper)] if batch else \
        [str(shell), "-NoProfile", "-ExecutionPolicy", "RemoteSigned", "-File", str(wrapper)]
    neighbor = original_neighbor(source, inputs)
    before = {str(path): history.file_digest(path) for path in (source, neighbor)}
    log = marker = execution = members = None
    cleanup_safe = True
    answers = "M\nN\n2,3\nY\nN\n\n" if fake_kind is None and not batch else "M\n"
    try:
        with paths.read_only_source(source):
            readonly_observed = bool(source.lstat().st_file_attributes & 1)
            try:
                process = history.run_entrypoint(command, cwd, environment=environment, stdin=answers, timeout=120)
            except history.EntryPointFailure as error:
                cleanup_safe = error.cleanup_safe
                raise
        frames = []
        if any(line.startswith("Log: ") for line in process["stdout"].splitlines()):
            log = launchers.exact_line(process["stdout"], "Log: ")
            require(log.is_absolute() and log.name == "console.log" and log.parent.parent == base
                    and re.fullmatch(r"\.WinBookSplit-console-[0-9a-f]{32}", log.parent.name), "Conversion console escaped explicit base")
            marker = json.loads((log.parent / ".WinBookSplit-console-owner.json").read_text(encoding="utf-8"))
            require(marker == {"run_id": log.parent.name.removeprefix(".WinBookSplit-console-"), "kind": "console"}, "Conversion console ownership differs")
            frames = [json.loads(line) for line in log.read_text(encoding="utf-8").splitlines() if line.startswith("{")]
        success = fake_kind is None and not batch
        expected_exit = 0 if success else 3 if batch else 4
        require(process["exit_code"] == expected_exit, "Actual conversion launcher exit differs: " + identifier + ": " + repr(process))
        if success:
            require(len(frames) == 1 and frames[0]["status"] == "success" and frames[0]["exit_code"] == 0, "Real conversion did not produce one successful result")
            execution = frames[0]["execution"]
            require(launchers.exact_line(process["stdout"], "Output: ") == Path(execution["final_directory"]), "Emitted conversion output differs from manifest")
            checked, members = check_publication(base, execution, reference["page_content_sha256"])
            preview = [{key: value for key, value in entry.items() if key not in {"sha256", "size_bytes", "page_count"}} for entry in execution["outputs"]]
            record = parity_record(source, base, execution, preview, checked, reference)
            record.update(engine_record=frames[0], source_read_only_attribute_observed=readonly_observed, source_attributes_restored=True,
                          page_content_observation="actual-converter-metadata-compared-with-independent-real-conversion-and-slices",
                          no_replanning_observation="supported-by-direct-api-and-shared-plan-controls",
                          preview_observation="expected-manual-ranges-and-execution-metadata; independent-preview-tested-directly")
        else:
            require(not any(line.startswith("Output: ") for line in process["stdout"].splitlines())
                    and "Done." not in process["stdout"] and "[Writing]" not in process["stdout"], "Failed converter announced output or success")
            record = {"outputs": [], "successful_final_count": 0, "no_success_summary": True, "workspace_absent_verified": True}
            if not batch:
                require(len(frames) == 1 and frames[0]["code"] == "conversion_output_invalid" and frames[0]["written_count"] == 0
                        and frames[0]["execution"] is None, "Zero-exit invalid PDF was not rejected by the conversion category: " + repr(frames))
                converter_observation = frames[0]["diagnostic"]["conversion"]
                require(converter_observation["exit_code"] == 0
                        and "WBS-FAKE-NATIVE-EXIT-0" in converter_observation["stdout_tail"], "Actual parent did not observe the native converter's zero status and final control")
                receipt = fake_receipt.read_text(encoding="utf-8").splitlines()
                require(receipt[0] == fake_kind and receipt[2] == "0" and not Path(receipt[1]).exists(), "Authored zero-exit converter did not reach its final native control or clean its workspace")
                record.update(engine_record=frames[0], converter_exit_code=0, converter_native_observation="actual-parent-record-and-authored-final-return-control",
                              converter_process=converter_observation, fake_receipt=receipt)
                remove_diagnostic(base, frames[0])
            else:
                record["scope"] = "actual-unchanged-BAT-missing-trusted-converter; successful-discovery-separately-covered-by-runtime-acceptance"
                require("Calibre" in process["stdout"] or "calibre" in process["stdout"], "BAT failure did not identify the missing conversion dependency")
                dependency_errors = [json.loads(line.removeprefix("[DEPENDENCY-ERROR] ")) for line in process["stdout"].splitlines()
                                     if line.startswith("[DEPENDENCY-ERROR] ")]
                require(log is None and frames == [] and len(dependency_errors) == 1
                        and dependency_errors[0].get("Code") == "converter_not_found"
                        and isinstance(dependency_errors[0].get("Attempts"), list)
                        and all(item.get("Accepted") is False and item.get("Probe") is None for item in dependency_errors[0]["Attempts"]),
                        "Actual BAT missing converter did not fail in dependency preflight before console/processing")
                require("Calibre 9.15.0" in process["stdout"] and "-CalibrePath" in process["stdout"]
                        and all(marker not in process["stdout"] for marker in ("[ENGINE]", "[CONVERSION]", "[DEPENDENCY] ")),
                        "Actual BAT missing converter omitted precise guidance or started processing")
                record.update(dependency_error=dependency_errors[0], preflight_before_console_output=True,
                              console_record_absent=True, engine_invocation_record_absent=True,
                              dependency_success_record_absent=True, conversion_started=False,
                              written_count=0, setup_guidance_verified=True)
        require(before == {path: history.file_digest(Path(path)) for path in before}, "Conversion changed original ebook or same-basename neighboring PDF")
        record.update(id=identifier, passed=True, actual_process=True, shell_executable=str(shell), command=command, cwd=str(cwd),
                      input_unchanged=True, neighbor_unchanged=True, stdin_utf8=answers, **process)
        if log is not None:
            text = log.read_text(encoding="utf-8")
            interaction = launchers.interactions.capture(text)
            summaries = [json.loads(line[len("[PROCESS] "):]) for line in text.splitlines() if line.startswith("[PROCESS] ")]
            launchers.interactions.validate(interaction, summaries,
                ["input_ready", "plan_ready"] if success else [],
                ["starts", "execute"] if success else [], starts="2,3", execution=execution)
            record.update(interaction=interaction, process_summaries=summaries)
    finally:
        if cleanup_safe:
            if execution is not None and members is not None:
                remove_publication(base, execution, members)
            if log is not None and marker is not None:
                launchers.remove_known_directory(log.parent, base, {".WinBookSplit-console-owner.json", "console.log"}, ".WinBookSplit-console-owner.json", marker)
    if not batch:
        require(not any(base.iterdir()), "Conversion left unexplained output-base members")
    require(log is None or not log.parent.exists(), "Owned conversion console remains")
    record["owned_outputs_removed"] = True
    return record


def characterize(work, shells, calibre):
    require(os.name == "nt", "Actual Windows conversion acceptance is required")
    manual.trusted_original_sources()
    before = runner.source_manifest()
    engine = load("wbs_current_conversion_engine", ROOT / "engine/winbooksplit_engine.py")
    provenance = fixtures.generate(work / "f", calibre)
    inputs = {str(work / "f" / item["name"]): item["sha256"] for item in provenance["files"]}
    source_by_format = {fmt: work / "f" / ("original-three-chapters." + fmt) for fmt in ("epub", "azw3")}
    references = {fmt: reference_pdf(source, work / (fmt + ".pdf"), calibre, work) for fmt, source in source_by_format.items()}
    probe = work / "Host.ps1"
    probe.write_text("[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)\n"
        "$total=0; $checked=@('WinBookSplit.ps1','engine/WinBookSplit.Paths.ps1','engine/WinBookSplit.Diagnostics.ps1')\n"
        "foreach($name in $checked){$tokens=$null;$errors=$null;[Management.Automation.Language.Parser]::ParseFile((Join-Path $env:WBS_PATHS_ROOT $name),[ref]$tokens,[ref]$errors)|Out-Null;$total+=$errors.Count}\n"
        "$policies=@(Get-ExecutionPolicy -List|Where-Object {$_.Scope -ne 'Process'}|ForEach-Object {[pscustomobject]@{scope=$_.Scope.ToString();policy=$_.ExecutionPolicy.ToString()}})\n"
        "[pscustomobject]@{host_major=$PSVersionTable.PSVersion.Major;host_version=$PSVersionTable.PSVersion.ToString();syntax_checked=$checked;syntax_error_count=$total;stored_policies=$policies}|ConvertTo-Json -Depth 6 -Compress\n", encoding="utf-8")
    hosts = [paths.host_observation(shell, work, probe) for shell in shells]
    require({host["id"] for host in hosts} == {"PS51", "PS7"}, "Both actual supported hosts are mandatory")
    by_host = {host["id"]: Path(host["shell_executable"]) for host in hosts}
    wrapper = work / "Invoke.ps1"
    wrapper.write_text("[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)\n"
                       "$keep=$env:WBS_CONVERSION_KEEP -eq '1'\n"
                       "& $env:WBS_CONVERSION_SCRIPT -InputFile $env:WBS_CONVERSION_INPUT -OutputDirectory $env:WBS_CONVERSION_BASE -CalibrePath $env:WBS_CONVERSION_CALIBRE -KeepConvertedPdf:$keep\n"
                       "exit $LASTEXITCODE\n", encoding="utf-8")
    direct, real, invalid = [], [], []
    index = 0

    def owned_case(fmt):
        nonlocal index
        index += 1
        cwd = work / ("c" + str(index))
        cwd.mkdir()
        source = cwd / ("Book [1] O'Neil & (章节) %!." + fmt)
        source.write_bytes(source_by_format[fmt].read_bytes())
        base = work / ("b" + str(index))
        base.mkdir()
        return source, base, cwd

    for fmt in ("epub", "azw3"):
        source, base, cwd = owned_case(fmt)
        direct.append(direct_case(fmt, engine, work, source, base, calibre, references[fmt], inputs))
    for host_id in ("PS51", "PS7"):
        for fmt in ("epub", "azw3"):
            for keep in (False, True):
                source, base, cwd = owned_case(fmt)
                identifier = host_id + "-" + fmt + "-" + ("keep" if keep else "default")
                real.append(launcher_case(identifier, by_host[host_id], source, base, cwd, wrapper, calibre,
                                          references[fmt], inputs, keep=keep))
    fake, zero, compilation = make_fake_converter(work, by_host["PS51"])
    for host_id in ("PS51", "PS7"):
        for kind in ("missing", "empty", "corrupt", "zero-page"):
            source, base, cwd = owned_case("epub")
            invalid.append(launcher_case(host_id + "-" + kind, by_host[host_id], source, base, cwd, wrapper, fake,
                                         None, inputs, fake_kind=kind, zero=zero))
    environment = history.clean_environment(work)
    observed = history.run([str(by_host["PS51"]), "-NoProfile", "-Command", "[Environment]::GetFolderPath('MyDocuments')"], work, environment=environment)
    require(observed["exit_code"] == 0 and observed["stdout"].strip(), "Actual BAT Documents base observation failed")
    documents = Path(observed["stdout"].strip()).resolve(strict=True)
    require(documents.is_dir() and not history.is_reparse(documents), "BAT default base is not an ordinary directory")
    batch_wrapper = work / "Invoke.cmd"
    batch_wrapper.write_text('@echo off\nsetlocal DisableDelayedExpansion\n"%WBS_CONVERSION_BAT%" "%WBS_CONVERSION_INPUT%"\n', encoding="ascii", newline="\r\n")
    source, unused_base, cwd = owned_case("epub")
    bat = launcher_case("BAT-missing-converter", by_host["PS51"], source, documents, cwd, batch_wrapper, calibre, None, inputs, batch=True)
    for host in hosts:
        after = paths.host_observation(Path(host["shell_executable"]), work, probe)
        require(after["stored_policies"] == host["stored_policies"], "Conversion changed stored execution policies")
        host["policies_after"] = after["stored_policies"]
    require(inputs == {path: history.file_digest(Path(path)) for path in inputs}, "Conversion acceptance changed an original fixture/input/neighbor")
    observation = history.probe_import(work)
    require(before == runner.source_manifest(), "Source changed during conversion acceptance")
    return {"schema_version": 1, "task_id": "M2-T03", "result": "CONVERSION_REGRESSION_PASSED", "success": True, "exit_code": 0,
            "acceptance_ids": runner.CONVERSION_ACCEPTANCE_IDS, "observed_at": datetime.now(timezone.utc).isoformat(),
            "calibre": {**provenance["calibre"], "size_bytes": calibre.stat().st_size}, "fixture_provenance": provenance,
            "reference_conversions": references, "host_cases": hosts, "real_cases": real, "direct_api_cases": direct,
            "invalid_converter_cases": invalid, "fake_converter_compilation": compilation, "bat_missing_converter_case": bat,
            "source_unchanged": True, "baseline_guards_preserved": True, "input_and_neighbor_unchanged": True,
            "machine_settings_unchanged": True, "immutable_original_commit": history.ORIGINAL_COMMIT,
            "tested_path_sha256": before, "input_sha256": inputs, "import_observation": observation,
            "environment": {"python": sys.version, "python_executable": sys.executable, "platform": platform.platform(), "pypdf": history.pypdf.__version__},
            "not_run": ["successful BAT converter discovery (M2-T04)", "Explorer drag/drop", "Calibre GUI", "network resource sandbox", "DRM removal", "release package"],
            "limits": ["Default launcher source-content vectors are converter metadata crosschecked with a separately generated real PDF and every output slice.",
                       "Direct API independently reads the captured source reader; keep cases independently reopen and hash the actual retained full PDF."]}


def main():
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
    sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--report", type=Path, required=True)
    parser.add_argument("--shell-path", type=Path, action="append", required=True)
    parser.add_argument("--calibre-path", type=Path, required=True)
    args = parser.parse_args()
    try:
        require(sys.flags.isolated and sys.dont_write_bytecode, "Use the explicit developer Python -I -B")
        report_path = runner.new_external_path(args.report)
        require(len(args.shell_path) == 2 and all(path.is_absolute() and path.is_file() for path in args.shell_path), "Both absolute actual supported shell hosts required")
        require(args.calibre_path.is_absolute() and args.calibre_path.is_file(), "An existing absolute actual pinned converter path is required")
        shells, calibre = [path.resolve() for path in args.shell_path], args.calibre_path.resolve()
        require(len(set(shells)) == 2, "Actual supported hosts must be distinct")
        with tempfile.TemporaryDirectory(prefix="C-") as directory:
            report = characterize(Path(directory).resolve(), shells, calibre)
        report["owned_temp_removed"] = not Path(directory).exists()
        runner.validate_conversion_report(report, [str(shell) for shell in shells], str(calibre))
        with report_path.open("x", encoding="utf-8", newline="\n") as stream:
            json.dump(report, stream, indent=2, ensure_ascii=True)
            stream.write("\n")
        print("Conversion acceptance passed: eight real PS format/retention runs, two captured-reader controls, eight invalid zero-exit converters and actual BAT dependency failure")
        return 0
    except (OSError, RuntimeError, ValueError, subprocess.TimeoutExpired) as error:
        print("Conversion acceptance failed: " + str(error), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
