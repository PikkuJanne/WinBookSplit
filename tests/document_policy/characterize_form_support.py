"""Native application regression for the intended static AcroForm behavior."""
from __future__ import annotations

import argparse
import importlib.util
import json
import os
from pathlib import Path
import platform
import re
import shutil
import subprocess
import sys
import tempfile

from PIL import Image, ImageChops
from pypdf import PdfReader, __version__ as pypdf_version
from pypdf.generic import ArrayObject, DictionaryObject, IndirectObject, StreamObject

ROOT = Path(__file__).resolve().parents[2]


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, ROOT / path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


policy = load("wbs_form_native_policy", "tests/document_policy/characterize_document_policy.py")
fixtures = load("wbs_form_fixtures", "tests/fixtures/generate_form_fixtures.py")
cli, require = policy.cli, policy.require
MATRIX = (
    ("PS51", "plain", "Manual"),
    ("PS7", "plain", "Auto2"),
    ("BAT", "plain", "Manual"),
    ("PS51", "acroform_empty", "Auto1"),
    ("PS7", "text", "Auto2"),
    ("BAT", "controls", "Manual"),
    ("PS51", "orphan_widget", "Manual"),
)
RANGES = {"Manual": [[0, 2], [2, 4]], "Auto1": [[0, 1], [1, 2], [2, 3], [3, 4]],
          "Auto2": [[0, 1], [1, 2], [2, 3], [3, 4]]}
PAGE_MARKER = re.compile(rb"SOURCE-PAGE-(\d+)\b")


def render(renderer, source, prefix, receipts):
    process = cli.run_unavailable_stdin(
        [str(renderer), "-r", "72", "-cropbox", "-png", str(source), str(prefix)],
        prefix.parent, os.environ.copy(),
    )
    receipts.append(process)
    require(process["exit_code"] == 0, "Poppler render failed: " + process["stderr"])
    paths = prefix.parent.glob(prefix.name + "-*.png")
    return sorted(paths, key=lambda path: int(path.stem.rsplit("-", 1)[1]))


def geometry(page):
    return list(page.mediabox), list(page.cropbox), page.rotation


def region(image, rect):
    # Checkpoint 1 seeds use an unrotated, zero-origin CropBox and render at 72 DPI.
    left, bottom, right, top = rect
    return image.crop((left, image.height - top, right, image.height - bottom))


def check_reference_oracle(renderer, work, sources, receipts):
    """Prove the pixel check distinguishes missing artwork and opposite states."""
    if "controls" not in sources:
        return None
    paths, _, fields = sources["controls"]
    renders = render(renderer, paths["static"], work / "controls-oracle", receipts)
    tiles = []
    for field in fields:
        with Image.open(renders[field["page"] - 1]) as image:
            tiles.append(region(image.convert("RGB"), field["rect_pdf"]))
    for tile in tiles:
        require(ImageChops.difference(tile, Image.new("RGB", tile.size, "white")).getbbox() is not None,
                "Field reference cannot detect dropped artwork")
    # The fixture orders the checkbox pair, then the radio pair, then a choice.
    for first, second in ((0, 1), (2, 3)):
        require(tiles[first].size == tiles[second].size
                and ImageChops.difference(tiles[first], tiles[second]).getbbox() is not None,
                "Field reference cannot distinguish opposite button states")
    return {"blank_regions_detected": len(tiles), "opposite_button_pairs_detected": 2}


def inspect_serialization(reader, chapter, pages, outside_values):
    """Count every serialized page and inspect streams even if unreachable."""
    ids = {(generation, number) for generation, rows in reader.xref.items() for number in rows if number}
    ids.update((0, number) for number in reader.xref_objStm)
    markers, page_count = set(), 0
    for generation, number in sorted(ids):
        value = reader.get_object(IndirectObject(number, generation, reader))
        if isinstance(value, DictionaryObject) and value.get("/Type") == "/Page":
            page_count += 1
        pending = [value]
        while pending:
            value = pending.pop()
            if isinstance(value, DictionaryObject):
                require(value.get("/Subtype") != "/Widget" and not {"/AcroForm", "/XFA", "/FT"}.intersection(value),
                        "Serialized form/widget data survived chapter export")
                pending.extend(value.values())
                if isinstance(value, StreamObject):
                    data = value.get_data()
                    markers.update(int(marker) for marker in PAGE_MARKER.findall(data))
                    require(all(item.encode("ascii") not in data for item in outside_values),
                            "Outside-chapter value survived in a decoded stream")
            elif isinstance(value, ArrayObject):
                pending.extend(value)
            elif isinstance(value, str):
                require(all(item not in value for item in outside_values),
                        "Outside-chapter value survived in a serialized string")
    require(page_count == len(pages) and markers == set(pages), "Serialized pages/content escaped the chapter range")
    return {"page_objects": page_count, "page_markers": sorted(markers), "sha256": cli.digest(chapter)}


def inspect_outputs(renderer, case_dir, base, result, case, reference, values, regions, source_hash):
    execution = result["execution"]
    start_end = [[row["start"], row["end"]] for row in execution["outputs"]]
    require(start_end == RANGES[case["mode"]] and result["written_count"] == len(start_end),
            "Wrong physical chapter partition")
    require(result["plan"]["source_identity"]["sha256"] == source_hash, "Plan source hash changed")
    final = Path(execution["final_directory"])
    members = {".WinBookSplit-owner.json", "WinBookSplit_Manifest.json", *[row["filename"] for row in execution["outputs"]]}
    require(final.parent == base and {path.name for path in final.iterdir()} == members, "Unexpected chapter directory contents")
    expected = PdfReader(reference)
    reference_renders = render(renderer, reference, case_dir / "reference", case["renders"])
    require(len(reference_renders) == len(expected.pages) == 4, "Wrong reference page count")
    all_values = {item for labels in values.values() for item in labels}
    for row in execution["outputs"]:
        start, end = row["start"], row["end"]
        chapter = final / row["filename"]
        reader = PdfReader(chapter)
        pages = list(range(start + 1, end + 1))
        chapter_renders = render(renderer, chapter, case_dir / f"chapter-{start}", case["renders"])
        require(len(reader.pages) == len(chapter_renders) == len(pages), "Wrong reopened chapter page count")
        inside_values = {item for page in pages for item in values.get(str(page), ())}
        record = {"range": [start, end], "regions": [], **inspect_serialization(reader, chapter, pages, all_values - inside_values)}
        case["outputs"].append(record)
        for local, page_number in enumerate(pages):
            page = reader.pages[local]
            text = page.extract_text() or ""
            require(PAGE_MARKER.findall(text.encode("utf-8")) == [str(page_number).encode("ascii")],
                    "Missing, reordered or foreign page marker")
            require(all(item in text for item in values.get(str(page_number), ())), "Expected visible form text is missing")
            require(geometry(page) == geometry(expected.pages[page_number - 1]), "Page geometry changed")
            with Image.open(reference_renders[page_number - 1]) as expected_image, Image.open(chapter_renders[local]) as observed_image:
                require(expected_image.size == observed_image.size, "Rendered page size changed")
                target, actual = expected_image.convert("RGB"), observed_image.convert("RGB")
                for field in regions:
                    if field["page"] == page_number:
                        rect = field["rect_pdf"]
                        difference = ImageChops.difference(region(target, rect), region(actual, rect)).getbbox()
                        record["regions"].append({**field, "difference_bbox": difference})
                        require(difference is None, f"Field region changed on source page {page_number}: {rect}")
                require(ImageChops.difference(target, actual).getbbox() is None,
                        f"Reference page pixels changed on source page {page_number}")


def run_case(application, runtime, renderer, work, host, kind, mode, sources, case):
    directory = work / f"{host['id']}-{kind}"
    directory.mkdir()
    base = directory / "output"
    base.mkdir()
    prior = base / "prior-output.txt"
    prior.write_bytes(b"Preserve earlier output\n")
    source = directory / "Original synthetic Å [1] & input.pdf"
    reference, values, regions = sources[kind]
    shutil.copyfile(reference["source"], source)
    environment = policy.support.controlled_environment(directory, host["shell_executable"])
    script = application / "WinBookSplit.ps1"
    batch = host["id"] == "BAT"
    if batch:
        app_copy = directory / "app"
        for name in policy.support.validator.APPLICATION:
            path = app_copy / name
            path.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(application / name, path)
        script_copy = app_copy / script.name
        script_copy.write_bytes(policy.validator.controlled_powershell(script_copy.read_bytes()))
        wrapper = directory / "invoke.cmd"
        wrapper.write_bytes(policy.validator.WRAPPER.encode("ascii"))
        environment.update(WBS_POLICY_BAT=str(app_copy / "WinBookSplit.bat"), WBS_POLICY_INPUT=str(source),
                           WBS_POLICY_OUTPUTDIRECTORY=str(base), WBS_POLICY_PYTHONPATH=str(runtime))
        command = [str(Path(os.environ["SystemRoot"]) / "System32/cmd.exe"), "/d", "/v:off", "/c", str(wrapper)]
        case["batch_controls"] = policy.validator.MODIFICATIONS
        case["batch_sha256"] = cli.digest(app_copy / "WinBookSplit.bat")
    else:
        command = [host["shell_executable"], "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "RemoteSigned",
                   "-File", str(script), "-InputFile", str(source), "-OutputDirectory", str(base),
                   "-PythonPath", str(runtime), "-NonInteractive", "-Mode", "Manual" if mode == "Manual" else "Auto"]
        command += ["-StartPages", "1,3"] if mode == "Manual" else ["-BookmarkLevel", mode[-1]]
    with policy.paths.read_only_source(source):
        before = cli.identity(source)
        case["source_before"] = before
        process = (policy.prompted.run_prompted(command, directory, environment, "M\n1,3\nY\nN\n") if batch
                   else cli.run_unavailable_stdin(command, directory, environment))
        case["process"] = process
        case["source_after"] = cli.identity(source)
        require(before == case["source_after"], "Read-only input identity changed")
        require(prior.read_bytes() == b"Preserve earlier output\n", "Earlier output changed")
        outcome = cli.outcome(process["stdout"])
        case["outcome"] = outcome
        log, frames, _, log_hash = cli.console_records(process["stdout"], base)
        case.update(engine_frames=frames, log_sha256=log_hash)
        result = outcome.get("engine_result")
        if process["exit_code"]:
            remaining = {path.name for path in base.iterdir()}
            require(remaining == {prior.name, *([log.parent.name] if log else [])},
                    "Rejected document left chapter or staging output")
            require(result and result["written_count"] == 0 and result["execution"] is None,
                    "Rejected document reported output")
        require(process["exit_code"] == 0 and result and result["code"] == "split_complete",
                f"Expected static split success, got {process['exit_code']}: "
                + (result["code"] + ": " + result["message"] if result else str(outcome)))
        inspect_outputs(renderer, directory, base, result, case, reference["static"], values, regions, before["sha256"])


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--app-root", type=Path, default=ROOT)
    parser.add_argument("--application-python-path", type=Path, required=True)
    parser.add_argument("--shell-path", type=Path, action="append", required=True)
    parser.add_argument("--renderer-path", type=Path, required=True)
    parser.add_argument("--report", type=Path, required=True)
    parser.add_argument("--case", choices=sorted({kind for _, kind, _ in MATRIX}), action="append")
    args = parser.parse_args()
    require(os.name == "nt" and sys.flags.isolated and sys.dont_write_bytecode, "Use Windows and developer Python -I -B")
    require(pypdf_version == "6.21.4", "Developer pypdf does not match the v1.0.4 lock")
    require(len(args.shell_path) == 2 and all(path.is_absolute() and path.is_file()
            for path in [*args.shell_path, args.renderer_path, args.application_python_path]),
            "Supply both shell executables, the application Python and Poppler as absolute paths")
    report_path = policy.runner.new_external_path(args.report)
    work = Path(tempfile.mkdtemp(prefix="WBS-forms-")).resolve()
    application = args.app_root.resolve(strict=True)
    probe = cli.process_tests.host_probe(work)
    hosts = {host["id"]: host for host in (policy.paths.host_observation(shell, work, probe)
                                         for shell in args.shell_path)}
    require(set(hosts) == {"PS51", "PS7"}, "Both native PowerShell hosts are required")
    hosts["BAT"] = {**hosts["PS51"], "id": "BAT"}
    runtime_probe = cli.run_unavailable_stdin([str(args.application_python_path), "-I", "-B", "-c",
        "import sys, pypdf; print(sys.version); print(pypdf.__version__)"], work, os.environ.copy())
    require(runtime_probe["exit_code"] == 0 and runtime_probe["stdout"].splitlines()[-1] == pypdf_version,
            "Application pypdf does not match the pinned developer version")
    report = {
        "workspace": str(work), "application": str(application),
        "application_sha256": {name: cli.digest(application / name) for name in policy.support.validator.APPLICATION},
        "checkout_head": subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=ROOT, text=True).strip(),
        "windows": platform.win32_ver(), "shells": hosts, "application_runtime": runtime_probe,
        "renderer": cli.run_unavailable_stdin([str(args.renderer_path), "-v"], work, os.environ.copy()), "cases": [],
    }
    selected = [row for row in MATRIX if not args.case or row[1] in args.case]
    sources = {}
    for kind in sorted({kind for _, kind, _ in selected}):
        source, reference, values, regions = fixtures.make_case(kind)
        paths = {"source": work / f"{kind}.pdf", "static": work / f"{kind}-reference.pdf"}
        for writer, key in ((source, "source"), (reference, "static")):
            with paths[key].open("xb") as stream:
                writer.write(stream)
        sources[kind] = paths, values, regions
    report["reference_oracle_renders"] = []
    report["reference_oracle"] = check_reference_oracle(args.renderer_path, work, sources, report["reference_oracle_renders"])
    for host, kind, mode in selected:
        case = {"host": host, "kind": kind, "mode": mode, "outputs": [], "renders": [],
                "static_reference_sha256": cli.digest(sources[kind][0]["static"])}
        try:
            run_case(application, args.application_python_path, args.renderer_path, work, hosts[host], kind, mode, sources, case)
        except Exception as error:
            case["failure"] = str(error)
            print(f"FAIL {host} {kind} {mode}: {error}", flush=True)
        else:
            print(f"PASS {host} {kind} {mode}", flush=True)
        report["cases"].append(case)
    with report_path.open("x", encoding="utf-8", newline="\n") as stream:
        json.dump(report, stream, indent=2, ensure_ascii=True)
        stream.write("\n")
    return int(any("failure" in case for case in report["cases"]))


if __name__ == "__main__":
    raise SystemExit(main())
