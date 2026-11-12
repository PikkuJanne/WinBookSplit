"""Native application regression for the intended static AcroForm behavior."""
from __future__ import annotations

import argparse
import importlib.util
import json
import math
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
BASELINE_MATRIX = (
    ("PS51", "plain", "Manual"),
    ("PS7", "plain", "Auto2"),
    ("BAT", "plain", "Manual"),
    ("PS51", "acroform_empty", "Auto1"),
    ("PS7", "text", "Auto2"),
    ("BAT", "controls", "Manual"),
    ("PS51", "orphan_widget", "Manual"),
)
CHECKPOINT_2_MATRIX = (
    ('PS51', 'acroform_null', 'Manual'),
    ('PS51', 'acroform_empty', 'Auto1'),
    ('PS7', 'fields_empty', 'Auto2'),
    ('PS7', 'resource_only', 'Manual'),
    ('PS51', 'xfa_null', 'Auto1'),
    ('PS7', 'xfa_empty_array', 'Auto2'),
    ('PS51', 'resource_only', 'Preview'),
    ('PS7', 'xfa_detection_only', 'Manual'),
    ('PS51', 'fields_malformed', 'Manual'),
    ('PS7', 'field_cycle', 'Preview'),
    ('PS51', 'empty_feature_trees', 'Manual'),
    ('PS7', 'document_javascript', 'Auto1'),
    ('PS51', 'document_attachment', 'Auto2'),
    ('BAT', 'script_and_attachment', 'Manual'),
    ('PS51', 'associated_attachment', 'Manual'),
    ('PS7', 'page_and_annotation_actions', 'Auto2'),
    ('PS51', 'named_navigation_and_features', 'Auto1'),
    ('PS7', 'named_navigation_and_features', 'Auto2'),
    ('BAT', 'named_navigation_and_features', 'Manual'),
    ('PS7', 'named_navigation_and_features', 'Preview'),
    ('PS51', 'broken_omitted_branches', 'Manual'),
    ('PS7', 'payload_aliases', 'Manual'),
    ('BAT', 'encrypted_metadata_and_features', 'Manual'),
    ('PS7', 'form_and_features', 'Manual'),
    ('PS51', 'encrypted_form_and_features', 'Preview'),
)
FORM_MATRIX = (
    ('PS51', 'text', 'Manual'),
    ('PS7', 'text', 'Auto2'),
    ('BAT', 'controls', 'Manual'),
    ('PS7', 'controls', 'Auto2'),
    ('PS51', 'shared_field', 'Auto1'),
    ('PS7', 'missing_appearance', 'Manual'),
    ('PS51', 'stale_appearance', 'Auto2'),
    ('PS51', 'orphan_widget', 'Manual'),
    ('PS7', 'rotated_crop', 'Auto2'),
    ('PS51', 'form_action', 'Manual'),
    ('PS7', 'encrypted_form', 'Auto1'),
    ('BAT', 'encrypted_form_and_features', 'Manual'),
    ('PS7', 'form_and_features', 'Auto2'),
    ('PS51', 'multiline_readonly', 'Manual'),
    ('PS7', 'multiselect', 'Auto1'),
    ('PS51', 'multiselect_missing_appearance', 'Auto2'),
    ('PS7', 'combo_missing_appearance', 'Manual'),
    ('PS51', 'winansi_missing_appearance', 'Manual'),
    ('PS51', 'hierarchy', 'Auto2'),
    ('PS7', 'transformed_appearance', 'Manual'),
    ('PS51', 'hidden_fields', 'Auto1'),
    ('PS7', 'signature_visible', 'Auto2'),
    ('PS51', 'button_stale_state', 'Manual'),
    ('PS7', 'signature_structure', 'Manual'),
    ('PS51', 'missing_appearance_font', 'Preview'),
    ('PS7', 'missing_appearance_font', 'Manual'),
)
MATRIX = tuple(dict.fromkeys(BASELINE_MATRIX + CHECKPOINT_2_MATRIX + FORM_MATRIX))
EXPECTED_FAILURES = {
    'xfa_detection_only': ('unsupported_document', 'xfa_rendering_required'),
    'fields_malformed': ('invalid_document', 'malformed_form'),
    'field_cycle': ('invalid_document', 'malformed_form'),
    'signature_structure': ('unsupported_document', 'form_appearance_required'),
    'missing_appearance_font': ('unsupported_document', 'form_appearance_required'),
}
CHECKPOINT_2_FAILURES = {
    **EXPECTED_FAILURES,
    'form_and_features': ('unsupported_document', 'form_appearance_required'),
    'encrypted_form_and_features': ('unsupported_document', 'form_appearance_required'),
}
FEATURE_WARNINGS = {
    'empty_feature_trees': set(),
    'document_javascript': {'pdf_javascript_omitted'},
    'document_attachment': {'pdf_attachment_omitted'},
    'script_and_attachment': {'pdf_javascript_omitted', 'pdf_attachment_omitted'},
    'associated_attachment': {'pdf_attachment_omitted'},
    'page_and_annotation_actions': {'pdf_javascript_omitted'},
    'named_navigation_and_features': {'pdf_javascript_omitted', 'pdf_attachment_omitted'},
    'payload_aliases': {'pdf_javascript_omitted', 'pdf_attachment_omitted'},
    'encrypted_metadata_and_features': {'pdf_javascript_omitted', 'pdf_attachment_omitted'},
    'form_and_features': {'pdf_javascript_omitted', 'pdf_attachment_omitted'},
    'encrypted_form_and_features': {'pdf_javascript_omitted', 'pdf_attachment_omitted'},
    'broken_omitted_branches': set(),
}
RANGES = {"Manual": [[0, 2], [2, 4]], "Auto1": [[0, 1], [1, 2], [2, 3], [3, 4]],
          "Auto2": [[0, 1], [1, 2], [2, 3], [3, 4]]}
PAGE_MARKER = re.compile(rb"SOURCE-PAGE-(\d+)\b")
OMITTED_KEYS = {'/AA', '/AF', '/AFRelationship', '/Collection', '/EF', '/EmbeddedFiles', '/FS',
                '/JavaScript', '/JS', '/OpenAction', '/Perms'}
OMITTED_TYPES = {'/EmbeddedFile', '/Filespec', '/Sig', '/DocTimeStamp'}
OMITTED_ACTIONS = {'/JavaScript', '/URI', '/Launch', '/GoToR', '/GoToE', '/SubmitForm',
                   '/ResetForm', '/ImportData', '/Rendition', '/RichMediaExecute'}
SECONDARY_CASES = {'controls', 'missing_appearance', 'stale_appearance', 'multiline_readonly', 'multiselect',
                   'multiselect_missing_appearance', 'combo_missing_appearance', 'winansi_missing_appearance',
                   'rotated_crop', 'transformed_appearance', 'signature_visible'}


def render(renderer, source, prefix, receipts, *, secondary=False):
    command = ([str(renderer), "-I", "-B", str(ROOT / 'tests/fidelity/render_pdfium.py'),
                "--input", str(source), "--output-prefix", str(prefix)] if secondary else
               [str(renderer), "-r", "72", "-cropbox", "-png", str(source), str(prefix)])
    process = cli.run_unavailable_stdin(
        command, prefix.parent, os.environ.copy(),
    )
    renderer_name = "PDFium" if secondary else "Poppler"
    require(process["exit_code"] == 0, renderer_name + " render failed: " + process["stderr"])
    paths = prefix.parent.glob(prefix.name + "-*.png")
    ordered = sorted(paths, key=lambda path: int(path.stem.rsplit("-", 1)[1]))
    receipts.append({"renderer": renderer_name, "process": process,
                     "images": {path.name: cli.digest(path) for path in ordered}})
    return ordered


def geometry(page):
    return list(page.mediabox), list(page.cropbox), page.rotation


def region(image, rect, page=None):
    """Map source PDF points to the 72 DPI rendered CropBox coordinates."""
    left, bottom, right, top = (float(x) for x in page.cropbox) if page is not None else (0, 0, *image.size)
    width, height = right - left, top - bottom
    rotation = page.rotation % 360 if page is not None else 0
    x0, y0, x1, y1 = rect
    def point(x, y):
        x, y = x - left, y - bottom
        if rotation == 90:
            return y, x
        if rotation == 180:
            return width - x, y
        if rotation == 270:
            return height - y, width - x
        return x, height - y
    corners = [point(x, y) for x in (x0, x1) for y in (y0, y1)]
    xs, ys = zip(*corners)
    crop = (math.floor(min(xs)), math.floor(min(ys)), math.ceil(max(xs)), math.ceil(max(ys)))
    require(0 <= crop[0] < crop[2] <= image.width and 0 <= crop[1] < crop[3] <= image.height,
            'Field region lies outside its rendered CropBox')
    return image.crop(crop)


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
    forbidden_bytes = {value.encode('utf-8') for value in outside_values}
    forbidden_bytes.update(fixtures.winansi_literal(value).encode('ascii') for value in outside_values)
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
                require(not OMITTED_KEYS.intersection(value)
                        and value.get('/Type') not in OMITTED_TYPES
                        and value.get('/S') not in OMITTED_ACTIONS
                        and value.get('/Subtype') != '/FileAttachment',
                        'Serialized attachment, executable action or relationship survived chapter export')
                action = value if value.get('/Type') == '/Action' else value.get('/A')
                if action is not None:
                    action = action.get_object()
                    require(isinstance(action, DictionaryObject) and action.get('/S') == '/GoTo'
                            and '/Next' not in action, 'Serialized action is not a projected local destination')
                pending.extend(value.values())
                if isinstance(value, StreamObject):
                    data = value.get_data()
                    markers.update(int(marker) for marker in PAGE_MARKER.findall(data))
                    require(all(item not in data for item in forbidden_bytes),
                            "Outside-chapter value survived in a decoded stream")
            elif isinstance(value, ArrayObject):
                pending.extend(value)
            elif isinstance(value, str):
                require(all(item not in value for item in outside_values),
                        "Outside-chapter value survived in a serialized string")
    require(page_count == len(pages) and markers == set(pages), "Serialized pages/content escaped the chapter range")
    return {"page_objects": page_count, "page_markers": sorted(markers), "sha256": cli.digest(chapter)}


def inspect_pixels(target, actual, source_image, expected_page, page_number, fields, record, renderer_name):
    require(target.size == actual.size, "Rendered page size changed")
    for field in fields:
        if field["page"] != page_number:
            continue
        rect = field["rect_pdf"]
        expected = region(target, rect, expected_page)
        observed = region(actual, rect, expected_page)
        difference = ImageChops.difference(expected, observed).getbbox()
        source_difference = None
        if source_image is not None:
            require(source_image.size == target.size, "Rendered source geometry changed")
            source_difference = ImageChops.difference(expected, region(source_image, rect, expected_page)).getbbox()
        record["regions"].append({**field, "renderer": renderer_name, "difference_bbox": difference,
                                  "source_difference_bbox": source_difference})
        require(difference is None, f"{renderer_name} field region changed on source page {page_number}: {rect}")
    require(ImageChops.difference(target, actual).getbbox() is None,
            f"{renderer_name} reference page pixels changed on source page {page_number}")


def open_rgb(path):
    with Image.open(path) as image:
        return image.convert("RGB")


def inspect_outputs(renderer, secondary, case_dir, base, result, case, source, reference, values, regions, source_hash):
    execution = result["execution"]
    start_end = [[row["start"], row["end"]] for row in execution["outputs"]]
    require(start_end == RANGES[case["mode"]] and result["written_count"] == len(start_end),
            "Wrong physical chapter partition")
    require(result["plan"]["source_identity"]["sha256"] == source_hash, "Plan source hash changed")
    final = Path(execution["final_directory"])
    members = {".WinBookSplit-owner.json", "WinBookSplit_Manifest.json", *[row["filename"] for row in execution["outputs"]]}
    require(final.parent == base and {path.name for path in final.iterdir()} == members, "Unexpected chapter directory contents")
    expected = PdfReader(reference)
    renderers = [('Poppler', renderer, False)]
    if secondary and case['kind'] in SECONDARY_CASES:
        renderers.append(('PDFium', secondary, True))
    references, source_renders = {}, {}
    for name, executable, is_secondary in renderers:
        references[name] = render(executable, reference, case_dir / f"reference-{name}", case["renders"],
                                  secondary=is_secondary)
        require(len(references[name]) == len(expected.pages) == 4, "Wrong reference page count")
        if case['kind'] in fixtures.FORM_CASES:
            source_renders[name] = render(executable, source, case_dir / f"source-{name}", case["renders"],
                                          secondary=is_secondary)
            require(len(source_renders[name]) == 4, "Wrong rendered source page count")
    all_values = {item for labels in values.values() for item in labels}
    feature_case = case['kind'] in fixtures.FEATURE_CASES
    for row in execution["outputs"]:
        start, end = row["start"], row["end"]
        chapter = final / row["filename"]
        reader = PdfReader(chapter)
        pages = list(range(start + 1, end + 1))
        chapter_renders = {}
        for name, executable, is_secondary in renderers:
            chapter_renders[name] = render(executable, chapter, case_dir / f"chapter-{start}-{name}", case["renders"],
                                           secondary=is_secondary)
            require(len(reader.pages) == len(chapter_renders[name]) == len(pages), "Wrong reopened chapter page count")
        inside_values = {item for page in pages for item in values.get(str(page), ())}
        forbidden = (all_values - inside_values) | set(fixtures.PRIVATE_FORM_MARKERS.get(case['kind'], ()))
        if feature_case:
            forbidden.update(fixtures.FEATURE_MARKERS)
        record = {"range": [start, end], "regions": [], **inspect_serialization(reader, chapter, pages, forbidden)}
        case["outputs"].append(record)
        for local, page_number in enumerate(pages):
            page = reader.pages[local]
            text = page.extract_text() or ""
            require(PAGE_MARKER.findall(text.encode("utf-8")) == [str(page_number).encode("ascii")],
                    "Missing, reordered or foreign page marker")
            require(all(item in text for item in values.get(str(page_number), ())), "Expected visible form text is missing")
            require(geometry(page) == geometry(expected.pages[page_number - 1]), "Page geometry changed")
            if feature_case and page_number == 2:
                resources = page['/Resources']
                layer = resources['/Properties'].raw_get('/FixtureLayer')
                require(resources['/XObject']['/FixtureImage'].get('/Subtype') == '/Image'
                        and layer in reader.root_object['/OCProperties']['/OCGs'],
                        'Image or optional-content group is missing')
                if case['kind'] in {'named_navigation_and_features', 'form_and_features', 'encrypted_form_and_features'}:
                    links = [item.get_object() for item in page['/Annots']
                             if item.get_object().get('/Subtype') == '/Link']
                    require(len(links) == 1 and links[0]['/Dest'][0] == page.indirect_reference,
                            'Named local link did not resolve to the copied physical page')
                if case['kind'] == 'associated_attachment':
                    stamps = [item.get_object() for item in page['/Annots']
                              if item.get_object().get('/Subtype') == '/Stamp']
                    require(len(stamps) == 1 and '/N' in stamps[0]['/AP'],
                            'Visible attachment artwork did not survive as a static annotation')
                if case['kind'] == 'page_and_annotation_actions':
                    require(any(item.get_object().get('/Subtype') == '/Square' for item in page['/Annots']),
                            'Static square annotation was dropped along with its action')
            for name, _, _ in renderers:
                inspect_pixels(open_rgb(references[name][page_number - 1]), open_rgb(chapter_renders[name][local]),
                               open_rgb(source_renders[name][page_number - 1]) if name in source_renders else None,
                               expected.pages[page_number - 1], page_number, regions, record, name)


def run_case(application, application_paths, runtime, renderer, secondary, work, host, kind, mode, sources, case,
             *, checkpoint_2):
    directory = work / f"{host['id']}-{kind}-{mode}"
    directory.mkdir()
    base = directory / "output"
    base.mkdir()
    prior = base / "prior-output.txt"
    prior.write_bytes(b"Preserve earlier output\n")
    neighbor = directory / 'neighbor.txt'
    neighbor.write_bytes(b'Preserve neighboring source files\n')
    source = directory / "Original synthetic Å [1] & input.pdf"
    reference, values, regions = sources[kind]
    shutil.copyfile(reference["source"], source)
    environment = policy.support.controlled_environment(directory, host["shell_executable"])
    script = application / "WinBookSplit.ps1"
    batch = host["id"] == "BAT"
    preview = mode == 'Preview'
    if batch:
        app_copy = directory / "app"
        for name in application_paths:
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
        command += ["-StartPages", "1,3"] if mode == "Manual" else ["-BookmarkLevel", "1" if preview else mode[-1]]
        if preview:
            command.append('-Preview')
    with policy.paths.read_only_source(source):
        before = cli.identity(source)
        neighbor_before = cli.identity(neighbor)
        prior_before = cli.identity(prior)
        input_members = {path.name for path in directory.iterdir()}
        case["source_before"] = before
        process = (policy.prompted.run_prompted(command, directory, environment, "M\n1,3\nY\nN\n") if batch
                   else cli.run_unavailable_stdin(command, directory, environment))
        case["process"] = process
        case["source_after"] = cli.identity(source)
        require(before == case["source_after"], "Read-only input identity changed")
        require(cli.identity(neighbor) == neighbor_before and cli.identity(prior) == prior_before,
                'Earlier output or neighboring source changed')
        require({path.name for path in directory.iterdir()} == input_members,
                'Unexpected file appeared next to the source or launcher')
        outcome = cli.outcome(process["stdout"])
        case["outcome"] = outcome
        log, frames, _, log_hash = cli.console_records(process["stdout"], base)
        case.update(engine_frames=frames, log_sha256=log_hash)
        result = outcome.get("engine_result")
        if process["exit_code"] or preview:
            remaining = {path.name for path in base.iterdir()}
            diagnostic = (result or {}).get('diagnostic') or {}
            record = diagnostic.get('record_path')
            allowed = {prior.name, *([log.parent.name] if log else [])}
            if record and Path(record).parent.parent == base and Path(record).parent.name.startswith('.WinBookSplit-failed-'):
                allowed.add(Path(record).parent.name)
            require(remaining == allowed,
                    "Preview or rejected document left chapter or staging output")
            require(result and result["written_count"] == 0 and result["execution"] is None,
                    "Preview or rejected document reported output")
        failures = CHECKPOINT_2_FAILURES if checkpoint_2 else EXPECTED_FAILURES
        if kind in failures:
            code, reason = failures[kind]
            require(result and result['code'] == code and process['exit_code'] == result['exit_code'] != 0
                    and (result.get('diagnostic') or {}).get('reason') == reason,
                    f'Expected {code}/{reason}, observed {result}')
            return
        if result and process['exit_code'] == 0:
            actual = {entry['code'] for entry in result['warnings']}
            if kind in FEATURE_WARNINGS:
                specific = {'pdf_javascript_omitted', 'pdf_attachment_omitted'}
                require(actual & specific == FEATURE_WARNINGS[kind], f'Wrong feature-specific warnings: {actual}')
            if kind in fixtures.FORM_CASES:
                warning = [entry['message'] for entry in result['warnings'] if entry['code'] == 'pdf_form_flattened']
                require(warning == ['Form fields will be preserved as static page content; chapter copies will not be fillable.'],
                        'Missing static-form disclosure')
                if kind == 'hidden_fields':
                    require('pdf_form_data_omitted' in actual, 'Missing nonvisual form data disclosure')
                if kind == 'signature_visible':
                    require('pdf_signature_omitted' in actual, 'Missing source signature disclosure')
                if kind in {'form_and_features', 'encrypted_form_and_features'}:
                    require('opening_destination_ignored' in actual, 'Local opening destination lost its expected notice')
        private_values = fixtures.PRIVATE_FORM_MARKERS.get(kind, ())
        if private_values:
            diagnostic_material = json.dumps(outcome, ensure_ascii=False) + process["stdout"] + process["stderr"]
            require(all(value not in diagnostic_material for value in private_values),
                    'Excluded form data escaped into native diagnostics')
        if preview:
            require(process['exit_code'] == 0 and result and result['code'] == 'preview_complete'
                    and result['plan']['source_identity']['sha256'] == before['sha256']
                    and result['plan']['ranges'] == RANGES['Auto1'], 'Expected a complete no-write preview plan')
            return
        require(process["exit_code"] == 0 and result and result["code"] == "split_complete",
                f"Expected static split success, got {process['exit_code']}: "
                + (result["code"] + ": " + result["message"] if result else str(outcome)))
        inspect_outputs(renderer, secondary, directory, base, result, case, source, reference["static"], values,
                        regions, before["sha256"])


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--app-root", type=Path, default=ROOT)
    parser.add_argument("--application-python-path", type=Path, required=True)
    parser.add_argument("--shell-path", type=Path, action="append", required=True)
    parser.add_argument("--renderer-path", type=Path, required=True)
    parser.add_argument("--secondary-python-path", type=Path, help="Isolated developer Python with PDFium")
    parser.add_argument("--report", type=Path, required=True)
    selection = parser.add_mutually_exclusive_group()
    selection.add_argument('--checkpoint-2', action='store_true', help='Run the historical checkpoint 2 matrix')
    selection.add_argument('--form-cases', action='store_true', help='Run only the static form acceptance cases')
    parser.add_argument("--case", choices=sorted({kind for _, kind, _ in MATRIX + CHECKPOINT_2_MATRIX}), action="append")
    args = parser.parse_args()
    require(os.name == "nt" and sys.flags.isolated and sys.dont_write_bytecode, "Use Windows and developer Python -I -B")
    require(pypdf_version == "6.21.4", "Developer pypdf does not match the application lock")
    require(len(args.shell_path) == 2 and all(path.is_absolute() and path.is_file()
            for path in [*args.shell_path, args.renderer_path, args.application_python_path]),
            "Supply both shell executables, the application Python and Poppler as absolute paths")
    report_path = policy.runner.new_external_path(args.report)
    work = Path(tempfile.mkdtemp(prefix="WBS-forms-")).resolve()
    application = args.app_root.resolve(strict=True)
    application_paths = policy.support.validator.APPLICATION
    if args.checkpoint_2:
        application_paths = tuple(name for name in application_paths if name != 'engine/winbooksplit_forms.py')
    if args.secondary_python_path:
        require(args.secondary_python_path.is_absolute() and args.secondary_python_path.is_file(),
                'Supply an absolute existing PDFium Python executable')
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
        "application_sha256": {name: cli.digest(application / name) for name in application_paths},
        "checkout_head": subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=ROOT, text=True).strip(),
        "windows": platform.win32_ver(), "shells": hosts, "application_runtime": runtime_probe,
        "renderer": cli.run_unavailable_stdin([str(args.renderer_path), "-v"], work, os.environ.copy()), "cases": [],
    }
    if args.secondary_python_path:
        report["secondary_renderer"] = cli.run_unavailable_stdin(
            [str(args.secondary_python_path), '-I', '-B', str(ROOT / 'tests/fidelity/render_pdfium.py'), '--probe'],
            work, os.environ.copy())
        require(report["secondary_renderer"]["exit_code"] == 0, "Secondary PDFium renderer unavailable")
    else:
        report["secondary_renderer"] = "NOT_RUN"
    matrix = CHECKPOINT_2_MATRIX if args.checkpoint_2 else FORM_MATRIX if args.form_cases else MATRIX
    selected = [row for row in matrix if not args.case or row[1] in args.case]
    require(selected, 'No cases match this matrix')
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
            run_case(application, application_paths, args.application_python_path, args.renderer_path,
                     args.secondary_python_path, work, hosts[host], kind, mode, sources, case,
                     checkpoint_2=args.checkpoint_2)
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
