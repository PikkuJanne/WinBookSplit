"""M4-T01 actual two-host page structure/navigation and two-renderer fidelity."""
from __future__ import annotations

import argparse
from datetime import datetime, timezone
from hashlib import sha256
import importlib.util
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
import time

from PIL import Image
from pypdf import PdfReader

ROOT = Path(__file__).resolve().parents[2]


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


support = load('wbs_fidelity_support', ROOT / 'tests/support/characterize_support.py')
cli, require, history, paths, conversion, launchers, runner = (support.cli, support.require, support.history,
    support.paths, support.conversion, support.launchers, support.runner)
fixtures = load('wbs_fidelity_fixtures', ROOT / 'tests/fidelity/generate_fidelity_fixtures.py')
validator = load('wbs_fidelity_validator', ROOT / 'tests/fidelity/validate_fidelity_report.py')
COMPLETED_CASES, LAST_CASE = [], None


def native(command, cwd, *, timeout=90):
    started, clock = datetime.now(timezone.utc).isoformat(), time.monotonic()
    child = subprocess.Popen(command, cwd=cwd, stdin=subprocess.DEVNULL, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    try:
        stdout, stderr = child.communicate(timeout=timeout)
    except subprocess.TimeoutExpired as error:
        if child.poll() is None:
            child.terminate()
        raise history.EntryPointFailure('Renderer/owned child deadline; tree stop unproved, workspace retained', cleanup_safe=False, pid=child.pid) from error
    return {'command': list(map(str, command)), 'cwd': str(cwd), 'stdin': 'DEVNULL', 'actual_process': True,
            'exit_code': child.returncode, 'timed_out': False, 'timeout_seconds': timeout, 'pid': child.pid,
            'started_at': started, 'finished_at': datetime.now(timezone.utc).isoformat(), 'elapsed_seconds': time.monotonic() - clock,
            'stdout': stdout.decode('utf-8', errors='strict'), 'stderr': stderr.decode('utf-8', errors='strict'),
            'stdout_sha256': sha256(stdout).hexdigest(), 'stderr_sha256': sha256(stderr).hexdigest()}


def render_probe(renderer, secondary_python, directory):
    primary = native([str(renderer), '-v'], directory)
    require(primary['exit_code'] == 0 and 'pdftoppm version 26.07.0' in primary['stderr'], 'Supported Poppler renderer version differs')
    script = ROOT / 'tests/fidelity/render_pdfium.py'
    secondary = native([str(secondary_python), '-I', '-B', str(script), '--probe'], directory)
    require(secondary['exit_code'] == 0, 'Secondary PDFium renderer unavailable')
    details = json.loads(secondary['stdout'])
    module = Path(details['module_path'])
    binaries = sorted(module.parent.parent.glob('pypdfium2_raw/*pdfium*.dll'))
    require(binaries, 'Actual PDFium native library hash unavailable')
    return {'primary': {'name': 'Poppler', 'path': str(renderer), 'version': '26.07.0', 'sha256': cli.digest(renderer), 'probe': primary},
            'secondary': {'name': 'PDFium', 'python_path': str(secondary_python), 'python_sha256': cli.digest(secondary_python),
                          'script_sha256': cli.digest(script), 'module_sha256': cli.digest(module),
                          'native_library_sha256': {item.name: cli.digest(item) for item in binaries}, **details, 'probe': secondary}}


def png_record(path, page):
    with Image.open(path) as image:
        rgba = image.convert('RGBA')
        return {'page': page, 'path': str(path), 'filename': path.name, 'bytes': path.stat().st_size,
                'width': rgba.width, 'height': rgba.height, 'png_sha256': cli.digest(path), 'pixels_sha256': sha256(rgba.tobytes()).hexdigest()}


def render_pdf(source, directory, prefix, renderer, secondary_python):
    expected = len(PdfReader(source).pages)
    primary_prefix = directory / (prefix + '-poppler')
    primary = native([str(renderer), '-r', '72', '-cropbox', '-png', str(source), str(primary_prefix)], directory)
    require(primary['exit_code'] == 0, 'Actual Poppler render failed')
    files = sorted(directory.glob(primary_prefix.name + '-*.png'), key=lambda path: int(path.stem.rsplit('-', 1)[1]))
    require(len(files) == expected, 'Actual Poppler page PNG count differs')
    primary_pages = [png_record(path, index + 1) for index, path in enumerate(files)]
    secondary_prefix = directory / (prefix + '-pdfium')
    secondary = native([str(secondary_python), '-I', '-B', str(ROOT / 'tests/fidelity/render_pdfium.py'),
                        '--input', str(source), '--output-prefix', str(secondary_prefix)], directory)
    require(secondary['exit_code'] == 0, 'Actual PDFium render failed')
    reported = json.loads(secondary['stdout'])
    require(reported['renderer'] == 'PDFium' and len(reported['pages']) == expected, 'Actual PDFium page PNG count differs')
    secondary_pages = [png_record(directory / entry['filename'], entry['page']) for entry in reported['pages']]
    require(all(all(actual[key] == declared[key] for key in ('page', 'width', 'height', 'png_sha256', 'pixels_sha256'))
                for actual, declared in zip(secondary_pages, reported['pages'])), 'Secondary pixel observation differs from actual retained image')
    return {'Poppler': {'process': primary, 'pages': primary_pages}, 'PDFium': {'process': secondary, 'pages': secondary_pages}}


def inspect_output(path, entry, reader_source, source_snapshots):
    reader = PdfReader(path)
    pages = [fixtures.page_snapshot(reader, index) for index in range(len(reader.pages))]
    expected = source_snapshots[entry['start']:entry['end']]
    raw_annots = []
    for local, page in enumerate(reader.pages):
        for annotation in page.get('/Annots', []):
            annotation = fixtures.resolved(annotation)
            p = annotation.get('/P')
            raw_annots.append({'page': local, 'name': str(annotation.get('/NM', '')), 'P_matches_page': p == page.indirect_reference,
                               'source_relations_absent': not any(key in annotation for key in ('/Popup', '/IRT', '/Parent')),
                               'unsupported_action_absent': '/A' not in annotation or fixtures.resolved(annotation['/A']).get('/S') == '/GoTo'})
    outline = []
    for node in reader.outline:
        require(not isinstance(node, list), 'Output chapter-start outline unexpectedly nested')
        outline.append({'title': node.title, 'page': reader.get_destination_page_number(node)})
    return {'filename': entry['filename'], 'start': entry['start'], 'end': entry['end'],
            'sha256': cli.digest(path), 'bytes': path.stat().st_size, 'pages': pages, 'expected_pages': expected,
            'inventory': fixtures.serialized_inventory(reader), 'metadata': dict(reader.metadata or {}), 'outline': outline,
            'annotations': raw_annots, 'catalog_keys': sorted(map(str, reader.trailer['/Root'].keys()))}


def application_case(work, host, kind, mode, originals, render_directory, renderer, secondary_python):
    global LAST_CASE
    identifier = host['id'] + '-' + kind + '-' + mode
    directory = work / identifier
    directory.mkdir()
    app, books, cwd, base = [directory / name for name in ('a', 'b', 'c', 'o')]
    copied = support.copy_application(app)
    for target in (books, cwd, base):
        target.mkdir()
    source = books / 'Authored Å 日本 [1] & %! $(literal).pdf'
    shutil.copyfile(originals[kind], source)
    neighbor, prior = books / 'authored-neighbor.txt', base / 'prior-output.pdf'
    neighbor.write_bytes(b'Authored neighbor remains unchanged.\n')
    prior.write_bytes(b'Authored previous output remains unchanged.\n')
    parameters = ['-InputFile', str(source), '-OutputDirectory', str(base), '-PythonPath', sys.executable,
                  '-Mode', 'Manual' if mode == 'manual' else 'Auto', '-NonInteractive']
    parameters += ['-StartPages', '1,4' if kind == 'rich' else '1,3'] if mode == 'manual' else ['-BookmarkLevel', mode]
    command = [host['shell_executable'], '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'RemoteSigned',
               '-File', str(app / 'WinBookSplit.ps1'), *parameters]
    artifact = render_directory / identifier
    artifact.mkdir()
    LAST_CASE = {'id': identifier, 'command': command, 'workspace': str(directory), 'render_directory': str(artifact)}
    with paths.read_only_source(source):
        before = {name: cli.identity(path) for name, path in (('source', source), ('neighbor', neighbor), ('prior', prior))}
        source_reader = PdfReader(source)
        source_snapshots = [fixtures.page_snapshot(source_reader, index) for index in range(len(source_reader.pages))]
        source_renders = render_pdf(source, artifact, 'source', renderer, secondary_python)
        observed = cli.run_unavailable_stdin(command, cwd, support.controlled_environment(cwd, host['shell_executable']), timeout=90)
        LAST_CASE.update(observed)
        final = cli.outcome(observed['stdout'])
        require(final['exit_code'] == observed['exit_code'] == 0 and final['status'] == 'success', 'Actual fidelity launcher failed: ' + identifier)
        log, engine_records, processes, log_hash = cli.console_records(observed['stdout'], base)
        require(log is not None, 'Actual fidelity launcher produced no finalized console')
        console_text = log.read_bytes().decode('utf-8')
        console = launchers.authenticate_console_manifest(log, final, source_path=source, source_observation=before['source'])
        engine = final['engine_result']
        execution, plan = engine['execution'], engine['plan']
        publication = Path(execution['final_directory'])
        output_records = []
        for number, entry in enumerate(execution['outputs']):
            path = publication / entry['filename']
            observation = inspect_output(path, entry, source_reader, source_snapshots)
            observation['renders'] = render_pdf(path, artifact, 'output-' + str(number + 1), renderer, secondary_python)
            output_records.append(observation)
        raw_manifest = (publication / 'WinBookSplit_Manifest.json').read_bytes().decode('utf-8')
        members = sorted(item.name for item in publication.iterdir())
        row = {'id': identifier, 'host_id': host['id'], 'kind': kind, 'mode': mode, 'actual_process': True,
               'host_version': host['host_version'], 'shell_executable': host['shell_executable'], 'parameters': parameters,
               'application_sha256': copied, 'input_path': str(source), 'output_base': str(base), 'source_read_only_observed': True,
               'source_observations_before': before, 'source_observations_after': {name: cli.identity(path) for name, path in (('source', source), ('neighbor', neighbor), ('prior', prior))},
               'source_pages': source_snapshots, 'source_metadata': dict(source_reader.metadata), 'source_renders': source_renders,
               'outcome': final, 'engine_records': engine_records, 'process_summaries': processes, 'console_log': console_text,
               'console_log_sha256': log_hash, 'console_evidence': console, 'plan': plan, 'publication': output_records,
               'publication_manifest': json.loads(raw_manifest), 'publication_manifest_raw': raw_manifest,
               'publication_manifest_sha256': cli.digest(publication / 'WinBookSplit_Manifest.json'), 'publication_members': members,
               'render_directory': str(artifact), 'application_unchanged': copied == {name: cli.digest(app / name) for name in copied}, **observed}
        if kind == 'rich' and mode == 'manual':
            summary = cwd / 'redacted-summary.json'
            export_command = [host['shell_executable'], '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'RemoteSigned', '-File',
                              str(app / 'Export-WinBookSplitDiagnostics.ps1'), '-ManifestPath', str(log.parent / 'WinBookSplit_Run.json'), '-OutputPath', str(summary)]
            exported = cli.run_unavailable_stdin(export_command, cwd, support.controlled_environment(cwd, host['shell_executable']), timeout=60)
            exported['actual_process'] = True
            require(exported['exit_code'] == 0 and summary.is_file(), 'Actual new warned run not exportable')
            summary_raw = summary.read_bytes().decode('utf-8')
            row['support_export'] = {'process': exported, 'summary_raw': summary_raw, 'summary_sha256': cli.digest(summary), 'summary': json.loads(summary_raw)}
        else:
            row['support_export'] = None
        LAST_CASE = row
        validator.validate_case(row, cleanup_complete=False)
        conversion.remove_publication(base, execution, set(members))
        launchers.remove_console_directory(log, base, final, source_path=source, source_observation=before['source'])
        row.update(source_observations_after_cleanup={name: cli.identity(path) for name, path in (('source', source), ('neighbor', neighbor), ('prior', prior))},
                   output_members_after_cleanup=sorted(item.name for item in base.iterdir()), passed=True)
        validator.validate_case(row)
    COMPLETED_CASES.append(row)
    return row


def characterize(work, shells, renderer, secondary_python, render_directory):
    require(os.name == 'nt' and len(shells) == 2, 'Actual Windows and two supported hosts required')
    before = runner.source_manifest()
    require(not render_directory.exists() and render_directory.is_absolute() and not render_directory.is_relative_to(ROOT), 'Render directory must be new and external')
    render_directory.mkdir()
    originals, provenance = fixtures.generate_fixtures(work / 'fixtures')
    hosts = [paths.host_observation(shell, work, cli.process_tests.host_probe(work)) for shell in shells]
    require({host['id'] for host in hosts} == {'PS51', 'PS7'}, 'Both supported distinct hosts required')
    renderers = render_probe(renderer, secondary_python, work)
    matrix = [('rich', mode) for mode in ('manual', '1', '2')] + [('scanned', 'manual')]
    cases = [application_case(work, host, kind, mode, originals, render_directory, renderer, secondary_python) for host in hosts for kind, mode in matrix]
    for host in hosts:
        host['policies_after'] = paths.host_observation(Path(host['shell_executable']), work, cli.process_tests.host_probe(work))['stored_policies']
    require(before == runner.source_manifest(), 'Sources changed during fidelity acceptance')
    return {'schema_version': 1, 'task_id': 'M4-T01', 'result': 'FIDELITY_REGRESSION_PASSED', 'success': True, 'exit_code': 0,
            'observed_at': datetime.now(timezone.utc).isoformat(), 'acceptance_ids': ['AC-066', 'AC-067', 'AC-068', 'AC-069'],
            'tested_path_sha256': before, 'source_unchanged': True, 'host_cases': hosts, 'cases': cases, 'case_count': len(cases),
            'fixture_provenance': provenance, 'renderers': renderers, 'render_directory': str(render_directory),
            'render_artifacts_retained': True, 'human_or_GUI_tested': False, 'environment': {'python': sys.version, 'python_executable': sys.executable},
            'limits': ['Automated same-renderer source/output exact RGBA pixel comparisons using two separate native renderer routes.',
                       'Root visual inspection is a separate observation; no human/Explorer/viewer/CI/package claim.', 'Generated original PDFs only; no private source documents.']}


def main():
    sys.stdout.reconfigure(encoding='utf-8', errors='backslashreplace')
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--report', required=True, type=Path)
    parser.add_argument('--shell-path', required=True, action='append', type=Path)
    parser.add_argument('--renderer-path', required=True, type=Path)
    parser.add_argument('--secondary-python-path', required=True, type=Path)
    parser.add_argument('--render-directory', required=True, type=Path)
    args = parser.parse_args()
    target = runner.new_external_path(args.report)
    temporary = tempfile.TemporaryDirectory(prefix='F-')
    work = Path(temporary.name).resolve()
    try:
        require(sys.flags.isolated and sys.dont_write_bytecode, 'Use pinned isolated developer Python -I -B')
        report = characterize(work, [path.resolve() for path in args.shell_path], args.renderer_path.resolve(), args.secondary_python_path.resolve(), args.render_directory.resolve())
        report['owned_temp_removed'] = False
        validator.validate_fidelity_report(report, list(map(str, map(Path.resolve, args.shell_path))), str(args.renderer_path.resolve()), str(args.secondary_python_path.resolve()), cleanup_complete=False)
        temporary.cleanup()
        report['owned_temp_removed'] = not work.exists()
        validator.validate_fidelity_report(report, list(map(str, map(Path.resolve, args.shell_path))), str(args.renderer_path.resolve()), str(args.secondary_python_path.resolve()))
        with target.open('x', encoding='utf-8', newline='\n') as stream:
            json.dump(report, stream, ensure_ascii=True, indent=2)
            stream.write('\n')
        print('Fidelity acceptance passed: ' + str(report['case_count']) + ' actual two-renderer native controls')
        return 0
    except BaseException as error:
        temporary._finalizer.detach()
        with target.open('x', encoding='utf-8', newline='\n') as stream:
            json.dump({'schema_version': 1, 'task_id': 'M4-T01', 'result': 'FIDELITY_REGRESSION_FAILED', 'success': False,
                       'exit_code': 1, 'error': str(error), 'workspace_retained': str(work) if work.exists() else None,
                       'render_directory': str(args.render_directory), 'owned_temp_removed': not work.exists(), 'cleanup_safe': False,
                       'completed_cases': COMPLETED_CASES, 'last_case': LAST_CASE}, stream, ensure_ascii=True, indent=2)
        print('Fidelity acceptance failed; authored workspace retained: ' + str(error), file=sys.stderr)
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
