"""M3-T05 actual two-host local manifests, warnings, redaction and finalization."""
from __future__ import annotations

import argparse
from copy import deepcopy
from datetime import datetime, timezone
from hashlib import sha256
import io
import json
import os
from pathlib import Path
import re
import shutil
import sys
import tempfile

from pypdf import PdfReader, PdfWriter
from pypdf.generic import NumberObject

ROOT = Path(__file__).resolve().parents[2]
import importlib.util


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


ux = load('wbs_support_ux', ROOT / 'tests/ux/characterize_ux.py')
cli, validator = ux.cli, load('wbs_support_validator', ROOT / 'tests/support/validate_support_report.py')
history, require, paths, conversion, launchers, runner = cli.history, cli.require, cli.paths, cli.conversion, cli.launchers, cli.runner
COMPLETED_CASES, EXPORT_CASES, LAST_CASE = [], [], None


def controlled_environment(cwd, shell):
    environment = history.clean_environment(cwd)
    system = Path(os.environ['SystemRoot'])
    environment.update(PATH=os.pathsep.join((str(Path(shell).parent), str(system / 'System32'), str(system))),
                       PSModulePath=str(Path(shell).parent / 'Modules'))
    return environment


def copy_application(app):
    app.mkdir()
    for name in validator.APPLICATION:
        destination = app / name
        destination.parent.mkdir(exist_ok=True)
        shutil.copyfile(ROOT / name, destination)
    return {name: cli.digest(app / name) for name in validator.APPLICATION}


def inject_fault(app, kind, original):
    if not kind.startswith(('log-fault-', 'manifest-fault-')):
        return None
    path = 'WinBookSplit.ps1' if kind.startswith('log-fault-') else 'engine/WinBookSplit.Logging.ps1'
    function = 'Complete-WinBookSplitConsoleLog' if path == 'WinBookSplit.ps1' else 'Write-WinBookSplitRunManifest'
    sentinel = 'Authored M3-T05 ' + ('log' if path == 'WinBookSplit.ps1' else 'manifest') + ' finalization failure.'
    source = app / path
    text = source.read_text(encoding='utf-8-sig')
    pattern = r'(function ' + function + r' \{\s*param\([^\n]*\)\r?\n)'
    changed, count = re.subn(pattern, lambda match: match[1] + "    throw '" + sentinel + "'\n", text)
    require(count == 1, 'Copied support fault seam changed: ' + function)
    source.write_text(changed, encoding='utf-8-sig', newline='\n')
    return {'path': path, 'function': function, 'sentinel': sentinel, 'original_sha256': original[path],
            'modified_sha256': cli.digest(source), 'scope': 'Copied application/helper throws before ordinary finalizer body; no production hook.'}


def authored_pdf(path, kind):
    generator = history.load_generator()
    if kind in {'failed-pdf', 'log-fault-primary', 'manifest-fault-primary'}:
        path.write_bytes(b'%PDF-1.7\nAuthored invalid PDF with no trailer or pages\n%%EOF\n')
        return None
    cli.expected_pdf(generator, path, flat=kind == 'no-plan')
    if kind == 'bookmark-warning':
        reader = PdfReader(path)
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        invalid = writer.add_outline_item('Authored invalid numeric destination', 1).get_object()
        invalid['/A']['/D'][0] = NumberObject(writer.pages[1].indirect_reference.idnum)
        writer.add_outline_item('Authored valid Unicode Å 日本 chapter', 3)
        buffer = io.BytesIO()
        writer.write(buffer)
        path.write_bytes(buffer.getvalue())
    if kind == 'parser-warning':
        raw = path.read_bytes()
        pointer = re.search(rb'startxref\s+(\d+)\s+%%EOF', raw)
        require(pointer is not None, 'Authored parser warning has no xref pointer')
        replacement = str(int(pointer[1]) + 1).encode('ascii')
        require(len(replacement) == len(pointer[1]), 'Authored startxref warning changed byte width')
        path.write_bytes(raw[:pointer.start(1)] + replacement + raw[pointer.end(1):])
    return conversion.page_content(PdfReader(path))


def publication(base, execution, expected):
    final = Path(execution['final_directory'])
    members = {'.WinBookSplit-owner.json', 'WinBookSplit_Manifest.json', *[item['filename'] for item in execution['outputs']]}
    require(final.parent == base and {item.name for item in final.iterdir()} == members, 'Support publication membership differs')
    raw = (final / 'WinBookSplit_Manifest.json').read_bytes().decode('utf-8', 'strict')
    require(validator.strict_json(raw) == execution['manifest'], 'Support actual chapter manifest differs')
    observed = []
    for item in execution['outputs']:
        target = final / item['filename']
        pages = conversion.page_content(PdfReader(target))
        require(pages == expected[item['start']:item['end']] and len(pages) == item['page_count']
                and cli.digest(target) == item['sha256'] and target.stat().st_size == item['size_bytes'],
                'Support retained chapter bytes/pages differ from the exact physical plan')
        observed.extend(pages)
    require(observed == expected, 'Support publication omitted/reordered/duplicated physical pages')
    return {'manifest': execution['manifest'], 'manifest_raw': raw, 'manifest_sha256': cli.digest(final / 'WinBookSplit_Manifest.json'),
            'members': sorted(members), 'content_sha256': observed}


def application_case(work, host, kind, ebook_sources, references, calibre):
    global LAST_CASE
    directory = work / (host['id'] + '-' + kind)
    directory.mkdir()
    app, books, cwd, base = [directory / name for name in ('a', 'b', 'c', 'o')]
    copied = copy_application(app)
    fault = inject_fault(app, kind, copied)
    actual_hashes = {name: cli.digest(app / name) for name in copied}
    for destination in (books, cwd, base):
        destination.mkdir()
    fmt = kind.removeprefix('success-') if kind in {'success-epub', 'success-azw3'} else 'pdf'
    source = books / ('Authored Å 日本 [1] & %! $(literal).' + fmt)
    if fmt == 'pdf':
        content = authored_pdf(source, kind)
    else:
        shutil.copyfile(ebook_sources[fmt], source)
        content = references[fmt]['page_content_sha256']
    neighbor, prior = books / 'authored-neighbor.txt', base / 'prior-output.pdf'
    neighbor.write_bytes(b'Authored neighboring data remains immutable.\n')
    prior.write_bytes(b'Authored earlier output remains immutable.\n')
    manual = kind not in {'no-plan', 'bookmark-warning'}
    parameters = ['-InputFile', str(source), '-OutputDirectory', str(base), '-PythonPath', sys.executable,
                  '-Mode', 'Manual' if manual else 'Auto']
    parameters += ['-StartPages', '1,2,3' if fmt != 'pdf' else '1,3,5'] if manual else ['-BookmarkLevel', '1']
    parameters += ['-NoPause'] if kind.endswith('cancelled') else ['-NonInteractive']
    if fmt != 'pdf':
        parameters += ['-CalibrePath', str(calibre)]
    if kind == 'preview':
        parameters += ['-Preview']
    if kind == 'preflight':
        parameters[parameters.index('-StartPages') + 1] = '1,invalid,5'
    command = [host['shell_executable'], '-NoProfile', *([] if kind.endswith('cancelled') else ['-NonInteractive']),
               '-ExecutionPolicy', 'RemoteSigned', '-File', str(app / 'WinBookSplit.ps1'), *parameters]
    LAST_CASE = {'id': host['id'] + '-' + kind, 'command': command, 'workspace': str(directory)}
    with paths.read_only_source(source):
        before = {name: cli.identity(item) for name, item in (('source', source), ('neighbor', neighbor), ('prior', prior))}
        if kind.endswith('cancelled'):
            native = ux.NativeDialogue(command, cwd, controlled_environment(cwd, host['shell_executable']))
            native.wait_text(ux.CONFIRM)
            hold = ux.held_without_chapters(native, base)
            native.send('C\n')
            observed = native.finish()
        else:
            hold = None
            observed = cli.run_unavailable_stdin(command, cwd, controlled_environment(cwd, host['shell_executable']), timeout=90)
        LAST_CASE.update(observed)
        final = cli.outcome(observed['stdout'])
        LAST_CASE['outcome'] = final
        log, engine_records, processes, log_hash = cli.console_records(observed['stdout'], base)
        text = None if log is None else log.read_bytes().decode('utf-8', 'strict')
        run_path = None if log is None else log.parent / 'WinBookSplit_Run.json'
        raw = run_path.read_bytes().decode('utf-8', 'strict') if run_path is not None and run_path.exists() else None
        manifest = None if raw is None else validator.strict_json(raw)
        console_members = [] if log is None else sorted(item.name for item in log.parent.iterdir())
        result = final.get('engine_result')
        plan_events = [] if text is None else validator.records(text, '[PLAN] ')
        plan = result.get('plan') if result else None
        if plan is None and plan_events:
            plan = plan_events[-1]['plan']
        parser = ((result or {}).get('diagnostic') or {}).get('parser_warnings') or {}
        published = publication(base, result['execution'], content) if result and result.get('execution') else None
        row = {'id': host['id'] + '-' + kind, 'kind': kind, 'host_id': host['id'], 'host_version': host['host_version'],
               'shell_executable': host['shell_executable'], 'actual_process': True, 'parameters': parameters,
               'application_sha256': copied, 'actual_application_sha256': actual_hashes, 'fault_injection': fault,
               'input_kind': fmt, 'input_path': str(source), 'output_base': str(base), 'source_read_only_observed': True,
               'source_observations_before': before, 'source_observations_after': {name: cli.identity(item) for name, item in (('source', source), ('neighbor', neighbor), ('prior', prior))},
               'outcome': final, 'console_log': text, 'console_log_sha256': log_hash, 'console_members': console_members,
               'run_manifest': manifest, 'run_manifest_raw': raw, 'run_manifest_sha256': None if raw is None else cli.digest(run_path),
               'engine_records': engine_records, 'process_summaries': processes, 'hold': hold,
               'expected_plan': plan, 'planning_warnings': (result or {}).get('warnings', []), 'parser_warnings': parser.get('records', []),
               'parser_suppressed_count': parser.get('suppressed_count', 0), 'parser_message_truncated_count': parser.get('message_truncated_count', 0),
               'expected_content_sha256': content, 'publication': published, **observed}
        LAST_CASE = row
        # Keep authored copies of real manifests for the subsequent actual export.
        if kind in {'success-pdf', 'failed-pdf', 'cancelled', 'log-fault-success', 'no-plan'}:
            require(raw is not None, 'Real export source manifest missing')
            (work / (host['id'] + '-export-input-' + kind + '.json')).write_bytes(raw.encode('utf-8'))
        if published:
            execution = result['execution']
            launchers.remove_known_directory(Path(execution['final_directory']), base, set(published['members']), '.WinBookSplit-owner.json',
                {'schema_version': 1, 'kind': 'run', 'run_id': execution['run_id']})
        if result:
            conversion.remove_diagnostic(base, result)
        if log is not None:
            expected_members = {'.WinBookSplit-console-owner.json', 'console.log'}
            if manifest is not None:
                expected_members.add('WinBookSplit_Run.json')
            require(set(console_members) == expected_members, 'Support cleanup refuses unexpected console members')
            marker = {'run_id': log.parent.name.removeprefix('.WinBookSplit-console-'), 'kind': 'console'}
            require(manifest is None or manifest['run_id'] == marker['run_id'], 'Support manifest does not authenticate its console owner')
            launchers.remove_known_directory(log.parent, base, expected_members, '.WinBookSplit-console-owner.json', marker)
        row.update(source_observations_after_cleanup={name: cli.identity(item) for name, item in (('source', source), ('neighbor', neighbor), ('prior', prior))},
                   application_unchanged=actual_hashes == {name: cli.digest(app / name) for name in actual_hashes},
                   output_members_after_cleanup=sorted(item.name for item in base.iterdir()), passed=True)
        validator.validate_application_case(row)
    COMPLETED_CASES.append(row)
    return row


def malicious_manifest(original):
    manifest = deepcopy(original)
    markers = ['MockPrivateUser_ZÅ日本', 'MOCK_SECRET_API_7efa4', 'MOCK_PRIVATE_DOCUMENT_TITLE_829c',
               'MOCK_PRIVATE_DOCUMENT_CONTENT_351d', 'MOCK_ENV_CREDENTIAL_f8eb']
    manifest['source_identity']['path'] = 'C:\\Users\\' + markers[0] + '\\' + markers[2] + '.pdf'
    manifest['source_identity']['resolved_path'] = manifest['source_identity']['path']
    manifest['settings']['normalized_inputs'] = {'authored_mock': markers[4]}
    manifest['outcome']['message'] = markers[1]
    manifest['engine_result']['message'] = markers[3]
    manifest['plan']['entries'][0]['title'] = markers[2]
    manifest['plan']['entries'][0]['filename'] = markers[0] + '.pdf'
    return manifest, markers


def export_case(work, host, kind):
    global LAST_CASE
    directory = work / (host['id'] + '-export-' + kind)
    directory.mkdir()
    app, cwd = directory / 'a', directory / 'c'
    copied = copy_application(app)
    cwd.mkdir()
    original_kind = {'failed': 'failed-pdf', 'cancelled': 'cancelled', 'incomplete': 'log-fault-success', 'no-plan': 'no-plan'}.get(kind, 'success-pdf')
    manifest = validator.strict_json((work / (host['id'] + '-export-input-' + original_kind + '.json')).read_text(encoding='utf-8'))
    markers = []
    if kind == 'redacted':
        manifest, markers = malicious_manifest(manifest)
    source, destination, prior = directory / 'WinBookSplit_Run.json', directory / 'support Å 日本 [1] &.json', directory / 'prior-support.json'
    raw = json.dumps(manifest, ensure_ascii=False, indent=2) + '\n'
    if kind == 'duplicate-keys':
        raw = raw.replace('"protocol": "winbooksplit.run",', '"protocol": "winbooksplit.run", "protocol": "winbooksplit.run",', 1)
    if kind == 'wrong-protocol':
        raw = raw.replace('winbooksplit.run', 'authored.wrong.protocol', 1)
    if kind == 'selected-secret':
        raw = raw.replace('"python": "3.14.8"', '"python": "MOCK_PRIVATE_VERSION_SECRET"', 1)
    if kind == 'pending-manifest':
        source = directory / 'WinBookSplit_Run.pending.json'
    source.write_bytes(raw.encode('utf-8') if kind != 'invalid-utf8' else b'\xff\xfe\x81invalid')
    prior.write_bytes(b'Authored earlier support data must not change.\n')
    if kind == 'existing':
        destination.write_bytes(prior.read_bytes())
    elif kind == 'same-path':
        destination = source
    elif kind == 'hardlink':
        os.link(source, destination)
    elif kind == 'reparse-output':
        target, junction = directory / 'target', directory / 'junction'
        target.mkdir()
        # The controlled junction is authored and scoped to this case only.
        import subprocess
        probe = subprocess.run([str(Path(os.environ['SystemRoot']) / 'System32/cmd.exe'), '/d', '/c', 'mklink', '/J', str(junction), str(target)],
                               capture_output=True, cwd=directory, timeout=15)
        require(probe.returncode == 0 and history.is_reparse(junction), 'Authored support junction creation failed')
        destination = junction / destination.name
    arguments = ['-ManifestPath', source.name if kind == 'relative-input' else str(source),
                 '-OutputPath', destination.name if kind == 'relative-output' else str(destination)]
    command = [host['shell_executable'], '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'RemoteSigned', '-File',
               str(app / 'Export-WinBookSplitDiagnostics.ps1'), *arguments]
    before, prior_before = cli.identity(source), cli.identity(prior)
    dest_before = cli.identity(destination) if destination.exists() else None
    LAST_CASE = {'id': host['id'] + '-export-' + kind, 'command': command, 'workspace': str(directory)}
    observed = cli.run_unavailable_stdin(command, cwd, controlled_environment(cwd, host['shell_executable']))
    LAST_CASE.update(observed)
    outcomes = validator.records(observed['stdout'], '[SUPPORT-EXPORT] ')
    require(len(outcomes) == 1, 'Support exporter did not emit exactly one terminal frame')
    good = kind in {'success', 'failed', 'cancelled', 'incomplete', 'no-plan', 'redacted'}
    summary_raw = destination.read_bytes().decode('utf-8', 'strict') if good and destination.exists() else None
    row = {'id': host['id'] + '-export-' + kind, 'kind': kind, 'host_id': host['id'], 'host_version': host['host_version'],
           'shell_executable': host['shell_executable'], 'actual_process': True, 'export_outcome': outcomes[0],
           'input_manifest': manifest, 'input_raw': raw, 'source_before': before, 'source_after': cli.identity(source),
           'prior_before': prior_before, 'prior_after': cli.identity(prior), 'private_markers': markers,
           'summary_raw': summary_raw, 'summary': None if summary_raw is None else validator.strict_json(summary_raw),
           'summary_sha256': None if summary_raw is None else cli.digest(destination), 'exporter_sha256': copied['Export-WinBookSplitDiagnostics.ps1'],
           'destination_created': destination.exists() and dest_before is None,
           'destination_before': dest_before, 'destination_after': cli.identity(destination) if destination.exists() else None,
           'application_unchanged': copied == {name: cli.digest(app / name) for name in copied}, **observed}
    LAST_CASE = row
    if not good:
        require(row['destination_before'] == row['destination_after'], 'Rejected support export created or overwrote a destination')
    elif destination.exists():
        destination.unlink()  # Authored external file; exact new output, no recursive source-directed cleanup.
    if kind == 'reparse-output':
        junction.rmdir()  # Remove only the authored reparse entry, never its target.
    row.update(owned_outputs_removed=True, passed=True)
    validator.validate_export_case(row)
    EXPORT_CASES.append(row)
    return row


def characterize(work, shells, calibre):
    require(os.name == 'nt' and len(shells) == 2, 'Actual Windows and both supported hosts required')
    before = runner.source_manifest()
    hosts = [paths.host_observation(shell, work, cli.process_tests.host_probe(work)) for shell in shells]
    require({host['id'] for host in hosts} == {'PS51', 'PS7'}, 'Actual distinct supported hosts required')
    ebooks = work / 'ebooks'
    provenance = conversion.fixtures.generate(ebooks, calibre)
    sources = {fmt: ebooks / ('original-three-chapters.' + fmt) for fmt in ('epub', 'azw3')}
    refs = {fmt: conversion.reference_pdf(source, work / ('reference-' + fmt + '.pdf'), str(calibre), work) for fmt, source in sources.items()}
    cases = [application_case(work, host, kind, sources, refs, calibre) for host in hosts for kind in validator.APPLICATION_KINDS]
    exports = [export_case(work, host, kind) for host in hosts for kind in validator.EXPORT_KINDS]
    for host in hosts:
        host['policies_after'] = paths.host_observation(Path(host['shell_executable']), work, cli.process_tests.host_probe(work))['stored_policies']
    require(before == runner.source_manifest(), 'Source changed during native support acceptance')
    return {'schema_version': 1, 'task_id': 'M3-T05', 'result': 'SUPPORT_REGRESSION_PASSED', 'success': True, 'exit_code': 0,
            'observed_at': datetime.now(timezone.utc).isoformat(), 'acceptance_ids': ['AC-063', 'AC-064', 'AC-065'],
            'tested_path_sha256': before, 'source_unchanged': True, 'host_cases': hosts, 'cases': cases, 'export_cases': exports,
            'case_count': len(cases) + len(exports), 'fixture_provenance': provenance, 'real_conversion_references': refs,
            'calibre_path': str(calibre), 'human_or_viewer_opening_tested': False,
            'environment': {'python': sys.version, 'python_executable': sys.executable, 'pypdf': history.pypdf.__version__},
            'limits': ['Redirected native console controls; no human, Explorer or viewer evidence.',
                       'Finalization faults throw only in disclosed copied app/helper; no production fault hook.',
                       'All mock private identifiers are authored test strings, not user data.']}


def main():
    sys.stdout.reconfigure(encoding='utf-8', errors='backslashreplace')
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--report', required=True, type=Path)
    parser.add_argument('--shell-path', required=True, action='append', type=Path)
    parser.add_argument('--calibre-path', required=True, type=Path)
    args = parser.parse_args()
    target = runner.new_external_path(args.report)
    temporary = tempfile.TemporaryDirectory(prefix='S-')
    work = Path(temporary.name).resolve()
    try:
        require(sys.flags.isolated and sys.dont_write_bytecode, 'Use isolated developer Python -I -B')
        report = characterize(work, [path.resolve() for path in args.shell_path], args.calibre_path.resolve())
        report['owned_temp_removed'] = False
        validator.validate_support_report(report, [str(path.resolve()) for path in args.shell_path], str(args.calibre_path.resolve()), cleanup_complete=False)
        temporary.cleanup()
        report['owned_temp_removed'] = not work.exists()
        validator.validate_support_report(report, [str(path.resolve()) for path in args.shell_path], str(args.calibre_path.resolve()))
        with target.open('x', encoding='utf-8', newline='\n') as stream:
            json.dump(report, stream, indent=2, ensure_ascii=True)
            stream.write('\n')
        print('Support acceptance passed: ' + str(report['case_count']) + ' actual native controls')
        return 0
    except BaseException as error:
        temporary._finalizer.detach()
        stopped = []
        for native in ux.ACTIVE_DIALOGUES:
            try:
                if native.child.poll() is None:
                    native.child.terminate()
                native.child.wait(timeout=5)
            except BaseException:
                pass
            stopped.append({'pid': native.child.pid, 'parent_exit_code': native.child.poll(), 'descendants_stopped': None,
                            'stdout_observed': native.text(), 'stderr_observed': native.text('stderr')})
        with target.open('x', encoding='utf-8', newline='\n') as stream:
            json.dump({'schema_version': 1, 'task_id': 'M3-T05', 'result': 'SUPPORT_REGRESSION_FAILED', 'success': False,
                       'exit_code': 1, 'error': str(error), 'workspace_retained': str(work) if work.exists() else None,
                       'owned_temp_removed': not work.exists(), 'cleanup_safe': False,
                       'completed_cases': COMPLETED_CASES, 'completed_export_cases': EXPORT_CASES, 'last_case': LAST_CASE,
                       'owned_parent_shutdown_attempts': stopped}, stream, indent=2, ensure_ascii=True)
        print('Support acceptance failed; retained authored workspace ' + str(work) + ': ' + str(error), file=sys.stderr)
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
