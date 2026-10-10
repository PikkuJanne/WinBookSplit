"""M4-T02 actual PS5.1/PS7/BAT fail-closed document policy acceptance."""
from __future__ import annotations

import argparse
from base64 import b64encode
from datetime import datetime, timezone
from hashlib import sha256
import importlib.util
import json
import os
from pathlib import Path
import shutil
import sys
import tempfile

from pypdf import PdfReader

ROOT = Path(__file__).resolve().parents[2]


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


support = load('wbs_document_policy_support', ROOT / 'tests/support/characterize_support.py')
prompted = load('wbs_document_policy_prompted', ROOT / 'tests/launcher/characterize_launcher.py')
validator = load('wbs_document_policy_guard', ROOT / 'tests/document_policy/validate_document_policy_report.py')
fixtures = validator.fixtures
cli, history, paths = support.cli, support.history, support.paths
conversion, launchers, runner, require = support.conversion, support.launchers, support.runner, support.require
COMPLETED_CASES, EXPORT_CASES, LAST_CASE = [], [], None


def publication(base, frame):
    execution = frame['execution']
    final = Path(execution['final_directory'])
    members = {'.WinBookSplit-owner.json', 'WinBookSplit_Manifest.json', *[item['filename'] for item in execution['outputs']]}
    require(final.parent == base and {item.name for item in final.iterdir()} == members, 'Unexpected publication membership')
    manifest_path = final / 'WinBookSplit_Manifest.json'
    raw = manifest_path.read_bytes().decode('utf-8', 'strict')
    outputs = []
    for item in execution['outputs']:
        target = final / item['filename']
        reader = PdfReader(target)
        outputs.append({'filename': item['filename'], 'sha256': cli.digest(target), 'bytes': target.stat().st_size,
            'page_ids': [int(page['/WBSFixturePage']) for page in reader.pages],
            'content_sha256': conversion.page_content(reader), 'page_actions_absent': all('/AA' not in page for page in reader.pages),
            'inert_action_marker_absent': b'ORIGINAL_INERT_POLICY_ACTION' not in target.read_bytes()})
    return {'publication': outputs, 'publication_manifest': validator.support.strict_json(raw), 'publication_manifest_raw': raw,
            'publication_manifest_sha256': cli.digest(manifest_path), 'publication_members': sorted(members)}


def export_case(application, app, log, environment):
    global LAST_CASE
    destination = Path(application['cwd']) / 'redacted-summary.json'
    command = [application['shell_executable'], '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'RemoteSigned', '-File',
        str(app / 'Export-WinBookSplitDiagnostics.ps1'), '-ManifestPath', str(log.parent / 'WinBookSplit_Run.json'),
        '-OutputPath', str(destination)]
    LAST_CASE = {'id': application['host_id'] + '-unsupported-export', 'command': command, 'workspace': str(app.parent)}
    observed = cli.run_unavailable_stdin(command, Path(application['cwd']), environment, timeout=60)
    LAST_CASE.update(observed)
    require(destination.is_file(), 'Explicit rejected-input export did not create its single summary')
    raw = destination.read_bytes().decode('utf-8', 'strict')
    record = {'id': application['host_id'] + '-unsupported-export', 'application_case_id': application['id'],
        'actual_process': True, 'summary_raw': raw, 'summary_sha256': cli.digest(destination),
        'summary': validator.support.strict_json(raw), **observed}
    validator.validate_export(record, application)
    # Exact newly authored export, never a source-directed recursive cleanup.
    destination.unlink()
    record.update(passed=True, owned_export_removed=not destination.exists())
    EXPORT_CASES.append(record)
    LAST_CASE = application
    return record


def application_case(work, host, kind, sources, *, batch=False):
    global LAST_CASE
    label = ('BAT' if batch else host['id']) + '-' + kind
    directory = work / ('case-' + str(len(COMPLETED_CASES) + 1))
    directory.mkdir()
    app, books, cwd, base = [directory / name for name in ('a', 'b', 'c', 'o')]
    copied = support.copy_application(app)
    if batch:
        ps = app / 'WinBookSplit.ps1'
        ps.write_bytes(validator.controlled_powershell(ps.read_bytes()))
    actual = {name: cli.digest(app / name) for name in copied}
    for destination in (books, cwd, base):
        destination.mkdir()
    source = books / 'Å日本 [1] & %WBS_POLICY_EXPAND%! $(literal).pdf'
    fixture = validator.fixture_kind(kind)
    shutil.copyfile(sources[fixture], source)
    neighbor, prior = books / 'authored-neighbor.txt', base / 'prior-output.pdf'
    neighbor.write_bytes(b'Authored neighboring bytes remain immutable.\n')
    prior.write_bytes(b'Authored earlier output remains immutable.\n')
    mode = validator.mode_for(kind)
    interactive = batch or kind.endswith('-interactive')
    environment = support.controlled_environment(cwd, host['shell_executable'])
    system = Path(os.environ['SystemRoot'])
    wrapper_text = wrapper_sha = None
    if batch:
        environment.update(WBS_POLICY_BAT=str(app / 'WinBookSplit.bat'), WBS_POLICY_INPUT=str(source),
            WBS_POLICY_OUTPUTDIRECTORY=str(base), WBS_POLICY_PYTHONPATH=sys.executable,
            WBS_POLICY_EXPAND='WRONG_EXPANDED_VALUE')
        wrapper = directory / 'invoke.cmd'
        wrapper.write_bytes(validator.WRAPPER.encode('ascii'))
        wrapper_text, wrapper_sha = validator.WRAPPER, cli.digest(wrapper)
        parameters = [str(source)]
        command = [str(system / 'System32/cmd.exe'), '/d', '/v:off', '/c', str(wrapper)]
        answers = 'M\n1,3\nY\nN\n' if kind == 'ordinary-manual' else 'M\n'
    else:
        parameters = ['-InputFile', str(source), '-OutputDirectory', str(base), '-PythonPath', sys.executable,
                      '-Mode', 'Manual' if mode == 'manual' else 'Auto']
        parameters += ['-StartPages', '1,3'] if mode == 'manual' else ['-BookmarkLevel', mode]
        parameters += ['-NoPause'] if interactive else ['-NonInteractive']
        if kind.endswith('-preview'):
            parameters.append('-Preview')
        command = [host['shell_executable'], '-NoProfile', *([] if interactive else ['-NonInteractive']),
                   '-ExecutionPolicy', 'RemoteSigned', '-File', str(app / 'WinBookSplit.ps1'), *parameters]
        answers = None
    with paths.read_only_source(source):
        before = {name: cli.identity(item) for name, item in (('source', source), ('neighbor', neighbor), ('prior', prior))}
        content = conversion.page_content(PdfReader(source)) if validator.expected_exit(kind) == 0 else None
        LAST_CASE = {'id': label, 'command': command, 'cwd': str(cwd), 'workspace': str(directory),
                     'input_path': str(source), 'source_observations_before': before}
        observed = prompted.run_prompted(command, cwd, environment, answers, timeout=90) if batch else \
            cli.run_unavailable_stdin(command, cwd, environment, timeout=90)
        LAST_CASE.update(observed)
        final = cli.outcome(observed['stdout'])
        LAST_CASE['outcome'] = final
        log, frames, transports, log_hash = cli.console_records(observed['stdout'], base)
        text = None if log is None else log.read_bytes().decode('utf-8', 'strict')
        console = None if log is None else launchers.authenticate_console_manifest(log, final,
            source_path=source, source_observation=before['source'])
        record = {'id': label, 'kind': kind, 'fixture_kind': fixture, 'host_id': 'BAT' if batch else host['id'],
            'host_version': host['host_version'], 'shell_executable': host['shell_executable'], 'python_executable': sys.executable,
            'cmd_executable': str(system / 'System32/cmd.exe'), 'mode': mode, 'parameters': parameters,
            'application_sha256': copied, 'actual_application_sha256': actual,
            'modifications': validator.MODIFICATIONS if batch else [], 'wrapper_text': wrapper_text, 'wrapper_sha256': wrapper_sha,
            'input_path': str(source), 'output_base': str(base), 'actual_process': True, 'source_read_only_observed': True,
            'source_observations_before': before,
            'source_observations_after': {name: cli.identity(item) for name, item in (('source', source), ('neighbor', neighbor), ('prior', prior))},
            'outcome': final, 'console_log': text, 'log_sha256': log_hash, 'console_evidence': console,
            'console_directory_name': None if log is None else log.parent.name,
            'engine_records': frames, 'process_summaries': transports,
            'interaction': None if text is None else launchers.interactions.capture(text),
            'expected_content_sha256': content, 'publication': None,
            'output_members_after': sorted(item.name for item in base.iterdir()),
            'stdin_utf8': None, **observed}
        LAST_CASE = record
        frame = final.get('engine_result')
        if frame and frame.get('execution'):
            record.update(publication(base, frame))
        # Validate all no-write and native evidence before acquiring any cleanup permission.
        validator.validate_case(record, cleanup_complete=False)
        if kind == 'encrypted-owner-empty' and not batch:
            require(log is not None, 'Rejected input has no finalized local export source')
            export_case(record, app, log, environment)
        if frame and frame.get('execution'):
            conversion.remove_publication(base, frame['execution'], set(record['publication_members']))
        if log is not None:
            launchers.remove_console_directory(log, base, final, source_path=source, source_observation=before['source'])
        after = {name: cli.identity(item) for name, item in (('source', source), ('neighbor', neighbor), ('prior', prior))}
        require(before == after == record['source_observations_after'], 'Source/neighbor/prior immutable identity changed')
        require({item.name for item in base.iterdir()} == {prior.name}, 'Unexplained output remains after exact authenticated cleanup')
        require(actual == {name: cli.digest(app / name) for name in actual}, 'Copied application bytes changed')
        record.update(source_observations_after_cleanup=after, output_members_after_cleanup=[prior.name],
                      application_unchanged=True, passed=True)
        validator.validate_case(record)
        COMPLETED_CASES.append(record)
        LAST_CASE = record
    return record


def characterize(work, shells):
    require(os.name == 'nt' and len(shells) == 2, 'Actual Windows and both supported hosts required')
    before = runner.source_manifest()
    probe = cli.process_tests.host_probe(work)
    hosts = [paths.host_observation(shell, work, probe) for shell in shells]
    require({host['id'] for host in hosts} == {'PS51', 'PS7'}, 'Two actual distinct supported hosts required')
    sources, provenance = fixtures.generate_fixtures(work / 'fixtures')
    cases = [application_case(work, host, kind, sources) for host in hosts for kind in validator.PS_KINDS]
    ps51 = next(host for host in hosts if host['id'] == 'PS51')
    cases += [application_case(work, ps51, kind, sources, batch=True) for kind in validator.BAT_KINDS]
    for host in hosts:
        host['policies_after'] = paths.host_observation(Path(host['shell_executable']), work, probe)['stored_policies']
        require(host['stored_policies'] == host['policies_after'], 'Stored execution policies changed')
    require(before == runner.source_manifest(), 'Repository source changed during native document-policy acceptance')
    return {'schema_version': 1, 'task_id': 'M4-T02', 'result': 'DOCUMENT_POLICY_REGRESSION_PASSED', 'success': True, 'exit_code': 0,
        'observed_at': datetime.now(timezone.utc).isoformat(), 'acceptance_ids': ['AC-070', 'AC-071', 'AC-072'],
        'tested_path_sha256': before, 'source_unchanged': True, 'host_cases': hosts, 'cases': cases, 'export_cases': EXPORT_CASES,
        'case_count': len(cases) + len(EXPORT_CASES), 'application_case_count': len(cases), 'export_case_count': len(EXPORT_CASES),
        'fixture_provenance': provenance,
        'batch_original_powershell_raw_base64': b64encode((ROOT / 'WinBookSplit.ps1').read_bytes()).decode('ascii'),
        'human_or_GUI_tested': False, 'environment': {'python': sys.version, 'python_executable': sys.executable, 'pypdf': history.pypdf.__version__},
        'limits': ['Generated original inert PDF fixtures only; no private input, human, Explorer or GUI evidence.',
                   'Signature dictionary/field detection only; no cryptographic signature-validity proof.',
                   'BAT remains shipped bytes; three exact copied PowerShell parameter defaults provide pinned runtime/output and suppress pause.',
                   'Policy rejections may persist authenticated console diagnostics; no extraction, working-PDF, plan, publication or fallback.',
                   'Ordinary and page-AA controls verify reopened physical-page identities/content and unchanged source/neighbor/prior files.']}


def main():
    sys.stdout.reconfigure(encoding='utf-8', errors='backslashreplace')
    sys.stderr.reconfigure(encoding='utf-8', errors='backslashreplace')
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--report', required=True, type=Path)
    parser.add_argument('--shell-path', required=True, action='append', type=Path)
    args = parser.parse_args()
    target = runner.new_external_path(args.report)
    temporary = tempfile.TemporaryDirectory(prefix='DP-')
    work = Path(temporary.name).resolve()
    try:
        require(sys.flags.isolated and sys.dont_write_bytecode, 'Use isolated developer Python -I -B')
        require(len(args.shell_path) == 2 and all(path.is_absolute() and path.is_file() for path in args.shell_path),
                'Both existing absolute supported host paths required')
        shells = [str(path.resolve()) for path in args.shell_path]
        report = characterize(work, [Path(path) for path in shells])
        report['owned_temp_removed'] = False
        validator.validate_document_policy_report(report, shells, cleanup_complete=False)
        temporary.cleanup()
        report['owned_temp_removed'] = not work.exists()
        validator.validate_document_policy_report(report, shells)
        with target.open('x', encoding='utf-8', newline='\n') as stream:
            json.dump(report, stream, indent=2, ensure_ascii=True)
            stream.write('\n')
        print('Document-policy acceptance passed: ' + str(report['case_count']) + ' actual native application/export controls')
        return 0
    except BaseException as error:
        temporary._finalizer.detach()
        with target.open('x', encoding='utf-8', newline='\n') as stream:
            json.dump({'schema_version': 1, 'task_id': 'M4-T02', 'result': 'DOCUMENT_POLICY_REGRESSION_FAILED', 'success': False,
                'exit_code': 1, 'error': str(error), 'workspace_retained': str(work) if work.exists() else None,
                'owned_temp_removed': not work.exists(), 'cleanup_safe': False, 'completed_cases': COMPLETED_CASES,
                'completed_export_cases': EXPORT_CASES, 'last_case': LAST_CASE}, stream, indent=2, ensure_ascii=True)
            stream.write('\n')
        print('Document-policy acceptance failed; retained authored workspace ' + str(work) + ': ' + str(error), file=sys.stderr)
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
