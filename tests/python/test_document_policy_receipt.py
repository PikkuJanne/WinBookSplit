"""Resealed synthetic policy-receipt controls; never native test evidence."""
from copy import deepcopy
from hashlib import sha256
import importlib.util
import json
from pathlib import Path
import sys
import unittest

ROOT = Path(__file__).resolve().parents[2]


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


guard = load('wbs_document_policy_unit_guard', ROOT / 'tests/document_policy/validate_document_policy_report.py')
console_fixture = load('wbs_document_policy_unit_console', ROOT / 'tests/python/console_receipt_fixture.py')


def compact(value):
    return json.dumps(value, ensure_ascii=True, separators=(',', ':'))


def reseal(row):
    if row['console_log'] is not None:
        text = '[ENGINE] ' + compact(row['interaction']['engine_invocations'][0]) + '\r\n'
        text += ''.join('[WBS-INTERACTION] ' + compact(item) + '\r\n' for item in row['interaction']['requests'])
        text += compact(row['outcome']['engine_result']) + '\r\n'
        text += '[PROCESS] ' + compact(row['process_summaries'][0]) + '\r\n'
        text += '[OPERATION-OUTCOME] ' + compact(row['outcome']) + '\r\n'
        row.update(console_log=text, log_sha256=sha256(text.encode()).hexdigest(), engine_records=[deepcopy(row['outcome']['engine_result'])])
        row['console_evidence'] = console_fixture.make(row['outcome'], row['log_sha256'], log_text=text,
            source_path=row['input_path'], source_observation=row['source_observations_before']['source'])
        row['console_evidence']['run_manifest']['settings'].update(non_interactive=True, no_pause=False)
        console_fixture.reseal(row['console_evidence'])
    row['stdout'] = row.get('display', '') + '[OUTCOME] ' + compact(row['outcome']) + '\r\n'
    for stream in ('stdout', 'stderr'):
        row[stream + '_sha256'] = sha256(row[stream].encode()).hexdigest()


def synthetic_rejection(*, preview=False):
    directory = ROOT / 'synthetic-policy-control-not-a-native-run'
    source, base, cwd = str(directory / 'b/source.pdf'), str(directory / 'o'), str(directory / 'c')
    kind = 'encrypted-owner-empty' + ('-preview' if preview else '')
    frame = {'protocol': 'winbooksplit.result', 'version': 1, 'mode': 'manual', 'status': 'unsupported',
             'code': 'unsupported_document', 'exit_code': 7, 'written_count': 0, 'execution': None, 'plan': None,
             'fallback_modes': [], 'warnings': [], 'diagnostic': None}
    final = {'protocol': 'winbooksplit.outcome', 'version': 1, 'mode': 'manual', 'status': 'unsupported',
             'code': 'unsupported_document', 'exit_code': 7, 'written_count': 0, 'final_directory': None, 'engine_result': frame}
    identity = {'sha256': 'a' * 64, 'size_bytes': 100, 'device': 1, 'inode': 2, 'attributes': 1}
    observed = {name: deepcopy(identity) for name in ('source', 'neighbor', 'prior')}
    parameters = ['-InputFile', source, '-OutputDirectory', base, '-PythonPath', str(directory / 'python.exe'),
                  '-Mode', 'Manual', '-StartPages', '1,3', '-NonInteractive', *(['-Preview'] if preview else [])]
    shell = str(directory / 'powershell.exe')
    process = {'JobAssigned': True, 'ParentStopped': True, 'DescendantsStopped': True, 'StreamsComplete': True,
               'ExitCode': 7, 'ResultRecordCount': 1}
    invocation = {'path': str(directory / 'python.exe'), 'arguments': ['-I', '-B', '-X', 'utf8',
                  str(directory / 'a/engine/winbooksplit_engine.py'), source, base, 'manual', '1,3']}
    row = {'id': 'PS51-' + kind, 'kind': kind, 'fixture_kind': 'encrypted-owner-empty', 'host_id': 'PS51', 'mode': 'manual',
           'source_read_only_observed': True, 'input_path': source, 'output_base': base, 'cwd': cwd,
           'source_observations_before': observed, 'source_observations_after': deepcopy(observed),
           'source_observations_after_cleanup': deepcopy(observed), 'output_members_after_cleanup': ['prior-output.pdf'], 'passed': True,
           'shell_executable': shell, 'python_executable': str(directory / 'python.exe'), 'parameters': parameters,
           'command': [shell, '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'RemoteSigned', '-File',
                       str(directory / 'a/WinBookSplit.ps1'), *parameters],
           'stdin': 'DEVNULL', 'stdin_utf8': None, 'modifications': [], 'actual_process': True, 'exit_code': 7,
           'pid': 123, 'timed_out': False, 'elapsed_seconds': .1, 'timeout_seconds': 90, 'stderr': '', 'outcome': final,
           'console_directory_name': None if preview else '.WinBookSplit-console-' + 'c' * 32,
           'output_members_after': ['prior-output.pdf'] if preview else ['prior-output.pdf', '.WinBookSplit-console-' + 'c' * 32],
           'publication': None, 'console_log': None if preview else '', 'console_evidence': None, 'log_sha256': None,
           'engine_records': [], 'process_summaries': [] if preview else [process],
           'interaction': None if preview else {'engine_invocations': [invocation], 'requests': [], 'replies': [], 'displayed_plans': []}}
    reseal(row)
    return row


def synthetic_export(application):
    manifest = application['console_evidence']['run_manifest']
    summary = {'protocol': 'winbooksplit.support', 'version': 1, 'application_version': '1.0.0',
        'runtime_versions': {**deepcopy(manifest['runtime_versions']), 'powershell': '5.1'},
        'settings': {key: manifest['settings'][key] for key in guard.support.SCALAR_SETTINGS}, 'plan': None,
        'outcome': {key: manifest['outcome'][key] for key in ('status', 'code', 'exit_code', 'written_count')},
        'warnings': deepcopy(manifest['warnings']),
        'log': {key: manifest['log'][key] for key in ('max_bytes', 'body_limit_bytes', 'bytes_written', 'limit_reached')}}
    directory = Path(application['input_path']).parent.parent
    frame = {'protocol': 'winbooksplit.support-export', 'version': 1, 'status': 'success',
             'code': 'support_export_complete', 'exit_code': 0, 'written_count': 1}
    row = {'id': 'PS51-unsupported-export', 'application_case_id': application['id'], 'cwd': application['cwd'],
        'command': [application['shell_executable'], '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'RemoteSigned', '-File',
                    str(directory / 'a/Export-WinBookSplitDiagnostics.ps1'), '-ManifestPath',
                    str(Path(application['output_base']) / application['console_directory_name'] / 'WinBookSplit_Run.json'),
                    '-OutputPath', str(Path(application['cwd']) / 'redacted-summary.json')],
        'actual_process': True, 'exit_code': 0, 'pid': 124, 'timed_out': False, 'elapsed_seconds': .1, 'timeout_seconds': 60,
        'stdin': 'DEVNULL', 'stdout': '[SUPPORT-EXPORT] ' + compact(frame) + '\r\n', 'stderr': '',
        'summary': summary, 'summary_raw': compact(summary)}
    for stream in ('stdout', 'stderr'):
        row[stream + '_sha256'] = sha256(row[stream].encode()).hexdigest()
    row['summary_sha256'] = sha256(row['summary_raw'].encode()).hexdigest()
    return row


def synthetic_deep_outline():
    row = synthetic_rejection()
    row.update(id='PS51-deep-outline', kind='deep-outline', fixture_kind='deep-outline', mode='1', exit_code=6)
    frame = row['outcome']['engine_result']
    frame.update(mode='1', status='invalid_input', code='invalid_outline', exit_code=6,
        warnings=[{'code': 'outline_limit', 'source_order': None, 'depth': 65,
                   'message': 'The outline tree exceeds the traversal limit.'}])
    row['outcome'].update(mode='1', status='invalid_input', code='invalid_outline', exit_code=6)
    row['process_summaries'][0]['ExitCode'] = 6
    row['parameters'][7:10] = ['Auto', '-BookmarkLevel', '1']
    row['command'] = [row['shell_executable'], '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'RemoteSigned',
        '-File', str(Path(row['input_path']).parent.parent / 'a/WinBookSplit.ps1'), *row['parameters']]
    row['interaction']['engine_invocations'][0]['arguments'][7:] = ['1', '']
    reseal(row)
    return row


class DocumentPolicyReceiptTests(unittest.TestCase):
    def test_metadata_only_rejection_and_no_write_preview_are_valid_synthetic_controls(self):
        for preview in (False, True):
            guard.validate_case(synthetic_rejection(preview=preview))

    def test_resealed_false_status_code_exit_and_plan_are_rejected(self):
        for values in ({'status': 'success'}, {'code': 'split_complete'}, {'exit_code': 0}, {'plan': {'total_pages': 4}},
                       {'fallback_modes': ['manual']}, {'execution': {}}, {'written_count': 1}, {'version': True}):
            row = synthetic_rejection()
            row['outcome']['engine_result'].update(values)
            for key in ('status', 'code', 'exit_code', 'written_count'):
                row['outcome'][key] = row['outcome']['engine_result'][key]
            reseal(row)
            with self.subTest(values=values), self.assertRaises(ValueError):
                guard.validate_case(row)

    def test_missing_authentic_log_or_engine_evidence_fails(self):
        for key, value in (('console_evidence', None), ('engine_records', []), ('process_summaries', [])):
            row = synthetic_rejection()
            row[key] = value
            with self.subTest(key=key), self.assertRaises(ValueError):
                guard.validate_case(row)
        for key in ('execution', 'warnings', 'diagnostic', 'fallback_modes'):
            row = synthetic_rejection()
            del row['outcome']['engine_result'][key]
            reseal(row)
            with self.subTest(result_field=key), self.assertRaises(ValueError):
                guard.validate_case(row)

    def test_resealed_claimed_capture_or_wrong_settings_are_rejected(self):
        for change in (lambda m: m['source_identity'].update(sha256='a' * 64),
                       lambda m: m['settings'].update(non_interactive=False),
                       lambda m: m['settings'].update(no_pause=True)):
            row = synthetic_rejection()
            change(row['console_evidence']['run_manifest'])
            console_fixture.reseal(row['console_evidence'])
            with self.assertRaises(ValueError):
                guard.validate_case(row)

    def test_resealed_success_label_or_password_prompt_fails(self):
        for display in ('Done.\n', 'Output: C:/invented\n', 'Enter password: ', guard.fixtures.USER_PASSWORD):
            row = synthetic_rejection()
            row['display'] = display
            reseal(row)
            with self.subTest(display=display), self.assertRaises(ValueError):
                guard.validate_case(row)

    def test_unowned_output_and_mutated_neighbor_are_rejected(self):
        row = synthetic_rejection()
        row['output_members_after'].append('.WinBookSplit-stage-invented')
        with self.assertRaises(ValueError):
            guard.validate_case(row)
        row = synthetic_rejection()
        row['source_observations_after']['neighbor']['sha256'] = 'b' * 64
        with self.assertRaises(ValueError):
            guard.validate_case(row)

    def test_resealed_engine_path_or_job_stop_contradiction_fails(self):
        for change in (lambda r: r['interaction']['engine_invocations'][0].update(path='C:/unrelated.exe'),
                       lambda r: r['interaction']['engine_invocations'][0]['arguments'].__setitem__(4, 'C:/unrelated.py'),
                       lambda r: r['process_summaries'][0].update(DescendantsStopped=False)):
            row = synthetic_rejection()
            change(row)
            reseal(row)
            with self.assertRaises(ValueError):
                guard.validate_case(row)

    def test_resealed_noninteractive_request_fails(self):
        row = synthetic_rejection()
        row['interaction']['requests'].append({'protocol': 'winbooksplit.interaction', 'stage': 'input_ready'})
        reseal(row)
        with self.assertRaises(ValueError):
            guard.validate_case(row)

    def test_preview_cannot_publish_diagnostics_or_request_confirmation(self):
        row = synthetic_rejection(preview=True)
        row['output_members_after'].append('.WinBookSplit-console-invented')
        with self.assertRaises(ValueError):
            guard.validate_case(row)
        row = synthetic_rejection(preview=True)
        row['display'] = 'Write these chapter PDFs? '
        reseal(row)
        with self.assertRaises(ValueError):
            guard.validate_case(row)

    def test_malformed_contract_remains_six_and_not_unsupported(self):
        frame = synthetic_rejection()['outcome']['engine_result']
        frame.update(status='read_error', code='unreadable_document', exit_code=6)
        guard.validate_result(frame, 'truncated')
        frame.update(status='invalid_input', code='invalid_document')
        guard.validate_result(frame, 'zero-pages')
        frame.update(code='invalid_outline', mode='1', warnings=[{'code': 'outline_limit', 'source_order': None,
            'depth': 65, 'message': 'The outline tree exceeds the traversal limit.'}])
        guard.validate_result(frame, 'deep-outline')
        with self.assertRaises(ValueError):
            guard.validate_result(frame, 'encrypted-user')

    def test_resealed_deep_outline_fixed_warning_is_valid_without_publication(self):
        guard.validate_case(synthetic_deep_outline())

    def test_resealed_deep_outline_warning_contradictions_are_rejected(self):
        for values in ({'code': 'navigation_link_dropped'}, {'depth': 64}, {'source_order': 1},
                       {'message': 'The outline was silently truncated.'}):
            row = synthetic_deep_outline()
            row['outcome']['engine_result']['warnings'][0].update(values)
            reseal(row)
            with self.subTest(values=values), self.assertRaisesRegex(ValueError, 'exact bounded traversal warning'):
                guard.validate_case(row)
        for warnings in ([], [synthetic_deep_outline()['outcome']['engine_result']['warnings'][0]] * 2):
            row = synthetic_deep_outline()
            row['outcome']['engine_result']['warnings'] = deepcopy(warnings)
            reseal(row)
            with self.subTest(warnings=warnings), self.assertRaisesRegex(ValueError, 'exact bounded traversal warning'):
                guard.validate_case(row)
        row = synthetic_rejection()
        row['outcome']['engine_result']['warnings'] = deepcopy(synthetic_deep_outline()['outcome']['engine_result']['warnings'])
        reseal(row)
        with self.assertRaisesRegex(ValueError, 'Preflight rejection falsely planned/warned chapter extraction'):
            guard.validate_case(row)

    def test_declared_batch_seam_changes_only_three_parameter_defaults(self):
        source = b'param(\r\n[string]$OutputDirectory,\r\n[string]$PythonPath,\r\n[switch]$NoPause,\r\n[switch]$NonInteractive\r\n)\r\nWrite-Host "authored"\r\n'
        changed = guard.controlled_powershell(source).decode('utf-8-sig')
        self.assertIn('[string]$OutputDirectory = $env:WBS_POLICY_OUTPUTDIRECTORY,', changed)
        self.assertIn('[switch]$NoPause = $true,', changed)
        self.assertTrue(changed.endswith('Write-Host "authored"\r\n'))
        self.assertNotIn('$NonInteractive =', changed)
        with self.assertRaises(ValueError):
            guard.controlled_powershell(source.replace(b'[switch]$NoPause,', b'[switch]$Different,'))

    def test_explicit_export_binds_argv_and_typed_result(self):
        application = synthetic_rejection()
        exported = synthetic_export(application)
        guard.validate_export(exported, application)
        exported['command'][6] = 'C:/unrelated.ps1'
        with self.assertRaises(ValueError):
            guard.validate_export(exported, application)
        for field in ('version', 'exit_code', 'written_count'):
            exported = synthetic_export(application)
            frame = json.loads(exported['stdout'][len('[SUPPORT-EXPORT] '):])
            frame[field] = bool(frame[field])
            exported['stdout'] = '[SUPPORT-EXPORT] ' + compact(frame) + '\r\n'
            exported['stdout_sha256'] = sha256(exported['stdout'].encode()).hexdigest()
            with self.subTest(field=field), self.assertRaises(ValueError):
                guard.validate_export(exported, application)


if __name__ == '__main__':
    unittest.main()
