"""Synthetic resealed support-receipt mutations; no native acceptance claim."""
from copy import deepcopy
from hashlib import sha256
import importlib.util
import json
from pathlib import Path
import unittest

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location('wbs_support_unit_validator', ROOT / 'tests/support/validate_support_report.py')
validator = importlib.util.module_from_spec(spec)
spec.loader.exec_module(validator)


def compact(value):
    return json.dumps(value, ensure_ascii=True, separators=(',', ':'))


def seal(row):
    row['console_log'] = ''.join(compact(item) + '\r\n' for item in row['engine_records'])
    row['console_log'] += ''.join('[PROCESS] ' + compact(item) + '\r\n' for item in row['process_summaries'])
    row['console_log'] += '[OPERATION-OUTCOME] ' + compact(row['outcome']) + '\r\n'
    row['console_log_sha256'] = validator.digest(row['console_log'])
    row['run_manifest']['log']['bytes_written'] = len(row['console_log'].encode())
    row['run_manifest_raw'] = compact(row['run_manifest']) + '\n'
    row['run_manifest_sha256'] = validator.digest(row['run_manifest_raw'])
    row['stdout'] = row.get('display', '') + '[OUTCOME] ' + compact(row['outcome']) + '\r\n'
    for name in ('stdout', 'stderr'):
        row[name + '_sha256'] = validator.digest(row[name])


def synthetic_success():
    source, base, folder = 'C:/synthetic/source.pdf', 'C:/synthetic/out', 'C:/synthetic/out/complete'
    source_identity = {'path': source, 'sha256': 'a' * 64, 'size_bytes': 100, 'binding': 'reader_snapshot'}
    plan = {'mode': 'manual', 'total_pages': 6, 'ranges': [[0, 2], [2, 4], [4, 6]],
            'coverage': {'complete': True, 'covered_pages': 6, 'section_count': 3},
            'normalized_inputs': {'manual_starts': [1, 3, 5]}, 'source_identity': deepcopy(source_identity)}
    plan['entries'] = [{'sequence': index, 'title': 'Authored section ' + str(index), 'start': start, 'end': end,
                        'parent_id': None, 'reason': 'manual', 'filename': str(index) + '.pdf', 'warnings': []}
                       for index, (start, end) in enumerate(plan['ranges'], 1)]
    outputs = [{**deepcopy(entry), 'page_count': entry['end'] - entry['start'], 'size_bytes': 100, 'sha256': 'f' * 64}
               for entry in plan['entries']]
    manifest = {'schema_version': 1, 'status': 'complete', 'mode': 'manual', 'run_id': 'b' * 32,
                'final_directory': folder, 'written_count': 3, 'total_pages': 6, 'source_identity': deepcopy(source_identity),
                'coverage': deepcopy(plan['coverage']), 'outputs': outputs}
    execution = {**deepcopy(manifest), 'manifest_filename': 'WinBookSplit_Manifest.json', 'manifest': deepcopy(manifest)}
    engine = {'protocol': 'winbooksplit.result', 'version': 1, 'mode': 'manual', 'status': 'success',
              'code': 'split_complete', 'exit_code': 0, 'written_count': 3, 'execution': execution, 'plan': deepcopy(plan), 'warnings': []}
    final = {'protocol': 'winbooksplit.outcome', 'version': 1, 'mode': 'manual', 'status': 'success', 'code': 'split_complete',
             'exit_code': 0, 'written_count': 3, 'final_directory': folder, 'engine_result': engine}
    identity = {'sha256': 'a' * 64, 'size_bytes': 100, 'device': 1, 'inode': 2, 'attributes': 1}
    observations = {name: deepcopy(identity) for name in ('source', 'neighbor', 'prior')}
    row = {'id': 'PS51-success-pdf', 'kind': 'success-pdf', 'host_id': 'PS51', 'host_version': '5.1.26100.9444',
           'actual_process': True, 'passed': True, 'pid': 123, 'exit_code': 0, 'elapsed_seconds': 1, 'timeout_seconds': 90,
           'timed_out': False, 'command': ['C:/synthetic/powershell.exe', '-File', 'C:/synthetic/app.ps1',
               '-InputFile', source, '-OutputDirectory', base, '-Mode', 'Manual', '-StartPages', '1,3,5', '-NonInteractive'],
           'parameters': ['-InputFile', source, '-OutputDirectory', base, '-Mode', 'Manual', '-StartPages', '1,3,5', '-NonInteractive'],
           'shell_executable': 'C:/synthetic/powershell.exe', 'stderr': '', 'display': 'Output: ' + folder + '\nDone.\n',
           'application_sha256': {name: 'c' * 64 for name in validator.APPLICATION}, 'fault_injection': None,
           'actual_application_sha256': {name: 'c' * 64 for name in validator.APPLICATION},
           'application_unchanged': True, 'source_read_only_observed': True, 'source_observations_before': observations,
           'source_observations_after': deepcopy(observations), 'source_observations_after_cleanup': deepcopy(observations),
           'output_members_after_cleanup': ['prior-output.pdf'], 'input_path': source, 'output_base': base,
           'input_kind': 'pdf', 'outcome': final, 'expected_plan': deepcopy(plan), 'planning_warnings': [], 'parser_warnings': [],
           'console_members': ['.WinBookSplit-console-owner.json', 'console.log', 'WinBookSplit_Run.json'],
           'engine_records': [deepcopy(engine)], 'process_summaries': [{'JobAssigned': True, 'ParentStopped': True, 'DescendantsStopped': True,
                'StreamsComplete': True, 'InputWriterStopped': True, 'ExitCode': 0}],
           'expected_content_sha256': ['d' * 64] * 6,
           'publication': {'manifest': manifest, 'manifest_raw': compact(manifest), 'manifest_sha256': validator.digest(compact(manifest)),
                           'content_sha256': ['d' * 64] * 6,
                           'members': ['.WinBookSplit-owner.json', 'WinBookSplit_Manifest.json', '1.pdf', '2.pdf', '3.pdf']}}
    row['run_manifest'] = {'protocol': 'winbooksplit.run', 'version': 1, 'application_version': '1.0.0-dev', 'run_id': 'e' * 32,
        'started_utc': '2026-10-10T00:00:00+00:00', 'finished_utc': '2026-10-10T00:00:01+00:00', 'diagnostics_finalized': True,
        'runtime_versions': {'powershell': row['host_version'], 'python': '3.14.8', 'pypdf': '6.19.0', 'calibre': None},
        'settings': {'mode': 'manual', 'input_kind': 'pdf', 'preview': False, 'non_interactive': True, 'no_pause': False,
                     'keep_converted_pdf': False, 'conversion_timeout': 1800, 'process_timeout': 3600, 'normalized_inputs': plan['normalized_inputs']},
        'source_identity': source_identity, 'plan': deepcopy(plan), 'engine_result': deepcopy(engine), 'outcome': deepcopy(final),
        'warnings': {'planning': [], 'parser': [], 'parser_suppressed_count': 0, 'parser_message_truncated_count': 0},
        'log': {'filename': 'console.log', 'max_bytes': 50331648, 'body_limit_bytes': 33554432, 'bytes_written': 0, 'limit_reached': False}}
    seal(row)
    return row


def summary(manifest):
    return {'protocol': 'winbooksplit.support', 'version': 1, 'application_version': '1.0.0-dev',
            'runtime_versions': {**deepcopy(manifest['runtime_versions']), 'powershell': '5.1'},
            'settings': {key: manifest['settings'][key] for key in validator.SCALAR_SETTINGS},
            'plan': None if manifest['plan'] is None else {key: deepcopy(manifest['plan'][key]) for key in ('total_pages', 'ranges', 'coverage')},
            'outcome': {key: manifest['outcome'][key] for key in ('status', 'code', 'exit_code', 'written_count')},
            'warnings': deepcopy(manifest['warnings']),
            'log': {key: manifest['log'][key] for key in ('max_bytes', 'body_limit_bytes', 'bytes_written', 'limit_reached')}}


class SupportReceiptTests(unittest.TestCase):
    def test_synthetic_positive_binds_raw_log_manifest_and_engine(self):
        row = synthetic_success()
        validator.validate_application_case(row)
        validator.validate_summary(summary(row['run_manifest']), row['run_manifest'])

    def test_resealed_execution_cannot_disagree_with_shared_plan(self):
        mutations = [lambda data: data['outputs'][0].update(start=1),
                     lambda data: data['source_identity'].update(sha256='e' * 64),
                     lambda data: data.update(mode='1'),
                     lambda data: data['coverage'].update(complete=1),
                     lambda data: data['outputs'][0].update(page_count=1),
                     lambda data: data.update(written_count=2)]
        for mutate in mutations:
            row = synthetic_success()
            engine = row['outcome']['engine_result']
            mutate(engine['execution'])
            # Independently reseal all duplicated records and byte hashes. The
            # unchanged plan is the separate authority for every actual output.
            engine['execution']['manifest'] = {key: deepcopy(value) for key, value in engine['execution'].items()
                                                if key not in {'manifest', 'manifest_filename'}}
            row['publication']['manifest'] = deepcopy(engine['execution']['manifest'])
            row['publication']['manifest_raw'] = compact(row['publication']['manifest'])
            row['publication']['manifest_sha256'] = validator.digest(row['publication']['manifest_raw'])
            row['engine_records'] = [deepcopy(engine)]
            row['run_manifest']['engine_result'] = deepcopy(engine)
            row['run_manifest']['outcome'] = deepcopy(row['outcome'])
            seal(row)
            with self.assertRaises(ValueError):
                validator.validate_application_case(row)

    def test_resealed_export_input_and_summary_cannot_agree_on_false_outcome(self):
        mutations = [lambda data: data['outcome'].update(status='failed'),
                     lambda data: data['outcome'].update(written_count=4),
                     lambda data: data.update(plan=None),
                     lambda data: data['outcome'].update(status='cancelled'),
                     lambda data: data['outcome'].update(status='read_error', code='invalid_document', exit_code=6, written_count=0)]
        for mutate in mutations:
            manifest = synthetic_success()['run_manifest']
            mutate(manifest)
            # Both selected input and exported JSON agree; only the operation
            # contract can reject this internally consistent false receipt.
            resealed_input = validator.strict_json(compact(manifest))
            resealed_summary = validator.strict_json(compact(summary(resealed_input)))
            with self.assertRaises(ValueError):
                validator.validate_summary(resealed_summary, resealed_input)

    def test_resealed_manifest_cannot_change_final_status_or_versions(self):
        for field, value in (('outcome', {**synthetic_success()['outcome'], 'status': 'failed'}),
                             ('runtime_versions', {'powershell': '7.6.5', 'python': '3.14.8', 'pypdf': '6.19.0', 'calibre': None}),
                             ('diagnostics_finalized', False)):
            row = synthetic_success()
            row['run_manifest'][field] = value
            seal(row)
            with self.subTest(field=field), self.assertRaises(ValueError):
                validator.validate_application_case(row)

    def test_resealed_manifest_cannot_invent_source_hash_or_normalized_inputs(self):
        for mutate in (lambda row: row['run_manifest']['source_identity'].update(sha256='f' * 64),
                       lambda row: row['run_manifest']['settings'].update(normalized_inputs={'manual_starts': [2, 4, 6]})):
            row = synthetic_success()
            mutate(row)
            seal(row)
            with self.assertRaises(ValueError):
                validator.validate_application_case(row)

    def test_resealed_plan_cannot_omit_first_page(self):
        row = synthetic_success()
        row['expected_plan']['ranges'][0][0] = 1
        row['run_manifest']['plan'] = deepcopy(row['expected_plan'])
        seal(row)
        with self.assertRaises(ValueError):
            validator.validate_application_case(row)

    def test_false_done_or_discarded_publication_on_incomplete_rejects(self):
        row = synthetic_success()
        row.update(kind='log-fault-success', id='PS51-log-fault-success', exit_code=6)
        sentinel = 'Authored M3-T05 log finalization failure.'
        row['actual_application_sha256']['WinBookSplit.ps1'] = 'f' * 64
        row['fault_injection'] = {'path': 'WinBookSplit.ps1', 'sentinel': sentinel, 'original_sha256': row['application_sha256']['WinBookSplit.ps1'],
                                  'modified_sha256': 'f' * 64}
        row['outcome'].update(status='incomplete', code='console_finalize_failed', exit_code=6, written_count=0)
        row['display'] = sentinel + '\nCompleted engine output retained: ' + row['outcome']['final_directory'] + '\n'
        row['run_manifest']['outcome'] = deepcopy(row['outcome'])
        seal(row)
        validator.validate_application_case(row)
        for mutate in (lambda changed: changed.update(display=changed['display'] + 'Done.\n'),
                       lambda changed: changed.update(publication=None)):
            changed = deepcopy(row)
            mutate(changed)
            seal(changed)
            with self.assertRaises(ValueError):
                validator.validate_application_case(changed)

    def test_log_byte_bounds_and_last_footer_are_authoritative(self):
        row = synthetic_success()
        row['console_log'] += '[OPERATION-OUTCOME] ' + compact({**row['outcome'], 'code': 'console_finalize_failed'}) + '\r\n'
        row['console_log_sha256'] = validator.digest(row['console_log'])
        row['run_manifest']['log']['bytes_written'] = len(row['console_log'].encode())
        row['run_manifest_raw'] = compact(row['run_manifest'])
        row['run_manifest_sha256'] = validator.digest(row['run_manifest_raw'])
        with self.assertRaises(ValueError):
            validator.validate_application_case(row)

    def test_parser_records_require_typed_truncation_severity_and_counts(self):
        row = synthetic_success()
        warning = {'category': 'pdf_parser', 'code': 'pypdf_parser_warning', 'severity': 'WARNING', 'message': 'Raw\nlocal message', 'truncated': 'false'}
        row['parser_warnings'] = [warning]
        row['run_manifest']['warnings']['parser'] = [warning]
        seal(row)
        with self.assertRaises(ValueError):
            validator.validate_application_case(row)
        row = synthetic_success()
        row['run_manifest']['warnings']['parser_suppressed_count'] = False
        seal(row)
        with self.assertRaises(ValueError):
            validator.validate_application_case(row)

    def test_export_whitelist_rejects_private_free_text_even_when_resealed(self):
        manifest = synthetic_success()['run_manifest']
        clean = summary(manifest)
        for mutation in (lambda data: data.update(source_identity={'path': 'C:/Users/MockPrivate/source.pdf'}),
                         lambda data: data['plan'].update(title='Mock private title'),
                         lambda data: data['outcome'].update(message='MOCK_SECRET'),
                         lambda data: data['settings'].update(normalized_inputs={'secret': 'MOCK_SECRET'})):
            bad = deepcopy(clean)
            mutation(bad)
            with self.assertRaises(ValueError):
                validator.validate_summary(bad, manifest)

    def test_export_cannot_alter_typed_outcome_or_physical_partition(self):
        manifest = synthetic_success()['run_manifest']
        for mutation in (lambda data: data['outcome'].update(exit_code=6),
                         lambda data: data['plan']['ranges'][0].__setitem__(0, 1)):
            bad = summary(manifest)
            mutation(bad)
            with self.assertRaises(ValueError):
                validator.validate_summary(bad, manifest)

    def test_summary_boolean_integer_aliases_are_rejected(self):
        manifest = synthetic_success()['run_manifest']
        mutations = [lambda data: data['settings'].update(non_interactive=1),
                     lambda data: data['plan']['coverage'].update(complete=1),
                     lambda data: data['warnings'].update(parser_suppressed_count=False),
                     lambda data: data['warnings'].update(parser_message_truncated_count=False),
                     lambda data: data['log'].update(limit_reached=0)]
        for mutate in mutations:
            bad = summary(manifest)
            mutate(bad)
            # Hashes cannot turn a JSON integer into a Boolean (or vice versa).
            raw = compact(bad)
            resealed = validator.strict_json(raw)
            with self.assertRaises(ValueError):
                validator.validate_summary(resealed, manifest)

    def test_fidelity_warning_categories_are_explicit_and_unknown_tokens_reject(self):
        categories = {'cross_chapter_link_dropped': 'navigation', 'navigation_link_dropped': 'navigation',
                      'article_navigation_dropped': 'navigation', 'annotation_dropped': 'annotation',
                      'annotation_relation_dropped': 'annotation', 'metadata_omitted': 'metadata',
                      'metadata_normalized': 'metadata'}
        manifest = synthetic_success()['run_manifest']
        manifest['warnings']['planning'] = [{'code': code, 'source_order': None, 'depth': None,
                                            'message': 'MOCK_PRIVATE_WARNING_TEXT'} for code in categories]
        clean = summary(manifest)
        clean['warnings']['planning'] = [{'code': code, 'category': category} for code, category in categories.items()]
        validator.validate_summary(clean, manifest)
        for index, code in enumerate(categories):
            bad = deepcopy(clean)
            bad['warnings']['planning'][index]['category'] = 'bookmark'
            with self.subTest(code=code), self.assertRaises(ValueError):
                validator.validate_summary(validator.strict_json(compact(bad)), manifest)
        # Input and summary agree on the unknown token; fixed schema still rejects it.
        manifest['warnings']['planning'][0]['code'] = 'unknown_navigation_token'
        clean['warnings']['planning'][0]['code'] = 'unknown_navigation_token'
        with self.assertRaisesRegex(ValueError, 'unknown selected token'):
            validator.validate_summary(validator.strict_json(compact(clean)), validator.strict_json(compact(manifest)))

    def test_duplicate_and_nonfinite_json_are_rejected(self):
        for raw in ('{"protocol":"a","protocol":"b"}', '{"count":NaN}'):
            with self.assertRaises(ValueError):
                validator.strict_json(raw)

    def test_changed_prior_and_lying_native_exit_reject(self):
        for mutation in (lambda row: row.update(exit_code=6),
                         lambda row: row['source_observations_after']['prior'].update(sha256='f' * 64)):
            row = synthetic_success()
            mutation(row)
            with self.assertRaises(ValueError):
                validator.validate_application_case(row)


if __name__ == '__main__':
    unittest.main()
