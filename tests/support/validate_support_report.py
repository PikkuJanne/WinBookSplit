"""Fail-closed validation of actual M3-T05 native support receipts.

Synthetic unit receipts exercise this validator; they are not native passes.
"""
from __future__ import annotations

from datetime import datetime
from hashlib import sha256
import json
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parents[2]
APPLICATION_KINDS = ('success-pdf', 'success-epub', 'success-azw3', 'failed-pdf', 'cancelled',
                     'no-plan', 'parser-warning', 'bookmark-warning', 'log-fault-success',
                     'manifest-fault-success', 'log-fault-primary', 'manifest-fault-primary',
                     'log-fault-cancelled', 'manifest-fault-cancelled', 'preview', 'preflight')
EXPORT_KINDS = ('success', 'failed', 'cancelled', 'incomplete', 'no-plan', 'redacted', 'existing', 'same-path', 'hardlink',
                'relative-input', 'relative-output', 'duplicate-keys', 'invalid-utf8', 'wrong-protocol', 'reparse-output', 'pending-manifest', 'selected-secret')
APPLICATION = ('VERSION', 'WinBookSplit.ps1', 'WinBookSplit.bat', 'Export-WinBookSplitDiagnostics.ps1', 'requirements.txt',
               'engine/WinBookSplit.Paths.ps1', 'engine/WinBookSplit.Diagnostics.ps1', 'engine/WinBookSplit.Runtime.ps1',
               'engine/WinBookSplit.Process.ps1', 'engine/WinBookSplit.Logging.ps1', 'engine/WinBookSplit.Support.ps1', 'engine/WinBookSplit.Outcomes.json',
               'engine/winbooksplit_engine.py', 'engine/winbooksplit_windows.py', 'engine/winbooksplit_conversion.py', 'engine/winbooksplit_job.py')
SCALAR_SETTINGS = ('mode', 'input_kind', 'preview', 'non_interactive', 'no_pause', 'keep_converted_pdf', 'conversion_timeout', 'process_timeout')


def need(condition, message):
    if not condition:
        raise ValueError(message)


def digest(text):
    return sha256(text.encode('utf-8')).hexdigest()


def strict_json(text):
    def pairs(items):
        result = {}
        for key, value in items:
            need(key not in result, 'Duplicate JSON property')
            result[key] = value
        return result
    return json.loads(text, object_pairs_hook=pairs, parse_constant=lambda value: (_ for _ in ()).throw(ValueError('Nonfinite JSON')))


def records(text, prefix):
    return [strict_json(line[len(prefix):]) for line in text.splitlines() if line.startswith(prefix)]


def partition(plan):
    need(isinstance(plan, dict) and type(plan.get('total_pages')) is int and plan['total_pages'] > 0,
         'Missing positive physical-page count')
    ranges = plan.get('ranges')
    need(isinstance(ranges, list) and bool(ranges), 'Missing nonempty shared ranges')
    cursor = 0
    for item in ranges:
        need(isinstance(item, list) and len(item) == 2 and all(type(number) is int for number in item)
             and item[0] == cursor and item[1] > cursor, 'Ranges omit, duplicate or reorder physical pages')
        cursor = item[1]
    coverage = plan.get('coverage')
    need(isinstance(coverage, dict) and type(coverage.get('complete')) is bool and type(coverage.get('covered_pages')) is int
         and type(coverage.get('section_count')) is int and cursor == plan['total_pages'] and coverage == {
        'complete': True, 'covered_pages': cursor, 'section_count': len(ranges)}, 'Shared coverage is incomplete')
    return ranges


def typed_outcome(final, engine=None):
    codes = json.loads((ROOT / 'engine/WinBookSplit.Outcomes.json').read_text(encoding='utf-8-sig'))['codes']
    cancelled_cleanup = (isinstance(final, dict) and final.get('code') == 'conversion_cleanup_failed'
                         and final.get('exit_code') == 130 and isinstance(engine, dict)
                         and (engine.get('diagnostic') or {}).get('primary_code') in {'conversion_cancelled', 'conversion_timeout'})
    need(isinstance(final, dict) and isinstance(final.get('code'), str) and final.get('code') in codes and type(final.get('exit_code')) is int
         and (final['exit_code'] == codes[final['code']] or cancelled_cleanup)
         and type(final.get('written_count')) is int and final['written_count'] >= 0,
         'Typed final outcome/code/count differs from shipped contract')
    need(final.get('status') in {'success', 'preview', 'failed', 'cancelled', 'timeout', 'incomplete', 'no_plan', 'error', 'read_error', 'write_error', 'invalid_input', 'unsupported'},
         'Unknown final status')


def validate_execution(row, result, publication):
    """Bind every retained output and physical-page range to the exact shared plan."""
    plan, execution = row.get('expected_plan'), result.get('execution')
    ranges = partition(plan)
    entries = plan.get('entries')
    need(isinstance(entries, list) and len(entries) == len(ranges), 'Shared entries/count differ from physical ranges')
    need(isinstance(execution, dict) and type(execution.get('schema_version')) is int and execution['schema_version'] == 1
         and execution.get('status') == 'complete' and isinstance(execution.get('run_id'), str)
         and re.fullmatch('[0-9a-f]{32}', execution['run_id'])
         and type(execution.get('total_pages')) is int and execution['total_pages'] == plan['total_pages']
         and execution.get('mode') == plan.get('mode') == result.get('mode')
         and execution.get('source_identity') == plan.get('source_identity')
         and execution.get('coverage') == plan['coverage'], 'Completed execution differs from captured shared plan/source')
    partition({'total_pages': execution['total_pages'], 'ranges': ranges, 'coverage': execution['coverage']})
    outputs = execution.get('outputs')
    need(isinstance(outputs, list) and len(outputs) == len(entries)
         and type(execution.get('written_count')) is int and execution['written_count'] == len(outputs)
         and type(result.get('written_count')) is int and result['written_count'] == len(outputs),
         'Execution count differs from complete shared-plan outputs')
    fields = ('sequence', 'title', 'start', 'end', 'parent_id', 'reason', 'filename', 'warnings')
    for index, (entry, output, bounds) in enumerate(zip(entries, outputs, ranges), 1):
        need(isinstance(entry, dict) and isinstance(output, dict)
             and all(field in entry and field in output and output[field] == entry[field] for field in fields)
             and type(entry.get('sequence')) is int and entry['sequence'] == index
             and all(type(item.get(field)) is int for item in (entry, output) for field in ('sequence', 'start', 'end'))
             and [entry['start'], entry['end']] == bounds
             and type(output.get('page_count')) is int and output['page_count'] == bounds[1] - bounds[0]
             and type(output.get('size_bytes')) is int and output['size_bytes'] > 0
             and isinstance(output.get('sha256'), str) and re.fullmatch('[0-9a-f]{64}', output['sha256']),
             'Published output identity/range differs from its shared-plan entry')
    manifest = execution.get('manifest')
    mirror = ('schema_version', 'status', 'run_id', 'final_directory', 'mode', 'total_pages', 'source_identity',
              'coverage', 'written_count', 'outputs')
    need(isinstance(manifest, dict) and all(field in manifest and manifest[field] == execution[field] for field in mirror)
         and execution.get('manifest_filename') == 'WinBookSplit_Manifest.json'
         and execution.get('original_ebook_identity') == plan.get('original_ebook_identity')
         and manifest.get('original_ebook_identity') == execution.get('original_ebook_identity'),
         'Published manifest/source differs from exact execution')
    members = ['.WinBookSplit-owner.json', 'WinBookSplit_Manifest.json'] + [entry['filename'] for entry in entries]
    need(isinstance(publication.get('members'), list) and len(publication['members']) == len(members)
         and len(set(members)) == len(members) and set(publication['members']) == set(members)
         and isinstance(publication.get('content_sha256'), list) and len(publication['content_sha256']) == plan['total_pages']
         and all(isinstance(value, str) and re.fullmatch('[0-9a-f]{64}', value) for value in publication['content_sha256']),
         'Actual published members/physical-page identities differ from complete plan')


def validate_native(row, prefix, final):
    need(row.get('actual_process') is True and row.get('passed') is True and row.get('timed_out') is False
         and type(row.get('pid')) is int and row['pid'] > 0 and type(row.get('exit_code')) is int
         and row['exit_code'] == final['exit_code'] and isinstance(row.get('command'), list) and bool(row['command'])
         and row['command'][0] == row.get('shell_executable') and type(row.get('elapsed_seconds')) in (int, float)
         and 0 <= row['elapsed_seconds'] < row.get('timeout_seconds', 0), 'Actual bounded native process/exit evidence missing')
    for name in ('stdout', 'stderr'):
        need(isinstance(row.get(name), str) and row.get(name + '_sha256') == digest(row[name]), 'Raw native stream/hash differs')
    need(records(row['stdout'], prefix) == [final], 'Final native frame missing/duplicate/conflicting')


def validate_run_manifest(manifest, row):
    final = row['outcome']
    need(isinstance(manifest, dict) and manifest.get('protocol') == 'winbooksplit.run'
         and type(manifest.get('version')) is int and manifest['version'] == 1
         and manifest.get('application_version') == '1.0.0'
         and isinstance(manifest.get('run_id'), str) and re.fullmatch('[0-9a-f]{32}', manifest['run_id']), 'Run manifest protocol/version/run identity invalid')
    need(set(manifest) == {'protocol', 'version', 'run_id', 'application_version', 'started_utc', 'finished_utc', 'diagnostics_finalized',
         'runtime_versions', 'settings', 'source_identity', 'plan', 'engine_result', 'outcome', 'warnings', 'log'}, 'Run manifest schema differs')
    need(manifest.get('outcome') == final and manifest.get('engine_result') == final.get('engine_result'),
         'Run manifest final outcome or engine result differs from actual operation')
    need(manifest.get('diagnostics_finalized') is True, 'Pending manifest was labelled final')
    need(all(isinstance(manifest[key], str) for key in ('started_utc', 'finished_utc')), 'Missing UTC timestamp text')
    try:
        started, finished = [datetime.fromisoformat(manifest[key].replace('Z', '+00:00')) for key in ('started_utc', 'finished_utc')]
    except ValueError as error:
        raise ValueError('Malformed UTC timestamp') from error
    need(started.utcoffset() is not None and finished.utcoffset() is not None
         and started.utcoffset().total_seconds() == finished.utcoffset().total_seconds() == 0 and started <= finished,
         'Run timestamps are not ordered UTC instants')
    settings = manifest.get('settings')
    need(isinstance(settings, dict) and settings.get('mode') == final.get('mode')
         and settings.get('input_kind') == row['input_kind'] and settings.get('preview') is False
         and settings.get('non_interactive') is (not row['kind'].endswith('cancelled'))
         and settings.get('no_pause') is row['kind'].endswith('cancelled')
         and settings.get('keep_converted_pdf') is False and type(settings.get('conversion_timeout')) is int
         and type(settings.get('process_timeout')) is int and settings['conversion_timeout'] == 1800
         and settings['process_timeout'] == 3600, 'Normalized run settings differ from literal native arguments')
    runtime = manifest.get('runtime_versions')
    need(runtime == {'powershell': row['host_version'], 'python': '3.14.8', 'pypdf': '6.20.0',
                     'calibre': '9.15.0' if row['input_kind'] != 'pdf' else None}, 'Run manifest versions differ from resolved runtimes')
    original = row['source_observations_before']['source']
    source = manifest.get('source_identity')
    processed = row.get('expected_plan') is not None
    need(isinstance(source, dict) and source.get('path') == row['input_path'] and source.get('size_bytes') == original['size_bytes']
         and (source.get('sha256') == original['sha256'] if processed else source.get('binding') == 'metadata_only' and source.get('sha256') is None),
         'Run manifest original source is not the immutable actual input')
    plan = manifest.get('plan')
    expected_plan = row.get('expected_plan')
    need(plan == expected_plan, 'Run manifest ranges/source differ from the held or executed shared plan')
    need(settings.get('normalized_inputs') == (None if plan is None else plan.get('normalized_inputs')),
         'Normalized inputs differ from the held shared plan')
    if plan is not None:
        partition(plan)
        captured = plan.get('original_ebook_identity', plan.get('source_identity'))
        need(isinstance(captured, dict) and captured.get('path') == row['input_path'] and captured.get('sha256') == original['sha256']
             and captured.get('size_bytes') == original['size_bytes'], 'Shared plan source is not the actual immutable original')
        engine = final.get('engine_result')
        need(isinstance(engine, dict) and engine.get('plan') == plan, 'Run shared plan differs from terminal engine plan')
    warnings = manifest.get('warnings')
    need(isinstance(warnings, dict) and warnings.get('planning') == row.get('planning_warnings', [])
         and warnings.get('parser') == row.get('parser_warnings', [])
         and type(warnings.get('parser_suppressed_count')) is int and warnings['parser_suppressed_count'] == row.get('parser_suppressed_count', 0)
         and type(warnings.get('parser_message_truncated_count')) is int
         and warnings['parser_message_truncated_count'] == row.get('parser_message_truncated_count', 0), 'Run warning categories/counts differ from captured engine diagnostics')
    need(len(warnings['parser']) <= 64, 'Parser warning bound exceeded')
    for warning in warnings['parser']:
        need(warning.get('code') in {'pypdf_parser_warning', 'pypdf_parser_error'} and warning.get('category') == 'pdf_parser'
             and warning.get('severity') in {'WARNING', 'ERROR'} and isinstance(warning.get('message'), str)
             and 0 < len(warning['message'].encode('utf-8')) <= 2048
             and type(warning.get('truncated')) is bool
             and warning['severity'] == ('WARNING' if warning['code'] == 'pypdf_parser_warning' else 'ERROR'),
             'Parser warnings are not bounded typed categorized messages')
    log = manifest.get('log')
    need(isinstance(log, dict) and log.get('filename') == 'console.log' and log.get('max_bytes') == 50331648
         and log.get('body_limit_bytes') == 33554432 and type(log.get('bytes_written')) is int
         and log['bytes_written'] == len(row['console_log'].encode('utf-8')) <= log['max_bytes']
         and log.get('limit_reached') is False, 'Run manifest log identity/byte bounds differ from finalized actual bytes')


def validate_application_case(row):
    kind, final = row.get('kind'), row.get('outcome')
    need(kind in APPLICATION_KINDS, 'Unexpected support application case')
    typed_outcome(final)
    need(final.get('protocol') == 'winbooksplit.outcome' and type(final.get('version')) is int and final['version'] == 1,
         'Application outcome protocol/version invalid')
    validate_native(row, '[OUTCOME] ', final)
    before = row.get('source_observations_before')
    need(isinstance(before, dict) and set(before) == {'source', 'neighbor', 'prior'}
         and before == row.get('source_observations_after') == row.get('source_observations_after_cleanup')
         and row.get('source_read_only_observed') is True and row.get('application_unchanged') is True
         and row.get('output_members_after_cleanup') == ['prior-output.pdf'], 'Immutable source/neighbor/prior/application or authenticated cleanup differs')
    need(all(isinstance(item, dict) and isinstance(item.get('sha256'), str) and re.fullmatch('[0-9a-f]{64}', item['sha256'])
             and type(item.get('size_bytes')) is int and item['size_bytes'] > 0 for item in before.values())
         and type(before['source'].get('attributes')) is int and before['source']['attributes'] & 1,
         'Actual file identities or source read-only attributes missing')
    parameters = row.get('parameters')
    need(isinstance(parameters, list) and all(isinstance(value, str) for value in parameters)
         and '-InputFile' in parameters and '-OutputDirectory' in parameters
         and parameters[parameters.index('-InputFile') + 1] == row['input_path']
         and parameters[parameters.index('-OutputDirectory') + 1] == row['output_base']
         and row['command'][-len(parameters):] == parameters
         and ('-NonInteractive' in parameters) is (not kind.endswith('cancelled'))
         and ('-NoPause' in parameters) is kind.endswith('cancelled'), 'Literal native arguments/settings differ from observed control')
    app = row.get('application_sha256')
    need(isinstance(app, dict) and set(app) == set(APPLICATION) and all(re.fullmatch('[0-9a-f]{64}', value) for value in app.values()),
         'Copied actual app byte hashes missing')
    actual = row.get('actual_application_sha256')
    need(isinstance(actual, dict) and set(actual) == set(app), 'Actual copied application hashes missing')
    faults = kind.startswith(('log-fault-', 'manifest-fault-'))
    need(row.get('fault_injection') is not None if faults else row.get('fault_injection') is None, 'Copied-helper fault scope was not disclosed')
    if faults:
        expected_path = 'WinBookSplit.ps1' if kind.startswith('log-fault-') else 'engine/WinBookSplit.Logging.ps1'
        need(row['fault_injection'].get('path') == expected_path
             and row['fault_injection'].get('sentinel') in row['stdout'] + row['stderr'] + row.get('console_log', '')
             and row['fault_injection'].get('original_sha256') == app[expected_path], 'Fault was not confined to the disclosed copied finalizer')
        need(actual[expected_path] == row['fault_injection'].get('modified_sha256') and actual[expected_path] != app[expected_path]
             and all(actual[name] == app[name] for name in app if name != expected_path), 'Copied fault changed other application paths')
    else:
        need(actual == app, 'Unmodified native control changed application bytes')
    if kind in {'preview', 'preflight'}:
        need(row.get('run_manifest') is None and row.get('console_log') is None and row.get('console_members') == []
             and row.get('publication') is None and 'Log: ' not in row['stdout'], 'Preview/preflight support records unexpectedly wrote files')
        need(final['exit_code'] == (0 if kind == 'preview' else 2)
             and final['status'] == ('preview' if kind == 'preview' else 'failed'), 'Preview/preflight outcome differs')
        return
    text = row.get('console_log')
    need(isinstance(text, str) and row.get('console_log_sha256') == digest(text), 'Actual finalized UTF-8 console log/hash missing')
    need(len(text.encode('utf-8')) <= 50331648, 'Console log exceeded its bound')
    footers = records(text, '[OPERATION-OUTCOME] ')
    need(bool(footers) and footers[-1] == final, 'Final corrected log footer differs from authoritative outcome')
    manifest = row.get('run_manifest')
    expected_members = {'.WinBookSplit-console-owner.json', 'console.log'}
    if manifest is not None:
        expected_members.add('WinBookSplit_Run.json')
    need(set(row.get('console_members', [])) == expected_members and len(row['console_members']) == len(expected_members),
         'Console ownership/membership differs')
    if kind.startswith('manifest-fault-'):
        need(manifest is None and row.get('run_manifest_raw') is None and row.get('run_manifest_sha256') is None,
             'Manifest failure fabricated a persisted success manifest')
    else:
        raw = row.get('run_manifest_raw')
        need(isinstance(raw, str) and strict_json(raw) == manifest and row.get('run_manifest_sha256') == digest(raw)
             and len(raw.encode('utf-8')) <= 33554432, 'Actual UTF-8 run manifest/hash differs')
        validate_run_manifest(manifest, row)
    result = final.get('engine_result')
    need(isinstance(result, dict) and result.get('protocol') == 'winbooksplit.result' and type(result.get('version')) is int and result['version'] == 1
         and row.get('engine_records') == [result], 'Actual terminal engine record does not bind to final outcome')
    processes = row.get('process_summaries')
    need(isinstance(processes, list) and len(processes) == 1, 'Actual process supervision record missing')
    process = processes[0]
    need(all(process.get(key) is True for key in ('JobAssigned', 'ParentStopped', 'DescendantsStopped', 'StreamsComplete'))
         and not any(process.get(key) for key in ('StartError', 'StopError', 'StreamError', 'ResultError', 'InteractionError', 'InputError'))
         and process.get('ExitCode') == result['exit_code'], 'Actual engine job/tree/EOF/writer proof incomplete')
    if kind.endswith('cancelled'):
        need(process.get('InputWriterStopped') is True, 'Interactive cancellation writer was not stopped')
    need(row.get('planning_warnings', []) == result.get('warnings', []), 'Planning warnings differ from actual terminal engine')
    parser = (result.get('diagnostic') or {}).get('parser_warnings') or {}
    need(row.get('parser_warnings', []) == parser.get('records', []) and row.get('parser_suppressed_count', 0) == parser.get('suppressed_count', 0)
         and row.get('parser_message_truncated_count', 0) == parser.get('message_truncated_count', 0), 'Parser warnings differ from terminal diagnostic')
    need(records(text, '[PROCESS] ') == processes and [strict_json(line) for line in text.splitlines() if line.startswith('{')] == [result],
         'Raw console stream/transport evidence differs')
    successful = kind.startswith('success-') or kind in {'parser-warning', 'bookmark-warning'}
    retained = kind.endswith('-success')
    if successful or retained:
        need(final['status'] == ('incomplete' if retained else 'success') and final['code'] == ('console_finalize_failed' if retained else 'split_complete')
             and final['exit_code'] == (6 if retained else 0), 'Published operation falsely green or incomplete result differs')
        result, publication = final.get('engine_result'), row.get('publication')
        need(isinstance(result, dict) and result.get('protocol') == 'winbooksplit.result' and type(result.get('version')) is int
             and result['version'] == 1 and result.get('status') == 'success' and result.get('code') == 'split_complete'
             and result.get('exit_code') == 0 and result.get('mode') == final.get('mode')
             and isinstance(publication, dict), 'Validated completed execution was not retained')
        execution = result.get('execution')
        need(isinstance(execution, dict) and type(execution.get('written_count')) is int and execution['written_count'] > 0
             and execution['written_count'] == result.get('written_count')
             and final['written_count'] == (0 if retained else execution['written_count'])
             and execution.get('final_directory') == final.get('final_directory') and Path(final['final_directory']).parent == Path(row['output_base'])
             and publication.get('manifest') == execution.get('manifest'), 'Final folder/count/manifest differs from validated execution')
        validate_execution(row, result, publication)
        need(publication.get('content_sha256') == row.get('expected_content_sha256') and bool(publication['content_sha256'])
             and strict_json(publication['manifest_raw']) == publication['manifest']
             and publication['manifest_sha256'] == digest(publication['manifest_raw']), 'Publication omitted/reordered physical pages or manifest bytes differ')
        need(('Done.' in row['stdout']) is successful, 'Human Done falsely announced an incomplete finalization')
        need((('Completed engine output retained: ' if retained else 'Output: ') + final['final_directory']) in row['stdout'],
             'Completed/retained actual folder was not named')
    else:
        need(row.get('publication') is None and final['written_count'] == 0 and final.get('final_directory') is None
             and 'Done.' not in row['stdout'], 'Failed/cancelled support case published or falsely succeeded')
        if kind.endswith('cancelled'):
            need(final['status'] == 'cancelled' and final['code'] == 'processing_cancelled' and final['exit_code'] == 130
                 and row.get('stdin_utf8') == 'C\n', 'Explicit held-plan cancellation differs')
        elif kind == 'no-plan':
            need(final['status'] == 'no_plan' and final['code'] == 'no_bookmarks' and final['exit_code'] == 5, 'No-plan primary outcome differs')
        else:
            need(final['exit_code'] == 6 and final['code'] in {'unreadable_document', 'invalid_document'} and final['status'] in {'failed', 'error', 'read_error'},
                 'Primary PDF failure/code was lost during support finalization')
    if kind == 'parser-warning':
        need(row.get('parser_warnings') and '[PYPDF WARNING]' in row['stdout'] + row['stderr'] and '[PYPDF WARNING]' in text,
             'Real pypdf warning was suppressed in visible or persisted diagnostics')
    if kind == 'bookmark-warning':
        need(any(warning.get('code') == 'invalid_destination' for warning in row.get('planning_warnings', []))
             and 'invalid_destination' in text and 'invalid_destination' in row['stdout'] + row['stderr'],
             'Real invalid bookmark warning was suppressed')


def validate_summary(summary, manifest):
    need(isinstance(manifest, dict) and manifest.get('protocol') == 'winbooksplit.run'
         and type(manifest.get('version')) is int and manifest['version'] == 1
         and manifest.get('diagnostics_finalized') is True, 'Export source is not a finalized run manifest')
    need(isinstance(summary, dict) and set(summary) == {'protocol', 'version', 'application_version', 'runtime_versions',
         'settings', 'plan', 'outcome', 'warnings', 'log'} and summary['protocol'] == 'winbooksplit.support'
         and type(summary['version']) is int and summary['version'] == 1 and summary['application_version'] == '1.0.0',
         'Support summary root whitelist/protocol invalid')
    versions = {**manifest['runtime_versions'], 'powershell': '5.1' if manifest['runtime_versions']['powershell'].startswith('5.1.') else '7'}
    need(summary['runtime_versions'] == versions and set(summary['runtime_versions']) == {'powershell', 'python', 'pypdf', 'calibre'},
         'Support runtime whitelist differs')
    need(versions['powershell'] in {'5.1', '7'} and versions['python'] == '3.14.8' and versions['pypdf'] == '6.20.0'
         and versions['calibre'] in {'9.15.0', None}
         and re.fullmatch(r'(?:5\.1|7)\.\d+(?:\.\d+){0,2}', manifest['runtime_versions']['powershell']),
         'Support runtime includes an unsupported or private token')
    settings = summary['settings']
    need(isinstance(settings, dict) and set(settings) == set(SCALAR_SETTINGS)
         and settings == {key: manifest['settings'][key] for key in SCALAR_SETTINGS}, 'Support settings copied identifiers or altered typed fields')
    need(settings['mode'] in ('', 'manual', '1', '2') and settings['input_kind'] in ('pdf', 'epub', 'azw3')
         and all(type(settings[key]) is bool for key in ('preview', 'non_interactive', 'no_pause', 'keep_converted_pdf'))
         and type(settings['conversion_timeout']) is int and 1 <= settings['conversion_timeout'] <= 86400
         and type(settings['process_timeout']) is int and 1 <= settings['process_timeout'] <= 172800,
         'Support settings use untyped or out-of-bound values')
    need(summary['outcome'] == {key: manifest['outcome'][key] for key in ('status', 'code', 'exit_code', 'written_count')},
         'Support outcome copied free text or differs')
    typed_outcome(summary['outcome'], manifest.get('engine_result'))
    expected_plan = None if manifest['plan'] is None else {key: manifest['plan'][key] for key in ('total_pages', 'ranges', 'coverage')}
    need(summary['plan'] == expected_plan, 'Support plan copied titles/identity or changed coverage')
    if expected_plan is not None:
        partition(summary['plan'])
    final = summary['outcome']
    status, code, native, written = [final[key] for key in ('status', 'code', 'exit_code', 'written_count')]
    need((status != 'success' or (code == 'split_complete' and written > 0 and expected_plan is not None
          and written == expected_plan['coverage']['section_count'] and not settings['preview']))
         and (status != 'preview' or (code == 'preview_complete' and written == 0 and settings['preview']))
         and (status != 'no_plan' or (native == 5 and code in {'no_bookmarks', 'no_usable_bookmarks', 'no_bookmarks_at_level'}))
         and (status != 'invalid_input' or code in {'input_invalid', 'invalid_document', 'invalid_outline', 'invalid_start_pages', 'invalid_mode', 'invalid_arguments'})
         and (status != 'read_error' or code == 'unreadable_document')
         and (status != 'write_error' or code == 'output_write_failed')
         and (status != 'unsupported' or code == 'unsupported_document')
         and (status != 'incomplete' or code in {'console_finalize_failed', 'output_handle_close_failed'})
         and (status not in {'failed', 'incomplete', 'invalid_input', 'read_error', 'write_error', 'error', 'unsupported'} or native not in {0, 130})
         and (status not in {'cancelled', 'timeout'} or (native == 130 and written == 0))
         and (status == 'success' or written == 0), 'Support summary status/code/count contradict the selected operation')
    warnings = summary['warnings']
    need(isinstance(warnings, dict) and set(warnings) == {'planning', 'parser', 'parser_suppressed_count', 'parser_message_truncated_count'},
         'Support warning whitelist invalid')
    need(all(isinstance(warnings[key], list) for key in ('planning', 'parser'))
         and all(type(warnings[key]) is int and 0 <= warnings[key] <= 2147483647
                 for key in ('parser_suppressed_count', 'parser_message_truncated_count')), 'Support warning arrays/counters are untyped')
    planning_categories = {
        **dict.fromkeys(('invalid_destination', 'external_destination', 'destination_error', 'duplicate_destination', 'outline_reordered',
                         'duplicate_parent_subtree', 'invalid_parent_subtree', 'child_outside_parent', 'malformed_outline', 'outline_cycle',
                         'outline_limit'), 'bookmark'),
        'output_handle_close_failed': 'output',
        'cross_chapter_link_dropped': 'navigation', 'navigation_link_dropped': 'navigation', 'article_navigation_dropped': 'navigation',
        'annotation_dropped': 'annotation', 'annotation_relation_dropped': 'annotation',
        'metadata_omitted': 'metadata', 'metadata_normalized': 'metadata',
    }
    need(isinstance(manifest['warnings']['planning'], list)
         and all(isinstance(item, dict) and isinstance(item.get('code'), str) and item['code'] in planning_categories
                 for item in manifest['warnings']['planning']), 'Support warning projection includes an unknown selected token')
    need(warnings['planning'] == [{'code': item['code'], 'category': planning_categories[item['code']]} for item in manifest['warnings']['planning']]
         and warnings['parser'] == [{key: item[key] for key in ('code', 'category', 'severity')} for item in manifest['warnings']['parser']]
         and all(warnings[key] == manifest['warnings'][key] for key in ('parser_suppressed_count', 'parser_message_truncated_count')),
         'Support warnings copied document details or altered categories/counts')
    need(len(warnings['planning']) <= 4096 and len(warnings['parser']) <= 64
         and all(item.get('code') in planning_categories for item in warnings['planning'])
         and all(item.get('category') == 'pdf_parser' and item.get('code') in {'pypdf_parser_warning', 'pypdf_parser_error'}
                 and item.get('severity') == ('WARNING' if item['code'] == 'pypdf_parser_warning' else 'ERROR') for item in warnings['parser']),
         'Support warning projection includes an unknown or contradictory selected token')
    need(summary['log'] == {key: manifest['log'][key] for key in ('max_bytes', 'body_limit_bytes', 'bytes_written', 'limit_reached')},
         'Support log copied paths or altered bounds')
    log = summary['log']
    need(all(type(log[key]) is int for key in ('max_bytes', 'body_limit_bytes', 'bytes_written')) and type(log['limit_reached']) is bool
         and 1 <= log['body_limit_bytes'] <= 33554432 and log['body_limit_bytes'] <= log['max_bytes'] <= 50331648
         and 0 <= log['bytes_written'] <= log['max_bytes'], 'Support log bounds/counts/flag are untyped')


def validate_export_case(row):
    kind, final = row.get('kind'), row.get('export_outcome')
    need(kind in EXPORT_KINDS and isinstance(final, dict) and final.get('protocol') == 'winbooksplit.support-export'
         and type(final.get('version')) is int and final['version'] == 1, 'Support export frame invalid')
    good = kind in {'success', 'failed', 'cancelled', 'incomplete', 'no-plan', 'redacted'}
    need(final.get('status') == ('success' if good else 'failed')
         and final.get('code') == ('support_export_complete' if good else 'support_export_failed')
         and type(final.get('exit_code')) is int and final['exit_code'] == (0 if good else 2)
         and type(final.get('written_count')) is int and final['written_count'] == (1 if good else 0), 'Export status/code/native-count contradiction')
    validate_native(row, '[SUPPORT-EXPORT] ', final)
    need(row.get('source_before') == row.get('source_after') and row.get('prior_before') == row.get('prior_after')
         and row.get('source_before') is not None and row.get('application_unchanged') is True
         and row.get('owned_outputs_removed') is True, 'Export altered its source/prior/app or cleanup unproved')
    if good:
        raw = row.get('summary_raw')
        need(isinstance(raw, str) and strict_json(raw) == row.get('summary') and row.get('summary_sha256') == digest(raw),
             'Actual redacted UTF-8 bytes/hash differ')
        validate_summary(row['summary'], row['input_manifest'])
        need(strict_json(row['input_raw']) == row['input_manifest'] and row['source_before']['sha256'] == digest(row['input_raw']),
             'Actual exporter source bytes differ from selected manifest')
        need(all(value not in raw + row['stdout'] + row['stderr'] for value in row.get('private_markers', [])),
             'Mock private identifiers/content/secrets escaped redaction into summary or console')
        if kind == 'redacted':
            need(len(row.get('private_markers', [])) >= 5 and all(value in row.get('input_raw', '') for value in row['private_markers']),
                 'Malicious authored fixture did not contain the claimed private fields')
    else:
        need(row.get('summary') is None and row.get('summary_raw') is None and row.get('summary_sha256') is None
             and row.get('destination_created') is False, 'Rejected export wrote or overwrote a summary')
        need(row.get('destination_before') == row.get('destination_after'), 'Rejected export altered an existing or aliased destination')


def validate_support_report(report, shells, calibre, *, cleanup_complete=True):
    need(report.get('schema_version') == 1 and type(report['schema_version']) is int and report.get('task_id') == 'M3-T05'
         and report.get('result') == 'SUPPORT_REGRESSION_PASSED' and report.get('success') is True
         and type(report.get('exit_code')) is int and report['exit_code'] == 0
         and report.get('acceptance_ids') == ['AC-063', 'AC-064', 'AC-065']
         and report.get('source_unchanged') is True and (not cleanup_complete or report.get('owned_temp_removed') is True)
         and report.get('human_or_viewer_opening_tested') is False, 'Support aggregate/source/cleanup evidence incomplete')
    hosts = report.get('host_cases')
    need(isinstance(hosts, list) and len(hosts) == 2 and {host.get('id') for host in hosts} == {'PS51', 'PS7'}
         and len(shells) == 2 and {str(Path(host['shell_executable']).resolve()).casefold() for host in hosts}
         == {str(Path(shell).resolve()).casefold() for shell in shells}, 'Actual supported distinct requested hosts missing')
    for host in hosts:
        need(host.get('passed') is True and host.get('exit_code') == 0 and host.get('host_major') == (5 if host['id'] == 'PS51' else 7)
             and host.get('host_version') == ('5.1.26100.9444' if host['id'] == 'PS51' else '7.6.6')
             and host.get('stored_policies') == host.get('policies_after'), 'Native host version/syntax/settings evidence differs')
    tested = report.get('tested_path_sha256')
    need(isinstance(tested, dict) and bool(tested) and all(name in tested for name in APPLICATION)
         and all(re.fullmatch('[0-9a-f]{64}', value) for value in tested.values()), 'Raw tested source hash map incomplete')
    apps, exports = report.get('cases'), report.get('export_cases')
    need(isinstance(apps, list) and len(apps) == 2 * len(APPLICATION_KINDS)
         and {row.get('id') for row in apps} == {host + '-' + kind for host in ('PS51', 'PS7') for kind in APPLICATION_KINDS}, 'Application control matrix missing/duplicated')
    need(isinstance(exports, list) and len(exports) == 2 * len(EXPORT_KINDS)
         and {row.get('id') for row in exports} == {host + '-export-' + kind for host in ('PS51', 'PS7') for kind in EXPORT_KINDS}
         and report.get('case_count') == len(apps) + len(exports), 'Export control matrix/count missing/duplicated')
    for row in apps + exports:
        host = next(host for host in hosts if host['id'] == row.get('host_id'))
        need(row.get('shell_executable') == host['shell_executable'] and row.get('host_version') == host['host_version'], 'Control host disagrees with native host observation')
        if row in apps:
            need(row.get('application_sha256') == {name: tested[name] for name in APPLICATION}, 'Copied native application differs from tested source map')
            validate_application_case(row)
        else:
            need(row.get('exporter_sha256') == tested['Export-WinBookSplitDiagnostics.ps1'], 'Actual exporter differs from tested source bytes')
            validate_export_case(row)
    refs = report.get('real_conversion_references')
    need(isinstance(refs, dict) and set(refs) == {'epub', 'azw3'} and report.get('calibre_path') == calibre,
         'Real format/reference converter evidence missing')
    provenance = report.get('fixture_provenance')
    need(isinstance(provenance, dict) and provenance.get('authored_original') is True and provenance.get('remote_resources') is False
         and provenance.get('calibre', {}).get('path') == calibre and provenance['calibre'].get('version') == '9.15.0'
         and provenance.get('azw3_generation', {}).get('exit_code') == 0, 'Authored offline EPUB/genuine AZW3 provenance incomplete')
    for fmt, reference in refs.items():
        need(reference.get('process', {}).get('exit_code') == 0 and reference['process'].get('argv', [None])[0] == calibre
             and reference.get('page_count') == (3 if fmt == 'epub' else 4) and reference.get('size_bytes', 0) > 0
             and isinstance(reference.get('sha256'), str) and re.fullmatch('[0-9a-f]{64}', reference['sha256'])
             and len(reference.get('page_content_sha256', [])) == reference['page_count'], 'Independent actual conversion content evidence differs')
        for row in apps:
            if row['input_kind'] == fmt:
                need(row.get('expected_content_sha256') == reference['page_content_sha256'], 'Native ebook publication differs from independent actual physical PDF')
