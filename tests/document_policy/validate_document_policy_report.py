"""Fail-closed pure guard for native unsupported-document policy receipts."""
from base64 import b64decode
from hashlib import sha256
import importlib.util
import json
import math
from pathlib import Path
import re
import sys

ROOT = Path(__file__).resolve().parents[2]


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


fixtures = load('wbs_policy_receipt_fixtures', ROOT / 'tests/document_policy/generate_document_policy_fixtures.py')
console = load('wbs_policy_receipt_console', ROOT / 'tests/manual/current_launchers.py')
support = load('wbs_policy_receipt_support', ROOT / 'tests/support/validate_support_report.py')
APPLICATION = support.APPLICATION
PS_KINDS = ('ordinary-manual', 'ordinary-level1', 'ordinary-level2', 'page-aa-exclusion',
            *fixtures.UNSUPPORTED, *fixtures.MALFORMED,
            'encrypted-user-preview', 'encrypted-owner-empty-preview', 'xfa-interactive')
BAT_KINDS = ('ordinary-manual', 'encrypted-user', 'encrypted-owner-empty', 'acroform', 'signature-perms',
             'orphan-widget', 'orphan-signature', 'catalog-openaction', 'embedded-attachment',
             'page-associated-file', 'page-3d', 'malformed-header', 'zero-pages')
RANGES = {'manual': [[0, 2], [2, 4]], '1': [[0, 2], [2, 4]], '2': [[0, 1], [1, 2], [2, 3], [3, 4]]}
MODIFICATIONS = ['copied parameter default OutputDirectory', 'copied parameter default PythonPath', 'copied NoPause default true']
WRAPPER = '@echo off\r\nsetlocal DisableDelayedExpansion\r\n"%WBS_POLICY_BAT%" "%WBS_POLICY_INPUT%"\r\n'
HEX = re.compile(r'[0-9a-f]{64}\Z')


def need(condition, message):
    if not condition:
        raise ValueError(message)


def fixture_kind(kind):
    if kind.startswith('ordinary-'):
        return 'ordinary'
    return kind.removesuffix('-preview').removesuffix('-interactive')


def mode_for(kind):
    return '1' if kind in {'ordinary-level1', 'deep-outline'} else '2' if kind == 'ordinary-level2' else 'manual'


def expected_exit(kind):
    fixture = fixture_kind(kind)
    return 7 if fixture in fixtures.UNSUPPORTED else 6 if fixture in fixtures.MALFORMED else 0


def controlled_powershell(raw):
    text = raw.decode('utf-8-sig')
    for name in ('OutputDirectory', 'PythonPath'):
        before = '[string]$' + name + ','
        need(text.count(before) == 1, 'Declared copied parameter-default seam changed')
        text = text.replace(before, '[string]$' + name + ' = $env:WBS_POLICY_' + name.upper() + ',')
    need(text.count('[switch]$NoPause,') == 1, 'Declared copied NoPause seam changed')
    return text.replace('[switch]$NoPause,', '[switch]$NoPause = $true,').encode('utf-8-sig')


def records(text, prefix):
    return [support.strict_json(line[len(prefix):]) for line in text.splitlines() if line.startswith(prefix)]


def validate_native(row, expected, command):
    need(row.get('actual_process') is True and type(row.get('exit_code')) is int and row['exit_code'] == expected
         and type(row.get('pid')) is int and row['pid'] > 0 and row.get('timed_out') is False
         and type(row.get('elapsed_seconds')) in (int, float) and math.isfinite(row['elapsed_seconds'])
         and type(row.get('timeout_seconds')) is int and 0 <= row['elapsed_seconds'] < row['timeout_seconds'] <= 90,
         'Finite actual native status/PID/deadline evidence missing')
    need(row.get('command') == command and all(isinstance(value, str) for value in command), 'Exact native policy argv differs')
    for stream in ('stdout', 'stderr'):
        need(isinstance(row.get(stream), str) and sha256(row[stream].encode()).hexdigest() == row.get(stream + '_sha256'),
             'Native stream raw hash differs')


def validate_result(frame, kind):
    expected = expected_exit(kind)
    need(isinstance(frame, dict) and {'protocol', 'version', 'mode', 'status', 'code', 'exit_code',
         'written_count', 'fallback_modes', 'execution', 'warnings', 'diagnostic'} <= set(frame)
         and frame.get('protocol') == 'winbooksplit.result'
         and type(frame.get('version')) is int and frame['version'] == 1 and frame.get('mode') == mode_for(kind)
         and type(frame.get('exit_code')) is int and frame['exit_code'] == expected
         and type(frame.get('written_count')) is int and isinstance(frame.get('warnings'), list)
         and frame.get('fallback_modes') == [],
         'Sole typed terminal policy result absent or contradictory')
    if expected == 7:
        need(frame.get('status') == 'unsupported' and frame.get('code') == 'unsupported_document'
             and frame['written_count'] == 0 and frame.get('execution') is None and frame.get('plan') is None,
             'Unsupported input falsely prepared, published, retried or succeeded')
    elif expected == 6:
        if kind == 'zero-pages':
            allowed = {('invalid_input', 'invalid_document')}
        elif kind == 'deep-outline':
            allowed = {('invalid_input', 'invalid_outline')}
        else:
            allowed = {('read_error', 'unreadable_document'), ('invalid_input', 'invalid_document')}
        need((frame.get('status'), frame.get('code')) in allowed and frame['written_count'] == 0
             and frame.get('execution') is None and frame.get('plan') is None,
             'Malformed/pathological document lost its bounded declared failure')
    else:
        need(frame.get('status') == 'success' and frame.get('code') == 'split_complete'
             and frame['written_count'] == len(RANGES[mode_for(kind)]) and isinstance(frame.get('execution'), dict)
             and isinstance(frame.get('plan'), dict), 'Ordinary input lost exact positive publication')
    diagnostic = frame.get('diagnostic')
    need(expected == 0 or frame['warnings'] == [], 'Preflight rejection falsely planned/warned chapter extraction')
    need(diagnostic is None or isinstance(diagnostic, dict) and not any(key in diagnostic for key in
         ('working_pdf', 'working_pdf_cleanup', 'conversion', 'retained_staging', 'owned_members', 'stage_identity')),
         'PDF preflight acquired conversion/staging/working-PDF ownership')


def validate_publication(row, frame):
    plan, execution = frame['plan'], frame['execution']
    bounds = RANGES[row['mode']]
    source = row['source_observations_before']['source']
    need(plan.get('mode') == row['mode'] and type(plan.get('total_pages')) is int and plan['total_pages'] == 4
         and plan.get('ranges') == bounds and [[entry['start'], entry['end']] for entry in plan.get('entries', [])] == bounds
         and plan.get('coverage') == {'complete': True, 'covered_pages': 4, 'section_count': len(bounds)}
         and plan['coverage']['complete'] is True and all(type(plan['coverage'][key]) is int for key in ('covered_pages', 'section_count')),
         'Ordinary plan omitted, duplicated or crossed physical pages')
    need(plan.get('source_identity', {}).get('sha256') == source['sha256']
         and plan['source_identity'].get('size_bytes') == source['size_bytes']
         and execution.get('source_identity') == plan['source_identity']
         and execution.get('mode') == plan['mode'] and execution.get('coverage') == plan['coverage']
         and execution.get('total_pages') == 4 and execution.get('written_count') == len(bounds)
         and execution.get('final_directory') == row['outcome']['final_directory']
         and Path(execution['final_directory']).parent == Path(row['output_base']), 'Publication changed captured source/plan/destination')
    outputs, observed = execution.get('outputs'), row.get('publication')
    need(isinstance(outputs, list) and isinstance(observed, list) and len(outputs) == len(observed) == len(bounds),
         'Reopened nonempty ordinary publication missing')
    expected_content = row.get('expected_content_sha256')
    need(isinstance(expected_content, list) and len(expected_content) == 4 and all(HEX.fullmatch(value) for value in expected_content),
         'Independent ordinary physical-content reference missing')
    for planned, output, opened, selected in zip(plan['entries'], outputs, observed, bounds):
        need(all(type(planned.get(key)) is int and type(output.get(key)) is int for key in ('start', 'end', 'sequence'))
             and all(output.get(key) == value for key, value in planned.items())
             and [output['start'], output['end']] == selected and type(output.get('page_count')) is int
             and output['page_count'] == selected[1] - selected[0] and HEX.fullmatch(output.get('sha256', ''))
             and type(output.get('size_bytes')) is int and output['size_bytes'] > 0
             and opened.get('filename') == output['filename'] and opened.get('sha256') == output['sha256']
             and opened.get('bytes') == output['size_bytes'] and opened.get('page_ids') == list(range(selected[0] + 1, selected[1] + 1))
             and opened.get('content_sha256') == expected_content[selected[0]:selected[1]]
             and opened.get('page_actions_absent') is True and opened.get('inert_action_marker_absent') is True,
             'Ordinary reopened output bytes/content/order or page-AA exclusion differs')
    need(row.get('publication_manifest') == execution.get('manifest') == support.strict_json(row['publication_manifest_raw'])
         and sha256(row['publication_manifest_raw'].encode()).hexdigest() == row['publication_manifest_sha256'],
         'Ordinary publication manifest raw/hash differs')
    need(set(row['publication_members']) == {'.WinBookSplit-owner.json', 'WinBookSplit_Manifest.json', *(output['filename'] for output in outputs)},
         'Ordinary publication contains unowned members')


def validate_case(row, *, cleanup_complete=True):
    host, kind = row.get('host_id'), row.get('kind')
    need(host in {'PS51', 'PS7', 'BAT'} and kind in (BAT_KINDS if host == 'BAT' else PS_KINDS)
         and row.get('id') == host + '-' + kind and row.get('fixture_kind') == fixture_kind(kind)
         and row.get('mode') == mode_for(kind) and row.get('source_read_only_observed') is True, 'Policy control ID/kind/mode/source mismatch')
    source_path, base, cwd = row['input_path'], row['output_base'], row['cwd']
    directory = Path(source_path).parent.parent
    need(Path(base) == directory / 'o' and Path(cwd) == directory / 'c', 'Policy native control escaped owned fixture layout')
    before = row.get('source_observations_before')
    need(isinstance(before, dict) and set(before) == {'source', 'neighbor', 'prior'} and before == row.get('source_observations_after')
         and before['source']['attributes'] & 1, 'Read-only source/neighbor/prior identity changed')
    for observation in before.values():
        need(isinstance(observation, dict) and HEX.fullmatch(observation.get('sha256', ''))
             and all(type(observation.get(key)) is int and observation[key] >= 0 for key in ('size_bytes', 'device', 'inode', 'attributes')),
             'Captured literal input identity untyped')
    if cleanup_complete:
        need(before == row.get('source_observations_after_cleanup') and row.get('output_members_after_cleanup') == ['prior-output.pdf']
             and row.get('passed') is True, 'Authenticated output cleanup or immutable input proof missing')
    preview, interactive = kind.endswith('-preview'), host == 'BAT' or kind.endswith('-interactive')
    if host == 'BAT':
        need(row.get('modifications') == MODIFICATIONS and row.get('wrapper_text') == WRAPPER
             and row.get('wrapper_sha256') == sha256(WRAPPER.encode('ascii')).hexdigest(), 'BAT declared source/wrapper seam differs')
        command = [row['cmd_executable'], '/d', '/v:off', '/c', str(directory / 'invoke.cmd')]
        need(row.get('stdin') == 'PIPE_UTF8' and row.get('stdin_utf8') == ('M\n1,3\nY\nN\n' if kind == 'ordinary-manual' else 'M\n'),
             'BAT literal menu/consent recipe changed')
        need(row.get('parameters') == [source_path], 'BAT changed the one literal input argument')
    else:
        parameters = ['-InputFile', source_path, '-OutputDirectory', base, '-PythonPath', row['python_executable'],
                      '-Mode', 'Manual' if row['mode'] == 'manual' else 'Auto']
        parameters += ['-StartPages', '1,3'] if row['mode'] == 'manual' else ['-BookmarkLevel', row['mode']]
        parameters += ['-NoPause'] if interactive else ['-NonInteractive']
        if preview:
            parameters += ['-Preview']
        command = [row['shell_executable'], '-NoProfile', *([] if interactive else ['-NonInteractive']),
                   '-ExecutionPolicy', 'RemoteSigned', '-File', str(directory / 'a/WinBookSplit.ps1'), *parameters]
        need(row.get('parameters') == parameters and row.get('modifications') == [] and row.get('stdin') == 'DEVNULL'
             and row.get('stdin_utf8') is None, 'Explicit native policy parameters/closed-input route differ')
    validate_native(row, expected_exit(kind), command)
    final = row.get('outcome')
    need(isinstance(final, dict) and final.get('protocol') == 'winbooksplit.outcome' and type(final.get('version')) is int
         and final['version'] == 1 and final.get('mode') == row['mode'] and type(final.get('exit_code')) is int
         and final['exit_code'] == row['exit_code'] and type(final.get('written_count')) is int
         and records(row['stdout'], '[OUTCOME] ') == [final], 'One authoritative native/final policy outcome missing')
    frame = final.get('engine_result')
    validate_result(frame, kind)
    need(all(final.get(key) == frame.get(key) for key in ('status', 'code', 'exit_code', 'written_count')),
         'PowerShell/BAT lost the primary document-policy result')
    combined = row['stdout'] + row['stderr'] + (row.get('console_log') or '')
    need(fixtures.USER_PASSWORD not in combined and fixtures.OWNER_PASSWORD not in combined
         and not re.search(r'(?:enter|provide|type)\s+(?:a\s+|the\s+)?password|(?:remove|bypass)\s+DRM', combined, re.I)
         and '[OPEN] ' not in combined, 'Policy processing solicited/leaked passwords or opened document data')
    if preview:
        need(row.get('console_evidence') is None and row.get('console_log') is None and row.get('log_sha256') is None
             and row.get('engine_records') == [] and row.get('process_summaries') == []
             and row['output_members_after'] == ['prior-output.pdf'], 'Rejected Preview created persisted diagnostics/output')
        need(not any(token in combined for token in ('[WBS-INTERACTION] ', '[INTERACTION-REPLY] ',
             'Write these chapter PDFs', 'Enter manual start pages', 'Open the completed folder')),
             'Rejected Preview solicited processing/consent')
    else:
        text = row.get('console_log')
        need(isinstance(text, str) and sha256(text.encode()).hexdigest() == row.get('log_sha256')
             and row.get('engine_records') == [frame], 'Actual finalized log/sole terminal result missing')
        console.validate_console_evidence(row.get('console_evidence'), final, row['log_sha256'],
            source_path=source_path, source_observation=before['source'], log_text=text)
        settings = row['console_evidence']['run_manifest']['settings']
        need(settings.get('input_kind') == 'pdf' and settings.get('preview') is False
             and settings.get('non_interactive') is (not interactive) and settings.get('no_pause') is interactive
             and settings.get('keep_converted_pdf') is False, 'Recorded policy settings differ from exact native route')
        transports = row.get('process_summaries')
        need(isinstance(transports, list) and len(transports) == 1 and all(transports[0].get(key) is True for key in
             ('JobAssigned', 'ParentStopped', 'DescendantsStopped', 'StreamsComplete'))
             and type(transports[0].get('ExitCode')) is int and transports[0]['ExitCode'] == expected_exit(kind)
             and type(transports[0].get('ResultRecordCount')) is int and transports[0]['ResultRecordCount'] == 1,
             'Actual policy engine/job/stream stop proof absent')
        interaction = console.interactions.capture(text)
        need(row.get('interaction') == interaction, 'Actual interaction capture differs from raw log')
        invocations = interaction['engine_invocations']
        need(isinstance(invocations, list) and len(invocations) == 1, 'One exact actual policy engine invocation required')
        expected_argv = ['-I', '-B', '-X', 'utf8', str(directory / 'a/engine/winbooksplit_engine.py'),
                         source_path, base, row['mode'], '' if host == 'BAT' or row['mode'] != 'manual' else '1,3']
        argv = invocations[0].get('arguments')
        need(invocations[0].get('path') == row['python_executable'] and isinstance(argv, list)
             and (argv == expected_argv if not interactive else len(argv) == 11 and argv[:9] == expected_argv
                  and argv[9] == '--interactive' and isinstance(argv[10], str) and re.fullmatch(r'[0-9a-f]{32}', argv[10])),
             'Recorded native policy engine argv/source/session differs')
        if interactive:
            stages = ['input_ready', 'plan_ready'] if expected_exit(kind) == 0 else []
            actions = ['starts', 'execute'] if expected_exit(kind) == 0 else []
            console.interactions.validate(interaction, transports, stages, actions, starts='1,3' if stages else None,
                                           execution=frame.get('execution'))
        else:
            console.interactions.require_noninteractive(interaction, transports)
            for field in ('InteractionCount', 'ReplyCount', 'QueuedReplyCount'):
                need(field not in transports[0] or type(transports[0][field]) is int and transports[0][field] == 0,
                     'NonInteractive policy control used interactive activity')
    if expected_exit(kind):
        need(final.get('final_directory') is None and row.get('publication') is None
             and 'Done.' not in row['stdout'] and 'Output: ' not in row['stdout'] and '[PLAN] ' not in combined
             and 'Write these chapter PDFs' not in combined and 'Open the completed folder' not in combined,
             'Rejected policy input acquired a plan/publication/folder prompt')
        if not preview:
            manifest = row['console_evidence']['run_manifest']
            need(manifest.get('plan') is None and manifest['source_identity'].get('binding') == 'metadata_only'
                 and manifest['source_identity'].get('sha256') is None, 'Rejected preflight falsely claimed a captured/extracted plan')
        need(set(row['output_members_after']) == {'prior-output.pdf'} | (set() if preview else {row['console_directory_name']}),
             'Rejected input created output staging/chapter/final directory')
    else:
        validate_publication(row, frame)
        need('Done.' in row['stdout'] and 'Output: ' + final['final_directory'] in row['stdout'], 'Ordinary truthful completion absent')
        if kind == 'page-aa-exclusion':
            need(any(item.get('code') == 'navigation_link_dropped' for item in frame.get('warnings', []))
                 and '[WARNING] navigation_link_dropped:' in row['stdout'] and 'navigation_link_dropped' in combined,
                 'Defined page-AA exclusion lost its visible fixed warning')


def validate_export(row, application):
    need(row.get('id') == application['host_id'] + '-unsupported-export' and row.get('application_case_id') == application['id'],
         'Explicit unsupported/null-plan export source binding differs')
    directory = Path(application['input_path']).parent.parent
    manifest = Path(application['output_base']) / application['console_directory_name'] / 'WinBookSplit_Run.json'
    command = [application['shell_executable'], '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'RemoteSigned', '-File',
        str(directory / 'a/Export-WinBookSplitDiagnostics.ps1'), '-ManifestPath', str(manifest),
        '-OutputPath', str(Path(application['cwd']) / 'redacted-summary.json')]
    validate_native(row, 0, command)
    need(row.get('cwd') == application['cwd'] and row.get('stdin') == 'DEVNULL', 'Explicit export cwd/stdin differs')
    raw = row.get('summary_raw')
    need(isinstance(raw, str) and sha256(raw.encode()).hexdigest() == row.get('summary_sha256')
         and support.strict_json(raw) == row.get('summary'), 'Actual support export raw/hash differs')
    support.validate_summary(row['summary'], application['console_evidence']['run_manifest'])
    expected = {'protocol': 'winbooksplit.support-export', 'version': 1, 'status': 'success',
                'code': 'support_export_complete', 'exit_code': 0, 'written_count': 1}
    actual = records(row['stdout'], '[SUPPORT-EXPORT] ')
    need(actual == [expected] and all(type(actual[0][key]) is int for key in ('version', 'exit_code', 'written_count')),
         'Actual explicit export result protocol/type differs')
    need(row['summary']['plan'] is None and row['summary']['outcome'] ==
         {'status': 'unsupported', 'code': 'unsupported_document', 'exit_code': 7, 'written_count': 0},
         'Unsupported metadata-only export invented a plan or positive output')
    need(all(secret not in raw for secret in (application['input_path'], fixtures.USER_PASSWORD, fixtures.OWNER_PASSWORD,
         'Original authored fixture author', 'WBS-POLICY-PAGE-', 'ORIGINAL_INERT_POLICY_')), 'Support export leaked fixture/private data')


def validate_document_policy_report(report, requested_shell_paths, *, cleanup_complete=True):
    need(isinstance(report, dict) and type(report.get('schema_version')) is int and report['schema_version'] == 1
         and report.get('task_id') == 'M4-T02' and report.get('result') == 'DOCUMENT_POLICY_REGRESSION_PASSED'
         and report.get('success') is True and type(report.get('exit_code')) is int and report['exit_code'] == 0
         and report.get('acceptance_ids') == ['AC-070', 'AC-071', 'AC-072'] and report.get('source_unchanged') is True
         and report.get('human_or_GUI_tested') is False and (not cleanup_complete or report.get('owned_temp_removed') is True),
         'Document-policy aggregate/source/cleanup scope invalid')
    hosts = report.get('host_cases')
    need(isinstance(hosts, list) and len(hosts) == 2 and {host.get('id') for host in hosts} == {'PS51', 'PS7'}
         and len(requested_shell_paths) == 2 and {str(Path(host['shell_executable']).resolve()).casefold() for host in hosts}
         == {str(Path(path).resolve()).casefold() for path in requested_shell_paths}, 'Both actual requested supported hosts missing')
    for host in hosts:
        major = 5 if host['id'] == 'PS51' else 7
        need(host.get('passed') is True and type(host.get('exit_code')) is int and host['exit_code'] == 0
             and type(host.get('host_major')) is int and host['host_major'] == major
             and isinstance(host.get('host_version'), str) and host['host_version'].startswith('5.1.' if major == 5 else '7.')
             and type(host.get('syntax_error_count')) is int and host['syntax_error_count'] == 0
             and host['stored_policies'] == host.get('policies_after'),
             'Host/settings evidence absent or policy changed')
    tested = report.get('tested_path_sha256')
    need(isinstance(tested, dict) and all(name in tested for name in APPLICATION) and all(HEX.fullmatch(value) for value in tested.values()),
         'Raw tested application/source map incomplete')
    raw = b64decode(report['batch_original_powershell_raw_base64'], validate=True)
    need(len(raw) <= 1048576 and sha256(raw).hexdigest() == tested['WinBookSplit.ps1'], 'Original BAT-control PS bytes differ from tested source')
    batch_hash = sha256(controlled_powershell(raw)).hexdigest()
    provenance = report.get('fixture_provenance')
    need(isinstance(provenance, dict) and provenance.get('private_data') is False and provenance.get('pdf_file_count') == 25
         and provenance.get('generator_sha256') == tested.get('tests/document_policy/generate_document_policy_fixtures.py')
         and set(provenance.get('fixtures', {})) == set(fixtures.KINDS), 'Original policy fixture provenance missing')
    cases, exports = report.get('cases'), report.get('export_cases')
    need(isinstance(cases, list) and len(cases) == 73 and {row.get('id') for row in cases} ==
         {host + '-' + kind for host in ('PS51', 'PS7') for kind in PS_KINDS} | {'BAT-' + kind for kind in BAT_KINDS}
         and report.get('application_case_count') == 73 and isinstance(exports, list) and len(exports) == 2
         and report.get('export_case_count') == 2 and type(report.get('case_count')) is int and report['case_count'] == 75,
         'Complete actual policy/app/BAT/export matrix absent or duplicated')
    for row in cases:
        host = next(host for host in hosts if host['id'] == ('PS51' if row['host_id'] == 'BAT' else row['host_id']))
        need(row['shell_executable'] == host['shell_executable'] and row['host_version'] == host['host_version']
             and row['application_sha256'] == {name: tested[name] for name in APPLICATION}
             and row['source_observations_before']['source']['sha256'] == provenance['fixtures'][row['fixture_kind']]['sha256']
             and row.get('application_unchanged') is True, 'Actual source/host/application provenance differs')
        expected = {name: tested[name] for name in APPLICATION}
        if row['host_id'] == 'BAT':
            expected['WinBookSplit.ps1'] = batch_hash
        need(row.get('actual_application_sha256') == expected, 'Copied BAT controls changed undeclared application bytes')
        validate_case(row, cleanup_complete=cleanup_complete)
    need({row.get('id') for row in exports} == {'PS51-unsupported-export', 'PS7-unsupported-export'}, 'Two actual unsupported exports missing')
    for exported in exports:
        application = next(row for row in cases if row['id'] == exported['application_case_id'])
        need(application['kind'] == 'encrypted-owner-empty' and application['host_id'] in {'PS51', 'PS7'}, 'Wrong null-plan export input selected')
        validate_export(exported, application)
