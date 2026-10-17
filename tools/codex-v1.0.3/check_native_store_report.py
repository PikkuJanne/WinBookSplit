#!/usr/bin/env python3
"""Offline consistency gate for the existing-owner Store scenario.

Reads a report plus bounded local receipt files. No subprocess/network/writes.
This checks declarations and artifact integrity, NOT native execution truth.
Run the maintained native/PDF validators and inspect receipts independently.
"""
from __future__ import annotations
import argparse
import hashlib
import json
import math
from pathlib import Path, PurePosixPath
import re
import stat
import sys
from typing import Any

FAMILY = 'PythonSoftwareFoundation.PythonManager_qbz5n2kfra8p0'
PUBLISHER = 'CN=4975D53F-AA7E-49A5-8B49-EA4FDC1BB66B'
SCHEMA_VERSION = 2
FORBIDDEN = (
    'python_installed_or_updated_by_kit', 'manager_installed_or_updated_by_kit',
    'private_runtime_fallback_attempted', 'private_runtime_fallback_used',
    'runtime_supplied_by_kit', 'runtime_copied_or_cloned', 'runtime_environment_injected',
    'explicit_python_path_supplied', 'hidden_runtime_selection_supplied',
    'parent_path_changed', 'child_path_injected', 'localappdata_overridden',
    'virtual_env_overridden', 'application_venv_prebuilt', 'alias_settings_changed',
)
CHECKS = (
    'registration_binding_verified', 'installed_only_listing',
    'bounded_probe_completed', 'isolated_interpreter_verified',
    'normal_setup_bat_passed', 'second_setup_noop', 'readiness_readonly',
    'ordinary_application_bat_passed', 'manual_pdf_valid', 'auto1_pdf_valid',
    'auto2_pdf_valid', 'full_page_coverage', 'source_and_neighbors_unchanged',
    'unique_output_no_overwrite', 'unrelated_cwd_passed', 'manager_state_unchanged',
    'global_packages_unchanged', 'parent_environment_unchanged',
)
KINDS = ('registration', 'discovery', 'setup', 'bat', 'pdf_validation', 'preservation')
APPLICATION_RECORD_MINIMUM = {'discovery': 1, 'setup': 3, 'bat': 1, 'pdf_validation': 4, 'preservation': 2}

class GateError(ValueError):
    """Malformed, missing, disqualifying or inconsistent evidence."""

def need(ok: bool, message: str) -> None:
    if not ok:
        raise GateError(message)

def digest(value: Any, length: int, field: str) -> str:
    need(isinstance(value, str) and re.fullmatch('[0-9a-f]{'+str(length)+'}', value) is not None,
         field + ': expected lowercase hexadecimal digest')
    return value

def unique_object(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        need(key not in result, 'duplicate JSON key: ' + key)
        result[key] = value
    return result

def parse_json(raw: bytes) -> Any:
    def bad_constant(value: str) -> None:
        raise GateError('non-finite JSON constant: ' + value)
    return json.loads(raw.decode('utf-8-sig'), object_pairs_hook=unique_object,
                      parse_constant=bad_constant)

def read_regular(path: Path, limit: int = 4 * 1024 * 1024) -> bytes:
    s = path.lstat()
    need(stat.S_ISREG(s.st_mode), 'not a regular evidence file: ' + path.name)
    need(not path.is_symlink() and not (getattr(s, 'st_file_attributes', 0) & 0x400),
         'linked/reparse evidence file: ' + path.name)
    need(0 < s.st_size <= limit, 'empty or oversized evidence file: ' + path.name)
    with path.open('rb') as handle:
        raw = handle.read(limit + 1)
    need(0 < len(raw) <= limit, 'invalid evidence length: ' + path.name)
    return raw

def artifact_path(root: Path, relative: Any) -> Path:
    need(isinstance(relative, str) and 0 < len(relative) <= 240,
         'invalid receipt path')
    need('\\' not in relative and ':' not in relative and
         not re.search(r'[\x00-\x1f\x7f]', relative), 'unsafe receipt path')
    pure = PurePosixPath(relative)
    need(not pure.is_absolute() and '..' not in pure.parts and
         '.' not in relative.split('/') and '' not in relative.split('/'),
         'receipt path must be relative with ordinary components')
    path = root
    for part in pure.parts:
        path = path / part
        s = path.lstat()
        need(not path.is_symlink() and not (getattr(s, 'st_file_attributes', 0) & 0x400),
             'receipt path contains link/reparse point')
    need(path.resolve().is_relative_to(root.resolve()), 'receipt escaped report directory')
    return path

def obj(value: Any, name: str) -> dict[str, Any]:
    need(isinstance(value, dict), name + ': expected object')
    return value

def ref(value: Any, field: str) -> str:
    need(isinstance(value, str) and 0 < len(value) <= 4096 and
         not re.search(r'[\x00-\x1f\x7f]', value), field + ': expected normalized identity token')
    return value

def runtime_inventory(value: Any, field: str) -> dict[str, dict[str, Any]]:
    need(isinstance(value, list) and 1 <= len(value) <= 32, field + ': invalid runtime inventory')
    out = {}
    for entry in value:
        entry = obj(entry, field)
        key = ref(entry.get('runtime_ref'), field + '.runtime_ref')
        need(key not in out, field + ': duplicate runtime reference')
        digest(entry.get('sha256'), 64, field + '.sha256')
        need(isinstance(entry.get('version'), str) and
             re.fullmatch(r'3\.[0-9]+\.[0-9]+', entry['version']) is not None,
             field + ': invalid stable version')
        need(entry.get('managed_by_store') is True, field + ': runtime lacks Store provenance')
        out[key] = entry
    return out

def expected_summaries(report: dict[str, Any]) -> dict[str, Any]:
    return {
        'registration': report['registration'],
        'discovery': report['discovery'],
        'setup': report['setup'],
        'bat': {key: report['checks'][key] for key in
                ('ordinary_application_bat_passed', 'unrelated_cwd_passed')},
        'pdf_validation': {key: report['checks'][key] for key in
                           ('manual_pdf_valid', 'auto1_pdf_valid', 'auto2_pdf_valid',
                            'full_page_coverage', 'source_and_neighbors_unchanged',
                            'unique_output_no_overwrite')},
        'preservation': {'before': report['before'], 'after': report['after'],
                         'forbidden_substitutions': report['forbidden_substitutions'],
                         'fallback_acquisition': report['fallback_acquisition'],
                         'inventory': report['inventory'], 'policy_component': report['policy_component']},
    }

def validate_receipt_records(records: list[Any], kind: str) -> None:
    """Check the kit's structured app-command contract, never guess from log text."""
    application_count = 0

    def visit(value: Any) -> None:
        nonlocal application_count
        if isinstance(value, list):
            for item in value:
                visit(item)
        elif isinstance(value, dict):
            for key in FORBIDDEN:
                if key in value:
                    need(value[key] is False, 'receipt contradicts no-substitution declaration: ' + key)
            if 'baseline_kind' in value:
                need(value['baseline_kind'] == 'pre_existing_owner_install',
                     'private or unproven runtime baseline in native receipt')
            if 'owner_regular315_usable' in value:
                need(value['owner_regular315_usable'] is True,
                     'receipt declares owner runtime unusable')
            if isinstance(value.get('command'), list):
                # Toolkit CLI --python-path names its already-created venv for
                # PDF verification. The PowerShell application -PythonPath and
                # its abbreviations are forbidden even in a mislabelled record.
                for argument in value['command']:
                    if isinstance(argument, str):
                        parameter = argument.casefold().split(':', 1)[0]
                        need(not (len(parameter) >= 2 and '-pythonpath'.startswith(parameter)),
                             'supplied PythonPath in native command receipt')
            if value.get('role') == 'application':
                application_count += 1
                need(value.get('environment') == {}, 'application receipt contains environment overrides')
                process = obj(value.get('process'), 'application process receipt')
                command = process.get('command')
                need(isinstance(command, list) and 1 <= len(command) <= 128,
                     'application receipt requires a literal argument vector')
                for argument in command:
                    ref(argument, 'application argument')
                    parameter = argument.casefold().split(':', 1)[0]
                    need(not (len(parameter) >= 2 and '-pythonpath'.startswith(parameter)) and
                         parameter != '--python-path' and not parameter.startswith('--python-path='),
                         'supplied PythonPath in application receipt')
                need(process.get('failure') is None and
                     process.get('environment_before_after_restored') is True,
                     'application process failed or environment restoration is unproved')
                probe = obj(process.get('probe'), 'application process probe')
                need(type(probe.get('Pid')) is int and probe['Pid'] > 0 and
                     type(probe.get('ExitCode')) is int and probe['ExitCode'] == 0,
                     'application receipt has no successful native process')
                need(all(probe.get(key) is True for key in
                         ('JobAssigned', 'ParentStopped', 'DescendantsStopped', 'StreamsComplete', 'InputWriterStopped')) and
                     all(probe.get(key) is False for key in
                         ('TimedOut', 'Cancelled', 'StdoutTruncated', 'StderrTruncated')) and
                     all(key in probe and probe[key] is None for key in
                         ('StartError', 'StopError', 'StreamError', 'ResultError', 'InteractionError', 'InputError')),
                     'application receipt lacks safe completed process supervision')
                elapsed = probe.get('ElapsedSeconds')
                need(type(elapsed) in (int, float) and 0 <= elapsed <= 608 and math.isfinite(elapsed),
                     'application receipt process duration is unbounded')
                redacted_bytes = obj(probe.get('redacted_stream_bytes'), 'redacted stream bytes')
                raw_hashes = obj(probe.get('raw_stream_sha256'), 'raw stream hashes')
                need(set(redacted_bytes) == set(raw_hashes) == {'stdout', 'stderr'},
                     'application receipt stream redaction provenance missing')
                for stream in ('Stdout', 'Stderr'):
                    # Native totals describe original local bytes. Redaction can
                    # change lengths, so never overwrite those measurements with
                    # the returned text's size or claim the two are equal.
                    name = stream.casefold()
                    need(type(probe.get(stream + 'TotalBytes')) is int and
                         0 <= probe[stream + 'TotalBytes'] <= 1048576,
                         'application receipt original stream byte bound is invalid')
                    digest(raw_hashes[name], 64, 'original stream hash')
                    need(isinstance(probe.get(stream), str) and type(redacted_bytes[name]) is int and
                         redacted_bytes[name] == len(probe[stream].encode('utf-8', errors='strict')),
                         'application receipt redacted stream bytes differ')
            for item in value.values():
                visit(item)

    visit(records)
    need(application_count >= APPLICATION_RECORD_MINIMUM.get(kind, 0),
         'structured application command receipts missing: ' + kind)

def validate_policy_records(records: list[Any]) -> None:
    """Correlate the local policy component's structured native operations."""
    stages: dict[str, list[dict[str, Any]]] = {}
    def visit(value: Any) -> None:
        if isinstance(value, list):
            for item in value: visit(item)
        elif isinstance(value, dict):
            if isinstance(value.get('stage'), str): stages.setdefault(value['stage'], []).append(value)
            for item in value.values(): visit(item)
    visit(records)
    def one(stage: str) -> dict[str, Any]:
        rows = stages.get(stage, [])
        need(len(rows) == 1, 'missing/duplicate policy operation: ' + stage)
        return rows[0]
    before, after = one('before'), one('after')
    need(before.get('role') == after.get('role') == 'policy_observation' and
         before.get('inbox_policy') == after.get('inbox_policy'), 'inbox policy observations differ')
    policy = obj(before.get('inbox_policy'), 'inbox policy observation')
    scopes = policy.get('scopes')
    names = {'MachinePolicy', 'UserPolicy', 'Process', 'CurrentUser', 'LocalMachine'}
    need(isinstance(scopes, list) and len(scopes) == 5 and all(isinstance(row, dict) for row in scopes) and
         {row.get('scope') for row in scopes} == names, 'inbox policy observation scopes incomplete')
    allowed = {'Undefined', 'Restricted', 'AllSigned', 'RemoteSigned', 'Unrestricted', 'Bypass', 'Default'}
    need(policy.get('effective') in allowed and all(row.get('policy') in allowed for row in scopes),
         'inbox policy observation values invalid')
    marks = one('authored_motw').get('marks')
    need(isinstance(marks, list) and len(marks) == 2, 'two actual entry-script Internet marks required')
    marked = {}
    for mark in marks:
        mark = obj(mark, 'Internet mark')
        path = ref(mark.get('path'), 'marked script path')
        leaf = path.replace('\\', '/').rsplit('/', 1)[-1]
        need(leaf in {'WinBookSplit.ps1', 'Setup-WinBookSplit.ps1'} and leaf not in marked and
             mark.get('ads') == '[ZoneTransfer]\r\nZoneId=3\r\n', 'invalid/duplicate Internet mark')
        digest(mark.get('sha256'), 64, 'marked script data hash')
        marked[leaf] = path
    denial = one('direct_file_remotesigned')
    process = obj(denial.get('process'), 'direct policy process')
    command = process.get('command')
    need(denial.get('role') == 'expected_policy_denial' and denial.get('environment') == {} and
         isinstance(command, list) and command[1:] == ['-NoProfile', '-NonInteractive', '-InputFormat', 'Text',
             '-OutputFormat', 'Text', '-ExecutionPolicy', 'RemoteSigned', '-File', marked['WinBookSplit.ps1'], '-Version'] and
         len(command) == 12 and isinstance(command[0], str) and process.get('failure') is None and
         process.get('environment_before_after_restored') is True, 'direct marked policy command differs')
    probe = obj(process.get('probe'), 'direct policy probe')
    need(type(probe.get('ExitCode')) is int and probe['ExitCode'] != 0 and type(probe.get('Pid')) is int and probe['Pid'] > 0 and
         all(probe.get(key) is True for key in ('JobAssigned', 'ParentStopped', 'DescendantsStopped', 'StreamsComplete', 'InputWriterStopped')) and
         all(probe.get(key) is False for key in ('TimedOut', 'Cancelled', 'StdoutTruncated', 'StderrTruncated')) and
         all(key in probe and probe[key] is None for key in ('StartError', 'StopError', 'StreamError', 'ResultError', 'InteractionError', 'InputError')) and
         isinstance(probe.get('Stderr'), str) and re.search('PSSecurityException|UnauthorizedAccess', probe['Stderr']) and
         isinstance(probe.get('Stdout'), str) and not re.search(r'(?m)^WinBookSplit ', probe['Stdout']),
         'direct policy refusal is not a safely completed native denial')
    elapsed = probe.get('ElapsedSeconds')
    need(type(elapsed) in (int, float) and 0 <= elapsed <= 38 and math.isfinite(elapsed),
         'direct policy refusal duration is unbounded')
    lengths = obj(probe.get('redacted_stream_bytes'), 'policy redacted stream bytes')
    hashes = obj(probe.get('raw_stream_sha256'), 'policy original stream hashes')
    need(set(lengths) == set(hashes) == {'stdout', 'stderr'}, 'policy stream provenance missing')
    for stream in ('Stdout', 'Stderr'):
        name = stream.casefold()
        need(type(probe.get(stream + 'TotalBytes')) is int and 0 <= probe[stream + 'TotalBytes'] <= 1048576 and
             type(lengths[name]) is int and lengths[name] == len(probe[stream].encode('utf-8', errors='strict')),
             'policy refusal original/redacted stream measurements differ')
        digest(hashes[name], 64, 'policy original stream hash')
    bat, action = one('marked_setup_bat_check'), one('exact_automatic_formatter_action')
    need(bat.get('role') == action.get('role') == 'application', 'marked positive policy operations missing')
    text = action.get('action')
    need(isinstance(text, str) and text.endswith(" -File '" + marked['Setup-WinBookSplit.ps1'].replace("'", "''") + "' -NonInteractive") and
         re.fullmatch(r"& '(?:[^']|'')+' -NoProfile -ExecutionPolicy Bypass -File '(?:[^']|'')+' -NonInteractive", text) is not None and
         '[EXACT_ACTION] ' + text in action['process']['probe']['Stdout'],
         'exact automatic formatter Action differs or was not retained')


def validate(report_path: Path, expected_commit: str, expected_zip_sha256: str) -> dict[str, Any]:
    digest(expected_commit, 40, 'expected commit')
    digest(expected_zip_sha256, 64, 'expected ZIP SHA256')
    report = obj(parse_json(read_regular(report_path, 2 * 1024 * 1024)), 'report')
    need(type(report.get('schema_version')) is int and report['schema_version'] == SCHEMA_VERSION,
         'unsupported report schema')
    need(report.get('scenario') == 'existing_owner_windows10_store_installation',
         'not the existing-owner Store scenario')
    need(report.get('evidence_kind') == 'native_execution' and report.get('status') == 'passed',
         'native gate not executed/passed')
    need(report.get('application_version') == '1.0.3', 'wrong application version')
    need(report.get('source_commit') == expected_commit, 'source commit mismatch')
    need(report.get('zip_sha256') == expected_zip_sha256, 'ZIP digest mismatch')
    run_id = report.get('run_id')
    need(isinstance(run_id, str) and re.fullmatch(r'[A-Za-z0-9_-]{8,80}', run_id) is not None,
         'invalid run ID')
    host = obj(report.get('host'), 'host')
    need(host.get('execution_kind') == 'native' and host.get('owner_desktop') is True,
         'owner desktop native execution missing')
    need(host.get('os') == 'Windows 10' and host.get('architecture') == 'x64',
         'wrong native OS/architecture')
    need(type(host.get('build')) is int and 19045 <= host['build'] < 22000,
         'not the targeted Windows 10 build range')
    need(isinstance(host.get('powershell_version'), str) and
         re.fullmatch(r'5\.1\.[0-9]+\.[0-9]+', host['powershell_version']) is not None,
         'actual inbox PS5.1 required')
    registration = obj(report.get('registration'), 'registration')
    need(registration.get('family') == FAMILY and registration.get('publisher') == PUBLISHER and
         registration.get('signature_kind') == 'Store' and registration.get('status') == 'Ok' and
         registration.get('architecture') == 'X64', 'not the verified Store registration tuple')
    need(isinstance(registration.get('version'), str) and
         re.fullmatch(r'[0-9]+\.[0-9]+\.[0-9]+\.[0-9]+', registration['version']) is not None,
         'manager version missing')
    denied = obj(report.get('forbidden_substitutions'), 'forbidden substitutions')
    need(set(denied) == set(FORBIDDEN), 'missing/unknown no-substitution observations')
    for key in FORBIDDEN:
        need(denied[key] is False, 'disqualifying or unknown substitution: ' + key)
    acquisition = obj(report.get('fallback_acquisition'), 'fallback acquisition')
    need(set(acquisition) == {'policy', 'attempted', 'used'} and
         acquisition.get('policy') == 'disabled_before_discovery' and
         acquisition.get('attempted') is False and acquisition.get('used') is False,
         'fallback acquisition must be disabled before discovery without any attempt or use')
    inventory = obj(report.get('inventory'), 'inventory')
    need(set(inventory) == {'purpose', 'supplied_to_application'} and
         inventory.get('purpose') == 'verification_only' and
         inventory.get('supplied_to_application') is False,
         'independent inventory must not supply runtime selection to the application')
    before = obj(report.get('before'), 'before')
    after = obj(report.get('after'), 'after')
    need(before.get('application_venv_exists') is False, 'application venv existed before discovery')
    before_runtimes = runtime_inventory(before.get('runtimes'), 'before.runtimes')
    after_runtimes = runtime_inventory(after.get('runtimes'), 'after.runtimes')
    need(before_runtimes == after_runtimes, 'pre-existing runtime inventory changed')
    discovery = obj(report.get('discovery'), 'discovery')
    need(discovery.get('mode') == 'automatic' and
         discovery.get('route') == 'registered_store_manager' and
         discovery.get('provenance') == 'pre_existing_owner_install',
         'discovery route is substituted or unproven')
    chosen = ref(discovery.get('runtime_ref'), 'chosen runtime')
    need(chosen in before_runtimes, 'chosen runtime absent from before inventory')
    need(discovery.get('version') == before_runtimes[chosen]['version'] and
         re.fullmatch(r'3\.15\.[0-9]+', discovery['version']) is not None,
         'owner 3.15 interpreter not selected')
    need(discovery.get('interpreter_probe_exit') == 0 and
         type(discovery.get('interpreter_probe_exit')) is int, 'interpreter probe failed')
    need(discovery.get('package_state') in ('ready', 'setup_required'),
         'compatible base/package state not distinguished')
    setup = obj(report.get('setup'), 'setup')
    need(setup.get('entry') == 'Setup-WinBookSplit.bat' and
         setup.get('runtime_selection') == 'automatic', 'normal automatic setup required')
    need(type(setup.get('exit_code')) is int and setup['exit_code'] == 0, 'setup failed')
    need(setup.get('base_runtime_ref') == chosen and setup.get('venv_base_runtime_ref') == chosen,
         'setup/venv base differs from discovered owner runtime')
    need(after.get('application_venv_exists') is True, 'normal setup did not create application environment')
    checks = obj(report.get('checks'), 'checks')
    need(set(checks) == set(CHECKS), 'missing/unknown native checks')
    for key in CHECKS:
        need(checks[key] is True, 'native check not passed: ' + key)
    policy = obj(report.get('policy_component'), 'policy component')
    need(policy.get('status') == 'PASS' and policy.get('scope') == 'OWNED_EXACT_EXTRACTION_HEALTHY_VENV',
         'exact-extraction local policy component not passed')
    for key in ('direct_remotesigned_denied', 'marked_setup_bat_noop', 'exact_formatter_action_noop',
                'source_bytes_and_marks_preserved', 'venv_unchanged', 'inbox_policy_unchanged'):
        need(policy.get(key) is True, 'local policy component check not passed: ' + key)
    organization = obj(policy.get('organization_policy_denial'), 'organization policy denial')
    need(organization.get('status') == 'NOT_RUN' and isinstance(organization.get('reason'), str)
         and bool(organization['reason']), 'local MOTW must not claim real organization-policy denial')
    artifacts = report.get('artifacts')
    need(isinstance(artifacts, list) and len(artifacts) == len(KINDS), 'incomplete native receipt set')
    expected = expected_summaries(report)
    seen_kinds: set[str] = set()
    seen_paths: set[str] = set()
    for artifact in artifacts:
        artifact = obj(artifact, 'artifact')
        kind = artifact.get('kind')
        need(kind in KINDS and kind not in seen_kinds, 'unknown/duplicate receipt kind')
        seen_kinds.add(kind)
        relative = artifact.get('path')
        need(isinstance(relative, str) and relative not in seen_paths, 'duplicate/missing receipt path')
        seen_paths.add(relative)
        path = artifact_path(report_path.parent, relative)
        need(path.resolve() != report_path.resolve(), 'report cannot be its own receipt')
        raw = read_regular(path)
        need(hashlib.sha256(raw).hexdigest() == digest(artifact.get('sha256'), 64, 'receipt hash'),
             'receipt byte hash mismatch: ' + relative)
        receipt = obj(parse_json(raw), 'receipt')
        need(type(receipt.get('schema_version')) is int and receipt['schema_version'] == SCHEMA_VERSION and
             receipt.get('kind') == kind and receipt.get('run_id') == run_id and
             receipt.get('source_commit') == expected_commit and
             receipt.get('zip_sha256') == expected_zip_sha256, 'receipt identity mismatch: ' + relative)
        need(receipt.get('summary') == expected[kind], 'receipt/report summary mismatch: ' + kind)
        need(isinstance(receipt.get('records'), list) and 1 <= len(receipt['records']) <= 256,
             'actual command/output receipt records missing: ' + kind)
        validate_receipt_records(receipt['records'], kind)
        if kind == 'preservation': validate_policy_records(receipt['records'])
    return {'status': 'CONSISTENT_NATIVE_REPORT', 'run_id': run_id,
            'source_commit': expected_commit, 'zip_sha256': expected_zip_sha256,
            'receipt_count': len(artifacts),
            'limitation': 'Checks declarations and receipt bytes, not physical execution truth. Review native receipts and application validators.'}

def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('report', type=Path)
    p.add_argument('--expected-commit', required=True)
    p.add_argument('--expected-zip-sha256', required=True)
    args = p.parse_args()
    try:
        result = validate(args.report, args.expected_commit, args.expected_zip_sha256)
    except (GateError, OSError, ValueError, TypeError, KeyError, RecursionError) as exc:
        print(json.dumps({'status': 'BLOCKED', 'reason': str(exc)}, ensure_ascii=True))
        return 2
    print(json.dumps(result, ensure_ascii=True))
    return 0

if __name__ == '__main__':
    raise SystemExit(main())
