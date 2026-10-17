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
from pathlib import Path, PurePosixPath
import re
import stat
import sys
from typing import Any

FAMILY = 'PythonSoftwareFoundation.PythonManager_qbz5n2kfra8p0'
PUBLISHER = 'CN=4975D53F-AA7E-49A5-8B49-EA4FDC1BB66B'
FORBIDDEN = (
    'python_installed_or_updated_by_kit', 'manager_installed_or_updated_by_kit',
    'private_runtime_fallback_used', 'runtime_copied_or_cloned',
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
                         'forbidden_substitutions': report['forbidden_substitutions']},
    }

def validate(report_path: Path, expected_commit: str, expected_zip_sha256: str) -> dict[str, Any]:
    digest(expected_commit, 40, 'expected commit')
    digest(expected_zip_sha256, 64, 'expected ZIP SHA256')
    report = obj(parse_json(read_regular(report_path, 2 * 1024 * 1024)), 'report')
    need(type(report.get('schema_version')) is int and report['schema_version'] == 1,
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
        need(type(receipt.get('schema_version')) is int and receipt['schema_version'] == 1 and
             receipt.get('kind') == kind and receipt.get('run_id') == run_id and
             receipt.get('source_commit') == expected_commit and
             receipt.get('zip_sha256') == expected_zip_sha256, 'receipt identity mismatch: ' + relative)
        need(receipt.get('summary') == expected[kind], 'receipt/report summary mismatch: ' + kind)
        need(isinstance(receipt.get('records'), list) and 1 <= len(receipt['records']) <= 256,
             'actual command/output receipt records missing: ' + kind)
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
