"""Observe one uninstrumented R shell experiment; never acceptance evidence."""
from datetime import datetime, timezone
from hashlib import sha256
import argparse
import importlib.util
import json
import os
from pathlib import Path, PurePosixPath
import re
import struct
import subprocess
import sys
import unicodedata

R = '001669a051077536864b3f787ca0ec111fe06013'
DIGEST = 'a80aa2f1b08d97d9e13fff39f2b08b079d40977120efe72df32a86f2cd7c50c6'
COORDINATOR = 'b44333f809851edac7fd84d0c59bcda100c2a9e8528ac240ab06120e56568567'
ENV_NAMES = ('TEMP', 'TMP', 'RUNNER_TEMP', 'GITHUB_WORKSPACE', 'GITHUB_SHA',
             'GITHUB_RUN_ID', 'GITHUB_RUN_ATTEMPT', 'ImageOS', 'ImageVersion',
             'PSModulePath', 'SystemRoot', 'USERPROFILE', 'LOCALAPPDATA',
             'WBS_TEST_WHEELHOUSE')


def ordinary(path: Path) -> dict:
    if not path.exists() or path.is_symlink() or path.is_junction():
        raise ValueError('Observed owned path is absent or a reparse point')
    value = path.stat()
    return {'path': str(path), 'device': value.st_dev, 'inode': value.st_ino,
            'is_directory': path.is_dir()}


def read_bound(path: Path, limit=64 * 1024 * 1024) -> tuple[bytes, dict]:
    if not path.is_file() or path.is_symlink() or path.is_junction():
        raise ValueError('Receipt/tool must be an ordinary file')
    before = path.stat()
    if not 0 < before.st_size <= limit:
        raise ValueError('Receipt/tool exceeds finite size bound')
    data = path.read_bytes()
    after = path.stat()
    if (before.st_dev, before.st_ino, before.st_size, before.st_mtime_ns) != \
            (after.st_dev, after.st_ino, after.st_size, after.st_mtime_ns):
        raise ValueError('Receipt/tool changed while read')
    return data, {'bytes': len(data), 'sha256': sha256(data).hexdigest()}


def write_new(path: Path, value: dict) -> None:
    with path.open('x', encoding='utf-8', newline='\n') as stream:
        json.dump(value, stream, indent=2)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--phase', choices=('before', 'after'), required=True)
    parser.add_argument('--reports', type=Path, required=True)
    parser.add_argument('--tool-root', type=Path, required=True)
    parser.add_argument('--temp-root', type=Path, required=True)
    parser.add_argument('--runner-exit', type=int)
    args = parser.parse_args()
    if sys.version_info[:3] != (3, 14, 8) or struct.calcsize('P') != 8 \
            or not sys.flags.isolated or not sys.dont_write_bytecode \
            or sys.version_info.releaselevel != 'final' \
            or bool(sysconfig_gil_disabled()):
        raise ValueError('Observer requires isolated regular final x64 3.14.8')
    if args.phase == 'after' and args.runner_exit is None:
        raise ValueError('After observation requires the actual runner exit')
    if os.environ.get('WBS_TEST_WHEELHOUSE') is not None:
        raise ValueError('Controlled original wheel preparation requires absent override')
    repo = Path.cwd().resolve()
    environment = {key: value for key, value in os.environ.items()
                   if not key.upper().startswith('GIT_')}
    environment.update(GIT_NO_REPLACE_OBJECTS='1', GIT_TERMINAL_PROMPT='0')

    def git(*arguments):
        result = subprocess.run(['git', *arguments], cwd=repo, env=environment,
                                stdin=subprocess.DEVNULL, capture_output=True,
                                timeout=60, check=True)
        if len(result.stdout) > 32 * 1024 * 1024:
            raise ValueError('Git observation exceeds finite bound')
        return result.stdout

    if git('remote', 'get-url', 'origin').decode().strip().removesuffix('.git') \
            != 'https://github.com/PikkuJanne/WinBookSplit' \
            or git('rev-parse', 'HEAD').decode().strip() != R \
            or git('status', '--porcelain', '--untracked-files=all') != b'':
        raise ValueError('Experiment requires clean exact R and fixed origin')
    runner_path = repo / 'tests/run_tests.py'
    if runner_path.read_bytes() != git('show', R + ':tests/run_tests.py'):
        raise ValueError('Maintained runner differs from R')
    spec = importlib.util.spec_from_file_location('exact_R_temp_observer', runner_path)
    runner = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(runner)
    source = runner.source_manifest()
    if len(source) != 207 or runner.manifest_digest(source) != DIGEST:
        raise ValueError('Experiment source map differs')
    for name, digest in source.items():
        if type(name) is not str or any(c in name for c in '\\:<>"|?*') \
                or any(unicodedata.category(c) in {'Cc', 'Cf', 'Cs'} for c in name) \
                or PurePosixPath(name).as_posix() != name \
                or any(part in ('', '.', '..') or part != part.strip()
                       or part.endswith('.') for part in name.split('/')):
            raise ValueError('Source name is not canonical')
        if sha256(git('show', R + ':' + name)).hexdigest() != digest:
            raise ValueError('Actual source differs from R Git blobs')
    original, original_id = read_bound(repo / 'tests/Invoke-ShellTests.ps1')
    if original_id['sha256'] != COORDINATOR \
            or original.count(b"$config.Output.Verbosity = 'None'") != 1:
        raise ValueError('Original uninstrumented coordinator differs')
    base = Path(os.environ['RUNNER_TEMP']).resolve() / 'WinBookSplit-ci'
    if args.reports.resolve() != base / 'reports' \
            or args.tool_root.resolve() != base / 'tools' \
            or args.temp_root.resolve() != base / 'shell-owned-temp' \
            or base.is_relative_to(repo):
        raise ValueError('Experiment roots must be the fixed owned RUNNER_TEMP layout')
    layout = {name: ordinary(path) for name, path in
              (('runner_temp', base.parent), ('root', base),
               ('reports', args.reports), ('tools', args.tool_root),
               ('shell_temp', args.temp_root))}
    if args.phase == 'before' and any(args.temp_root.iterdir()):
        raise ValueError('Process-local shell temp root must start empty')
    python_data, python_id = read_bound(args.reports / 'python.json')
    predecessor = json.loads(python_data)
    if predecessor['layer'] != 'python' or predecessor['success'] is not True \
            or predecessor['source']['base_commit'] != R \
            or predecessor['source']['tested_path_sha256'] != source \
            or predecessor['source']['source_unchanged'] is not True:
        raise ValueError('The original full Python predecessor did not pass exact R')
    units = predecessor['steps'][0]
    match = re.search(r'Ran (\d+) tests? in ([0-9.]+)s', units['stderr'])
    if units['name'] != 'stdlib-unittest' or units['exit_code'] != 0 \
            or units.get('timed_out') is not False or match is None \
            or int(match[1]) != 670 or not re.search(r'^OK\s*$', units['stderr'], re.M) \
            or 'skipped=' in units['stderr']:
        raise ValueError('The 670-test predecessor is incomplete or failed')
    inbox = Path(os.environ['SystemRoot']) / 'System32/WindowsPowerShell/v1.0/powershell.exe'
    tool_files = sorted(path for path in args.tool_root.rglob('*') if path.is_file())
    if not 0 < len(tool_files) <= 4096:
        raise ValueError('Pinned tool closure exceeds finite member bound')
    tools = {path.relative_to(args.tool_root).as_posix(): read_bound(path)[1]
             for path in tool_files}
    binaries = {'python': read_bound(Path(sys.executable))[1],
                'inbox': read_bound(inbox)[1]}
    receipt_ids = {name: read_bound(args.reports / name)[1] for name in
                   ('os.json', 'tools-bootstrap.json', 'package-inputs.json', 'python.json')}
    record = {'schema_version': 1, 'kind': 'DIAGNOSTIC_ONLY_UNMODIFIED_TEMP',
              'acceptance_claimed': False, 'phase': args.phase,
              'observed_at_utc': datetime.now(timezone.utc).isoformat(),
              'orchestration_commit': os.environ['GITHUB_SHA'],
              'run_id': os.environ['GITHUB_RUN_ID'],
              'run_attempt': os.environ['GITHUB_RUN_ATTEMPT'],
              'source_commit': R, 'source_digest': DIGEST, 'source_sha256': source,
              'original_coordinator': original_id, 'output_verbosity': 'None',
              'source_clean': True, 'environment': {key: os.environ.get(key) for key in ENV_NAMES},
              'actual_python': {'version': sys.version, 'executable': sys.executable,
                               'isolated': True, 'no_bytecode': True, 'bits': 64},
              'owned_layout': layout, 'tool_closure': tools, 'binaries': binaries,
              'receipt_identities': receipt_ids,
              'python_predecessor': {'tests': 670, 'seconds': float(match[2]),
                                     'actual_result': 'PASS'},
              'controlled_change': 'Only the shell runner child process TEMP/TMP point to the fresh owned RUNNER_TEMP directory; original coordinator/tests/rules and limits unchanged',
              'maintained_native_timeout_seconds': 300,
              'source_derived_launch_settings': {'runner_TEMP': str(args.temp_root), 'runner_TMP': str(args.temp_root), 'WBS_TEST_WHEELHOUSE': None, 'child_TEMP_TMP': 'Maintained shell_command sets both to its reported native cwd', 'child_PSModulePath': str(inbox.parent / 'Modules'), 'native_environment_directly_observed': False, 'corroboration': 'After receipt checks reported owned-work parent equals requested temp root and exact original shell argv'},
              'meaning': 'One candidate experiment only; the two original required reference failures remain FAIL'}
    if args.phase == 'after':
        data, identity = read_bound(args.reports / 'temp-before.json')
        before = json.loads(data)
        for name in ('source_sha256', 'tool_closure', 'binaries', 'receipt_identities',
                     'owned_layout', 'environment'):
            if before[name] != record[name]:
                raise ValueError('Before/after source, tool, layout or parent context changed')
        shell_data, shell_id = read_bound(args.reports / 'shell-ps51.json')
        shell = json.loads(shell_data)
        if shell['layer'] != 'shell' or shell['source']['base_commit'] != R \
                or shell['source']['tested_path_sha256'] != source \
                or shell['source']['source_unchanged'] is not True:
            raise ValueError('Original shell runner source binding differs')
        step = shell['steps'][-1]
        work = Path(shell['owned_work_directory'])
        if work.parent != args.temp_root.resolve() or step['name'] != 'shell-0' \
                or Path(step['cwd']) != work / 'shell-0' \
                or step['timeout_seconds'] != 300:
            raise ValueError('Observed native work root or maintained bound differs')
        command, unused = runner.shell_command(inbox, work / 'shell-0',
                                               work / 'shell-0/report.json',
                                               args.tool_root, False)
        if step['command'] != command:
            raise ValueError('Original native invocation differs')
        record.update(before_receipt=identity, runner_exit=args.runner_exit,
                      shell_report=shell_id, actual_owned_work_directory=str(work),
                      actual_native_cwd=step['cwd'], native_transport=step,
                      original_shell_result=shell['success'],
                      native_exit_code=None if step['timed_out'] or 'evidence_error' in step
                      else step['exit_code'],
                      complete_original_shell_report=shell)
    write_new(args.reports / ('temp-' + args.phase + '.json'), record)
    print(json.dumps({'kind': record['kind'], 'phase': args.phase,
                      'acceptance_claimed': False, 'source_unchanged': True}))
    return 0


def sysconfig_gil_disabled():
    import sysconfig
    return sysconfig.get_config_var('Py_GIL_DISABLED')


if __name__ == '__main__':
    raise SystemExit(main())
