"""Four fixed PS7 checks on an already authenticated final extraction.

The native SDK owns every child. This adapter only prepares literal recipes and
reuses the frozen kit's package, input, publication and console validators.
"""
from __future__ import annotations
import argparse
from contextlib import ExitStack
from hashlib import sha256
import importlib.util
import json
import os
from pathlib import Path
import re
import stat
import sys

IDS = ('version', 'setup-check', 'manual', 'auto2')
OS_KEYS = ('Caption','Version','BuildNumber','ProductType','processor_architectures','os_bits','process_bits')
MAXIMUM = 16 * 1024 * 1024


def need(value, code):
    if not value:
        raise ValueError(code)


def bounded(path):
    path = Path(path)
    need(path.is_absolute() and path.resolve(strict=True) == path, 'ordinary_path_required')
    for item in [*reversed(path.parents),path]:
        info = item.lstat()
        need(not stat.S_ISLNK(info.st_mode) and not getattr(info,'st_file_attributes',0) & 0x400,
             'reparse_path_rejected')
    need(path.is_file() and path.stat().st_nlink == 1 and 0 < path.stat().st_size <= MAXIMUM,
         'bounded_ordinary_file_required')
    with path.open('rb') as stream:raw = stream.read(MAXIMUM+1)
    need(0 < len(raw) <= MAXIMUM, 'bounded_json_required')
    return raw


def read(path):
    raw = bounded(path)
    def pairs(items):
        result, seen = {}, set()
        for key, value in items:
            need(key.casefold() not in seen, 'duplicate_json_key')
            seen.add(key.casefold())
            result[key] = value
        return result
    return json.loads(raw.decode('utf-8', 'strict'), object_pairs_hook=pairs,
                      parse_constant=lambda _: (_ for _ in ()).throw(ValueError('nonfinite_json')))


def load_pinned(path, row, name):
    raw = bounded(path)
    need(type(row) is dict and set(row) == {'bytes','sha256'} and type(row['bytes']) is int
         and type(row['sha256']) is str and re.fullmatch(r'[0-9a-f]{64}',row['sha256'])
         and len(raw) == row['bytes'] and 0 < len(raw) <= MAXIMUM
         and sha256(raw).hexdigest() == row['sha256'], 'helper_pin_differ')
    module = importlib.util.module_from_spec(importlib.util.spec_from_loader(name, loader=None))
    module.__file__ = str(path)
    sys.modules[name] = module
    exec(compile(raw, str(path), 'exec'), module.__dict__)
    return module


def accepted_prior(prior, bindings, bindings_digest, case_ids):
    need(prior.get('protocol') == 'winbooksplit.native-acceptance' and type(prior.get('schema_version')) is int
         and prior['schema_version'] == 1 and prior.get('success') is True and prior.get('cleanup_safe') is True
         and prior['identities']['bindings_sha256'] == bindings_digest
         and prior['identities']['production'] == bindings['production']
         and prior['os']['platform_scope'] == 'NATIVE_WINDOWS11_DRY_RUN'
         and tuple(row['id'] for row in prior['oracle']['cases']) == case_ids
         and all(row['status'] == 'PASS' for row in prior['oracle']['cases']), 'accepted_same_zip_win11_receipt_required')


def native_facts(facts, plan, prior, expected_version):
    need(set(plan['prior_os']) == set(OS_KEYS) and plan['prior_os'] == {key:prior['os'][key] for key in OS_KEYS}
         and all(facts.get(key) == value for key,value in plan['prior_os'].items()), 'same_native_workstation_required')
    need(Path(facts['shell_executable']) == Path(plan['pwsh_path'])
         and re.fullmatch(r'7\.[0-9]+\.[0-9]+', facts['powershell_version'])
         and facts['powershell_version'] == plan['expected_pwsh_version'] == expected_version
         and type(facts['process_bits']) is int and facts['process_bits'] == 64
         and type(facts['os_bits']) is int and facts['os_bits'] == 64
         and type(facts['ProductType']) is int and facts['ProductType'] == 1
         and facts['processor_architectures'] and all(type(value) is int and value == 9 for value in facts['processor_architectures'])
         and re.fullmatch(r'Microsoft Windows 11 [A-Za-z ]+',facts['Caption']) and int(facts['BuildNumber']) >= 22000
         and facts['Version'] == '10.0.'+facts['BuildNumber'] and re.fullmatch(r'[0-9]+\.[0-9]+\.[0-9]+(?:\.[0-9]+)?',facts['clr']),
         'actual_native_win11_ps7_identity_required')


def native_row(row, case, plan, smoke):
    probe = smoke.native_probe(row['probe'])
    need(row['command'] == [plan['pwsh_path'],*case['arguments']] and row['cwd'] == case['cwd']
         and row['answers'] == [] and probe['ExitCode'] == 0 and not probe['Stderr'], 'ps7_literal_native_call_differ')
    return probe


def context(args):
    need(os.name == 'nt' and sys.flags.isolated and sys.dont_write_bytecode,
         'native_isolated_driver_required')
    root = args.kit_root
    bindings_path = root / 'kit-bindings.json'
    need(re.fullmatch(r'[0-9a-f]{64}', args.bindings_sha256)
         and sha256(bounded(bindings_path)).hexdigest() == args.bindings_sha256, 'bindings_pin_differ')
    bindings = read(bindings_path)
    helper = bindings['helpers']
    windows = load_pinned(root/'engine/winbooksplit_windows.py', helper['engine/winbooksplit_windows.py'], 'wbs_ps7_windows')
    smoke = load_pinned(root/'tools/compatibility/native_smoke.py', helper['tools/compatibility/native_smoke.py'], 'wbs_ps7_oracle')
    verifier = load_pinned(root/'tools/release/frozen_assets.py', helper['tools/release/frozen_assets.py'], 'wbs_ps7_assets')
    smoke.ordinary(root, directory=True)
    need(re.fullmatch(r'[0-9a-f]{64}', args.prior_report_sha256)
         and sha256(bounded(args.prior_report)).hexdigest() == args.prior_report_sha256, 'prior_report_pin_differ')
    prior = smoke.read_json(args.prior_report)
    accepted_prior(prior, bindings, args.bindings_sha256, smoke.CASE_IDS)
    original = smoke.read_json(args.native_plan)
    need(sha256(bounded(args.native_plan)).hexdigest() == prior['oracle']['plan_sha256'], 'accepted_native_plan_differ')
    extraction = smoke.read_json(args.extraction_receipt)
    app = smoke.ordinary(Path(original['application_root']), directory=True)
    need(extraction.get('protocol') == 'winbooksplit.native-extraction' and extraction.get('success') is True
         and extraction['production'] == bindings['production'] and extraction['destination'] == str(app)
         and original['policy_sha256'] == bindings['production']['policy_sha256'], 'same_extraction_required')
    manifest, payload = verifier.verify_frozen_assets(root/'assets', bindings['production'])
    need(extraction['payload_sha256'] == manifest['payload_sha256']
         and set(extraction['files']) == set(payload), 'extraction_members_differ')
    files = []
    for name, raw in payload.items():
        path = smoke.ordinary(app/name)
        row = extraction['files'][name]
        need(row == {'bytes':len(raw),'sha256':sha256(raw).hexdigest()} and path.read_bytes() == raw,
             'extracted_source_differ')
        files.append(smoke.observation(path))
    smoke.bind_application(app, bindings['production']['policy_sha256'])
    return bindings, original, app, files, smoke, windows, prior


def prepare(args):
    validated = context(args)
    smoke, windows = validated[4:6]
    work = smoke.ordinary(args.work_root, directory=True)
    with ExitStack() as guards:
        for ancestor in [*reversed(work.parents),work]:guards.enter_context(windows.DirectoryGuard(ancestor))
        return prepare_recipes(args, validated, guards)


def prepare_recipes(args, validated, guards):
    bindings, original, app, files, smoke, windows, prior = validated
    work = smoke.ordinary(args.work_root, directory=True)
    need(not any(work.iterdir()) and not work.is_relative_to(app), 'fresh_external_work_required')
    pwsh = smoke.ordinary(args.pwsh_path)
    runtime_probe = smoke.read_json(args.runtime_probe)
    need(smoke.native_probe(runtime_probe)['ExitCode'] == 0, 'runtime_probe_failed')
    runtime = smoke.unique_json(runtime_probe['Stdout'])
    need(runtime == original['runtime'] and smoke.compatibility.validate_interpreter(runtime, smoke.RUNTIME_POLICY)['ok']
         and runtime.get('pypdf_version') == smoke.RUNTIME_POLICY['pypdf']['version'], 'runtime_changed')
    python = smoke.ordinary(Path(original['python_path']))
    need(Path(runtime['executable']) == python and python.is_relative_to(app/'.venv'), 'application_local_runtime_required')
    cases, immutable = [], [*files, smoke.observation(pwsh), smoke.observation(python), smoke.observation(Path(runtime['pypdf_path']))]
    base = ['-NoProfile','-NonInteractive','-ExecutionPolicy','RemoteSigned','-File']
    for identifier in IDS:
        cwd = work/identifier;cwd.mkdir()
        guards.enter_context(windows.DirectoryGuard(cwd))
        arguments = [*base, str(app/'WinBookSplit.ps1'), '-Version'] if identifier == 'version' else [*base,str(app/'Setup-WinBookSplit.ps1'),'-Check','-NonInteractive']
        case = {'id':identifier,'arguments':arguments,'cwd':str(cwd),'expected_exit':0,'timeout_seconds':120}
        if identifier in {'manual','auto2'}:
            name = 'simple-10.pdf' if identifier == 'manual' else 'nested-12.pdf'
            fixture = smoke.ordinary(args.kit_root/'fixtures'/name)
            with windows.FileGuard(fixture) as guard:
                fixture_bytes = bounded(fixture)
                need(sha256(fixture_bytes).hexdigest() == bindings['fixtures'][name]['sha256'], 'fixture_pin_differ')
                source = cwd/('Original Å 日本 %WBS_KIT_X% ! book.pdf')
                with source.open('xb') as stream:stream.write(fixture_bytes)
                fixture_observed = smoke.observation(fixture)
                need(fixture_observed['size_bytes'] == len(fixture_bytes)
                     and fixture_observed['sha256'] == sha256(fixture_bytes).hexdigest(), 'fixture_observation_differ')
                guard.assert_unchanged()
            neighbor = source.with_suffix('.neighbor.pdf')
            with neighbor.open('xb') as stream:stream.write(b'Authored neighbor; preserve exact bytes.\n')
            output = cwd/'out';output.mkdir();guards.enter_context(windows.DirectoryGuard(output))
            marker = output/'prior-output.pdf'
            with marker.open('xb') as stream:stream.write(b'Authored prior; preserve exact bytes.\n')
            mode = 'Manual' if identifier == 'manual' else '2'
            arguments = [*base,str(app/'WinBookSplit.ps1'),'-InputFile',str(source),'-OutputDirectory',str(output),
                         '-Mode',mode,'-PythonPath',str(python),'-NonInteractive','-NoPause']
            if identifier == 'manual':arguments += ['-StartPages','1,4,7']
            case.update(arguments=arguments,source_path=str(source),output_base=str(output),prior_members=[marker.name],
                        ranges=[[0,3],[3,6],[6,10]] if identifier == 'manual' else [[0,2],[2,3],[3,6],[6,8],[8,10],[10,12]])
            immutable += [smoke.observation(path) for path in (source,neighbor,marker)]
            immutable.append(fixture_observed)
        cases.append(case)
    return {'protocol':'winbooksplit.final-zip-ps7-plan','schema_version':1,'bindings_sha256':args.bindings_sha256,
        'production':bindings['production'],'application_root':str(app),'pwsh_path':str(pwsh),'pwsh_sha256':sha256(pwsh.read_bytes()).hexdigest(),
        'prior_report_sha256':args.prior_report_sha256,'expected_pwsh_version':args.expected_pwsh_version,
        'prior_os':{key:prior['os'][key] for key in OS_KEYS},
        'python_path':str(python),'runtime':runtime,'immutable_files':immutable,'cases':cases}


def validate(args):
    bindings, original, app, _, smoke, windows, prior = context(args)
    plan = smoke.read_json(args.plan);raw = smoke.read_json(args.process_report)
    need(plan['production'] == bindings['production'] and plan['application_root'] == str(app)
         and plan['bindings_sha256'] == args.bindings_sha256 and plan['prior_report_sha256'] == args.prior_report_sha256
         and tuple(row['id'] for row in plan['cases']) == IDS, 'ps7_plan_binding_differ')
    need(set(raw) == {'protocol','schema_version','plan_sha256','facts','runtime_after','environment_unchanged','cases'}
         and raw['protocol'] == 'winbooksplit.final-zip-ps7-process' and type(raw['schema_version']) is int and raw['schema_version'] == 1
         and raw['plan_sha256'] == sha256(args.plan.read_bytes()).hexdigest()
         and tuple(row['id'] for row in raw['cases']) == IDS, 'ps7_process_binding_differ')
    facts_probe = smoke.native_probe(raw['facts']);facts = smoke.unique_json(facts_probe['Stdout'])
    need(facts_probe['ExitCode'] == 0, 'facts_probe_failed')
    native_facts(facts, plan, prior, args.expected_pwsh_version)
    need(smoke.native_probe(raw['runtime_after'])['ExitCode'] == 0
         and smoke.unique_json(raw['runtime_after']['Stdout']) == plan['runtime'] == original['runtime']
         and raw['environment_unchanged'] is True, 'runtime_or_environment_changed')
    for item in plan['immutable_files']:smoke.unchanged(item)
    publications, results = set(), []
    with ExitStack() as stack:
        for item in plan['immutable_files']:stack.enter_context(windows.FileGuard(Path(item['path'])))
        for case, row in zip(plan['cases'],raw['cases']):
            probe = native_row(row,case,plan,smoke)
            if case['id'] == 'version':need(probe['Stdout'].strip() == 'WinBookSplit '+smoke.APP_VERSION, 'version_differ')
            elif case['id'] == 'setup-check':
                setup = smoke.fixed_frame(probe['Stdout'],'[SETUP] ')
                need(setup['code'] == 'setup_ready' and setup['status'] == 'success' and setup['exit_code'] == 0
                     and setup['changed'] is False and setup['selected_python'] == plan['python_path'], 'readonly_setup_check_differ')
            else:
                final = smoke.outcome(probe['Stdout']);frame = final['engine_result'];source = Path(case['source_path']);base = Path(case['output_base'])
                need(final['status'] == 'success' and final['code'] == 'split_complete' and final['exit_code'] == 0
                     and frame['status'] == 'success' and final['written_count'] == len(case['ranges']), 'split_outcome_differ')
                published = smoke.validate_publication(base,frame['execution'],smoke.pdf_snapshot(source),case['ranges'],False,smoke.APP_VERSION)
                smoke.publication_once(published,publications)
                logs = [Path(line[5:].strip()) for line in probe['Stdout'].splitlines() if line.startswith('Log: ')]
                need(len(logs) == 1 and logs[0].parent.parent == base, 'owned_console_path_differ')
                receipt = smoke.authenticate_console_manifest(logs[0],final,source_path=source,source_observation=smoke.observation(source),expected_python_version=plan['runtime']['version'])
                dependency = smoke.fixed_frame(logs[0].read_text(encoding='utf-8'),'[DEPENDENCY] ')['Runtime']
                need(dependency['Path'] == plan['python_path'] and dependency['Details'] == plan['runtime']
                     and smoke.unique_json(smoke.native_probe(dependency['Probe'])['Stdout']) == plan['runtime'], 'runtime_console_binding_differ')
                need(receipt['run_manifest']['runtime_versions']['powershell'] == facts['powershell_version']
                     and final['final_directory'] == str(published)
                     and {path.name for path in base.iterdir()} == set(case['prior_members'])|{published.name,logs[0].parent.name}, 'ps7_receipt_or_base_members_differ')
            results.append({'id':case['id'],'status':'PASS','exit_code':0,'probe_sha256':sha256(json.dumps(probe,sort_keys=True).encode()).hexdigest()})
        for item in plan['immutable_files']:smoke.unchanged(item)
    return safe_result(bindings,plan,facts,results,sha256(bounded(args.plan)).hexdigest(),sha256(bounded(args.process_report)).hexdigest())


def safe_result(bindings, plan, facts, results, plan_digest, process_digest):
    need(tuple(row.get('id') for row in results) == IDS and all(set(row) == {'id','status','exit_code','probe_sha256'}
         and row['status'] == 'PASS' and type(row['exit_code']) is int and row['exit_code'] == 0
         and type(row['probe_sha256']) is str and re.fullmatch(r'[0-9a-f]{64}',row['probe_sha256']) for row in results),
         'safe_case_table_differ')
    return {'protocol':'winbooksplit.final-zip-ps7','schema_version':1,'success':True,'evidence_kind':'NATIVE_AUTOMATED',
        'production':bindings['production'],'bindings_sha256':plan['bindings_sha256'],'prior_report_sha256':plan['prior_report_sha256'],
        'plan_sha256':plan_digest,'process_report_sha256':process_digest,
        'ps7':{key:facts[key] for key in ('powershell_version','clr','process_bits','os_bits','Caption','Version','BuildNumber','ProductType','processor_architectures')},
        'pwsh_sha256':plan['pwsh_sha256'],'runtime_versions':{'python':plan['runtime']['version'],'pypdf':plan['runtime']['pypdf_version']},
        'cases':results,'source_unchanged':True,'limits':['Same authenticated extraction as the inbox5.1 acceptance.',
            'This supplemental report does not satisfy native Windows10 or repeat ebook/BAT/Explorer cases.']}


def main():
    parser=argparse.ArgumentParser();parser.add_argument('command',choices=('prepare','validate'))
    for name in ('kit-root','prior-report','native-plan','extraction-receipt','output'):parser.add_argument('--'+name,type=Path,required=True)
    for name in ('bindings-sha256','prior-report-sha256','expected-pwsh-version'):parser.add_argument('--'+name,required=True)
    for name in ('pwsh-path','work-root','runtime-probe','plan','process-report'):parser.add_argument('--'+name,type=Path)
    args=parser.parse_args();result=prepare(args) if args.command == 'prepare' else validate(args)
    need(args.output.is_absolute() and not args.output.exists(), 'new_literal_report_required')
    with args.output.open('x',encoding='utf-8',newline='\n') as stream:json.dump(result,stream,ensure_ascii=True,indent=2);stream.write('\n')


if __name__ == '__main__':main()
