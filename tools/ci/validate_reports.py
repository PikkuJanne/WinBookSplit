"""Fail hosted CI when a required report is absent, failed or from another source."""
import argparse
import importlib.util
import json
from pathlib import Path
import re

MAX_REPORT = 16 * 1024 * 1024
ROOT = Path(__file__).resolve().parents[2]

def require(condition, message):
    if not condition: raise ValueError(message)

def unique(pairs):
    result = {}
    for key, value in pairs:
        require(key not in result, 'Duplicate report key')
        result[key] = value
    return result

def read(path):
    require(path.is_file() and not path.is_symlink() and 0 < path.stat().st_size <= MAX_REPORT, 'Missing or oversized report: ' + path.name)
    with path.open('rb') as stream: raw=stream.read(MAX_REPORT+1)
    require(len(raw)<=MAX_REPORT,'Oversized report: '+path.name)
    record=json.loads(raw.decode('utf-8-sig'),object_pairs_hook=unique,
                      parse_constant=lambda value: (_ for _ in ()).throw(ValueError('Nonfinite report number')))
    require(type(record) is dict,'Report must be an object')
    return record

def application_units(unit, candidate, version):
    # Reuse the same closed receipt contract used during native execution.
    path=ROOT/'tests/compatibility/application_matrix.py'
    spec=importlib.util.spec_from_file_location('wbs_ci_application_matrix',path)
    module=importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module.validate_units(unit,Path(candidate),version)

def validate(role, commit, reports, python_version=None):
    require(re.fullmatch('[0-9a-f]{40}',commit) is not None,'Full source commit required')
    require(role in ('reference','application'),'Unknown report role')
    names=['os.json','package-inputs.json']+(['python.json','shell-ps51.json','shell-ps7.json','tools-bootstrap.json'] if role=='reference' else ['application.json'])
    records={name:read(reports/name) for name in names}
    os=records['os.json']
    require(os.get('source_commit')==commit and os.get('product_type') in (2,3)
            and type(os.get('product_type')) is int and os.get('process_bits')==64
            and type(os.get('process_bits')) is int and os.get('architecture')=='64-bit'
            and all(type(os.get(key)) is str and os[key] for key in ('caption','version','build','powershell','clr')),
            'Actual Windows Server OS/source evidence required')
    package=records['package-inputs.json']
    require(package.get('success') is True and type(package.get('exit_code')) is int
            and package['exit_code']==0 and package.get('source_commit')==commit
            and package.get('working_tree_inputs_used') is False and package.get('package_built') is False,
            'Package source report failed or has another source')
    for name in names:
        if name in ('os.json','package-inputs.json'): continue
        record=records[name]
        require(record.get('success') is True,'Failed required report: '+name)
        if name=='tools-bootstrap.json':
            require(type(record.get('exit_code')) is int and record['exit_code']==0
                    and record.get('result')=='CI_TOOL_BOOTSTRAP_PASSED','Shell bootstrap failed')
            continue
        source=record.get('source',{})
        require(source.get('base_commit')==commit and source.get('source_unchanged') is True,
                'Test report source differs: '+name)
        require(record.get('synthetic_input_unchanged') is True and record.get('owned_work_directory_removed') is True,
                'Test inputs or owned cleanup unproved: '+name)
        steps=record.get('steps')
        require(type(steps) is list and steps,'Missing required test steps: '+name)
        for step in steps:
            require(type(step.get('exit_code')) is int and step['exit_code']==0
                    and (step.get('timed_out') is False if step.get('name')!='setup-test-wheel' else True),
                    'Failed test step: '+name)
        if name.startswith('shell-'):
            require(record.get('layer')=='shell','Declared shell route differs')
            require([step['name'] for step in steps]==['setup-test-wheel','shell-0'],'Required shell steps differ')
            shells=[step['evidence'] for step in steps if step['name'].startswith('shell-')]
            require(len(shells)==1,'Missing shell report')
            pester=shells[0]['pester']
            version=shells[0].get('host',{}).get('version')
            expected=r'5\.1(?:\.[0-9]+){0,2}' if name=='shell-ps51.json' else r'7\.[0-9]+\.[0-9]+(?:\.[0-9]+)?'
            require(type(version) is str and re.fullmatch(expected,version) is not None,'Declared shell host differs')
            require(type(pester.get('total')) is int and type(pester.get('passed')) is int
                    and pester['total']>0 and pester['total']==pester['passed']
                    and all(type(pester.get(key)) is int and pester[key]==0 for key in ('failed','skipped','not_run','failed_containers')),
                    'Incomplete Pester run')
            for field,values in (('syntax','errors'),('scaffold_analysis','findings')):
                rows=shells[0].get(field)
                require(type(rows) is list and rows and all(type(row) is dict and type(row.get('path')) is str
                        and row['path'] and type(row.get(values)) is list and not row[values] for row in rows),
                        'Shell syntax/analysis coverage missing or failed')
        elif name=='application.json':
            require(type(python_version) is str and re.fullmatch('3\\.(13|14)\\.[0-9]+',python_version) is not None,'Exact matrix test target required')
            require(len(steps)==1 and steps[0]['name']=='application-matrix','Application route missing')
            matrix=steps[0]['evidence']
            units=matrix['runtime_only_units']
            application_units(units,matrix['application_runtime']['candidate'],python_version)
            require(matrix.get('success') is True and matrix.get('cleanup_safe') is True
                    and matrix.get('timed_out') is False and matrix['application_runtime']['full_probe']['version']==python_version
                    and units.get('python_version')==python_version and type(units.get('tests_run')) is int and units['tests_run']>0
                    and units.get('skipped')==[] and all(type(units.get(key)) is int and units[key]==0 for key in ('failures','errors','expected_failures','unexpected_successes')),
                    'Application target/unit run differs')
            require(len(matrix['hosts'])==2 and all(type(host.get('product_type')) is int and host['product_type'] in (2,3) and host['build']==os['build'] for host in matrix['hosts']),
                    'Application native hosts differ from actual Server')
            host_versions=[host.get('shell_version') for host in matrix['hosts']]
            require(all(type(version) is str for version in host_versions)
                    and re.fullmatch(r'5\.1(?:\.[0-9]+){0,2}',host_versions[0]) is not None
                    and re.fullmatch(r'7\.[0-9]+\.[0-9]+(?:\.[0-9]+)?',host_versions[1]) is not None,
                    'Distinct inbox5.1 and PS7 application hosts required')
            expected={'manual-front-matter','auto-level1','auto-level2','late-invalid-token','no-bookmarks','zero-page','wrapper-PS5','wrapper-PS7'}
            require(len(matrix['cases'])==len(expected) and {case['id'] for case in matrix['cases']}==expected
                    and all(case['status']=='PASS' for case in matrix['cases']),
                    'Required PDF/application cases missing')
        elif name=='python.json':
            require(record.get('layer')=='python' and len(steps)==1 and steps[0]['name']=='stdlib-unittest'
                    and re.search(r'Ran [1-9][0-9]* tests in [0-9.]+s',steps[0].get('stderr',''))
                    and re.search(r'^OK\s*$',steps[0]['stderr'],re.M) and 'skipped=' not in steps[0]['stderr'],'Incomplete Python reference suite')
    return names

def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--role',choices=('reference','application'),required=True)
    parser.add_argument('--commit',required=True)
    parser.add_argument('--reports',type=Path,required=True)
    parser.add_argument('--python-version')
    args=parser.parse_args()
    try:
        names=validate(args.role,args.commit,args.reports,args.python_version)
        print('CI_REPORTS_PASSED: '+', '.join(names))
        return 0
    except (ValueError,OSError,UnicodeError,KeyError,TypeError) as error:
        print('CI_REPORTS_FAILED: '+str(error))
        return 1

if __name__=='__main__': raise SystemExit(main())
