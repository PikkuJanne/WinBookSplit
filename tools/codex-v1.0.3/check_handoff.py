#!/usr/bin/env python3
"""Validate bundle integrity or imported handoff links; no Windows/app tests."""
from __future__ import annotations
import argparse
import hashlib
import json
from pathlib import Path
import re
import sys

def check(root: Path, mode: str = 'auto') -> dict:
    if mode == 'auto':
        mode = 'bundle' if (root/'HANDOFF_MANIFEST.json').exists() else (
            'checkout' if (root/'AGENTS.md').is_file() else 'bundle')
    if mode not in ('bundle', 'checkout'):
        raise ValueError('mode must be bundle or checkout')
    doc=root/'docs/codex-v1.0.3'
    def load(name):return json.loads((doc/name).read_text(encoding='utf-8'))
    tasks=load('TASKS.json')['tasks']; cases=load('ACCEPTANCE_CASES.json')['cases']
    ids=[x['id'] for x in tasks]; cids=[x['id'] for x in cases]
    assert len(ids)==8 and len(set(ids))==8, 'invalid task set'
    assert len(cids)==48 and len(set(cids))==48, 'invalid acceptance set'
    assigned=[]; seen=set()
    for task in tasks:
        assert set(task['depends_on']) <= seen, 'invalid task dependency order'
        seen.add(task['id'])
        assert (doc/task['brief']).is_file(), 'missing task brief'
        assigned+=task['acceptance_ids']
        for cid in task['acceptance_ids']:
            case=next(c for c in cases if c['id']==cid)
            assert case['task']==task['id'], 'case/task mismatch'
            assert cid in (doc/task['brief']).read_text(), 'brief missing case'
    assert len(assigned)==len(set(assigned)) and set(assigned)==set(cids), 'case mapping incomplete'
    incident=load('fixtures/incident.json')
    assert incident['pypdf_in_base']=='UNKNOWN', 'invented package observation'
    assert incident['ps51_python_probe']['status']=='COLLECTOR_FAILED', 'invented PS5.1 pass'
    assert len(incident['local_file_sha256'])==4, 'missing source fingerprints'
    assert all(re.fullmatch('[a-f0-9]{64}',h) for h in incident['local_file_sha256'].values())
    assert incident['python_probe']['version']=='3.15.0'
    assert len(incident['original_attempts'])==4
    fixture=load('fixtures/registration-cases.json')['cases']
    assert len(fixture)==12 and all(x['kind']=='SYNTHETIC_TEST_INPUT' for x in fixture)
    template=load('templates/native-store-report.json')
    assert template['status']=='not_run' and template['evidence_kind']=='template', 'prefilled native pass'
    for source in load('SOURCE_MAP.json')['sources']:
        if 'local' in source: assert (doc/source['local']).is_file()
    must=['docs/codex-v1.0.3/SCOPE.md','docs/codex-v1.0.3/ROOT_CAUSE.md',
          'docs/codex-v1.0.3/IMPLEMENTATION.md','docs/codex-v1.0.3/NATIVE_ACCEPTANCE.md',
          'docs/codex-v1.0.3/RELEASE_RUNBOOK.md','tools/codex-v1.0.3/check_native_store_report.py']
    if mode == 'bundle':
        must += ['START_HERE.md','CODEX_START_PROMPT.md','CODEX_CONTINUE_PROMPT.md','CODEX_RELEASE_PROMPT.md']
    else:
        must += ['AGENTS.md']
        assert 'docs/codex-v1.0.3/SCOPE.md' in (root/'AGENTS.md').read_text(encoding='utf-8'), 'inactive checkout routing'
    assert all((root/p).is_file() for p in must), 'missing required handoff member'
    # Historical evidence and other repository files are outside this import.
    scan_roots = [root] if mode == 'bundle' else [doc, root/'tools/codex-v1.0.3']
    for path in (p for base in scan_roots for p in base.rglob('*')):
        if path.is_file() and path.suffix in ('.md','.json','.txt','.py'):
            text=path.read_text(encoding='utf-8')
            # Username from the uploaded report must not escape sanitization.
            # Split literal avoids storing the sensitive token in this helper.
            assert ('jt'+'vuo') not in text, 'unredacted local account name'
    manifest=root/'HANDOFF_MANIFEST.json'
    integrity='not_checked_manifest_absent'
    if mode == 'bundle' and manifest.exists():
        members=json.loads(manifest.read_text())['files']
        listed={x['path'] for x in members}
        actual={p.relative_to(root).as_posix() for p in root.rglob('*') if p.is_file()
                and p.name!='HANDOFF_MANIFEST.json' and '__pycache__' not in p.parts}
        assert listed==actual, 'manifest member set mismatch'
        for member in members:
            b=(root/member['path']).read_bytes()
            assert len(b)==member['bytes'] and hashlib.sha256(b).hexdigest()==member['sha256'], 'handoff hash mismatch'
        integrity='passed'
    if mode == 'checkout':
        integrity='not_applicable_mutable_checkout'
    return {'status':'HANDOFF_STRUCTURALLY_VALID','mode':mode,'tasks':len(tasks),'acceptance_requirements':len(cases),
            'synthetic_registration_inputs':len(fixture),'manifest_integrity':integrity,
            'limitation':'Not application, Windows, or native execution acceptance.'}

def main():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('root',nargs='?',type=Path,default=Path(__file__).resolve().parents[2])
    p.add_argument('--mode',choices=('auto','bundle','checkout'),default='auto')
    a=p.parse_args()
    try: result=check(a.root, a.mode)
    except (AssertionError,OSError,ValueError,KeyError,StopIteration) as e:
        print(json.dumps({'status':'INVALID_HANDOFF','reason':str(e)}));return 2
    print(json.dumps(result));return 0
if __name__=='__main__':raise SystemExit(main())
