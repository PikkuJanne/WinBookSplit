#!/usr/bin/env python3
"""Validate handoff task/case structure only; never certifies application or release."""
from __future__ import annotations
import sys
sys.dont_write_bytecode = True
import argparse
import json
from pathlib import Path
import re
sys.path.insert(0, str(Path(__file__).resolve().parent))
from handoff_common import HandoffError, load_json, safe_path

TASK_STATES = {'NOT_STARTED', 'IN_PROGRESS', 'BLOCKED', 'DONE'}
CASE_STATES = {'NOT_RUN', 'PASS', 'FAIL', 'BLOCKED'}


def validate_evidence(root, name):
    require(isinstance(name, str) and name.startswith('evidence/') and name != 'evidence/README.md',
            'Evidence must name a new actual record under evidence/, not a template/guide')
    path = safe_path(root, name)
    require(path.is_file() and path.stat().st_size > 0, 'Evidence record missing or empty')
    if path.suffix.lower() == '.json':
        data = load_json(path)
        require(not (isinstance(data, dict) and data.get('template_not_evidence') is True),
                'Template is not executed evidence')
    elif path.suffix.lower() == '.md':
        text = path.read_text(encoding='utf-8')
        require('TEMPLATE — not executed evidence' not in text, 'Template is not executed evidence')


def validate_scenarios(root):
    doc = load_json(root / 'RUNTIME_POLICY_CASES.json')
    require(doc.get('schema_version') == 1 and doc.get('target_release') == 'v1.0.2', 'Wrong policy-fixture schema/release')
    require(doc.get('not_runtime_configuration') is True, 'Policy fixtures must not be a production policy')
    fixtures = doc.get('cases')
    require(isinstance(fixtures, list) and fixtures, 'Missing policy fixtures')
    seen = set()
    for row in fixtures:
        require(isinstance(row, dict) and isinstance(row.get('id'), str) and row['id'] not in seen, 'Duplicate/invalid policy fixture')
        seen.add(row['id'])
        require(isinstance(row.get('expected_ok'), bool) and isinstance(row.get('synthetic_future_patch'), bool), 'Invalid policy fixture booleans')
        require(row.get('expected_reason') in {'supported','python_security_floor','unsupported_python','unstable_python','malformed_runtime'}, 'Invalid expected fixture reason')
        require((row['expected_reason'] == 'supported') == row['expected_ok'], 'Contradictory policy fixture')
    require(any(r.get('input') == '3.15.0' and r['expected_ok'] for r in fixtures), 'Required Python3.15.0 case missing')
    incidents = load_json(root / 'INCIDENT_CASES.json')
    require(incidents.get('schema_version') == 1 and incidents.get('target_release') == 'v1.0.2', 'Wrong incident fixture schema/release')
    rows = incidents.get('cases')
    require(isinstance(rows, list) and rows, 'Missing incident fixtures')
    seen = set()
    for row in rows:
        require(isinstance(row, dict) and isinstance(row.get('id'), str) and row['id'] not in seen, 'Duplicate/invalid incident fixture')
        seen.add(row['id'])
        require(isinstance(row.get('given'), str) and bool(row['given']) and isinstance(row.get('expected'), str) and bool(row['expected']), 'Missing incident conditions')
        require(row.get('status') in CASE_STATES, 'Invalid incident state')
    return len(fixtures), len(rows)



def require(value, message):
    if not value:
        raise HandoffError(message)


def validate(plan_root: Path) -> dict:
    root = plan_root.resolve(strict=True)
    taskdoc, casedoc = load_json(root / 'TASKS.json'), load_json(root / 'ACCEPTANCE_CASES.json')
    require(taskdoc.get('schema_version') == 1 and casedoc.get('schema_version') == 1, 'Unknown plan schema')
    require(taskdoc.get('target_release') == casedoc.get('target_release') == 'v1.0.2', 'Wrong target release')
    tasks, cases = taskdoc.get('tasks'), casedoc.get('cases')
    require(isinstance(tasks, list) and tasks and isinstance(cases, list) and cases, 'Empty/invalid task or case list')
    tids, cids = {}, {}
    for task in tasks:
        require(isinstance(task, dict), 'Invalid task object')
        ident = task.get('id')
        require(isinstance(ident, str) and re.fullmatch(r'V102-\d{2}', ident), 'Invalid task ID')
        require(ident not in tids, 'Duplicate task ID')
        require(task.get('status') in TASK_STATES, 'Invalid task state')
        require(isinstance(task.get('depends_on'), list) and isinstance(task.get('acceptance'), list) and task['acceptance'], 'Missing dependencies/cases')
        require(len(set(task['depends_on'])) == len(task['depends_on']), 'Duplicate task dependency')
        require(len(set(task['acceptance'])) == len(task['acceptance']), 'Duplicate task acceptance')
        require(safe_path(root, task['brief']).is_file(), 'Missing task brief')
        for name in task.get('read', []):
            require(safe_path(root, name).is_file(), 'Missing task reading')
        require(isinstance(task.get('evidence'), list), 'Invalid task evidence list')
        for name in task['evidence']:
            validate_evidence(root, name)
        tids[ident] = task
    for case in cases:
        require(isinstance(case, dict), 'Invalid case object')
        ident = case.get('id')
        require(isinstance(ident, str) and re.fullmatch(r'V102-AC-\d{3}', ident), 'Invalid case ID')
        require(ident not in cids, 'Duplicate acceptance ID')
        require(case.get('task') in tids, 'Unknown case owner')
        require(case.get('status') in CASE_STATES, 'Invalid case state; no skipped-as-pass state')
        require(isinstance(case.get('expected'), str) and case['expected'].strip(), 'Empty expected result')
        require(isinstance(case.get('evidence'), list), 'Invalid case evidence list')
        for name in case['evidence']:
            validate_evidence(root, name)
        if case['status'] == 'PASS':
            require(bool(case['evidence']), 'PASS requires actual referenced evidence; template is not acceptance')
        cids[ident] = case
    owners = {}
    for ident, task in tids.items():
        for dep in task['depends_on']:
            require(dep in tids and dep != ident, 'Unknown/self dependency')
        for cid in task['acceptance']:
            require(cid in cids and cids[cid]['task'] == ident, 'Case ownership mismatch')
            require(cid not in owners, 'Acceptance mapped more than once')
            owners[cid] = ident
        if task['status'] == 'DONE':
            require(bool(task['evidence']), 'DONE requires actual evidence references')
            require(all(cids[c]['status'] == 'PASS' for c in task['acceptance']), 'DONE has an unpassed acceptance case')
            require(all(tids[d]['status'] == 'DONE' for d in task['depends_on']), 'DONE has an incomplete dependency')
    require(set(owners) == set(cids), 'Unmapped acceptance case')
    visiting, visited = set(), set()
    def visit(ident):
        require(ident not in visiting, 'Cyclic task dependencies')
        if ident in visited:
            return
        visiting.add(ident)
        for dep in tids[ident]['depends_on']:
            visit(dep)
        visiting.remove(ident)
        visited.add(ident)
    for ident in tids:
        visit(ident)
    gates = load_json(root / 'RELEASE_GATES.json')
    require(gates.get('schema_version') == 1 and gates.get('target_release') == 'v1.0.2' and gates.get('all_acceptance_cases_required') is True, 'Invalid release gates')
    require(isinstance(gates.get('gates'), list) and gates['gates'], 'Missing release gates')
    gate_ids = set()
    for gate in gates['gates']:
        require(isinstance(gate.get('id'), str) and gate['id'] not in gate_ids, 'Duplicate/invalid release gate')
        gate_ids.add(gate['id'])
        require(isinstance(gate.get('required_cases'), list) and gate['required_cases'], 'Empty release gate')
        require(all(cid in cids for cid in gate['required_cases']), 'Unknown release-gate case')
        require(gate.get('status') in CASE_STATES, 'Invalid release-gate state')
        if gate['status'] == 'PASS':
            require(all(cids[cid]['status'] == 'PASS' for cid in gate['required_cases']), 'Gate PASS before associated case PASS')
    fixture_count, incident_count = validate_scenarios(root)
    ready = [i for i, t in tids.items() if t['status'] != 'DONE' and all(tids[d]['status'] == 'DONE' for d in t['depends_on'])]
    return {'status': 'STRUCTURE_VALID', 'target_release': 'v1.0.2', 'tasks': len(tids), 'cases': len(cids),
            'ready_tasks': ready, 'release_gates': len(gate_ids), 'synthetic_policy_fixtures': fixture_count, 'incident_fixtures': incident_count, 'scope': 'Structure/references only. Does not run tests, inspect evidence truth, check GitHub or authorize a release.'}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--plan', type=Path, default=Path(__file__).resolve().parents[2] / 'docs/codex-v1.0.2')
    args = parser.parse_args()
    try:
        print(json.dumps(validate(args.plan), indent=2))
        return 0
    except (HandoffError, OSError, ValueError, TypeError, KeyError, RecursionError) as exc:
        print('Plan invalid: ' + str(exc), file=sys.stderr)
        return 2

if __name__ == '__main__':
    raise SystemExit(main())
