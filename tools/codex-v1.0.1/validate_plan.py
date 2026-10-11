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


def require(value, message):
    if not value:
        raise HandoffError(message)


def validate(plan_root: Path) -> dict:
    root = plan_root.resolve(strict=True)
    taskdoc, casedoc = load_json(root / 'TASKS.json'), load_json(root / 'ACCEPTANCE_CASES.json')
    require(taskdoc.get('schema_version') == 1 and casedoc.get('schema_version') == 1, 'Unknown plan schema')
    require(taskdoc.get('target_release') == casedoc.get('target_release') == 'v1.0.1', 'Wrong target release')
    tasks, cases = taskdoc.get('tasks'), casedoc.get('cases')
    require(isinstance(tasks, list) and tasks and isinstance(cases, list) and cases, 'Empty/invalid task or case list')
    tids, cids = {}, {}
    for task in tasks:
        require(isinstance(task, dict), 'Invalid task object')
        ident = task.get('id')
        require(isinstance(ident, str) and re.fullmatch(r'V101-\d{2}', ident), 'Invalid task ID')
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
            require(safe_path(root, name).is_file(), 'Task evidence missing')
        tids[ident] = task
    for case in cases:
        require(isinstance(case, dict), 'Invalid case object')
        ident = case.get('id')
        require(isinstance(ident, str) and re.fullmatch(r'V101-AC-\d{3}', ident), 'Invalid case ID')
        require(ident not in cids, 'Duplicate acceptance ID')
        require(case.get('task') in tids, 'Unknown case owner')
        require(case.get('status') in CASE_STATES, 'Invalid case state; no skipped-as-pass state')
        require(isinstance(case.get('expected'), str) and case['expected'].strip(), 'Empty expected result')
        require(isinstance(case.get('evidence'), list), 'Invalid case evidence list')
        for name in case['evidence']:
            require(safe_path(root, name).is_file(), 'Case evidence missing')
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
    ready = [i for i, t in tids.items() if t['status'] != 'DONE' and all(tids[d]['status'] == 'DONE' for d in t['depends_on'])]
    return {'status': 'STRUCTURE_VALID', 'target_release': 'v1.0.1', 'tasks': len(tids), 'cases': len(cids),
            'ready_tasks': ready, 'scope': 'Structure/references only. Does not run tests, inspect evidence truth, check GitHub or authorize a release.'}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--plan', type=Path, default=Path(__file__).resolve().parents[2] / 'docs/codex-v1.0.1')
    args = parser.parse_args()
    try:
        print(json.dumps(validate(args.plan), indent=2))
        return 0
    except (HandoffError, OSError, ValueError, TypeError, KeyError, RecursionError) as exc:
        print('Plan invalid: ' + str(exc), file=sys.stderr)
        return 2

if __name__ == '__main__':
    raise SystemExit(main())
