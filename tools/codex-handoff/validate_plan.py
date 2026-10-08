#!/usr/bin/env python3
"""Validate handoff task references and split oracles, not application correctness."""
from __future__ import annotations
import sys
sys.dont_write_bytecode = True
import argparse
import json
from pathlib import Path
import re
from handoff_common import HandoffError, load_json, safe_path

ALLOWED_STATUSES = {"todo", "in_progress", "blocked", "done"}


def unique_rows(rows: list, label: str) -> dict:
    result = {}
    for row in rows:
        identity = row.get("id")
        if not isinstance(identity, str) or not identity or identity in result:
            raise HandoffError("Invalid/duplicate " + label + " ID")
        result[identity] = row
    return result


def validate(root: Path) -> dict:
    root = root.resolve(strict=True)
    tasks = unique_rows(load_json(root / "TASKS.json")["tasks"], "task")
    cases = unique_rows(load_json(root / "ACCEPTANCE_CASES.json")["cases"], "acceptance")
    oracle_file = load_json(root / "PLAN_ORACLES.json")
    oracle_list = oracle_file["manual"] + oracle_file["bookmarks"]
    oracles = unique_rows(oracle_list, "oracle")
    if not tasks or not cases or not oracles:
        raise HandoffError("An empty plan cannot be validated")
    visited, visiting = set(), set()
    def visit(identity):
        if identity in visiting:
            raise HandoffError("Task dependency cycle")
        if identity in visited:
            return
        visiting.add(identity)
        for dep in tasks[identity]["depends_on"]:
            if dep not in tasks:
                raise HandoffError("Unknown task dependency")
            visit(dep)
        visiting.remove(identity)
        visited.add(identity)
    claimed = set()
    for identity, t in tasks.items():
        if t["status"] not in ALLOWED_STATUSES:
            raise HandoffError("Invalid task status")
        visit(identity)
        if not safe_path(root, "tasks/" + identity + ".md").is_file():
            raise HandoffError("Missing task brief")
        for spec in t["specs"]:
            if not safe_path(root, spec).is_file():
                raise HandoffError("Missing referenced task specification")
        ids = t["acceptance_ids"]
        if not ids or len(ids) != len(set(ids)):
            raise HandoffError("Missing/duplicate task acceptance references")
        for cid in ids:
            if cid not in cases or cases[cid]["task_id"] != identity:
                raise HandoffError("Broken acceptance/task mapping")
            claimed.add(cid)
        if t["status"] == "done":
            if not re.fullmatch(r"[0-9a-f]{40}|[0-9a-f]{64}", t.get("implementation_commit") or ""):
                raise HandoffError("A done task needs an actual implementation commit")
            if not t.get("evidence_paths"):
                raise HandoffError("A done task needs evidence references")
            for evidence in t["evidence_paths"]:
                if not evidence.startswith("evidence/") or not safe_path(root, evidence).is_file():
                    raise HandoffError("Done-task evidence file missing/outside evidence directory")
    if claimed != set(cases):
        raise HandoffError("Unmapped acceptance cases")
    for case in cases.values():
        if case["task_id"] not in tasks or not case.get("procedure") or not case.get("expected"):
            raise HandoffError("Incomplete acceptance case")
        if case.get("required") is not True:
            raise HandoffError("Unexpected mandatory-case scope change; reconcile the plan explicitly")
    for o in oracles.values():
        ranges, error = o.get("expected_ranges"), o.get("expected_error")
        if (ranges is None) == (error is None):
            raise HandoffError("Oracle needs either exact ranges or an error")
        if ranges is not None:
            n = o["pages"]
            if not isinstance(n, int) or n <= 0 or not ranges:
                raise HandoffError("Invalid valid-plan oracle size")
            end = 0
            for pair in ranges:
                if len(pair) != 2 or any(type(v) is not int for v in pair):
                    raise HandoffError("Invalid oracle range types")
                start, stop = pair
                if start != end or not (0 <= start < stop <= n):
                    raise HandoffError("Oracle contains a gap, overlap or invalid bound")
                end = stop
            if end != n:
                raise HandoffError("Oracle omits trailing pages")
    return {"status": "VALID", "tasks": len(tasks), "acceptance_cases": len(cases),
            "split_oracles": len(oracles), "scope": "Structural consistency only; requirements are not test results."}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--plan-root", type=Path, default=Path(__file__).resolve().parents[2] / "docs/codex-v1.0.0")
    args = parser.parse_args()
    try:
        print(json.dumps(validate(args.plan_root), indent=2))
        return 0
    except (HandoffError, OSError, ValueError, KeyError, TypeError) as exc:
        print(json.dumps({"status": "INVALID", "reason": str(exc)}), file=sys.stderr)
        return 2

if __name__ == "__main__":
    raise SystemExit(main())
