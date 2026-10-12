#!/usr/bin/env python3
"""Read-only, point-in-time local/GitHub branch equality check. No Git writes."""
from __future__ import annotations
import sys
sys.dont_write_bytecode = True
import argparse
from datetime import datetime, timezone
import json
from pathlib import Path
import re
from handoff_common import EXPECTED_REPO, HandoffError, git, inspect_repository


def check(repo: Path) -> tuple[int, dict]:
    report = {"repository": EXPECTED_REPO, "observed_at": datetime.now(timezone.utc).isoformat(),
              "status": "UNKNOWN", "read_only": True,
              "scope": "Branch/worktree equality only; not CI, review, tags or release validation."}
    try:
        state = inspect_repository(repo)
        report.update({"branch": state["branch"], "local_head": state["head"], "worktree_clean": not state["dirty"]})
        ref = "refs/heads/" + state["branch"]
        # Use the independently validated single push URL, never a cached origin ref.
        result = git(state["root"], ["ls-remote", "--exit-code", state["push_url"], ref])
        if result.returncode == 2:
            report.update(status="UNSYNCED", reason="The live remote branch is absent")
            return 1, report
        if result.returncode:
            raise HandoffError("Live remote check failed (network/authentication); cached refs are not evidence")
        rows = [line.split() for line in result.stdout.splitlines() if line.strip()]
        matches = [row[0] for row in rows if len(row) == 2 and row[1] == ref and re.fullmatch(r"[0-9a-f]{40}|[0-9a-f]{64}", row[0])]
        if len(rows) != 1 or len(matches) != 1:
            raise HandoffError("Unexpected live ref response; branch equality is unknown")
        report["remote_head"] = matches[0]
        # Re-read local state so changes during the network call don't certify stale HEAD/cleanliness.
        after = inspect_repository(repo)
        if (after["head"], after["branch"], after["dirty"], after["push_url"]) != (
            state["head"], state["branch"], state["dirty"], state["push_url"]
        ):
            raise HandoffError("Local state changed during verification; run again")
        if state["dirty"] or state["head"] != matches[0]:
            report.update(status="UNSYNCED", reason="Dirty worktree and/or local HEAD differs from live branch")
            return 1, report
        report.update(status="SYNCED", reason="Clean local HEAD equals fresh live branch at observation time")
        return 0, report
    except (HandoffError, OSError, ValueError) as exc:
        report.update(status="UNKNOWN", reason=str(exc))
        return 2, report


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repo", type=Path, default=Path.cwd())
    args = parser.parse_args()
    code, report = check(args.repo)
    print(json.dumps(report, ensure_ascii=True, indent=2))
    return code

if __name__ == "__main__":
    raise SystemExit(main())
