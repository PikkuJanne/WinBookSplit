# Active v1.0.2 handoff

Start with `SCOPE.md`, `STATUS.md`, `NEXT_SESSION.md`, `TASKS.json`, then the selected task. `ROADMAP.md` gives the sequence. `ACCEPTANCE_CASES.json` is the test ledger; all delivered cases are NOT_RUN. `RELEASE_GATES.json` names the release blockers. Specs describe behavior, not an implementation already delivered.

This task supersedes earlier release-only routing and repeated manual programmes, not safety controls. Historical docs remain historical. The new setup BAT, Python 3.15 integration, automated test extensions and release are work for Codex.

`tools/codex-v1.0.2/validate_plan.py` validates references/states only. `check_sync.py` checks live branch equality only. Neither proves compatibility or authorizes skipping release gates.
