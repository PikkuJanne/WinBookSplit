# Next session — M2-T04

Next: **M2-T04 — Implement interpreter and converter preflight**.
C `0128377099e47f51bfad7887f421b52999d522ad` is clean/live **SYNCED** at `2026-10-09T11:06:07.454548+00:00` on `codex/winbooksplit-v1-m2`. [Draft continuation PR](https://github.com/PikkuJanne/WinBookSplit/pull/13). E's own receipt is external/PR/thread;
recheck actual live state before work. PR12 was already merged at T03 start;
normal fast-forward incorporated 8a622dca. Never reset history or infer M2 acceptance.

Read AGENTS.md, STATUS.md, this file, TASKS.json, SCOPE_AND_DECISIONS.md,
tasks/M2-T04.md, specs/PROCESS_AND_PATHS.md, specs/OUTPUT_AND_CONVERSION.md,
SECURITY.md, TESTING.md and GITHUB_WORKFLOW.md. Inspect actual Git/origin/branch/
worktree, live PR/protection/CI/tags/releases. Preserve unrelated work; reconcile
normally and reuse suitable M2 branch/PR.

Final unrelated-CWD eleven-stage gate passed 182 Python tests, nine Pester tests per actual PowerShell host, and all inherited regressions. Tested 61-path raw digest: `2dedae126955a871fff7812145a7bc78d38d258aa0c8e1e3918e87d225ffb4f9`. All raw tested bytes match clean C; no clean-C full rerun is claimed. See evidence/M2-T03-conversion.md/json for actual cases/receipts/limits.
Real Calibre 9.15.0 conversion passes for EPUB (three physical pages) and genuine
AZW3 (four pages), both hosts and retention choices. Independent captured-reader
and retained-byte checks, eight invalid zero-exit controls and BAT dependency
failure (exit 1) pass. Seventeen native units and scoped owned-tree timeout/
injected-cancel controls pass; human Ctrl+C/Explorer remains untested. Fresh pinned
CPython 3.14.8/pypdf 6.19.0 venv, all 66 tools and unchanged stored policies pass.
Failed and partial attempts are retained separately.

Preserve complete immutable reader-bound plan, exact coverage/names/source binding,
flat known-file ownership and authenticated cleanup. prepare_ebook with output_base,
calibre_path, keep_converted_pdf=False and conversion_timeout=1800 removes its unique
conversion workspace before returning memory snapshot; original ebook identity is
separate and its stem supplies naming. Optional WinBookSplit_Converted.pdf is exact
captured bytes in successful run, separate manifest record/chapter count. PDF-only
schemas/historical tests/oracles/guards remain strict and unchanged.

Keep suspended Windows job assignment before resume, unbuffered bounded 64 KiB dual
drains, actual exit/context and proved-stop before cleanup. Unknown/replaced/reparse
members or unproved shutdown preserve stage with primary cause. No arbitrary PID
kill, unguarded fallback, recursive Calibre-temp adoption or sandbox claim. Existing
check/create/lock-close/crash limitations persist. Interactive preview remains M3.

T04 must reproduce/fix current interpreter/converter discovery/preflight. PS still
runs bare python; converter selection is explicit CalibrePath or three old locations,
without exact runtime/import probes. Portable Calibre 9.15 is outside those locations;
successful BAT discovery is untested. Safely validate documented explicit/venv/PATH
Python+pypdf candidates and trusted converter/version; reject broken/shadowed/wrong/
missing candidates, keep PDF Calibre-free, never auto-execute document-dir decoys.
Installation stays opt-in/project-local; no global policy/elevation/settings change.
Run AC-043/044/045 and affected eleven-stage regressions using actual hosts,
fresh venv/tools/reports and real Calibre when needed. Keep 0/1/55 until M3.

Independent review, intended-file C/E commits, normal pushes and fresh clean/live
HEAD comparison required. Stop after T04; T05 owns remaining process/encoding and
cumulative M2 review. No milestone merge/tag/interim release here.
