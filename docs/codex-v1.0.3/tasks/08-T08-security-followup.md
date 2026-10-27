# T08 - proposed minimal security follow-up after 27 October review

## Latest authority and active work

The owner ordered completion on 27 October and closed all further owner work.
OWNER_DECISION_20261027.md supersedes earlier manual-test/publication blockers in
this brief. Implement/test623, freeze and verify newbytes, complete reviewed
ordinaryintegration/tag/publication/anonymousbytes/finalmainclosure. Freshowner
Windows10Store and genuinecompany-policy scenarios remain NOT_RUN/owner-waived,
not PASS; disclose them and retain all guards. Do not request another owner run.
This task is now IN_PROGRESS; historical/proposed details below retain context.


Status: **REMEDIAL_SCOPE_REVIEWED; NOT_IMPLEMENTED/NOT_TESTED/NOT_ACCEPTED**.
This is the next separate task, not an acceptance waiver or an implemented patch.
The current checkpoint only reviews the returned packet and current gates.

The supplied return tests original R/pypdf6.20.0, not F/pypdf6.21.1. Current
upstream GHSA-x3rh-gjgm-fmqc also establishes a relevant xref-stream initialization
issue affecting F, fixed in 6.21.3. Hold further F owner runs. Preserve original
T07, R, F, archives, contracts and returned bytes. See
../evidence/T08-owner-return-20261027.md/JSON and the three independent audit files.

The concrete proposed scope is the smallest dependency/data/maintained-test
update to verified official pypdf6.21.3. The non-yanked wheel is 409554 bytes,
SHA256 c8ca4d9a360dbd5d4b9bb7a8389e2a9ae6a106bee722b496f918040ba2daa17e,
Requires-Python>=3.9, BSD-3-Clause; license bytes equal retained6.21.1.
Metadata feasibility is not runtime compatibility, tests, native acceptance or
general security clearance. Recheck official evidence before implementation.

For that separate implementation task:

1. Reproduce the relevant missing guard safely using upstream code and reduced
   limits with small positive controls; never allocate an unbounded hostile PDF.
2. Update exact requirements/policy/pins and necessary current test/tooling/docs
   only. Retain supported regular x64 Python3.13/3.14/3.15, Calibre floor, shared
   engine/discovery/setup/origin/process/file protections and earlier guards.
3. Run meaningful supported-runtime, PDF/Calibre, both-shell, setup/BAT/remedy,
   authored MOTW and exact-source CI checks with honest scopes/failures/skips.
4. Review and freeze a new committed candidate. Use guarded builders for new
   immutable assets/kit/carrier. Preserve R/F; do not alter existing archives or
   retrofit their acceptance. Validate the actual one-launch/report flow.
5. Only then request one consolidated fresh Windows10 existing-Store test for
   the new exact bytes. No interpreter/Manager install/update, private fallback,
   selected PythonPath, injected environment, prebuilt venv or policy change.

Separate STORE-024 genuine already-enforced policy evidence still remains
NOT_RUN. New dependency/native work cannot silently clear it. No release/tag/main
merge until every required real gate passes. Follow RELEASE_RUNBOOK and ordinary
commit/push/live-equality discipline. Stop after the next task checkpoint.
