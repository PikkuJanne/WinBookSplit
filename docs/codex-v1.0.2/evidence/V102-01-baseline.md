# V102-01 baseline and reproduced incidents — 12 October 2026

Actual checkout is the existing `PikkuJanne/WinBookSplit` Git worktree. Clean
local main and fresh live main both identified
`8fd38e27925985b86d60ddf0950e15da41189ed6`, tree
`4e21f710e99a46f7339e91a4b6b9851b9f6e7417`. No unrelated edits, nested AGENTS,
or pending Git operations were present. Created ordinary scoped branch
`codex/v1.0.2-setup-python315`; no reset, stash, source-install change, or
historical evidence modification. Live main is not protected according to the
read-only API observation; no settings were changed.

Bundle ZIP SHA256 is
`89ff1ed4124c8f5ffd2e43772c0a648bd7121d09e9d3b26b787240a993dbc5a8`.
Extracted outside the checkout after checking member paths and duplicates.
Existing trusted regular x64 CPython 3.14.7 ran the standard-library helpers;
this is helper execution, not supported application-runtime acceptance.
`validate_bundle.py` passed: 53 manifested files, six Python files,
12 tasks/72 cases/10 gates. `import_bundle.py --repo <actual-checkout>`
preview reported 44 CREATE/no conflicts; only then `--apply` created exactly
the new docs/tools directories. Root AGENTS was deliberately reconciled,
including every active start/authority/completion paragraph. Safety rules and
old reports/tags/assets remain unchanged.

Fresh public release metadata and live tag/peeled targets are in the JSON:
v1.0.0 source `73c90d10dc16b1c321737afbdfbe959d61a24573`, v1.0.1 source
`6f2c2e5d4d5fc484bcb1530474f09957e09775ef`. Both are public non-prerelease,
with all three assets recorded by ID/name/size/API digest. This checkpoint is
an identity snapshot, not new anonymous-download byte verification.

Actual external owned `reproduce_baseline.py` (Python `-I -B`) called the
shipped evaluator with synthetic regular stable 3.15.0: rejected as
`unsupported_python`, distinct from a security-floor rejection. No PDF or
3.15 binary execution is claimed. On native Windows 11 Pro x64 build 26300,
inbox PS5.1, a copy of shipped setup with an entry marker was given
`Zone.Identifier` ZoneId=3. In the actual child `-ExecutionPolicy RemoteSigned`
context, direct invocation exited 1 with UnauthorizedAccess/not digitally
signed; neither marker nor setup result was produced. Child effective and
Process policy were observed as RemoteSigned. Parent PS5 policy scopes before
and after were unchanged. This confirms rejection before setup entered, not
a pypdf/3.15 engine failure. The child PSModulePath used inbox modules to avoid
inheriting the desktop tool's PS7 module precedence; no persistent policy was
changed.

Commands/results:

- `python -I -B -m unittest discover -s <bundle>/bundle_tests -v`: 47 tests,
  46 passed and one ERROR because this non-elevated Windows session cannot
  create symlinks (WinError 1314). This is not a full 47-test pass.
- Independent ordinary NTFS junction control through imported `safe_path`:
  PASS, rejects linked ancestor as symlink/junction/reparse. Substitute coverage
  for the existing reparse guard; no elevation or security setting change.
- `python -I -B -m unittest discover -s tests/python -p test_compatibility.py`:
  all 45 current evaluator/origin tests passed in 2.902 seconds.
- Handoff structure validation rerun after evidence/status updates: required
  before commit; terminal checkpoint supplies the actual outcome.

Reviewed root routing and evidence for safety/privacy; review identified the
missing child policy observation, which was added and rerun successfully.
AC001–006 are supported by this record and its JSON. Actual application/native
acceptance for the future candidate remains NOT_RUN.

Implementation map: retain shared Compatibility JSON/Python/PowerShell
evaluators; extend policy/tests in V102-02, real runtime/setup in V102-03,
discovery in V102-04, setup BAT in V102-05 and shared remedy/selected-reason
formatting in V102-06. Reuse unittest/Pester, existing source CI and package
allowlist, then extend `tools/compatibility` with MOTW/BAT/remedy tests and
Python 3.15 frozen-package kit. Existing per-user reviewed 3.13.16/3.14.8,
PS7 7.6.6, dev fixture tools and Calibre 9.15.0 are available. No native Windows
10 executor is accessible in this session; its old owner-launched report is
history. Prepare one self-contained redacted report kit after independent work.

Next: V102-02 primary-source security/artifact review and shared policy update.
Commit/push/live-clean equality is checked after commit to avoid self-reference.
