# V102-02 shared compatibility policy — 12 October 2026

Policy 3 prefers regular stable Windows x64 GIL CPython 3.15, then 3.14,
then 3.13. The new 3.15 compatibility/security/recommended patch is 0;
later stable patches remain accepted numerically. Existing 3.14.8 and
3.13.16 security floors, exclusions, pypdf 6.20.0 wheel hash and Calibre
9.15.0 policy are unchanged. Production validators/selection code required
no change: existing shared numeric evaluation and branch-order preference
already implement the intended behavior.

The independent primary-source review supports the 3.15.0 floor: final-tag
beta-3 NEWS carries gh-151544's Windows landmark fix; final-tag rc3 NEWS and
the official regular Windows x64 SBOM identify OpenSSL 3.5.9, which addresses
the existing reviewed HTTPS-setup issue. The stable release page confirms
regular Windows x64 artifact availability. These new exact URLs and limited
relevance statements are recorded in the shipped policy as Release3150,
Changelog315, OpenSSL3150 and Python315Sbom. Retained older advisory evidence
continues to apply to the retained older floors. This review does not turn
artifact metadata into an application/native acceptance pass.

The focused Python tests now assert the new production policy and run all
24 supplied RUNTIME_POLICY_CASES through the actual evaluator. PowerShell
tests exercise those same fixtures through its version and identity validators,
then compare every decision and message with the actual Python evaluator.
Synthetic 3.15.1/3.15.10/3.15.99 and unknown 3.16 records are fixture tests,
not installed future binaries. Unknown minors remain unsupported; malformed,
unstable, architecture/flavor and older security-floor reasons stay distinct.
Discovery tests verify preferred 3.15 over high older-branch patches and numeric
9/10 comparison. Existing precise blocked-patch/floor mutations explicitly
target 3.14 after its policy-list position changed. A malformed first-branch
blocked-recommendation fixture now blocks patch 0, preserving its guard.
Outcome's unsupported-minor fixture uses 3.16 rather than accepted 3.15.

Actual commands/results are recorded in V102-02-policy.json with source-base
commit/tree and exact tested-file hashes; runtime and user paths use descriptive
aliases in this public evidence. No application setup, diagnostic formatter or
PDF/conversion implementation changed in this task.

- Existing isolated developer CPython 3.14.8 ran
  `-I -B -m unittest discover -s tests/python -p test_compatibility.py -v`:
  46 tests passed, zero skips, exit 0, 1.967 seconds.
- Parent independently ran the same focused suite on actual isolated
  CPython 3.13.16, 3.14.8 and newly acquired 3.15.0: 46/46 each,
  zero skips, exit 0 (2.062/2.269/1.860 seconds). All observed policy hash
  `57811a77c837bd891750ea126d19c876e0ad7099793b490445d126720baf4582`.
- Parent's actual 3.15.0 full helper probe accepted policy 3 and verified
  locked pypdf 6.20.0 from that selected runtime's trusted roots, with GIL,
  isolated/no-bytecode identity as required. This proves runtime/dependency
  identity; it does not substitute for real PDF or final-ZIP acceptance.
- Native Windows 11 inbox PowerShell 5.1.26100.9444 and PowerShell 7.6.6
  each ran focused Compatibility/Discovery/Outcome Pester containers:
  67/67 passed, zero skipped/not-run/failed containers, exit 0. Pester 6.2.0
  and PSScriptAnalyzer 1.25.0 were imported from existing reviewed external
  tooling. Changed shell tests had zero parser/analyzer findings. Process
  policy was RemoteSigned; no persistent policy change was requested.
- Initial focused shell invocations each failed nine existing discovery
  assertions because the external runner supplied forward-slash TestPython
  paths instead of their canonical ProviderPath. They recorded 58/67 passes
  and nine failures. Correcting only external wrapper argument normalization
  and rerunning in fresh owned directories produced the passes above. The
  initial failed reports remain private and are not counted as passes.
- The existing source guard now includes the new consumed policy-fixture
  JSON. An actual source_manifest invocation returned its exact current hash;
  fixture mutations during tests cannot evade the normal change guard.
- `-I -B tools/codex-v1.0.2/validate_plan.py --plan docs/codex-v1.0.2`
  returned STRUCTURE_VALID/exit 0. This is reference/state validation only.

No future-patch binary execution, native final-package Windows 10/11 acceptance,
release publication or synchronization is claimed by this record. Those later
gates remain separate. Parent and independent security reviewer examined this
focused policy/test diff with no actionable findings; checkpoint commit/push
and fresh clean/live equality follow outside this evidence-writing operation.
