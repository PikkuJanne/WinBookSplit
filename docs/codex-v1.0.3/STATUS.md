# Status

State: T01 COMPLETE — baseline/reproduction only; STOP at this checkpoint.
Target: v1.0.3. Branch: `codex/v1.0.3-store-discovery`.
Source before checkpoint: `4cc6b00b99685607f3214e731edaaec7d65926df`, tree
`8a5a246d2d18b3e2b3fd746a9347f77fb95f6e08`; initially clean main matched live origin.
Baseline release source: `483baf858569bc838ac8a809d329b2595cdacf57` (v1.0.2).

Root routing is reconciled to the new owner request; only the two new handoff
directories were imported after conflict/manifest review. Historical evidence,
all production files and VERSION remain unchanged. All four captured hashes match
worktree/HEAD/tag/anonymous v1.0.2 ZIP exactly; all nine existing release assets
match their published size/digests. There is no v1.0.3 tag/release collision.

STORE-001 through STORE-006 PASS their T01 reproduction requirements. Actual
PS5.1.26100.9711 and PS7.6.5 each passed six selected synthetic baseline tests and
failed all three desired assertions, proving query exclusion, independent
downstream rejection and the misleading missing_python summary. These failures
remain RED. New fallback tests: seven pass, one expected failure; historical
private-runtime evidence cannot satisfy the new Store scenario. The checker
currently misses attempted acquisition when fallback was not used; T05 must close
that explicit RED requirement. Imported helper suite: 48 passed, one symlink skip.
See `evidence/T01.md`, `T01.json`, baseline, focused shell and fallback receipts.

No application repair/setup pass or native Windows 10 acceptance is claimed.
The owner's exact-ZIP Store gate remains NOT_RUN; publication is BLOCKED until
all real gates pass. Captured base pypdf remains UNKNOWN; the original PS5.1
nested collector failure remains COLLECTOR_FAILED, not incompatibility evidence.
This task used a Windows 11 developer host, synthetic fixtures and unchanged
production functions; no desktop Python/Manager/PATH/alias or working-document
changes occurred. No new release/tag was created and no existing release edited.

Next task: T02, official manager identity and entry points. Verify narrow alias
binding/actual supervised launch questions; do not request another broad report.
Record the actual checkpoint commit/push/live-remote comparison in the external
handover after committing these files, avoiding self-reference here.
