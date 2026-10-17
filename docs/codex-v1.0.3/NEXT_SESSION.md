# Next session

T01 is COMPLETE for baseline/reproduction only. STOP after its normal commit,
push and fresh live-branch equality check; the actual new commit belongs in the
external user/terminal handover, not in its own contents. The scoped branch is
`codex/v1.0.3-store-discovery`; its parent source was
`4cc6b00b99685607f3214e731edaaec7d65926df`. Do not reset to that baseline.

The next task is T02. Import and active routing reconciliation are already done.
Read `evidence/T01.md` and `T01.json`, the baseline and fallback evidence, and
the four focused Pester receipts. Production/VERSION remain v1.0.2 and unchanged.
Old history and all three releases remain immutable. Do not repeat the import
over existing destinations or request a new broad owner diagnostic report.

T01 tests: PS5.1 and PS7 each have six baseline passes and three intentional
desired-contract failures. Query-only changes still fail downstream validation;
missing_python remains misleading. Use `StoreDiscovery.Tests.ps1`'s separate
T01-Store/T01-Store-Desired tags, with DesiredStoreBehavior enabled for desired
assertions. The ordinary shell runner currently does not select this opt-in
reproduction file. Promote corrected desired contracts into maintained regression
coverage during the repair rather than retaining bug-expecting tests as policy.

`test_store_gate.py` has seven passes and one expected failure: attempted fallback
acquisition can evade the new checker when not used. T05 must disable acquisition
for the Store scenario and require/reject the attempt observation. Do not label
this known failure PASS or repurpose the old private-runtime report as new proof.

T02's narrow investigation is the official tuple/registration-to-family-alias
binding and actual packaged alias launch under the bounded supervisor. Confirm
the current supported installed-only command syntax and concrete runtime
provenance as needed without changing the owner installation or weakening general
path checks. The captured report did not establish these launch mechanisms.

The actual owner Windows 10 exact-ZIP Store gate remains NOT_RUN; publication is
BLOCKED until all required gates pass. The present host is Windows 11. Preserve
the captured pypdf UNKNOWN and original PS5.1 COLLECTOR_FAILED distinctions.
Consolidate owner-only verification into T07's tested kit when needed.

Exact continuation prompt:

```text
Continue T02 only in D:\projects\WinBookSplit-main on codex/v1.0.3-store-discovery.
Read root/nested AGENTS.md, docs/codex-v1.0.3/{SCOPE,STATUS,NEXT_SESSION,ROOT_CAUSE,IMPLEMENTATION,TESTING}.md,
TASKS.json, tasks/02-T02.md and T01 evidence. Inspect live Git/origin/worktree/tags/releases;
preserve current work and historical records. T01 already reproduced both Store
exclusions, all four incident attempts and missing_python on PS5.1/PS7; production
is unchanged. Implement T02's paired official manager identities and verified
package-family entry points without global guard relaxation. Verify real bounded
alias launch where authorized; keep unavailable native evidence explicit and
defer owner-only cases to T07. Do not install/update desktop Python or Manager,
change PATH/aliases, inject a selected runtime, or repeat broad diagnostics.
Keep the fallback-attempt validator gap scheduled for T05. After T02 update all
evidence/status/next-session records, review, stage intended paths, commit, push
normally, verify clean HEAD against freshly queried live remote, then STOP with
a verbose handover and exact T03 continuation prompt. Do not publish early.
```

Handover after each task must state:
- What changed, with paths and rationale.
- Exact tests run, host/source/artifact identities, actual passes/failures/skips.
- Whether the owner's native Store gate ran; no substituted-runtime equivalent.
- Commit/branch, push result, live remote comparison and unrelated work preserved.
- Unresolved facts/blockers, and one precise next-task continuation prompt.

Do not request Python reinstallation or another broad diagnostic collection.

A native-only pending case can be deferred explicitly to T07 while independent
implementation proceeds. Never convert DEFERRED_NATIVE to PASS or publish with
pending cases merely to satisfy task ordering.
