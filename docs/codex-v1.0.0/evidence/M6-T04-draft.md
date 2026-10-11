# M6-T04: fixed tag and inspected draft

**AC-101/102 PASS; draft only.** The annotated `v1.0.0` tag is fixed at accepted **R=73c90d10dc16b1c321737afbdfbe959d61a24573**, tree **62da4fe06b2d07374b27743dd252aab492993c35**. Direct tag object **97df3d364e9bde0b2c2c6c3655d073b7b56b5dd2** and remote peeled R agree with independent GitHub tag-object API proof. Only that named tag was pushed. Sole [draft409510027](https://github.com/PikkuJanne/WinBookSplit/releases/tag/untagged-6252f4e92fae63237bfa) is draft=true, prerelease=false, published_at=null. This authenticated draft link is not a public release URL.

Starting clean local/live main **1751f8ecfe2ddc68b6308f47dde1309f3eaeba47**, tree **c8130d3480eb3d0090a751ea1f85daf8b122fe82**, correct public origin, no ongoing merge/rebase, no open PR, no local/remote tags or releases were inspected before writing. R is its ancestor; all176 guarded Git/raw paths retain digest **a428c67c95acf794a15767975079d34a79ca2a5b25af38f951cafa04d6d2715b**. No runtime/build/dependency/test or frozen asset change occurred. This is one excluded documentation checkpoint, following the existing M6 main route.

## Actual operations and review

The [machine record](M6-T04-draft.json) retains exact manual operation commands, offset-qualified start/end times, native exits, captured stdout/stderr hashes, source/asset checks and sealed external reviewer receipts. Both streams were captured through bounded `communicate`, `shell=False`, without interpolation of book/path text. Native operations succeeded; no process stream was discarded.

| Actual operation | Outcome | Scope |
|---|---|---|
| Fresh `check_sync.py --repo .`, Git/origin/ancestry/worktree/ref and authenticated release inventory |native0/SYNCED; no prior tags/releases|Precondition only; no cached-ref or latest-HEAD tag target|
| API query R CI38097123491 and starting-main CI38121273790 |completed/success, exact commits|Actual existing runs queried; no new workflow/full-suite invocation or artifact rerun|
| Independent frozen source/ZIP review |892/892 read-only checks PASS|Exact28members/order/metadata/bytes,16runtime/threebuilder provenance, dependencies/checksums,176guarded paths; no private/development payload or common secret markers, not a comprehensive scanner|
| `git tag -a v1.0.0 73c90d10... -m 'WinBookSplit v1.0.0'`; explicit `git push origin refs/tags/v1.0.0:refs/tags/v1.0.0` |both native0|Annotated new fixed R tag; no other version/retarget/delete/force push|
| Live direct/peeled `git ls-remote` and tag ref/object API |native0; annotated object and R agree|Does not infer tag identity from release target_commitish|
| `gh release create v1.0.0` with three literal asset paths, `--draft --verify-tag --target R`, title and reviewed notes file |native0; sole draft409510027|No publication or wildcard/clobber operation|
| Authenticated `gh release download v1.0.0` with three literal patterns to a new external directory |native0; exact accepted bytes|Draft byte retrieval only; anonymous/public download and Windows smoke not run|
| Root retained/API/download audit |64/64 checks/native0|Metadata, exact body, remote tag, stream/native receipts, assets name/size/state/API digest and byte equality|
| Independent reviewed notes and live draft audit |PASS/no blocker|Truthful owner-attested/waived GUI scope, unsigned/privacy/platform/fidelity limits, exact metadata/body/assets/tag/downloaded bytes|

Audit tools: regular Python3.14.8 external development venv, Git2.56.0.windows.1, GitHub CLI2.97.0 on Windows11Pro x64; outer task shell PS7.6.5. No application acceptance is attributed to that outer shell. Prior actual application PS5.1/PS7.6.6/pypdf6.20.0/Calibre9.15/Poppler scope stays in M6-T03. Full/native/Calibre/render suites were not rerun because no product bytes changed and no defect was found.

## Exact draft assets and notes

| Asset | Asset ID | Bytes | Accepted/API/retrieved SHA-256 |
|---|---:|---:|---|
| WinBookSplit-v1.0.0.zip |630123440|414497|e78a1d0b976223832fcff1c7f17d0d53473112324b05eeec3ec7f74e575ccdc3|
| release-manifest.json |630123444|11547|ea5c29faaaeda9132ae75546e1fe0751d8f6417b109b23047c7590875a2f182f|
| SHA256SUMS.txt |630123442|178|1bdff00a4589049e41901417784d58b43b31371ca2b74cd98f14a5c4f535548b|

Exactly three uploaded assets, no missing/extra/duplicate names. Authenticated retrieved bytes equal final-R-01 originals. [Reviewed draft notes](M6-T04-release-notes.md) are7547bytes/SHA256 **a311b928a0b46034276e60e5f6edd2c84a43827cc584c82215a562d5b58cfe66**; server body equals the complete file. Only two contextual paragraphs differ from preserved M6-T02 prepared notes. Frozen README/changelog/manifest candidate labels are historical and stay byte-for-byte unchanged.

M6-T03 AC099/GateC remains accepted in recorded actual native plus owner-attested scope. F01-F07 are USER_ATTESTED_PASS; extra six-run GUI artifact audit is WAIVED_BY_USER_NOT_RUN. Draft notes explicitly disclose that distinction. No GUI logs/output totals/hashes/captured native BAT exits/run times/independent viewer proof are invented; no more human evidence is requested. Closed M3/M5 D13 and G04-G16 SUPERSEDED/native-only mappings remain intact. Workstation/scanner/renderer/UNKNOWN/historical failure/newline limits remain. Tag and package are unsigned; matching checksums are not publisher authentication.

The read-only GitHub tag-namespace inventory initially returned HTTP404 because the namespace was absent; successful exact/all `git ls-remote` supplied absence proof. That API operation is retained as a preparation observation, never a PASS. Original independent reviewer trial01 failed because PowerShell comma/operator precedence formed one expected ref string instead of two; a separately preserved parenthesized-array correction also fixes expected checksum lines. Original script/API receipts and FAILED-premise.json remain sealed; successful review02 does not relabel trial01 PASS. These are preparation/reviewer-premise observations, with no product or tag/draft defect. GitHub uses an untagged placeholder in draft URLs even though tag_name is v1.0.0; read the real public URLs anew after M6-T05 publication.

## Checkpoint and next task

M6-T04 changes only intended task/evidence/status/handoff/testing/release-record documentation. Structural validation passes36tasks/107cases/30oracles/native0 and diff checking passes/native0, with ordinary LF-to-CRLF warnings retained. Two focused documentation reviews pass/no blockers, including96traceability checks. The traceability reviewer original trial failed27/28 because its expected command-object comparison omitted the audit-added label field; separate v02 compares original fields and passes96/96. Original failed script/receipt remains preserved, with no product/release defect. Reviewed-document seals identify the snapshot before adding these review summaries; they never assert their own final commit hash. The actual commit/push/current CI/fresh clean-live receipt follows externally after it exists, avoiding self-referential hashes. R is the task's accepted implementation/tag anchor, not the documentation commit. No source rebuild or retag to bookkeeping HEAD.

Next **M6-T05 — Publish and verify anonymous downloads**, AC103/104/105. AC106/107/final-main closure belong to M6-T06. This project is IN PROGRESS; tag/draft/upload/authenticated retrieval do not complete publication.
