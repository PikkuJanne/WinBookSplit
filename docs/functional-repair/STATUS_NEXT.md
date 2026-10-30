# Functional repair — checkpoint 1

Active owner request: 30 October 2026. Scope: first-run application-local pypdf
setup with interactive consent/resume; empty-password PDF authentication;
validated benign local opening destinations. Preserve the architecture, Store
discovery fix and unrelated safety checks. The attached pack is supporting task
material, not additional publication authority.

Checkpoint 1 is RECORDED / INTENTIONALLY RED: PDF application regressions fail
at the confirmed rejection sites; missing-package classification fails on both
PowerShell hosts. Ordinary unchanged BAT ran, but the missing-pypdf scenario is
NOT_RUN because natural discovery finds Python 3.14.7 below policy. Existing
PDF-policy checks (22) and relevant harness checks (15) pass. Details and
commands: [checkpoint evidence](CHECKPOINT_1.md).
v1.0.3 and earlier releases/evidence remain immutable. No publication is planned
in this checkpoint. Obsolete blanket encryption and processing-time setup bans
are superseded only within this focused repair; current v1.0.3 behavior is not
yet repaired. The private book is unavailable and remains NOT_RUN. Stop after
this checkpoint's commit/push and clean live-branch verification; report that
receipt in the handover rather than writing a self-referential commit hash here.

Next exact task after checkpoint 1 is committed and synchronized:

```text
Continue checkpoint 2 on codex/functional-repair in PikkuJanne/WinBookSplit.
Read AGENTS.md, docs/functional-repair/STATUS_NEXT.md and CHECKPOINT_1.md, then
the repair pack's CODEX_TASK.md checkpoint 2 (ZIP in the owner's Downloads;
SHA256 in CHECKPOINT_1.md). Inspect local and live remote state.
Fix the shared PDF opener/access decision
using public pypdf empty-password authentication; handle already-authenticated
still-is_encrypted readers consistently in preview/execution/reuse/conversion.
Narrowly validate and ignore benign local OpenAction destinations in both
catalog and object traversal, warn once, and omit them from output catalogs.
Preserve bounds, source bytes, nonempty-password refusal, active-content and
all unrelated safety checks. Replace obsolete blanket-ban tests with the new
positive/negative contract. Do not change setup behavior yet, provision a
private Python, prepare the ordinary-test app environment, or publish. Run
relevant application regressions with locked pypdf 6.21.3, record actual results,
update this status/evidence, commit intended changes, push normally, verify
clean HEAD equals the live remote branch, give the checkpoint-3 handover and STOP.
```
