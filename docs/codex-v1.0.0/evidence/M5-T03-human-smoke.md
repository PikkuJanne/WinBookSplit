# M5-T03 — Current Explorer progress

**G01 PASS; AC-088/M5-T03 still BLOCKED. G02-G16 remain NOT RUN.** This is an additive human progress record, separate from the immutable canonical 38-case native report and the failed historical trials. Date: 2026-10-10, Europe/Berlin; tester: the user directly in this chat, with no separate name supplied. See [machine record](M5-T03-human-smoke.json) and [smoke matrix](M5-T03-manual-smoke.md).

The user reported: “Input & [Å_日本] manual.pdf split successfully. PASS.” The user then answered “Yes” to whether this was G01 (Level1, `1 → Y → O`), the output folder opened and the first/last PDFs showed `WBS-PAGE-001/002` and `WBS-PAGE-012`. This records user-confirmed procedure/visible markers; runtime versions, output counts and complete coverage are independently audited findings rather than additional user claims.

The user did not paste paths. A narrowly filtered recent authored-prefix lookup found one candidate publication. Its exact source path/hash, completion manifest, confirmed Level1 plan and finalized matching console outcome bind it to the prepared `automated-candidate02/app51` run. Only the identified authored publication/console directories and known candidate/source were audited. Personal absolute paths and raw logs remain local.

| Evidence | Observed result |
|---|---|
| User procedure report | G01 Level1/confirm/open, folder opening and expected first/last visible markers confirmed |
| Authored source | `Input & [Å_日本] manual.pdf`, 10560 bytes, SHA256 `b96c858519a581277a543b5bc94a14b5fa850153a763f8e8b5d700c1c41ee830`; unchanged |
| Source/candidate binding | Prepared run02/app51, 28 shipped payload files from C `a282cf4412494063b91f2a71cb21b283731569f1`; unchanged |
| Actual confirmed plan/publication | Level1 ranges `[0,2), [2,8), [8,11), [11,12)`; four PDFs with 2/6/3/1 pages; all12 physical pages covered once |
| Reopened actual output | Every page's expected text/content/boxes/rotation matches its source slice; actual file hashes match the completion manifest |
| Run/log integrity | Finalized console/run/publication ownership, exact confirmed execution and complete application-owned child-tree/stream transport; application/engine recorded exit0 |
| Logged actual runtime | PowerShell5.1.26100.9444, application-venv isolated CPython3.14.8/pypdf6.19.0; no converter for PDF; current1019 dependency files stable during the read-only audit |
| Independent review | Separate read-only G01 spot review22/22; no contradictions |

No application invocation, rerender, GUI automation, cleanup, source change or new unit/Pester/full-suite run occurred during this audit. The outer BAT native exit was **not directly observed**; a finalized application/engine outcome is not an outer process receipt. Actual Explorer/folder/visible-page behavior relies on the user's procedure confirmation. No additional readability, runtime or every-page inspection is attributed to the short “Yes.” Historical M3 human testing remains absolutely closed.

The previous evidence checkpoint `735c1d06bb59e9681ee5bbbf534d30400adf6101` was clean/live synchronized at this continuation's start. [Draft PR29](https://github.com/PikkuJanne/WinBookSplit/pull/29) remains the partial task vehicle. This progress checkpoint's own commit/push/live receipt follows externally/thread after it happens, without embedding a future/self hash. No merge/tag/release or final acceptance is claimed.

Next: **G02**, the same authored PDF dropped onto the same BAT, choices `2 → Y → N`, each followed by Enter. Expected seven physical ranges: `1-2 / 3 / 4-6 / 7-8 / 9-10 / 11 / 12`. Inspect the proposed plan before confirming, retain Log/Output paths and record representative parent/front-matter boundaries. Then continue G03-G16 using the prepared local STEPS/RESULTS. Preserve G01's source/log/output for the G16 repeat comparison.
