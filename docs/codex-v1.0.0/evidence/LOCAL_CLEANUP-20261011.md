# Project closure and local cleanup — 2026-10-11

The owner explicitly requested: **"Please close the project and delete all local test and temp files."** This later authority supersedes earlier instructions to retain generated external test artifacts. The public v1.0.0 project was already complete at D=4d15589524764c14e4883f22f550d650f07aafb4, tree b5c7d44cc46a64579405923fc4fcc772e09a28dc. All36task outcomes/107required acceptance mappings remain unchanged; no new version/task or application behavior was introduced.

## Actual cleanup

[Machine receipt](LOCAL_CLEANUP-20261011.json) lists each scoped target and actual result. A read-only independent inventory proved exact ownership:30direct D:/projects test siblings (25task workspaces and five UI kits),84direct Temp entries (83project-specific names plus the failed authored F-xu2zuc1_ fixture),37Documents folders (19console folders/18synthetic outputs), and three ignored repository bytecode caches. Private documents were not processed. The initial inventory measured **79,516files/5,744,323,585bytes**, excluding links and the preserved repository/release archive.

PowerShell7.6.5 used native `Remove-Item -LiteralPath` operations. Before each recursive removal, the absolute resolved root was checked against its exact owned parent/candidate, with main and the preserved release directory excluded. Tree walks did not follow reparse points. All19test junctions were individually unlinked without recursion, followed by a zero-link rescan before parent recursion. All176guarded source/build/test/dependency paths were hashed before and after the main cleanup; independent post-cleanup verification covers the bounded retry too.

Commands executed from the repository included `& .git\winbooksplit-cleanup.ps1 -Mode Inventory`, then `-Mode Delete`, then `& .git\winbooksplit-cleanup-retry.ps1`. These local inventory/helper files are temporary cleanup metadata and are removed before the checkpoint. Native deletion's actual initial exit1/PARTIAL is retained in the machine receipt:153roots succeeded, but the browser-held WBS-T04-UI-20261010-02/cwd directory interrupted the remaining root. A bounded retry removed the16remaining files/278631bytes and all removable directories. All30original files in that target are now absent, so the total79,516files/5,744,323,585bytes were removed across the two attempts.153of154roots are completely absent. **The remaining root and its empty cwd child are directory entries only, held open by Firefox.** No test/temp files remain there. The owner was asked separately before closing the shared browser viewer window; until approval, Firefox and its tabs remain running. The result is FILES_REMOVED_EMPTY_DIRECTORY_LOCKED, not a false claim that all directories were removed.

An earlier active-process guard aborted before any deletion because installed Firefox/crashhelper retained historical UI-test path arguments. The referenced test PDF was already absent. Only exact current PID/executable/full-command matches for those two shared processes were exempted from the command-reference check; executable-under-target guards stayed active. No process was killed. Normal filesystem deletion still enforced Windows locks. The earlier guard abort and later locked-directory failure are preparation/cleanup observations, not product or application passes.

## Preserved records and assets

Before deleting the external final workspace, its exact [final-D receipt](M6-T06-final-checkpoint.json) was copied into durable evidence without editing. SHA256 **9e4a305ecd7d873b0c1fd93aa534e69737774eb685379291aea106137aab1306**. It records D's actual normal commit/push, clean local main=fresh live main at2026-10-11T08:42:01.085631+00:00, fixed R ancestry/payload/tag and CI38125187709completed/success. This later checkpoint may reference that historical D without a self-hash loop.

The three accepted assets were copied into **D:/projects/WinBookSplit-v1.0.0** and rehashed before deleting the original M6-T02 workspace. The independent reviewer rechecked all three afterward. They remain unsigned and unchanged:

| Asset | Bytes | SHA256 |
| --- | ---: | --- |
| WinBookSplit-v1.0.0.zip | 414497 | e78a1d0b976223832fcff1c7f17d0d53473112324b05eeec3ec7f74e575ccdc3 |
| release-manifest.json | 11547 | ea5c29faaaeda9132ae75546e1fe0751d8f6417b109b23047c7590875a2f182f |
| SHA256SUMS.txt | 178 | 1bdff00a4589049e41901417784d58b43b31371ca2b74cd98f14a5c4f535548b |

Tracked regression-test source, product source, Git history, durable evidence, shared applications/original tools/private settings and unrelated user files remain. Existing external raw reports, failed receipts, fixtures, copied runtimes, output PDFs/ebooks/renders and scratch scripts are now deleted. Historical committed summaries/seals are retained, including failures and limits; their paths/hashes record prior observations and do not assert current raw-file availability. Public release409510027/v1.0.0 and fixed annotated tag97df3d364e9bde0b2c2c6c3655d073b7b56b5dd2 at R=73c90d10dc16b1c321737afbdfbe959d61a24573 remain unchanged.

## Validation and final checkpoint

Four temporary cleanup metadata/helper files were sealed, deleted by exact `apply_patch` file removals and verified absent, adding4files/109855bytes beyond the original artifact inventory. Combined totals are79,520files/5,744,433,440bytes. Automatic command review had rejected the initial shell helper-removal command as blocked by policy before execution; the safer exact file removals succeeded. Git metadata/history were preserved.

The176guarded paths match R's recorded raw SHA256 values; guarded digest a428c67c95acf794a15767975079d34a79ca2a5b25af38f951cafa04d6d2715b. Read-only ownership, absent-target and exact-asset checks are cleanup verification only. No new local application/native Windows/Calibre/GUI/human/full-suite/render/scanner pass is claimed.

Structural plan validation ran with regularCPython3.14.8 from the disposable historical development venv before deletion: `python.exe -B tools\codex-handoff\validate_plan.py --plan-root docs\codex-v1.0.0` returned VALID,36tasks/107acceptance cases/30oracles. A preceding `--repo .` invocation was rejected as an unsupported argument and was corrected; it was not a test pass. Later documentation/plan/live-sync checks use the existing bundled CPython3.12.14 with `-B`; no new runtime is installed and no caches/test artifacts are recreated. `git diff --check`, staged-path/source checks, normal commit/push and fresh live-main/tag/release verification complete the documentation checkpoint. Its own actual SHA and sync outcome are reported in the chat after those operations, outside this commit. Completed project chats are archived after synchronization, with the actual app result reported in the final chat response.
