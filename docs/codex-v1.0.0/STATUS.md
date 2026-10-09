# Project status

Project: WinBookSplit first public v1.0.0; IN PROGRESS.
M0/M1/M2 reviewed/merged; M3 is complete through **M3-T01** only.
Next action: **M3-T02 — Add the small non-interactive CLI**.
Branch/upstream: codex/winbooksplit-v1-m3 / origin/codex/winbooksplit-v1-m3.
Origin: https://github.com/PikkuJanne/WinBookSplit.git.
C `b8c6ed0f9e1153e20ba9b265615c6c3cab8d1a7e` is clean/live **SYNCED** at `2026-10-09T14:49:29.758828+00:00`.
E/continuation draft PR receive their actual external/thread receipts after push/creation;
recheck live state rather than treating this historical C receipt as a reset target.
No public release; fresh observed tags/releases remain empty.

## Current-build human checks — complete

[Human evidence](evidence/M3-T01-human.md) and [machine record](evidence/M3-T01-human.json)
record the user-operated Explorer/visual/repeat/fallback/error checks on source
E `9983191ed3105d0edfec35be19888190d5acca33`. These use synthetic fixtures and
controlled dependency/output defaults; they are not final-package acceptance.
Four physical Ctrl+C controls passed under PS5.1/PS7. The fresh BAT engine and
conversion retries both passed with native/final cancellation 130 and N then Enter
at CMD's termination prompt. The saved operator receipt and this thread's user
confirmation close the current-build human testing. Preserve the original BAT
native-255 failure and unlaunched conversion case as historical evidence.
No shipped runtime changes were needed. Exact-package acceptance remains a later
release gate, separate from these completed current-build observations.

PR16 is now merged at main `981894db7a91d7c2d3f0e3af258323ddeef38106`
(`2026-10-09T15:03:08Z`). The M3 branch was normally fast-forwarded to that merge
before this documentation-only supplement. Its own push receipt follows externally.

## M3-T01 evidence and limits

[Outcome evidence](evidence/M3-T01-outcomes.md) and [machine record](evidence/M3-T01-outcomes.json)
satisfy AC-050/051/052. One shipped JSON exit map now serves Python/PowerShell; final
fallback, explicit cancellation, native failures and BAT status agree. Final success
waits for log flush/close; secondary errors preserve primary failures, and completed
publication is retained with incomplete 6 when finalization fails. Stdout OUTCOME is
authoritative; the log's OPERATION-OUTCOME is provisional. UTF-8 preserves BAT Unicode
paths; redirected PS7 skips cosmetic clearing. During supervision a rooted native
Ctrl+C/Break handler preserves the enclosing pipeline and owns only its job shutdown.

Frozen unrelated-CWD fourteen-stage full gate passed 217 Python tests,
49 Pester tests per actual PS5.1.26100.9444/PS7.6.5 host, all inherited regressions,
real Calibre 9.15.0 EPUB/genuine AZW3, 26 native plus 9 process applications and 51 outcome
applications. All 85 tested raw paths match materialized clean C; digest
`ac6c5b910891ef8a86f078db43b2677f66cd515022dd17251ed10bbc9b3d87a0`. No clean-C full rerun is claimed. Zero skips/syntax/new scaffold
findings; 51 legacy application findings per host remain separate. Fresh pinned dev
venv/imports/pip check, 66 shell-tool hashes, stored policies, source/neighbor/prior
identity/content and known final suite cleanup pass. Independent source/strict receipt/
raw acceptance/public evidence review passed; no submitted GitHub approval or CI pass.

Six actual OS signals, token cancellation and engine/converter deadlines prove owned
PID/tree shutdown and unrelated Python survival. Any automated CMD prompt decision
is recorded separately; human Ctrl+C/Explorer were NOT RUN at that automated checkpoint.
The later partial human supplement above records the actual subsequent checks. Earlier failed/unknown receipts
remain honest: lost fifth partial receipt, retained sixth/seventh unsafe workspaces,
Unicode and redirected-header fixes, rejected-report labeling and inherited process
footer-oracle fixes. An external capture collision left an unknown authored console
log location; no broad scan/adoption/deletion was attempted. Do not retry old cleanup.
Outer interruption retains marked staging without an engine ledger. Existing same-
account race/crash limits are not a sandbox. Packaging must include Outcomes.json
and Process.ps1. M3-T02 version/CLI, M3-T03 menu/preview UX, remaining human checks, broader
fidelity, CI, package and public release gates remain open. No tag/interim release.

PR15 was already merged at thread start; clean main 85916466 was reconciled normally
before the new M3 branch. Earlier sections below are historical observations.

## Historical checkpoint before M3-T01

Project: WinBookSplit first public v1.0.0; IN PROGRESS.
M0/M1 reviewed/merged; cumulative M2 review and acceptance passed through **M2-T05**.
Next dependency-ready task: **M3-T01 — Propagate outcomes and cancel safely**.
Branch/upstream: codex/winbooksplit-v1-m2 / origin/codex/winbooksplit-v1-m2.
Origin: https://github.com/PikkuJanne/WinBookSplit.git.
C `877dd4bd7189d10d3dd9bbfd7efee8d42e631bae` is clean/live **SYNCED** at `2026-10-09T13:15:01.396075+00:00`.
[Cumulative M2 continuation PR](https://github.com/PikkuJanne/WinBookSplit/pull/15); E and the delegated merge receive their
own external/PR/thread receipts after their actual push/merge. Recheck live state.
No public release; fresh observed tags/releases are empty.

## M2-T05 evidence and limits

[Process evidence](evidence/M2-T05-process.md) and [machine record](evidence/M2-T05-process.json)
satisfy AC-046/047/048/049. The narrow shipped PowerShell helper starts a direct
executable suspended, assigns its own non-breakaway Windows job before resume,
inherits only three selected handles, and drains strict UTF-8 stdout/stderr concurrently.
Bounded 64 KiB human tails preserve blank lines/no-newline tails; independent bounded
machine records preserve results beyond those tails. Parent, descendant and pipe EOF
proofs bound timeout/cancellation. Preflight uses the same owned-tree transport.
Isolated Python explicitly uses -I -B -X utf8. BAT remains byte-identical.

Final unrelated-CWD thirteen-stage full gate passed 192 Python tests, 40 Pester tests
per actual supported host, and all ten inherited/current regression stages including
real Calibre 9.15.0 EPUB/genuine AZW3 conversion. Tested 76-path raw digest:
`7ae333509eaa80abf92529e5ed720c65e7a3cf307d92bb0c18e4fc0c43e689e7`. All raw tested bytes match clean C; no clean-C full rerun is claimed.
Process acceptance has 26 actual native controls and nine actual PS5.1/PS7/BAT
applications: bounded dual-stream floods, fast nonzero exit/tails, Unicode boundaries,
literal hostile arguments, independent large result records and protocol failures,
owned descendants/pipe inheritance on timeout/token cancellation, unrelated process
survival, and complete three-page synthetic PDF splits/content/manifests/source/neighbor
checks. Ordinary and deeply nested targeted runners pass the final destination-aware
filename oracle and independently identified fixture process cleanup.

Fresh hash-pinned CPython 3.14.8/pypdf 6.19.0 developer venv, all 66 tool hashes,
actual PS5.1.26100.9444/PS7.6.5 and stored-policy comparisons passed. Zero skipped
tests/new scaffold static findings; 43 legacy application analyzer observations per
host remain separately recorded. Earlier failed/unknown attempts and the full01
test-oracle failure remain separate honest receipts. Independent cumulative M2
source/strict receipt/public evidence review passed. No submitted GitHub approval
or CI pass is claimed; fresh live main was unprotected with no required checks/runs.

PR14 was already merged at 12:32:34Z (main c48a9dfe); ordinary fetch/fast-forward
reconciled it before T05. Existing implementation commits/history were preserved.
The prior parent-only preflight descendant-proof gap is now closed. Original engine
drains already avoided deadlock; reproduced defects were unbounded logs, blank-line
loss and unbounded inherited-pipe waits, not a fabricated baseline deadlock.
Transport owns no filesystem cleanup; interrupted engine staging is explicitly retained
because the console has no engine file ledger. Current compatible 0/1/55 codes remain
for M3's final outcome/fallback mapping. Token cancellation is tested; human Ctrl+C,
Explorer drag/drop, PDF feature/fidelity, expanded ebook/DRM/remote cases, extracted
package, CI and public release gates remain open. Packaging must include the new helper.
No tag/interim release. Earlier sections below are historical observations.

## Historical checkpoint before M2-T05

Project: WinBookSplit first public v1.0.0; IN PROGRESS.
M0/M1 reviewed/merged; cumulative M2 remains open. Completed through **M2-T04**.
Next dependency-ready task: **M2-T05 — Supervise subprocesses and UTF-8 streams**.
Branch/upstream: codex/winbooksplit-v1-m2 / origin/codex/winbooksplit-v1-m2.
Origin: https://github.com/PikkuJanne/WinBookSplit.git.
C `e9bcef45d826a5f94679c9e24a0cc67dd6017e43` is clean/live **SYNCED** at `2026-10-09T12:19:12.811580+00:00` on `codex/winbooksplit-v1-m2`.
[Draft continuation PR](https://github.com/PikkuJanne/WinBookSplit/pull/14); E receives its own external/PR/thread receipt.
No public release; fresh tags/releases empty at this task's observed audit.

## M2-T04 evidence and limits

[Runtime evidence](evidence/M2-T04-runtime.md) and [machine record](evidence/M2-T04-runtime.json)
satisfy AC-043/044/045. Explicit/application venv/read-only py listing/trustedPATH
select exact regular Windows x64 CPython 3.14.8/pypdf 6.19.0 and isolated -I -B engine.
Missing/broken/shadowed/wrong dependencies reject before console/conversion/output;
precise selected paths/version/imports/setup are reported. PDF never probes Calibre;
ebooks validate explicit/trustedPATH/known-location9.15.0 including portableinstall.
No auto-install/registration/elevation/globalpolicy/PATH/settings change.

Final unrelated-CWD twelve-stage gate passed 186 Python tests, 28 Pester tests per actual supported host and all nine inherited/current regression stages. Tested 67-path raw digest: `8ea3467230d88e2438b2daa7da6b61ea79465ad3e177e083a57d1a00e4f41e66`. All tested raw bytes match clean C; no clean-C full rerun is claimed.
Runtime 33 actual cases pass:16 early refusals,17 successes/51 chapter PDFs, five real
portable EPUB successes including unchanged BAT and two authored listing controls.
Complete physical page/content/manifest/source/neighbor/decoy/policy/known cleanup
checks pass. Fresh pinned venv/imports and all 66 tool hashes pass. Focused19 runtime
Pester/host and four synthetic validator tests have honest separate scope. Failed
full/focused/development attempts and corrected defects remain separate receipts.

PR13 was already merged at 11:23:44Z at main/merge c7af98e8; normal fast-forward
reconciled main/M2 without reset or milestone acceptance inference. Newdraft continues
T04 only; no merge/tag/release. Earlier sections below are historical receipts.
Preflight bounds parent/EOF/streams but DescendantsStopped=null; conversion owned job
supervision is unchanged. Authored listing is not actual manager registration;
wrong pypdf declaration is an ownedcopy. Trusted probe is not authentication/sandbox.
Two earlier partial sample workspaces remain external after rejectedcleanup; application
outputs and final suite known cleanup pass separately. M2-T05 owns remaining engine
supervision/UTF-8/cumulative review; M3 version/CLI/preview and remaining fidelity/
Explorer/PDF features/CI/package/publicrelease gates remain open. 0/1/55 persists.

## M2-T03 evidence and limits

[Conversion evidence](evidence/M2-T03-conversion.md) and [machine record](evidence/M2-T03-conversion.json)
satisfy AC-040/041/042. Original ebooks and same-name PDFs stay unchanged; parent-owned
conversion PDF is validated, captured and cleaned before shared immutable preparation
returns. Optional exact full PDF has separate manifest metadata in successful output
only. Suspended job assignment, concurrent bounded drains, actual status/context and
proved-stop cleanup preserve failures. Unknown/replaced objects/unproved stop retain
stage explicitly. PDF schemas/historical guards and complete page coverage remain.

Final unrelated-CWD eleven-stage gate passed 182 Python tests, nine Pester tests per actual PowerShell host, and all inherited regressions. Tested 61-path raw digest: `2dedae126955a871fff7812145a7bc78d38d258aa0c8e1e3918e87d225ffb4f9`. All raw tested bytes match clean C; no clean-C full rerun is claimed.
Eight real Calibre 9.15.0 EPUB/genuine AZW3 host/retention launches, two independent
captured-reader controls, eight invalid native-zero cases and BAT dependency failure
pass. All three EPUB/four AZW3 physical pages, content/markers, identities, manifests, retained
bytes, known cleanup/policies pass. Seventeen focused units and separate owned-tree
timeout/cancel controls pass. Fresh pinned venv/imports and all 66 tool hashes pass.
Earlier failed/partial attempts have separate honest receipts. Independent review
passed; no GitHub approval/CI/human Explorer claim.

PR12 was already merged at 10:19:22Z with main/merge 8a622dca; normal fast-forward
reconciled main/M2. No milestone acceptance inferred, no merge/tag/release here.
Earlier sections retain historical checkpoint observations. Runtime/discovery,
remaining process/encoding, preview/CLI, rendered fidelity/features, human Explorer,
expanded/remote/DRM cases, CI/package and release remain open. No sandbox/crash or
arbitrary long-path/UNC claim; current 0/1/55 persists until M3. Next is M2-T04 only.

## M2-T02 evidence and limits

[Path evidence](evidence/M2-T02-paths.md) and
[machine record](evidence/M2-T02-paths.json) satisfy AC-036 through AC-039.
Literal FileSystem readable-file validation rejects directory/provider inputs,
fixes wildcard metadata and accepts observed read-only synthetic inputs. PS accepts
an existing literal output base. Native exclusive console reservation and caught
marker/log setup preserve meaningful failures. Existing BAT/process drains and
ownership cleanup remain unchanged.

Every mode shares destination-aware safe Unicode/fallback filenames and widened
numbering. Bound preparation freezes shortened titles/run stem before preview;
execution rejects a changed base, consumes exact names and retains source binding.
Unbound plans never silently rename at execution. Conservative complete-path limits
are 259/247/255 UTF-16 units for files/created directories/components.

Final unrelated-CWD ten-stage gate passed 162 Python tests, nine Pester tests
per actual supported host, all inherited mode/plan/diagnostic/output checks, three
actual literal PS5.1/PS7/BAT successes, eight actual rejection paths, nine title
cases, 120 real outputs per mode, a 259-unit final path and five destination failures.
Direct API no-replan/reopen traps, immutable preview, source/content/manifest/neighbor
checks and authenticated cleanup pass. 14 focused filename units pass; separate
controlled console-init before/after and native reservation receipts are accurately
scoped. Fresh pinned venv/imports and all 66 shell-tool hashes pass. Stored policies
remain unchanged in each host's actual path-stage comparison.
Tested 55-path digest: `9da651cc0c13c961cf9e0a0d96d939792771ced264a7f37c0e0107636123f5d9`.
All final raw tested bytes match clean C; no clean-C full rerun is claimed.
Independent source/harness/raw/public evidence review passed; no CI/GitHub approval
is claimed. Earlier failures/passing snapshots remain separate historical receipts.

At start, PR11 was already live MERGED at 09:08:41Z, main/merge 8a21e49. Normal
fetch/fast-forward reconciled M2 and local main, then normal M2 push; no reset or
M2 milestone-acceptance inference. PR11 and earlier status text below are historical.

Actual sharing denials/ENOSPC injection are not ACL/full-disk claims. Existing
same-account check/create/lock-close intervals are not a sandbox. Human Explorer,
UNC/arbitrary long paths, real Calibre, remaining discovery/process/CLI work,
PDF fidelity, package, CI and public release remain open. No merge/tag/release.

## M2-T01 evidence and limits

[Output evidence](evidence/M2-T01-output.md) and
[machine record](evidence/M2-T01-output.json) satisfy AC-032 through AC-035.
Every mode executes its existing immutable source-bound plan into a unique owned
stage. Exclusive files are registered immediately, reopened and checked for
positive expected page counts, exact membership, sizes and hashes. A complete
manifest and Windows same-volume no-replace rename publish a unique final child.
Repeat/concurrent runs preserve earlier runs. The output argument is an existing
base; PS uses actual Documents and prints explicit final/log paths. Console logs
and bounded failed diagnostics have separate marked directories. BAT is unchanged.

Cleanup authenticates held directory/file identities, marker and exact flat
members. It removes only this run's known owned objects; unknown, replaced or
reparse members are retained with an explicit error. Edited manifest paths do
not drive deletion. Post-publication close failures preserve complete success
with a warning; pre-publication failures remain nonzero and never announce Done.

Final unrelated-CWD nine-stage gate passed 146 Python tests, nine Pester tests
per actual PS5.1.26100.9444/PS7.6.5 host, and inherited manual/bookmark/plan/
diagnostic regressions. Output checks cover four repeat/same-basename CLI runs,
four simultaneous identical-clock process runs, five failure injections and
eight ownership controls including actual junctions. There are 26 focused output
units, six actual success launchers (three concurrent), and 12 additional actual
controlled PS5.1/PS7/BAT failure probes. Sources/neighbors/prior output and owned
cleanup were checked. Fresh hash-pinned CPython3.14.8/pypdf6.19.0 dev venv and
all66 saved shell-tool hashes passed; stored execution policies are unchanged.
Tested51-path digest: `eb9c4dee60c60e022bd5fafbacb1acad6894a2312f7240bfe3fd7932d1bb1e10`.
Every final tested raw path byte matched clean C; no clean-C full rerun is claimed.
Independent source/harness/native/raw/public evidence review passed;
no submitted GitHub approval or CI pass is claimed.

Compatible FILE_WRITE_ATTRIBUTES can mutate an empty directory despite sharing
locks. Already observed mutation rejects before owner creation; check/create is
not atomic. Windows requires closing child locks before final check/rename.
These intervals do not sandbox an arbitrary same-account process. Cleanup never
follows the target. Earlier probe snapshots and their defects are distinct from
final source acceptance. No crash/power-loss durability is promised.

Explorer/full interactive UX, early ordinary discovery and conversion paths,
real Calibre conversion, PDF fidelity/features, release package, CI and public
v1.0.0 remain open. M2-T02 owns full literal-path/filename handling; M2-T03/T04/T05
own conversion, runtime and remaining process/encoding work. 0/1/55 compatibility
codes remain until M3's CLI surface. No tag/release or milestone merge here.
M0/M1 sections below retain their historical observations and merge receipts.

## M1-T06 evidence and limits

[Diagnostic evidence](evidence/M1-T06-diagnostics.md) and
[machine record](evidence/M1-T06-diagnostics.json) satisfy AC-030/031 and record
the cumulative M1 review/merge. Distinct structured results replace the hidden
sentinel. The real handler offers only applicable explicit retries, rejects bad
protocol/success records and preserves failed/cancelled exits through PS/BAT.
Native quoting and both drains preserve literal data, stderr and the final result.

Final unrelated-CWD eight-stage gate passed 118 Python tests, nine Pester
tests per actual supported host, and all inherited manual/bookmark/shared-plan
regressions. New diagnostics cover 23 engine records, three real writer controls,
21 category records × 11 choices and 18 malformed-protocol rejections per host,
literal native arguments/Unicode and 200K dual streams. Twelve controlled actual
PS5.1/PS7/BAT failure probes passed after the final UTF-8 source change.
Sources/neighbors/prior output/owned cleanup were checked. Earlier failures and
redirected stderr/prompt limitations remain in evidence; no human/Explorer,
Calibre, CI, full ordinary discovery or release-package pass is claimed.

At the M1-T06 checkpoint output failure could still leave partial files;
M2-T01 now provides staging, validation, publication and owned cleanup as above. Current 0/1/55 compatibility codes remain
until M3 final CLI binding. Earlier evidence below is historical, including
earlier incomplete-milestone/PR states; fresh live state is authoritative.

## M1-T05 evidence and limits

[Shared plan evidence](evidence/M1-T05-plan.md) and
[machine record](evidence/M1-T05-plan.json) satisfy AC-027 through AC-029.
Existing logical mode planners stay intact. One validator freezes complete
ordered coverage/entries/ranges/metadata. Reader-bound `prepare_split`, read-only
`preview_plan` and `execute_split` share the same plan without recalculation.
Changed/repointed/deleted paths safely retain the original captured reader;
ordinary plan/reader replacement and stream mutation reject before writing.
Source/output alias and existing files reject before slices; exclusive creation
refuses later collisions. Full transactions/rollback/publication remain later.

Final unrelated-CWD seven-stage gate passed 105 Python tests, nine Pester tests
per actual PS5.1.26100.9444/PS7.6.5, zero skips/syntax/scaffold findings.
New plan evidence: nine structural/two preview/three actual writer modes/two
source aggregates plus deletion, exact entries/filenames/ranges/IDs/content,
300 seeded plans (100/mode), and 18 stable-source units. Twenty runner units
reject incomplete promised evidence. Manual/Level 1/Level 2 regressions preserve
22/4/4 targets, 8/5/6 aggregates, 250/150/150 plans and 25/10/10 writer samples.
Six controlled launchers include three concurrent; both PS7 BM-03 probes retain
corrected six-section complete bytes/IDs. Sources/neighbors/decoys/shared TEMP
are intact; only owned outputs/temp are removed.

Fresh supported hash-pinned dev venv and all 66 external shell-tool file hashes
passed; stored policies unchanged. Tested digest `73f7c3848cf35dcd3444cdd5af861231e81efe775d9afff298eac22abc2defae`;
all 41 tested actual bytes matched clean C, without claiming a clean-C full rerun.
Pre-change coverage/stale-plan/alias/overwrite probes, failed setup capture,
earlier unit source-consistency failure and pre-final-guard passing full receipt
remain distinct from final acceptance. Original guards/fixtures/oracles/older
evidence and launchers are unchanged. Independent source/harness/raw/public
review found no blocker; no GitHub submitted review or CI pass is claimed.

Seven pypdf NullObject warnings, PS5.1 CLIXML progress, PS7 launcher RawUI stderr
and 60 legacy analyzer observations per host remain recorded limits. Ordinary
discovery/all launcher/error paths, corrected Level 2 PS5.1/BAT/Level 1 console,
interactive preview/fallback/Explorer, Calibre, PDF feature/fidelity, output
transactions, extracted package, CI and release/download gates remain open.
No merge/tag/release created here. Prior sections retain historical scope.

## M1-T04 evidence and limits

[Level 2 evidence](evidence/M1-T04-level2.md) and
[machine record](evidence/M1-T04-level2.json) satisfy AC-023 through AC-026.
`plan_level2` reuses normalized Level 1 intervals and selects only each retained
parent's own direct children. It filters outside children before first-source
deduplication/sorting, keeps first-parent subtrees, warns on ignored alias/invalid
ancestors and retains deeper lineage without creating boundaries. Front matter,
parent openings and intact no-child fallbacks produce complete ordered coverage;
no child crosses its own parent end. At least one usable retained direct child
is required; otherwise no_bookmarks_at_level/sentinel/55 rejects before writing.
Empty/no usable parents stay distinct; malformed/zero inputs fail. Shared
coverage validation/source binding is M1-T05; full interactive fallback is M3.

Pre-fix evidence distinguishes four actual PDF CLI cases from 11 callable mocks.
Final unrelated-CWD six-stage full gate passed 85 Python tests, nine Pester tests
per actual PS5.1.26100.9444/PS7.6.5 host, zero skips/syntax/scaffold findings.
Level 2: four corrected CLI targets/six hierarchy aggregates/150 seeded plans/
ten actual writer samples, checking every slice/flattened page identity and
selected-parent ownership. Manual: 22 targets/eight extras/250 plans/25 writers;
Level 1: four targets/five aggregates/150 plans/ten writers. Additional real-PDF
unit cases preserve unusable ancestor lineage, ignore orphan subtrees and reject
alias/grandchild-only Level 2. Seven pypdf NullObject unit warnings are retained.

Current manual historical bookmark comparisons are empty. Its corrected six-
section BM-03 CLI reference supplies the unchanged owned launcher helper. Six
actual launchers include three concurrent; both PS7 Level 2 probes match the
corrected complete PDF bytes/IDs, while other probes cover manual. Inputs,
neighbors/decoys/shared TEMP remain intact and only owned GUID/marker Documents
outputs are removed. Both PS7 probes retain inherited RawUI stderr; 60 legacy
analyzer observations per host remain limitations. Actual corrected Level 2
PS5.1/batch and corrected Level 1 console paths remain unclaimed.

Fresh supported regular GIL CPython 3.14.8 x64/pypdf 6.19.0 dev venv passed
hash-pinned install, exact versions/import origins and pip check. All 66 shell-
tool files matched prior hashes; stored policies stayed unchanged. Tested digest:
`fca9a3b30142da14f0403a971b9d5ee02ee7e81d0cf3899ed0cc28f7609c1003`.
All 38 tested raw paths matched clean C; no clean-C full rerun is claimed.
Original/extraction guards, fixture generator/oracles and older evidence remain
unchanged. Independent source/harness/raw-evidence/privacy review found no
remaining M1-T04 blocker; no submitted GitHub review or CI pass is claimed.

Shared plan invariants/preview/source parity, all launcher/discovery/error/output
transactions, interactive fallback/Explorer, Calibre conversion, rendering,
extracted package, CI (M4-T05) and release/download gates remain open. No merge/
tag/release was created here. Prior sections retain their historical scope.

## M1-T03 evidence and limits

[Level 1 evidence](evidence/M1-T03-level1.md) and
[machine record](evidence/M1-T03-level1.json) satisfy AC-019 through AC-022.
The engine preserves source order/depth/lineage, warns and skips recoverable
invalid/external destinations, keeps first-source duplicates and warns on
physical reordering. Front matter plus parent intervals cover every physical
page. Raw internal page references/fit validation prevents pypdf's number-as-
object-ID and unknown-fit fallback errors. Raw outline/name trees and iterative
normalization reject unreadable/cyclic/reused/over-limit structures before writes:
depth 64, 10,000 raw nodes per tree/named definitions, and 10,000 normalized
entries plus child-list containers. No-plan Level 1 keeps sentinel/exit 55;
malformed/zero-page inputs fail. Valid named/indirect destinations are covered.

Pre-fix evidence distinguishes three actual generated PDF writer cases from
16 callable mocks. Final unrelated-CWD full gate passed 63 Python tests, nine
Pester tests per actual PS5.1.26100.9444/PS7.6.5 host, zero skips/syntax/scaffold
findings, four corrected Level 1 CLI targets, five normalization aggregates,
150 seeded plans and ten real writer samples. Every slice/flattened sequence
preserves exact page identities/order. Manual regressions retain 22 targets,
eight extra CLI cases, 250 plans/25 writer samples; only historical Level 2
BM-03 is compared with the original now. Six owned launcher probes including
three concurrent preserve inputs/neighbors/decoys/shared TEMP and remove only
GUID/marker-owned Documents outputs. They cover manual and historical Level 2;
corrected Level 1 console/Explorer remains unclaimed. Both PS7 probes retain
RawUI stderr; 60 inherited analyzer observations per host remain limitations.

Fresh hash-pinned regular GIL CPython 3.14.8 x64/pypdf 6.19.0 passed exact
versions/import origins/pip check. All 66 shell-tool files matched prior hashes;
stored policies stayed unchanged. Final tested-path digest:
`b937df04cc8035a30d01dbae9737114fe3646fc937033970c416b32e769cc468`.
Every tested raw byte matched clean C; no clean-C full rerun is claimed. Three
earlier development harness failures remain distinct from passing final source.
Original/extraction guards/oracles/generator and historical evidence are intact.
Independent source/harness/raw-evidence/privacy review found no remaining
M1-T03 blocker; no GitHub submitted review or CI pass is claimed.

Level 2 still crosses parents and is M1-T04. Full launcher/error/output safety,
Explorer, Calibre conversion, rendering, extracted package, CI (M4-T05) and
release/download gates remain open. No merge/tag/release was created in this
task. Prior sections below retain their historical scope and observations.

## M1-T02 evidence and limits

[Manual acceptance evidence](evidence/M1-T02-manual.md) and
[machine record](evidence/M1-T02-manual.json) satisfy AC-013 through AC-018.
The engine validates all ASCII decimal tokens/bounds before writing, normalizes
valid starts with notices and always includes physical page 1. `1` now writes
one whole-document PDF; explicit/implicit first-page inputs preserve identical
complete ranges. Empty/mixed-invalid/non-ASCII/out-of-range tokens fail entirely;
zero-page manual PDFs fail. Huge numbers/leading zeros are bounded lexically
before integer conversion. Bookmark/writer/launcher behavior remains scoped to
later tasks.

Fresh hash-pinned regular GIL x64 CPython 3.14.8/pypdf 6.19.0 passed exact
versions/import origins and pip check. Final unrelated-CWD full gate passed
36 Python tests, nine Pester tests per actual PS5.1.26100.9444/PS7.6.5 host,
zero syntax/scaffold findings/skips, all 22 manual CLI oracles, eight extra
CLI cases, 250 seeded valid plans and 25 real writer samples. Every slice's
page identities/order and flattened complete coverage were checked. Six bounded
owned launcher probes, including three concurrent launches, matched references
and preserved inputs/neighbors/decoys/shared TEMP. GUID/marker-owned Documents
outputs were removed safely without reading private entries. Both PS7 probes
retain inherited RawUI cursor stderr; console/error acceptance is unclaimed.
Sixty inherited application analyzer observations per host remain limitations.

The current full route deliberately uses corrected manual targets. Immutable
original guards, extraction harness, fixture generator/oracles and historical
M1 evidence are unchanged. Historical AC-011 ASTs read immutable M1-T01 C;
current safe import and three known-bad bookmark observations are separate.
A reproduced new evidence-checker Windows path casing defect was fixed and
regression-tested; earlier failed receipts remain distinct from final digest
`27cf0208484d831d5e818e64b1f2478a5066b96f436e2fb9da53f8029a3fe1ad`.
Every tested raw byte matched clean C; no clean-C full rerun is claimed.

Independent source/harness/evidence/privacy review found no remaining M1-T02
blocker; no GitHub submitted review or CI pass is claimed. Live main/feature
were unprotected with no rulesets/workflows/runs/tags/releases; no settings were
changed. Calibre remains installed but conversion/discovery integration is open.
Corrected bookmarks, all launcher/error/output-safety paths, Explorer, rendering,
extracted-package and release gates remain NOT RUN. No merge/tag/release was
created in this task. Older sections below retain their historical observations.

## M1-T01 evidence and limits

[Extraction/install evidence](evidence/M1-T01-extraction.md) and
[machine record](evidence/M1-T01-extraction.json) satisfy AC-011/012. The shipped
engine preserves all three processing-function ASTs, with import-safe callable
functions and a guarded CLI entry. PowerShell resolves it from `$PSScriptRoot`
and no longer generates/removes a shared TEMP engine. The six unedited original
files, including batch, still match their raw baseline blobs. Manual/bookmark/
zero-output bugs are intentionally preserved; behavior fixes start M1-T02.

Final unrelated-directory full run passed 20 Python tests, nine Pester tests in
each actual PS5.1.26100.9444/PS7.6.5 host, zero syntax/new-scaffold findings,
all 14 complete original/extracted known-bad observations and safe import.
Six actual PS5.1/PS7/batch probes, including three concurrent launches, matched
reference PDFs/page identities and preserved inputs, synthetic neighbors,
counterfeit CWD engine and the shared TEMP sentinel. Actual Documents outputs
were exclusively GUID/marker-owned and observed removed safely; private entries
were not enumerated/read. A reproduced descendant-after-timeout test-harness
defect and uncertain-pipe-close classifier were fixed and tested. Unknown
ownership or uncertain process state preserves outputs. Two PS7 piped probes
still emitted legacy RawUI cursor errors; console/error acceptance is unclaimed.
Earlier failed/pre-hardening receipts remain distinct from the final passing
digest `2a456dfc50d0ae62e93eaa52c10717ed9a90a3d819ab80560b5a5006f6387054`.
Clean C independently matched every tested byte; no clean-C full rerun claimed.

User-requested official portable Calibre 9.15.0 is now installed per-user at
`$UserProfile\Apps\Calibre915\Calibre Portable`. Published SHA-256/SHA-512 and
Valid Windows signatures were verified; installer and converter version probe
returned 0. No elevation or global PATH/policy change. The trusted nonstandard
path is outside unchanged launcher discovery; integrate discovery in M2-T04.
No GUI interaction or actual ebook conversion ran. Fresh regular GIL x64
CPython 3.14.8 dev venv/hash-pinned dependencies/import origins passed.

Independent source/raw-evidence/privacy review found no remaining blocker.
Live main had no protections/rulesets/required checks, workflows/runs or tags/
releases; no GitHub submitted review or CI execution is claimed. Canonical CI
owner is **M4-T05**; M0 prose assigning it M5-T01 was a documentation error.
Full repaired application/Explorer/Calibre/rendering/extracted-package/release
gates remain NOT RUN. Unsupported platforms stay unclaimed; no merge/tag/release
was created. M0 sections below retain their historical observations, including
the then-unmodified application and absent Calibre.

## M0-T04 evidence and limits

[Harness evidence](evidence/M0-T04-scaffolding.md) and [machine record](evidence/M0-T04-harness.json) satisfy AC-009/010. The small local runner has targeted/full stdlib unittest, exact-version shell tools, syntax/new-scaffold static checks, source/evidence hashes, sticky native/test exit status, exclusive external reports and owned temporary fixtures. Two full unrelated-directory runs passed with identical tested source digests: 11 Python tests, 9 Pester tests in each actual 5.1.26100.9444/7.6.5 host, 14 original known-bad engine cases and two early launcher boundaries. Native exit 23, failed Python unittest and one failed Pester test per host all retained runner exit 1 after later success. All checkout/input/neighbor bytes were preserved and owned temporary directories removed.

Fresh regular x64 CPython 3.14.8 dev venv/hash-pinned dependencies/import origins passed. Isolated Pester 6.2.0 and PSScriptAnalyzer 1.25.0 imported from absolute manifests in both hosts; all 66 saved tool file hashes/sizes were independently checked. Developer children use process-only RemoteSigned and their own built-in module paths; stored machine/user policies are unchanged. No global module/dependency installation occurred. Invalid shell preflight report writes and incomplete child-evidence acceptance were reproduced, fixed and covered by regressions. The external byte-matched handoff suite passed 56 tests with one symlink-creation skip.

Independent cumulative M0 source/evidence/privacy review found no remaining blocker. Actual GitHub main was unprotected with no rulesets, required checks, workflows/runs or submitted reviews; CI is NOT RUN and remains M5-T01. Exact-head merge used the normal process without admin bypass/deletion. Main was fast-forwarded normally and tested raw bytes rechecked. The seven original application files remain byte-identical; its 68 per-host analyzer findings are legacy observations, and known runtime defects remain. Full repaired splitting, Explorer, Calibre, current rendering and release-package gates are NOT RUN; unsupported platforms stay unclaimed. No tag/release was created.

## M0-T03 evidence and limits

[Support/setup matrix](SUPPORT_AND_SETUP.md), [task evidence](evidence/M0-T03-support.md) and [machine record](evidence/M0-T03-environment.json) satisfy AC-007/008's selection/setup obligations. Frozen application Python is regular GIL CPython 3.14.8 x64 only (minimum = selected), plain pypdf 6.19.0; Calibre 9.15.0 is the mandatory ebook-conversion test target, optional for PDF-only use. Hashed runtime/dev requirements are separate. Developer pins are ReportLab 5.0.1/Pillow 12.3.0/charset-normalizer 3.5.2; stdlib unittest/Pester 6.2.0/PSScriptAnalyzer 1.25.0 are selected for M0-T04. Current primary vendor requirements/advisories were reviewed with precise source links; no automated scanner pass is claimed.

Official Python was extracted without runtime registration/global installs. Fresh runtime/dev venv installs, pip checks and import-origin/version/absence assertions passed. The complete documented runtime setup ran under both actual shell hosts. The Windows PowerShell 5.1 native quoting defect found during review was fixed and rerun; execution policy remained unchanged. All 14 original engine cases and two unchanged early launcher error paths reproduced on the final dependency set, with exact page identities and source/input/neighbor preservation. This remains known-bad characterization, not repaired-engine acceptance. All seven original application files remain byte-identical.

Actual workstation remains Windows 11 Pro 26H2 build 26300.9457 x64; hosts 5.1.26100.9444 and 7.6.5. Existing registered Python 3.14.7 and bundled 3.12.14 are historical/unchanged, not current supported pins. Calibre was not found; no conversion ran. External shell tooling and selected newer renderer are not installed/tested. Full PS5.1/7 splitting, Explorer, real EPUB/AZW3, repaired-engine and release-package checks remain NOT RUN. Windows10/ARM/UNC and other Python versions remain unclaimed. Independent source/evidence review found no remaining blocker; no GitHub review submission or CI execution is claimed. M0-T04 was open at that checkpoint; its record above now closes the cumulative gate.

## M0-T02 evidence and limits

[Baseline evidence](evidence/M0-T02-baseline.md) and [machine report](evidence/M0-T02-characterization.json) record 14 original embedded-engine cases and two actual unchanged Windows launcher error paths. The engine was executed with generated page-marked PDFs; each output was reopened and checked by exact page identity/order. Manual first-page loss, success with zero files, invalid-token filtering, bookmark front-matter omission and Level 2 parent crossing were reproduced. Known-bad observations are separate from the corrected `PLAN_ORACLES.json` targets. This completes AC-004/005/006's reproduction/provenance obligations; it does not pass fixed-engine acceptance.

The original synthetic 10/12-page fixtures reproduce byte-for-byte with the same source/dependencies. Full anonymous fixed metadata, source/license hashes and exact outlines are recorded. All 22 pages were rendered and visually inspected. Original source, inputs and sentinel neighbors remained unchanged. No real documents, generated PDFs or private artifacts were committed.

Actual fixture/engine environment: existing Windows 11 workstation build 26300.9457; explicit bundled isolated Python 3.12.14, pypdf 6.10.0, ReportLab 4.4.9. Orchestrating shell was PowerShell 7.6.5; the bounded launcher probe used actual Windows PowerShell 5.1.26100.9444. Poppler was 26.07.0. These are development observations; supported versions remain M0-T03's decision. No global dependency installation occurred.

Full Windows PowerShell 5.1/7 splitting workflow, human Explorer drag/drop, Calibre conversion, release-package checks and fixed-engine acceptance: **NOT RUN**. Independent read-only source/evidence review found no blocker. No GitHub review submission or CI pass is claimed. M0-T04 was open at that checkpoint; its record above now closes the cumulative gate.

## Repository reconciliation and earlier work

At M0-T02's start, clean local/live feature checkpoint `4f550a2cdfb9736be8c2fd35ab83cc65c1f61ffc` was freshly SYNCED. GitHub showed PR #1 already merged at `2026-10-08T16:51:44Z`, with live main `1c8c681cae0019dc2f413803e150951a962287ff`. A normal fetch and fast-forward merge incorporated that actual history into the existing M0 branch. Draft PR #2 continued that task. No history rewrite or runtime edit occurred.

At M0-T03's start, clean local/live feature checkpoint `869218c3cc4c8b7f6b47d057a5e8e4a414924173` was freshly SYNCED. PR #2 was already merged at `2026-10-08T17:16:55Z`; live main `a66c8f95f922c36c58b47b6cbfba82be399a552a`. A normal fetch/fast-forward incorporated that history before implementation. New draft PR #3 continues the remaining M0 work; closed PR #2 is not reused. The earlier PR merges do not complete the M0-T04 gate.

M0-T01 implementation C was `ac09c024589158e1402ae63263818466d3e77d5c`. Its [handoff evidence](evidence/M0-T01-handoff.md) and [live audit](evidence/M0-T01-live-audit.json) retain the missing initial Git metadata, safe reconciliation/import, seven preserved original blobs and actual helper checks (57 tests: 56 passed, one symlink skip). Its final `4f550a2` receipt is historical and was reverified before M0-T02. The recorded original baseline remains `6edbed7c1a0c94968999882c5a46d90492d3c327`; all seven matched through M0. M1-T01 intentionally changes only PowerShell for extraction and adds the shipped engine; the other six remain byte-identical.

## Finish rule

Only RELEASE_RUNBOOK.md's verified public non-draft/non-prerelease v1.0.0 release, anonymously downloaded matching assets, fixed tag and final synchronized main allows COMPLETE. No tag/release was created in M0-T02. Unsigned status must be disclosed if signing is omitted.
