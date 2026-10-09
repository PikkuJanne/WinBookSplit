# M3-T01 current-build human checks — complete

The current-build human checks are **complete**. Explorer, visual, repeat, fallback
and error checks passed; all six physical Ctrl+C controls passed across the four
original PowerShell cases and two fresh BAT retries. The user confirmed: “Both
retries passed. Human part of testing is now absolutely done”.
[Machine record](M3-T01-human.json) contains the distinct raw receipt hashes.

## Source and environment

Tested copied application: E `9983191ed3105d0edfec35be19888190d5acca33`, containing
implementation C `b8c6ed0f9e1153e20ba9b265615c6c3cab8d1a7e`. The 85 previously tested
raw paths still match, digest `ac6c5b910891ef8a86f078db43b2677f66cd515022dd17251ed10bbc9b3d87a0`.
Windows 11 x64 build 26300.9457; Windows PowerShell 5.1.26100.9444, PowerShell 7.6.5,
regular CPython 3.14.8, pypdf 6.19.0 and real portable Calibre 9.15.0.

`<human-kit>` is the retained external copied-application workspace. It has a fresh
application venv and authored PDF/EPUB/genuine AZW3 inputs, with no private books.
Explorer PS defaults alone select its output base and real converter. Shipped BAT,
engine and supporting runtime bytes are unchanged. This is current-workstation
testing of controlled copies, not an extracted final package or clean OS.

## Human observations and disk verification

The user confirmed all six Explorer drag/drop rows, first/last-page visual checks,
and a repeat run. Independently reopened output PDFs, page content, physical ranges,
manifests and preserved identities support these results:

| Check | Result |
| --- | --- |
| Simple PDF Level 1 | 3 PDFs, 10 physical pages |
| Nested PDF Level 2 | 6 PDFs, 12 physical pages |
| Simple PDF manual 1,4,7 | 3 PDFs, 10 physical pages |
| Flat PDF manual 1,4,7 | 3 PDFs, 10 physical pages |
| Real EPUB manual 2,3 | 3 PDFs, 3 physical pages |
| Genuine AZW3 manual 2,3 | 3 PDFs, 4 physical pages |
| Repeat simple Level 1 | New 3-PDF/10-page run; all 45 earlier saved files unchanged |
| Flat PDF automatic → explicit manual fallback | Initial no-bookmarks 5, retry 0; 3 PDFs/10 pages |
| Explicit fallback cancel | Human final stdout and native BAT exit 130; no chapter publication |
| Invalid manual token | Human final stdout and native BAT exit 2; no Done/publication |

The initial six rows total 21 PDFs/49 pages. Their native BAT exits and final stdout
were not independently captured; success is human-reported and supported by disk
and provisional logs. The first flat-PDF row used direct manual mode; the later
separate fallback repeat supplies actual automatic-to-manual evidence. Visual
inspection is user-reported, not an automated rendering/fidelity pass. Sources,
controlled application files, neighbors and earlier outputs remained unchanged.

## Failed helpers and successful repairs

The first two error-control helpers failed before application launch with native
9009 and CMD errors including `ho`/`ho.`/`ll`. Original UTF-8 LF-only helper bytes
and failed observations remain archived. Actual native CMD reproduction isolated
the Unicode/chcp line-ending boundary; changing only LF to CRLF reached the exact
Unicode argument and preserved expected 130/2. Scripted diagnostic application
controls then returned 130/2; separate human retries confirmed both exits.
`wrapper-repair-receipt.json` records the two superseded helper hashes. The original
preparation manifest remains historical and unchanged. No shipped source fix or
human pass is inferred from those earlier failed attempts.

## Physical Ctrl+C coverage

| Control | Saved native exit | Current result |
| --- | --- | --- |
| PS5.1 engine | 130 | PASS |
| PS5.1 conversion control | 130 | PASS |
| PS7 engine | 130 | PASS |
| PS7 conversion control | 130 | PASS |
| Original BAT engine | 255 | Failed launcher check; application cancellation was 130 |
| Original BAT conversion | Not launched | NOT RUN |
| Fresh BAT engine retry | 130 | PASS |
| Fresh BAT conversion retry | 130 | PASS |

The user initially reported six successes, then explicitly answered **“Y, then
Enter”** when asked about CMD's termination prompt. The saved series contains only
five rows and stopped after native 255. Answering Y terminated that batch job;
there is no sixth measurement or original human-report.json. A new followup receipt
binds the clarification and first-four attribution without changing historical
NOT_CONFIRMED records or turning the failed BAT attempt into a pass.

The four passed controls prove final/provisional cancellation 130, stopped owned
parent/descendant processes, pipe EOF, unchanged source/prior/neighbor identities,
retained marked incomplete stages and unrelated Python survival at measurement.
Independent review also checked the original BAT shutdown and retained stage, but
its native-exit assertion remains failed. Later reused numeric PIDs are not old
process identities and were not adopted or killed.

These cancellation fixtures use real authored PDF writing and a native authored
converter control, with finite 120-second parent/child holds. They do not prove
physical cancellation of real Calibre. The observer sends no signals or keyboard
events and never automatically kills the application target.

## Completed BAT retry and limits

Fresh `<human-kit>/cancellation-bat-retry/Start-Cancel-controls.cmd` ran only
BAT-engine and BAT-conversion. Both saved native exits and final stdout outcomes
are 130. The operator's separate `human-report.json`, confirmed at
`2026-10-09T16:47:29.813297+00:00`, attests physical Ctrl+C in both READY windows
and **N**, then Enter at CMD's termination prompts. The latest user message
confirms both retries passed.

Read-only independent receipt/disk review passed: both real console measurements,
exact source/default/helper/runtime hashes, stopped owned parent/descendants,
completed streams, no chapter publication, unchanged source/neighbor/prior
identities and content, retained marked stages and unrelated Python survival at
measurement. All 85 tested repository paths still match the sealed digest.
The review compares recorded process observations; it does not adopt later reused
numeric PIDs. Original failed/NOT_CONFIRMED receipts remain byte-identical.
The converter-stage empty PDF is retained incomplete output, not a successful PDF.

`completed-retry-independent-review-v2.json` records the successful audit and exact
retained files. Its preserved first review had an overbroad “no processes” field;
v2 corrects it to no **application** processes/signals because read-only `git show`
processes did run. No application was relaunched or human action automated.

Runtime source was unchanged; no application regression rerun is claimed by this
documentation supplement. Structural plan and whitespace checks gate the checkpoint.
Next is M3-T02 CLI. M3-T03 UX, wider PDF fidelity, CI, exact-package acceptance and
public release remain open; this current-build completion does not certify future
package bytes. PR16 merged at main `981894db7a91d7c2d3f0e3af258323ddeef38106`;
normal fast-forward reconciliation preserved history. This supplement's own commit,
push and live-sync receipt are recorded externally after they actually happen.
