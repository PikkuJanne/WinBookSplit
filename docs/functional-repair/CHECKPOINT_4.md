# Checkpoint 4 — workflow and exact-package verification

30 October 2026. Feasible checks and the consolidated native kit are recorded
here. Ordinary supported-Store consent/setup/resume acceptance remains NOT_RUN.
No publication, release, tag, version or dependency change is authorized.

## Source and package

Read AGENTS.md, STATUS_NEXT.md, CHECKPOINT_1 through CHECKPOINT_3 and their JSON
maps, the historical closure/testing references, and the Downloads repair
pack's CODEX_TASK.md checkpoint 4 and ACCEPTANCE.md. The pack still hashes to
`f33ab17eab774f4a65cf6e008949012e810cb0c10ce4c9bb9796f0d6524dc70d`.

Started clean on `codex/functional-repair`, HEAD and freshly queried live branch
both `ca29d143cf104e6802eeededbb42b817db051845`, tree
`cd1803383c0f57ea1bbd4fdd29b6e1aa35a3b6d0`. Origin remains
`https://github.com/PikkuJanne/WinBookSplit.git`; one worktree. Main was
`bafa1868da10ab6ba65f112fe29b4bf4deb909f8`.

The existing builder first refused `builder_source_mismatch`: the Windows
checkout's builder files have CRLF while their committed blobs have LF. Exact
committed builder/helper bytes were materialized outside the repository and
the unchanged builder was run with the explicit original repository/commit.
Its guard was preserved. That attempt then exposed `unshipped_document_link`:
three shipped documentation pages linked to unshipped functional status.
Only those links were changed to explicit repository URLs and committed as
`7837503e7cbab27d0bd6c8bf90e95460581158d8`, tree
`331c8599fd882e0114d578eb974d0943b4c162f0`. Behavioral documentation remains
checkpoint 5. Both original failures are retained.

The existing builder's build and verify now pass against that source. The
**unpublished checkpoint test archive**, still named by the existing builder
`WinBookSplit-v1.0.3.zip`, is 960,774 bytes with SHA256
`451cd7659fac0b95fc033cebb3bca2798057a19332284021f69cf52ce2ea5dd3`.
It is a local test artifact, not the published v1.0.3 asset or a new release.
Its manifest SHA256 is
`ba28c8b111185cd87c7c635fe7d4c24fd0ab4e411afba27d9ac0327f9c384d75`.
All 34 archive members and their fresh extraction match the source Git blobs
byte for byte. Compared with the Windows checkout, 21 members are exact and
13 differ only by CRLF-to-LF normalization. The independent member-map SHA256
is `1bae7a86e31f30486ad1e7ae4689ba2e8d097e122ebc4448d44938c107cd9866`.
The allowlist excludes environments, developer caches, private documents and
chapter outputs. Final checkpoint tools/evidence are outside this payload;
CHECKPOINT_4.json records their separate byte identities.
The canonical changed tool/test map SHA256 is
`10fa7fb4f74f0d817da09ce886d4654242ad00923551002f27a67ffc8098f234`;
it distinguishes physical working bytes from staged Git blobs and excludes
documentation to avoid self-reference.

## Executed results

Existing tools only: Windows 11 build 26300, regular x64 CPython 3.14.8 with
locked pypdf 6.21.3, actual PowerShell 5.1 and existing PowerShell 7 hosts.
The developer interpreter is a labelled test driver, never application
configuration for ordinary discovery. No Python, Manager or global package
was installed or changed. Explicit offline component fixtures are separate
from the fresh ordinary-test application.

| Check | Actual result and boundary |
| --- | --- |
| Exact package build/verify/member audit | PASS as described above. No published assets changed. |
| Existing functional tests against fresh extracted engine | PASS, 11 methods, zero failures/errors/skips, 27.883s. The unchanged trusted harness's engine/policy/root globals select the extracted payload. Manual/Auto1/parent-aware Auto2 preview/execution, captured readers, interactive confirmation, converted validation and password/active-content negatives run against those exact engine bytes. |
| Actual packaged healthy application, repeated | PASS, six launches: two each on PS5.1.26100.9711, PS7.6.5 and documented PS7.6.6, using the explicit pre-existing developer runtime and NonInteractive mode. Unicode/bracket/ampersand source/output paths and unrelated CWD; combined empty-password/local-GoTo input; three chapters and six independently authored page-content markers in order; exact manual ranges; one ignored-view warning; unencrypted outputs without OpenAction. Source, neighbor, prior output and all earlier chapters preserved. No setup frame/prompt or app `.venv` creation. This is healthy NO_INSTALL coverage, not BAT cold-start/Store consent acceptance. |
| Native oracle on actual packaged outputs | PASS on all six healthy runs. Independent wrong-page-order, changed-source and wrong-range controls are rejected. Helper checks do not establish native Store acceptance. |
| Existing package guard suite | PASS, 26 methods, 305.534s. Uses the existing builder/allowlist and retains its source, membership, collision and provenance guards. |
| Fixed final application-unit roster | PASS, 171 methods across ten required modules, zero failures/errors/skips, 196.181s. Final outcomes included. Source is the byte-preserved starting checkout; the package freeze changes only three documentation links. |
| Planning/output/ebook/fault unit controls | PASS, 172 methods, zero failures/skips, 80.923s. Existing manual, outline levels, plan, preview, output, diagnostics, ebook assets/behavior, converter cache and fault invariant/receipt modules. |
| Actual native manual characterization | PASS, 118.106s, on PS5.1 and PS7.6.6 plus the existing BAT route. Explicit existing developer driver; independently expected ranges/page order, literal paths, invalid/manual-input and output/process preservation controls. This is separate from ordinary Store first run. |
| Actual authored EPUB/AZW3 conversion | PASS, 157.727s, existing Calibre 9.15.0 and Poppler 26.07.0 on PS5.1/PS7.6.6. Real conversion and authored render comparisons; source/input/tool/prior-output preservation. No private book, GUI, installation or arbitrary-network-isolation claim. |
| FirstRun isolated components | PASS, 38/38 on actual PS5.1.26100.9711 and 38/38 on PS7.6.6, zero skips. Decline/EOF/cancel/setup failure/repair/read-only/healthy/literal-request controls remain component evidence. |
| Focused failure reruns before recipe-helper correction | PS7.6.6 Discovery: 40/40 PASS. PS5.1 Discovery/Setup/Support: 84/87 PASS, three native recipe failures retained and diagnosed separately below. Original successful Runtime/Manager/Outcome scopes retain their original run boundary; no final 215-case rerun is claimed. |
| Final changed-helper native setup recipes | PASS, three executed tests per host on PS5.1 / PS7.6.6, 118.838s / 119.251s, zero failures/skips. Actual emitted setup action, backup-preserving BAT repair/healthy check, and authored MOTW/RemoteSigned denial. Both are launched through Python with the inherited environment. The other 24 Setup methods per host were filtered NOT_RUN in these focused reruns. Final helper PSScriptAnalyzer 1.25.0: zero findings. |
| Unchanged native process suite, sequential | PASS, 35 actual controls (26 supervisor and nine PS5.1/PS7/BAT application controls), 71.716s; source unchanged. The first failed attempt is retained below. |
| Ordinary unchanged source-copy BAT | Desired missing-package case SKIPPED / NOT_RUN, one test, zero failures. Actual BAT returns exit 3 / runtime_not_found for naturally discovered Python 3.14.7 / python_security_floor. Input/application bytes preserved, no `.venv` or chapters. |

The workstation has the existing **python.org** Manager registration, family
`PythonSoftwareFoundation.PythonManager_3847v3x7pw1km`, version 26.4.240.0.
It does not have the required Store registration. Natural discovery reaches
the existing concrete 3.14.7 and correctly preserves the 3.14.8 security floor.
The existing developer Python was not substituted, copied, provisioned or
injected into PATH/PythonPath. No ordinary-test application environment was
pre-created.

The first broad PowerShell component attempt remains FAIL: 215 cases per host,
PS5.1 188 PASS / 27 FAIL and PS7.6.6 207 PASS / 8 FAIL. Its external controller
used the parent of WorkRoot as CWD; Discovery's deliberately external executable
fixtures then fell inside the excluded CWD and correctly failed the guard.
The PS5.1 combined session also reported duplicate ObjectSecurity type data
and unavailable Get-FileHash/Get-ExecutionPolicy commands. A fresh host resolves
those commands; a particular container's responsibility for the module state
was not established. The corrected external driver uses CWD=WorkRoot, imports
that host's actual builtin modules explicitly and separates FirstRun into a
fresh session. At this rerun, production files, tests and assertions were unchanged. Only
failed scopes and the separated FirstRun components are rerun; their final
results are recorded separately.

The three remaining PS5.1 failures occurred in the generated **test recipe**
before application actions. A read-only replay through Python with the inherited
Codex PS7 module search path confirmed exit 1, empty stdout and
`CouldNotAutoloadMatchingModule Microsoft.PowerShell.Security` at the recipe's
Get-ExecutionPolicy call. The setup frontend was not reached. The narrow
`Setup.Tests.ps1` helper repair explicitly imports the actual child host's
`$PSHOME` builtin Security module before observing policy. It preserves the
inherited environment, actual emitted commands, MOTW, policy expectations and
all assertions; it changes no application module or execution policy. The
three affected recipes are rerun separately on both hosts through the failing
Python/inherited-environment route. Their new test-byte scope is separate from
the original source runs above. Both final three-test runs pass. The helper's
final working SHA256 is
`85772563b3487eda373a6286878d752c9f122873d776de84946650b32dd4ddff`;
its original working SHA256 was
`139b652d6327932ca72df471f87ebf2f853f7ba273c50b50a02ff3cf88945df9`.

The original native process attempt also remains FAIL. Its authored one-second
detached-pipe control timed out at 1.0286s during the concurrent run before the
child PID marker appeared. Job assignment, parent/descendant stop and complete
stream drainage were proved, with no leaked child; the attempt did not finish
the process acceptance suite. The unchanged sequential 35-control rerun passes,
without changing its deadline or expectations.

The 54 PDFs found in the checked local Downloads/Desktop/Documents paths are
historical authored fixtures/chapters; the supplied private 335-page book was
not found. Its Auto1/Auto2, original hash and rendered-page comparisons remain
NOT_RUN. Current host facts identify Windows 11; no accessible Windows 10
environment was found. Neither unavailable scenario is waived into a pass.

## Native kit and continuation

The focused kit reuses the existing bounded process/session, Store inventory,
hash/JSON and redaction helpers. It contains this exact archive, a combined
public fixture, independently authored expected page markers, one numbered
guide and `START-HERE.bat`. It provisions no runtime and passes no selected
interpreter to ordinary setup/BAT. A future qualifying run checks consent,
fresh app-runtime/base/package provenance, original-source continuation,
reopened chapters, page order/coverage and a repeated BAT run with no setup,
reinstall or overwrite. The kit explicitly leaves network activity NOT_RUN
without a network trace; absence of setup logs is not claimed as zero network.
Raw artifacts remain local and the transferable result is one redacted
`SEND-THIS-REPORT.zip`, including a missing-prerequisite result.

Final kit is `%USERPROFILE%\Downloads\WinBookSplit-Functional-Native-Checkpoint4.zip`,
1,578,321 bytes, SHA256
`8068d08b0f1207bee9c0cbbee2b644386847601844b2840d01cbb77f70f7758f`.
Its bindings SHA256 is
`b56af8c69c7e38b70419afe60e3f0ce8278ee209f9113468c2fe0e290c189124`.
The actual generated START launcher, from a fresh extraction and unrelated CWD
with the inherited environment, returns **3 / NOT_RUN**, empty stderr and one
redacted report ZIP. The packaged BAT itself returns 3 for the observed floor
rejection. Independent durable-archive inspection confirms all 24 bound files,
exact nested test ZIP, current controller/oracle bytes and redacted report;
this is a tooling rehearsal, not successful Store setup. Final native-oracle
units: six PASS, 5.478s, including actually rewritten reordered/encrypted chapter
negatives. Final PS7.6.5 parser/configured analyzer: zero findings.

Earlier kit rehearsals are retained as FAIL: empty pause replies are rejected by
the existing session transport; failure records needed literal message text;
the first strict output guard rejected the application's legitimate finite
failure-console records. The corrected kit uses nonempty pause replies, saves
the controller exit across its own pause, and permits only those diagnostic
members while continuing to reject chapter/staging output. Its final launcher
and report path were executed after these corrections.

The retained package audit is under
`%TEMP%\WBS-functional-cp4-package-6130d7916065476db29c14124ab8c162`.
Its commands use exact Git-blob builder files and:

```powershell
& $DevPython -I -B $CommittedBuilder build --repo $Repo --commit 7837503e7cbab27d0bd6c8bf90e95460581158d8 --output $FreshAssets
& $DevPython -I -B $CommittedBuilder verify --repo $Repo --commit 7837503e7cbab27d0bd6c8bf90e95460581158d8 --output $FreshAssets
& $DevPython -I -B tools/compatibility/build_functional_kit.py --package-dir $FreshAssets --destination $FreshKit
& $DevPython -I -B tests/python/test_functional_native_kit.py
```

`$DevPython` is the pre-existing
`%TEMP%\WBS-623-finish-q6dwc7hk\dev14\Scripts\python.exe`.
The final kit rehearsal is under `%TEMP%\WBS-functional-cp4-kit-tested-fresh`;
its checks are under `%TEMP%\WBS-functional-cp4-kit-checks`.
The original ordinary BAT receipt is `%TEMP%\WBS-first-run-9q3aans4\evidence.json`.
Other regressions and their exact argv/host/source maps are under
`%TEMP%\WBS-functional-cp4-regressions-1568ca9035be494f820d089f08e01295`.
`compact-results.json` binds the completed scopes and retained failed attempts;
`final-recipe-summary.json` records the changed-helper runs. Source-stability
checks pass for the original, corrected and final focused runs. These use an
isolated same-commit source copy with the initial checkout's exact working
bytes, keeping concurrent native-kit additions out of their source manifests.
CHECKPOINT_4.json binds local raw receipts without committing their logs or
generated PDFs. No helper self-test or explicit healthy run fills a missing
ordinary acceptance row.

The ordinary Store first-run and its repeated-launch package-network
observation, Windows 10 and private-book acceptance remain NOT_RUN. Published
tags/assets and historical evidence are preserved. No processing, discovery,
setup, lock/floor or safety implementation changes are made in this checkpoint.

Commit/push intended paths normally, verify clean local HEAD against the freshly
queried live branch, report that receipt in the handover, then STOP. Continue
checkpoint 5 only in the next chat; the exact prompt is in STATUS_NEXT.md.
