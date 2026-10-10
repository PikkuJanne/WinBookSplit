# M3-T04 — plans and converted-PDF page guidance

AC-060/061/062 scoped acceptance passed. Actual clean C1 full 17-stage plus clean C2 affected Python/native UX and source delta are verified; implementation C2 is normally pushed and clean/live synchronized. Public evidence E receives its own external commit/push receipt after publication of these documentation bytes.

Interactive processing prepares one immutable physical-page plan, displays titles, ranges, filenames, destination, count, coverage and front-matter/opening/fallback notices, then requires explicit confirmation. Ebook manual selection offers the real generated PDF before physical starts; the same owned working copy remains available through confirmation. Success reports the real final directory and chapter/page counts and offers opt-in literal folder opening. Failure never announces success; NonInteractive/Preview preserve their no-prompt/open/pause rules.

## Source and verification

Clean C1 `cd9b4c359a042d6ef3fc8900cb3cf0f373bbe01f`, tree `2ea85a67d9a4c74b9e899f6cc05fb66f69de31ba`, supplied the full seventeen-stage matrix. Its 106 raw code/test/tool paths have digest `97a09181887d5b9ed0d80ea08b73e60bf063ca0eac80ddd7bd1a312a5770c4ee`. Full02 result/native exit/counts/source stability/cleanup: **PASS: outer native 0/all17 stages0; 258 Python tests; 88 Pester tests per host, zero failure/skip/not-run; 72 CLI controls/70 launcher controls/28 UX controls (26 host-app plus 2 engine rows); exact C1 source unchanged and owned workspace removed**.

Final implementation C2 **32b64bbd2884e76e3f0128403217d21f7018ec57**, tree **a4a3a4e039937ebee121de32c0047aa5f41b901a**,106-path digest **f40cc1c22671997434ee29c5b8810c467f4b3c37908684ca3a6203468277ee11**. C2 changes only two identified validator/unit-test paths: **tests/ux/validate_ux_report.py + tests/python/test_ux_receipt.py**. The final raw comparison proves 104/106 files and all 12 runtime files unchanged. Clean C2 Python/actual supported-host UX reports and counts: **PASS: clean C2 268 Python tests/zero failures/errors/skips and actual both-host native 28 UX controls (26 app + 2 engine), both outer/stage native 0, source unchanged/owned cleanup**. This combines the real full C1 matrix with affected C2 validation; it does not claim an unrun full C2 matrix.

Mechanical R1 `a0135fe7bf746719c92fdd487f215cbdec2be00b` separately extracts the existing conversion/completion blocks. Its exact isolated mechanical snapshot passed 166 old-module tests with AST/extracted-block comparison. That scope is separate from the interactive feature and full-gate proof.

Exact full02 outer argv flags/order are independently sealed by live task process metadata in `full02-command-metadata01.json`. The relative `tests/run_tests.py` launch occurs at the repository; the runner separately creates an owned unrelated workspace for test children. Win32_Process exposes command line, not cwd; completed per-step receipts prove native 0 and the unrelated workspace (shell children use shell subdirectories). Opener04's `RemoteSigned` is verified directly from its actual native-result argv, with cwd `<owned-opener-kit>/cwd`.

Runtime versions observed in real supplements: Python3.14.8, pypdf6.19.0, Calibre9.15.0. Final actual Windows/PS51/PS7 versions: **Windows-11-10.0.26300-SP0; Windows PowerShell5.1.26100.9444 and PowerShell7.6.5**. Exact executed commands/cwd/streams remain in sealed external receipts. Public command templates below redact owned/runtime paths and are explicitly not new execution receipts.

```text
REDACTED TEMPLATE ONLY — actual full02 command/outer native exit: full02.json + full02-outer-completion.json + full02-command-metadata01.json
"<python-runtime>" -I -B tests/run_tests.py --layer full --tool-root "<isolated-shell-modules>" --shell-path "<Windows-PowerShell51>" --shell-path "<PowerShell7>" --calibre-path "<real-Calibre-ebook-convert>" --report "<new-external-report>"
outer launch cwd: <repository> (root's actual launch description)
runner test children cwd: <runner-owned-unrelated-workspace>; completed full02 per-step receipts prove unrelated cwd/native0

REDACTED TEMPLATE OF ACTUAL OPENER04 ARGV — original bytes: opener04-native-result.json
"<Windows-PowerShell51>" -NoProfile -ExecutionPolicy RemoteSigned -File "<owned-app>/WinBookSplit.ps1" -InputFile "<authored-epub>" -OutputDirectory "<owned-output-base>" -PythonPath "<python-runtime>" -CalibrePath "<real-Calibre-ebook-convert>" -NoPause
staged stdin: M, O, 1,2,3, Y, O, each followed by LF
```

## Acceptance and method

AC-060 (canonical manual_windows): **PASS both actual hosts: exact 5-section, 12-page nested plan with front-matter/opening/childless-fallback notices; confirmed hash matches reopened outputs**. Actual native cases prove readable titles/ranges/filenames/count/destination/coverage and all three notices. Confirmed immutable plan/hash matches reopened outputs exactly.

AC-061: opener04 actual native 0 requested the generated PDF through unchanged, unmocked `Invoke-Item -LiteralPath`; entered physical starts1,2,3; confirmed the displayed plan. The real 3-page, 58,264-byte working PDF retained identical hash/native file identity through selection and confirmation. Three1-page outputs retain authored markers001/002/003, exact content hashes and complete coverage; the owned working copy and converter stage are absent after success. Source/neighbor/prior native identities remained unchanged. All12 tested runtime bytes subsequently match C1; C2 equivalence: **PASS all 12 runtime files identical to C1 and C2**.

AC-062: opener04 reports the actual 3 PDFs/3 physical-page/final-folder result and requests native folder opening only after the explicit final O. No launch error or stderr occurred. C1/C2 no-plan/failure/NI/Preview evidence: **PASS actual both-host strict no-plan/failure/cancel/NonInteractive/Preview controls; no false success or prompt/open in NI/Preview**. Fresh signal07 PS51/PS7 Ctrl+C and Ctrl+Break controls all exited 130 while a genuine4-page/2-section plan awaited consent, with zero replies/chapter PDFs, sole standalone OUTCOME equal to the log footer, preserved input/prior identities and proven owned child stop/pipe EOF/input-writer stop.

The opening evidence is actual Windows native request plus authored file/lifecycle/content proof. An earlier opener02 Sky listWindows exposed a Firefox title; subsequent Firefox getWindow hit Computer Use browser URL policy and computer use stopped. No current rendered-PDF screenshot, viewer-content or Explorer-window visual pass is claimed. Opener04 contains no mocked Invoke-Item or application instrumentation. These are disclosed assistant/native observations, not new human results.

Inherited authored fault/deadline/terminal-frame tests declare their engine substitutions and complete noninteractive/copied-PS choices in exact raw argv/seam receipts. They prove transport/protocol behavior. Real PDF/ebook/opening controls use the actual engine and real converter; failed-before-plan controls contain no fabricated consent.

## Preserved attempts and limits

- Baseline reproduction is old b316cd9 behavior, four actual cases, never a current feature pass.
- Inherited smoke01 preserved real mixed-reader EOF cancellation; smoke02 passed 15 selected development controls and is not the final full matrix.
- Signals01/02 helper rejected an additional private conhost before signaling. Retained parent handle was stopped; descendant shutdown/cleanup is not adopted as a pass.
- Signals03 partial run has no completed aggregate: a legitimate engine-cancelled record tripped the helper's old zero-record expectation; Break also interrupted the controller. Retain partial workspace and claim no aggregate/cleanup pass.
- Signals04/05 are prior app-byte observations; use fresh07 as primary C-bound cancellation evidence.
- Signals06 all four applications exited 130, but PS51 Ctrl+C appended OUTCOME to its pending prompt. Strict 3/4; failed receipt/review retained. Root's line-boundary fix preceded fresh07 strict 4/4.
- Opener01 actual Get-FileHash unavailability reached application log outcome6 and no viewer. Native app exit was not separately captured; helper failure is not that exit receipt.
- Opener02 helper AttributeError ux.digest occurred after Done; native application exit was not properly sealed. Prior Sky listWindows Firefox title belongs only to this attempt, with no screenshot/04 visual claim.
- Opener03 failed before app launch: locale-default reading of UTF8 creator text transcribed the authored source path incorrectly. Root command-chunk/helper-source provenance receipt exists; original stderr bytes were not separately saved. No native app pass.
- Full01 failed both shell stages (one optional-Diagnostic StrictMode test per actual host) and the UX stage whose old harness demanded a persisted log for no-write Preview. Source/helper fixes and Preview stdout-only receipt rules preceded clean-C full02; full01 remains failed.
- Mechanical isolated snapshot construction had prior UTF8/Git-setup failures; only the exact R1 snapshot's actual 166 old-module unit success is adopted for the mechanical refactor, not final feature acceptance.

Human M3-T01 checks remain absolutely done by user attestation. Their historical receipts and disclosed owned-kit defaults stay closed, with no new-source/package relabeling or further human request. All failed/partial authored workspaces remain retained; no broad or private-document cleanup claim.

## Additional scoped UX and receipt-oracle review

UX targeted01 failed the old pending-timeout oracle when a legitimate graceful engine-cancelled frame raced supervisor shutdown. Actual130/processor_timeout/stopped tree was preserved. The narrow corrected oracle accepts zero or one cancelled/no-output frame; four explicitly synthetic positive-output/different-frame/reply mutations are rejected.

UX targeted02 subsequently passed 28 actual controls (26 supported-host app rows plus 2 direct-engine working-copy lifetime rows) on unchanged 106 raw bytes matching C1 digest. It preceded the clean-C1 full02 gate and does not replace that full gate.

Preview-focused correction has4 actual PS51/PS7 Preview/NI controls,8 explicitly synthetic Preview receipt mutations, and a separate actual PS51 no-log reproduction. Preview deliberately has no persisted console/transport file; native stdout outcome/embedded plan/streams are authoritative, while NI still requires its real disk log and output proofs. No viewer/Explorer pass is inferred.

C1 cancellation/other-branch resealed data mutations exposed receipt-oracle blind spots, not actual application failures. The external two-file C2 candidate first adds status/code/output/mode guards. Additional pure candidate reviews exposed terminal-count/mode, NI immutable-source binding and destination binding gaps. Final candidate03 has 14 pure methods,28 retained actual rows revalidated and 14 resealed contradictions rejected; fresh clean-C2 268 Python tests and actual 28 UX controls pass; no external synthetic test is relabeled as a native rerun.

## Checkpoint and next task

Final C2 clean/live synchronization: **2026-10-10T04:41:41.349780+00:00; C2-live-sync01.json (normal push/native 0, clean local HEAD equals fresh live feature branch)**. Fresh continuation PR/live-main/check state: **https://github.com/PikkuJanne/WinBookSplit/pull/20; PR20 OPEN draft at headC2 32b64bbd2884e76e3f0128403217d21f7018ec57/base main b316cd9eee8d2fade933b31e2bbf60be1ec99e2c; merge CLEAN, no GitHub review decision/check rollup; main unprotected with no required checks, check-runs0/combined status pending with0 statuses/actions0/releases0/tags0; no CI pass inferred**. Final source/raw/public privacy review: **PASS scoped C1/full and C2/affected source/raw audits; final public byte hashes are recorded separately in the external review receipt**. E receives a separate external normal-push/clean-live receipt after commit; it does not claim its own future hash.

Mark M3-T04 done only after real full02, affected clean C2 reports and104/106+12/12 delta proofs, reviewed evidence and verified checkpoint. Next is **M3-T05 — Produce useful local manifests and redacted diagnostics**. No T05 export implementation, upload, release/tag, CI/package/fidelity certification is included here; first public v1.0.0 remains in progress.

## Raw receipt seals

Original files remain external. Names below are redacted labels; local user/runtime/source/output paths are intentionally omitted. Add completed full02, C2 affected/delta/sync, fresh PR and final review rows only after those receipts exist.

| Receipt label | Bytes | SHA256 |
| --- | ---: | --- |
| baseline-native-ux-reproduction.json | 212673 | f7b8fbb74a3ec59076003c83c23be6620ea56c2029a24cf0f5e7f652873e4977 |
| inherited-smoke-01.json | 56630 | f30da368913332afabb50e430a27b191acec19d375e3408787aed23b31437d1d |
| inherited-smoke-02.json | 868281 | a499eb7bc119c754b7773d983b48f275ac26aa25493de67ade37ff628297d459 |
| pending-session-signals-01.json | 35201 | 46152e5cd61ea5ef468fa77e3b938b1be09d71508aa8a4c4ab950d9ebccbc124 |
| pending-session-signals-02.json | 46173 | 6448c3fbddd5b971146a33591cc00004487d9884a2ab0f88fb9d39d444850ea5 |
| pending-session-signals-04.json | 184365 | e62f333a1ae6bd51f3b36c0952166d8052f7e8e78d4acc978fa0826b59d84328 |
| pending-session-signals-05.json | 46358 | 8835a787feef12506d07f94fe173d0a526d36e946d40f678fad62275d6494bb5 |
| pending-session-signals-06.json | 183656 | a0321a8c5d824cba8f617d2baedc25b23019ebae4cb9f9700d59233c0a742257 |
| pending-session-signals-attempt06-review.json | 4655 | d6e33a158d7df82f65c76ec09c4a6a0c66359813e6880653573a0b5401816cfe |
| pending-session-signals-07.json | 184743 | 71e7a7eac089822350a98ba3a4bfa5e2c643afa90ea77b1779ff45dfe5b3c158 |
| pending-session-signals-final07-review.json | 9440 | 655cdcd32243d180ba39026db37f4065117fab8fc707021612341dcba5bb0c5a |
| signal-app-freeze07.json | 1902 | 9b5263f6d91ed3893655ac6acb15576ec690ff921669c0cc0674f9b085937857 |
| actual-viewer01-failure.json | 20058 | 46bc80d450fd9fef598fe19186b216ddde02948eee22d14c2d6de367193a49e5 |
| opener02-helper-failure.json | 19961 | c03d393431dc048f7316d86e450ba1f04fee7164c44044a0cdfb3faa99553762 |
| actual-viewer03-prelaunch-failure.json | 1095 | 5d241c23e0b5022daf09c27117ee5facd23893772662aff094e298003d38cf54 |
| opener04-viewer-ready.json | 4834 | 0d2ec8d21b3d8d7b052a361a2d37d8ee74f7e31407b57c3a46d2926aef0c98fa |
| opener04-confirmation-ready.json | 23395 | 25dda49f89c858fd24d510b263d14331cd7cb90bce70e2207220bb36e37515d3 |
| opener04-output-ready.json | 33030 | 50e048febf85d990757f2beb50403038e3052e70349e20c85f8652c7332954dc |
| opener04-native-result.json | 18952 | ece7692bdbdf8710a6a12f9d796f7d10910e53d0ededb7508ee03e5e5df06d40 |
| opener04-native-complete.json | 105502 | ba9e0e0e3db281130b4b949a76be49dfc7333a95a6019eef31c9c44743044e76 |
| actual-native-opener04-independent-review-v2.json | 9374 | 4971c1dac38a9aa94c8b4758dd9218481e9250ce692042e4a7b75f27d071117c |
| full01.json | 15529849 | 9901211412fbee0e83ca8399ba1e258da1c377dcfb1895b008bade5274e3ef85 |
| implementation-commit.json | 11905 | 15ad05c5a58c5af605bd5f0421c2e8746df29f51ddea220cd1bdfddd421f2cad |
| mechanical-r1-commit.json | 694 | 4319f59fd42d9c45982f21aa91c29acf9d91c12fa4c9b7063759d9a0bbf157c0 |
| mechanical-construction.json | 1271 | 9df376d62a414e5752a992eb5e788056e399faaf2fb09e7973d98ab4bb266365 |
| mechanical-unit-report03.json | 2702 | 543cb98d14a12d67494db56f0a88993858730cfc3c29dd7f32df76c09cbaed71 |

Additional preserved review and launch metadata seals (final candidate and actual affected-report seals follow below):

| Receipt label | Bytes | SHA256 |
| --- | ---: | --- |
| full02-command-metadata01.json | 7452 | 6896d4c74bc6a05a7bda8c68b9b90b87fcf2267c973a210f60b480ca6a1cf690 |
| ux-targeted01.json | 1151663 | 61ffa1e85b572b3419725f976937e03bc2aab9149c85f50472644cd3b6235416 |
| ux-targeted02.json | 2431102 | 36e1a258607301081bf45a88c42db34404a9acd02ddb582b36d7b50f2d054965 |
| ux-targeted02-independent-review.json | 5157 | 746fc56e7be66d09ba54543d22f316cb630d660cc201ac169dd9406e6bf05519 |
| ux-preview-focused01.json | 176739 | abd7f1060cb944e6ca50bdb4d333f9fef6867665a9bcac33a19b219317a19e1f |
| ux-preview-no-log-reproduction01.json | 18064 | 8cadbb95a079d3cf742e404b018a85a420010a4b81ebded8776dbb91f055c197 |
| ux-timeout-oracle-review01.json | 1660 | 53c7091374de9e6e7109239f27b2ee093ac6c75ec870a8258dbc15299c0b8e0b |
| postfix-source02-review.json | 3471 | af612496bc3eee669c3254da0f0024ac3ab2dba42f302eab932405c8bca6a3bc |
| pester-diagnostic01-review.json | 2308 | 49b4376332f6e8e073cd313f624cfcb3acc560b92eec8379ccfb6eafebd0b39f |
| cancel-status-C01-review.json | 2521 | 9fadf1bd5612293997bd9b807a6738c9d46b0616187775ee9694d4324f88db5f |
| other-ux-branches-C01-review.json | 605808 | 3286d2dcd2e389a6824238a8db19ad6ef6094c4891e97c495f3779960029ae8c |
| cancel-guard-C02-candidate.json | 3721 | 4a3cabc94b77711d45cca4e9c919a2d2db1c27f7bc1683287127c69d1de3a9a7 |
| C02-extra-receipt-review01.json | 1054 | 0afbbf33301e5882c86bc1db25262eade35cb2db68a578fae927c94ca5361623 |
| C02-extra-destination-review01.json | 531 | da84f5d20dc386f5912e61c257c786341c28a6699a2d30c00420c966a566605f |

Actual final source/gate detail: each supported host parsed 13 syntax files and 12 scaffold files without findings. The 70 legacy application PSScriptAnalyzer observations per host (67 warnings/3 information) are disclosed observations, not an application analyzer-clean claim. Fresh read-only Windows metadata is Windows 11 Pro x64 build 26300.9457/26H2; legacy registry ProductName remains Windows 10 Pro and is not an additional OS test.

Redacted actual clean-C2 launch templates (outer cwd `<repository>`; child cwd independently sealed):

```text
<python-runtime> -I -B tests/run_tests.py --layer python --report <new-external-report>
<python-runtime> -I -B tests/run_tests.py --layer ux --tool-root <isolated-shell-modules> --shell-path <Windows-PowerShell51> --shell-path <PowerShell7> --calibre-path <real-Calibre-ebook-convert> --report <new-external-report>
```

Additional actual reviewed raw seals:

| Raw receipt | Bytes | SHA-256 |
|---|---:|---|
| full02.json | 16891033 | d251dc56d1c5fdf705c4b5038ad29e1244aba8cf8a73b41ba5cdf1378b6f2d80 |
| full02-outer-completion.json | 807 | 83e2b114b22d2c0c45625d26f78f92ddbd44ff04dab50ecfe2806ea7377ed2cf |
| full02-independent-audit01.json | 11868 | 4ef7c0614dec0fe50888d46d863effe343d4441c5cb1ad555c3f3bc57fd8f1e8 |
| full02-independent-detail01.json | 5639 | d0f866cbb2bde50bda3105073ced4d297dea0f2fe7c26f7dfda2b63b71f13ed6 |
| candidate-review03.json | 9784 | 17582bb284441a349dc000f78201452c18aa0e0ee980cbdadef9983729f92be2 |
| C02-candidate03-independent-review.json | 6196 | d7213fdfe0ff8a896379e8260d9c1b12847a398b4fd226a4a4b47dc2f3619895 |
| C2-patch-apply01.json | 642 | 0c4b4cbd511559d25ec874e6e26415e4f2938f4393edf6d35c10461b12e3fd2d |
| implementation-C2-commit.json | 729 | a8315778065f111e5878e14e8b00db328ef957418da7781ab2f61c001982d6bc |
| C2-source-and-C1-revalidation01.json | 17539 | 704806e80e2905f578ffeefce9290efa3244e211b920d537c1398a4e10ad9a72 |
| python-post-review01.json | 81018 | 6643c41c1229f49c5d0df4521d308b84aee86446326819f47af233366f6c53d4 |
| python-post-review01-outer.json | 833 | 125c062689a295190cedf92ce1a479e809d4844674eaeba4cda567d4790902d2 |
| ux-post-review01.json | 2431350 | b3c3678489c5c39e2dc0886e29c30ff1a99a04d45e2586200cc2bdb3256cfca1 |
| ux-post-review01-outer.json | 1280 | 1c15ddb9582bbae2313866ce8706270b6567159e23e7e1b69547bcf0c1210a0d |
| environment-final01.json | 1213 | 1e8e629d996cfa5685298e04ca40bf714dfa3f384618f63c97580461a777ec1e |

Actual C2 checkpoint and post-review seals:

| Raw receipt | Bytes | SHA-256 |
|---|---:|---|
| post-review-C2-independent-audit01.json | 4948 | ed2da510e6a54860e5d58a46582cc748fa6949cda6963e27c5339d0cdbf7b735 |
| C2-normal-push01.json | 644 | 98b487dbd4090ae7bda5af6610e5252c5827cc6443532eeb5bf2d905895be16d |
| C2-live-sync01.json | 1211 | fe2a7f104360a3e2426838c36d46e86ff4e3eebe430236adcb5a357a33297fc6 |
| draft-PR-create01.json | 876 | 23a8204597e8a00ed586f1959586cf8c62207b53dcea99e15055c4c2e9a8bf7f |
| draft-PR-inspect-C2.json | 1186 | be05117d6e81f8de9609482f299ba16774dfd523577efca7d5cb1f6360d93825 |
| github-main-C2.json | 842 | f7075d30c0c2ac03ab7380e13cccacd361a631eec8149649641ef14512741180 |
| github-checkruns-C2.json | 744 | 9dace63d728505a51d33d9fa17c22dd986e9128b20c3d3b595142446bf4e9a5e |
| github-status-C2.json | 760 | 27051f11c9018290fb5fc95dc261f7d989d92df965b93e0370eae8b62f969fa4 |
| github-actions-C2.json | 640 | f5172abbe178cb9b26847eaca5f9077e711ceeeb6b94a75799a7b73ff98d9bd9 |
| github-releases-C2.json | 622 | ee58766f0b94802f9bced12e65f1c23fb29efa0dcb6e1060ba47ad5f5e7c0640 |
| github-tags-C2.json | 604 | 1cd525334e3b8fbc5a8453c1a61815ebee5b2f12831abf69b21938a88f7f0ef0 |
| live-refs-C2.json | 754 | dbfe270102db4a53a6175d031983b9bb8f063d3b59a21040c67df4de44dc44ee |

Exact observed C1 native row exit histograms: CLI0×26/2×36/3×4/5×6; launcher0×40/2×3/130×27; UX application0×16/130×10 plus 2 direct-engine 130 controls. Fresh C2 UX has the same 16 success/10 application-cancel/two direct-engine-cancel distribution. These expected negative controls pass their validators; outer gates and every test stage exit0.

Checkpoint metadata validation initially returned native 2: root supplied repository-relative TASKS evidence paths. The two paths now follow the existing evidence-directory-relative convention, and the repeated structural validator returned native 0/VALID. Both receipts are retained. This documentation correction changes no application bytes or runtime acceptance result.

| Checkpoint receipt | Bytes | SHA-256 |
| --- | ---: | --- |
| E-plan-validation01.json | 729 | 0ca8716e2022dd0272ad7f3202796ceb3ab369c0fabffddf46d68d6710e1a7af |
| E-plan-validation02.json | 826 | 0d1e1f262c97fa1f01cee569f51e737d8bf4498c9b55954763c133fa97b82dd2 |
