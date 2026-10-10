# M5-T02 — Allowlisted release package

Completed scoped package integrity at implementation **3785d731e66692c462b3c50cafaef67aaa7f524c**, tree **89a0448bf942acaef232d2b21ce3e14cc8288112**. Date: 2026-10-10, Europe/Berlin; machine receipts use UTC. The **175** mapped current source files equal Git/raw bytes, digest **97aa8a37e04db67a56e86935d04666d8a3ecd62d2964a785ab07ba17607ea4e0**. All 16 runtime files remain byte-identical to starting merged main `d720fdf7a2edad584f7e8521d3e75737c4938f96`. See [machine evidence](M5-T02-package.json).

## Changes and actual package

`tools/release/build_package.py` reads an explicit full commit through the reused dependency checker; `payload.json` allows exactly 28 runtime/current-support files. It excludes handoff/tests/tools/venvs/caches/poster/private inputs/bundled dependencies, uses fixed ZIP ordering/timestamps/regular modes and checks every shipped relative Markdown link. The current user docs link excluded developer/history references to their source repository. Source Git replacement refs are disabled after a real before-fix reproduction.

Builds require fresh ordinary absolute external directories, refuse existing output, and retain an incomplete marker on post-write failure. Verification requires a complete existing asset directory and independently recomputes Git-derived manifest/file bytes, checksum lines, member safety and exact archive recipe. Bounded asset reads and early fixed-format/footer checks refuse ZIP64 before archive metadata allocation. The builder creates no publisher.

Actual committed-C build, repeat build and CLI verification each returned 0 and preserved source hashes. An independent auditor that imports no project code passed **295/295** normal member/provenance/dependency/hash checks. Both builds' three assets are byte-identical in Python 3.14.8/Git 2.56.0.windows.1:

| Asset | Bytes | SHA-256 |
|---|---:|---|
| WinBookSplit-v1.0.0.zip | 413958 | 1f29981128c5fa74d17faca9e97d67f2605f008e80068d1f76809cbb79218645 |
| release-manifest.json | 11546 | f6822612e10100ef6cf40bc36a7c96f977c57a702ca620703d31032e0a595c48 |
| SHA256SUMS.txt | 178 | 7a26bbdca856cd060962662ccdc6dbd4e992e7625bea490c69bd0eb9d2d77f72 |

The external manifest includes the ZIP hash and never its own hash; sums list only ZIP/manifest. These are unsigned unpublished candidate assets, retained as `EXTERNAL/candidate-C01`; they are not the final M6 assets.

## Actual verification

| Command / scope | Source / outcome |
|---|---|
| DEV -I -B -m unittest discover -s ROOT/tests/python -p test_release_package.py -v | Final focused trial 25/25, zero skips, native 0,173.282s; authored Git fixtures and forged asset hashes |
| BASE -I -B tests/python/test_ci_package_inputs.py -v | Preliminary stdlib 16/16, zero skips, native 0 on actual base 3.14.7; supported 3.14.8 shared run below includes them |
| DEV -I -B -m unittest discover -s ROOT/tests/python -p test_runner.py -v | Preliminary 40/40 native 0, bounded timeout/failure guards; included again below |
| DEV -I -B tests/run_tests.py --layer python --report EXTERNAL/python-C01.json | Clean C **509 methods**, zero skips, outer/child native 0, bounded 600s; source/synthetic input stable and owned workspace removed |
| DEV -I -B tests/run_tests.py --layer shell --tool-root TOOLS --shell-path PS51 --shell-path PS7 --report EXTERNAL/shell-C01.json | Clean C **116/116 per host**, zero skips/not-run, syntax/scoped gates and native 0 |
| Committed builder build/build/verify with --commit C and fresh external paths | Native0 each; exact 28 files/all 16 runtime dependencies; repeated assets identical |
| Independent stdlib/Git audit, actual normal assets | 295 checks/native 0; no project imports, extraction or application execution |
| [Branch CI](https://github.com/PikkuJanne/WinBookSplit/actions/runs/38076575517), [PR CI](https://github.com/PikkuJanne/WinBookSplit/actions/runs/38076586544) | Actual completed passing reports/source/artifacts independently audited; hosted scope remains separate |

Local environment: Windows 11 x64 build 26300, regular CPython 3.14.8, plain pypdf 6.19.0; PS 5.1.26100.9444 and PS 7.6.5. Developer packages ReportLab 5.0.1/Pillow 12.3.0/charset-normalizer 3.5.2, Pester 6.2.0/PSSA 1.25.0 remain unshipped. Current shell tools were reused from retained hash-verified setup. Hosted fresh setup/version/source details are in machine evidence.

**AC-084/085/086 PASS** within package-integrity scope. Tests cover missing committed engine despite a worktree rescue, bad VERSION/allowlist, source/provenance/tool mismatches, excluded generated decoys, unsafe/duplicate/nonregular/encrypted/compressed members, rehashed tampering, source binding, owned no-overwrite/incomplete output and small ZIP64 parser-not-invoked control. Asset-only negatives reuse a captured immutable Git-verified source to avoid repeated identical subprocess reads; public build/validate/CLI/source tests remain separate.

## Retained defects, trials and limits

Two package trials were deliberately interrupted (tool native -1), with 5 and 19 completed methods respectively; neither is a suite pass. Ownership/tree-stop receipts and corrected zero-remaining Python/Git audits are retained. The first broad audit predicate self-matched its own shell and was corrected. The raw Git string0 fixture assumption failed before correction; the corrected before-fix replacement test then reproduced the real original-SHA/different-tree defect. The small ZIP64 fixture exposed an early metadata-limit gap; final guard/test pass before C. No large allocation was attempted.

The first overbroad 297-file raw auditor failed on historical CRLF evidence. Independent 23-file checks prove newline-only differences; those unrelated historical bytes are preserved. All 175 current mapped source files, 28 payload blobs and 3 bound tool files are exact committed bytes. Raw receipts and actual argv remain outside Git; only safe relative-file/hash projections are committed. No private book/document was processed or uploaded.

The Python child deadline changes from 300 to bounded 600 seconds for the enlarged real-Git suite; the historical M5-T01 timeout remains failed evidence. Historical full/unknown results and absolutely closed human M3-T01 are untouched. This task makes no new Explorer/human/Calibre/render/clean-OS/full22/final-package/publication/signing claim. M5-T03 owns actual extracted-candidate Windows acceptance; M5-T04 owns broader security/advisory review, M5-T05 cumulative scope integration, and M6 final source/assets/publication/download gates.

## Checkpoint and handoff

Live PR27 was already merged when this task started; the clean milestone branch fast-forwarded to real main before edits. C was normally pushed and clean/live SYNCED at **2026-10-10T18:38:14.346131+00:00**. [Draft PR28](https://github.com/PikkuJanne/WinBookSplit/pull/28) holds this task. The evidence checkpoint's own later SHA/push/equality receipt stays external/thread/PR after those actions, without a self-reference. No branch/tag reset, settings/protection change, overwrite or release occurred.

Next: **M5-T03 — Verify the extracted candidate in a clean Windows setup**, AC-087/088/089. Reinspect live Git/PR/check state, rehash all candidate assets and bind actual Windows/Explorer/Calibre results to C and the exact ZIP. Stop at that task boundary.
