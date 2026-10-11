# V101-03 package integration receipt

Reviewed on 2026-10-11. This is external development evidence, not a Windows 10 or release acceptance report.

Source baseline HEAD: dccb460be65d2352b81d43a977b20aaf4ce325f3; tree: 6ca042651bcd67f0624a08df8b261a3bdde29c08. Tests used the current scoped working files in authored isolated real-Git repositories, while other agents also had authorized uncommitted task-03 source changes. Exact changed file digests are below. No root Git commit, push, index, status document, release, or installed system runtime was changed by this work.

Actual runtime: regular Windows AMD64 CPython 3.14.8 (tags/v3.14.8:8e6e75d, Sep 30 2026, 18:19:33), MSC v.1944; pypdf 6.20.0. Host platform Windows-11-10.0.26300-SP0. This cannot establish Windows 10 compatibility. The executable was the authorized per-user isolated dev14\Scripts\python.exe under %USERPROFILE%\.cache\WinBookSplit-Dev\v1.0.1-20261011 (actual user path redacted here). All commands ran from <checkout> with -I -B.

## Actual reproduction before edits

A dev14 Python stdin script loaded tools/ci/check_package_inputs.py and checked baseline committed prerequisites: rejected missing_input, WinBookSplit_icon.ico. The same AST checker on the current task-03 engine rejected undeclared_runtime_dependency, engine/winbooksplit_engine.py. On the new Python compatibility helper it rejected unresolved_runtime_dependency, engine/winbooksplit_compatibility.py. The policy JSON was appended only to an in-memory runtime-path tuple to isolate the dynamic-import rejection; that reproduction did not edit files.

Live GitHub run reports independently showed the same pre-existing missing-icon input failure on both current main be1fa1e271aa2bd031d907e9dbc9e7bfc32bbb62 and task-02 dccb460. Original failure evidence was retained externally under ci-security-review-20261011. No older tag or asset was modified.

## Final scoped changes

- Required shipped closure includes WinBookSplit.Compatibility.ps1, WinBookSplit.Compatibility.json and winbooksplit_compatibility.py.
- Existing intended artwork path WinBookSplit.ico replaces the nonexistent WinBookSplit_icon.ico in current checker/payload only. Artwork bytes unchanged.
- Dynamic-import audit allows only importlib.import_module('pypdf') with one literal positional argument and no keywords in exactly engine/winbooksplit_compatibility.py. Other targets, variable targets, alias imports, alternate receivers, __import__, direct import_module calls and the same call in another application source remain rejected.
- Payload paths remain lexically sorted and version remains 1.0.0 until the separately authorized freeze/version task.
- Existing fixtures assert 19 runtime inputs, 13 support inputs, exact closure edges and all three committed compatibility members. Missing committed compatibility inputs cannot be rescued by a present working file. Release test asserts 31 members and the current icon/compatibility members; ZIP format checks remain in place.

## Commands and real outcomes

1. dev14\Scripts\python.exe -I -B tests\python\test_ci_package_inputs.py
   Exit 0; Ran 18 tests in 255.956s; OK. These used authored real-Git fixtures and actual source checks.
2. dev14\Scripts\python.exe -I -B tests\python\test_release_package.py (first attempt)
   Exit 1; setUpClass failed before any tests with invalid_payload_allowlist because the changed icon path was out of lexical order. Corrected ordering in payload.json.
3. dev14\Scripts\python.exe -I -B tests\python\test_release_package.py (corrected payload run)
   Exit 1; Ran 25 tests in 225.825s; 24 passed, one failed solely because the old test asserted 28 members when the actual intended payload had 31. Corrected the expected count and added explicit new-member/icon assertions. Other 24 cases completed in this run and are not represented as a fresh full rerun.
4. dev14\Scripts\python.exe -I -B tests\python\test_release_package.py ReleasePackageTests.test_payload_members_and_independent_git_blob_hashes_are_exact -v
   Exit 0; Ran 1 test in 17.098s; OK. This reran the exact affected real-Git/blob/ZIP method after the assertion edit.
5. dev14\Scripts\python.exe -I -B tests\python\test_release_security.py
   Exit 0, but no tests were run because this module has no __main__ unittest entry point. This is not counted as a test pass.
6. dev14\Scripts\python.exe -I -B -m unittest discover -s tests/python -p test_release_security.py -v
   Exit 0; Ran 5 tests in 0.825s; OK. Three actual Windows BAT shell decoy/literal-argument checks and two actual pypdf stream resource-limit checks; no skips.
7. git diff --check -- tools/ci/check_package_inputs.py tools/release/payload.json tests/python/test_ci_package_inputs.py tests/python/test_release_package.py
   Exit 0, no whitespace errors; Git emitted only LF-to-CRLF working-copy normalization notices.

No newly authored test simulated a runtime version or platform to claim compatibility. Existing package tests contain targeted mocks for builder failure/ZIP parsing safety; those tests are package-logic checks, not Windows compatibility evidence.

## Final SHA-256 identities

- tools/ci/check_package_inputs.py: 239a49d5a8daa0d52eb2a3394383cff7f9b9050e2a5b61649926058ef6d2badd
- tools/release/payload.json: 1974511a131147da260ae55d7e94aeff8baff66eeb35c214affdde4c153b16e3
- tests/python/test_ci_package_inputs.py: a9fc82087011f932dc73c6f2498d34aa2a0effc51d08f7c9bccfc3abbe117793
- tests/python/test_release_package.py: b148c67b32e31e2ee314a3ce48ed73891a9bb50a4bbe0a2639c328b77c7239d7
- WinBookSplit.ico: 0d7df83ea710192bf568e3347af4551640058407e121a03f5d50d627e548dc42; 370070 bytes, unchanged.

Read-only review of the companion PowerShell WithIdentity change found that its SHA-256 is computed from the exact byte array that the shared policy parser validates, closing the old two-read identity mismatch. This uses System.Security.Cryptography builtins and does not need another shipped PowerShell module.

Checkpoint: task-03 package integration changes and focused verification complete; parent owns final full CI/native verification and documentation/commit synchronization. Native Windows 10 evidence remains a later explicit acceptance requirement.
