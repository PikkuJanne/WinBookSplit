# T01 / STORE-006: historical acceptance is not Store-route acceptance

Observed 2026-10-17 11:18 UTC. Developer-only contract checks on Windows 11
build 26300, CPython 3.12.14, x64, isolated/no-bytecode flags. No application,
setup BAT, native manager, installer or acquisition helper was executed by these
tests. This is **not** Windows 10 native/application acceptance.

Source HEAD before the T01 checkpoint: `4cc6b00b99685607f3214e731edaaec7d65926df`;
committed tree: `8a5a246d2d18b3e2b3fd746a9347f77fb95f6e08`. Tests run with the
imported, uncommitted v1.0.3 handoff checker and new regression file, while the
parent agent independently reconciled routing. SHA256 fingerprints at execution:

| Input | SHA256 |
| --- | --- |
| `tests/python/test_store_gate.py` | `c2acee8357b19e908fd8f257a7be7f73757c2edb24ec45efdb65ac7047b98fc0` |
| `tools/codex-v1.0.3/check_native_store_report.py` | `fb028b5531541e0ad64a7f7ee975d61d842d20c58cde6d01e77adbaa0647f298` |
| `tools/compatibility/native_smoke.py` | `70ad7cb9f13d9aa81e85bb773fcc9ccc20673ad2960f8ba5841b4e05f307569e` |
| `tools/compatibility/Invoke-NativeAcceptance.ps1` | `9747a3754e37fe6a342af9f38a58cce49682a2854c90abc2e7d0de06e9c9ef07` |
| Historical `V102-11-windows10-review.json` | `3a5deafcac3fe00db8f26b09b6189d652f7dce47909853471d6fd1617726b90a` |
| Historical `V102-11-windows10.json` | `b60000035c22b88507e3e27af35e06cd7af183e0752bc170a0cee1cf531b9d09` |

## Command and result

`<DEV_PYTHON>` denotes the existing bundled developer Python executable used for
these checks; its local account path is intentionally omitted from public evidence.

```powershell
& '<DEV_PYTHON>' -I -B tests/python/test_store_gate.py
```

Exit 0; 8 tests, 7 passed and 1 expected failure. The expected failure is an
unimplemented requirement, not a successful Store acceptance gate.

- The real immutable historical returned report still passes the maintained
  `native_smoke.export_report` consumer under its original contract. It reports
  `owner_regular315_usable=false`, `owner_selected=false`,
  `baseline_kind=reviewed_private3150_archive`, and explicit runtime identity.
- Both historical JSON records and the original embedded returned-report object
  fail the new checker with `not the existing-owner Store scenario`. Their bytes
  were checked unchanged by the tests. They were not relabelled or adapted.
- A complete, explicitly labelled **SYNTHETIC** contract with six hash-bound
  temporary **SYNTHETIC** receipt files is a positive consistency control only.
- Used fallback, missing/unknown usage flags, hidden runtime selection, explicit
  runtime paths, copied/cloned runtime, child PATH injection and a prebuilt app
  venv disqualify otherwise complete synthetic contracts.
- A narrow checker gap is reproduced: top-level `fallback_acquisition` and a
  hash-bound preservation receipt record can declare
  `private_runtime_fallback_attempted=true` with
  `private_runtime_fallback_used=false`; the supplied checker accepts that
  declaration set. `FORBIDDEN` currently requires a usage observation only and
  receipt `records` have no attempt-specific semantics. The future maintained
  kit/checker change must explicitly record acquisition disabled/no attempt and
  reject an attempt before any fallback provisioning. No checker change is part
  of T01.

The failing assertion was also run without `expectedFailure` handling:

```powershell
& '<DEV_PYTHON>' -I -B -c 'import runpy, unittest; scope = runpy.run_path("tests/python/test_store_gate.py"); case = scope["StoreExistingGateTests"]; method = "test_attempted_private_acquisition_disqualifies_even_when_not_used"; getattr(case, method).__unittest_expecting_failure__ = False; result = unittest.TextTestRunner(verbosity=2).run(unittest.TestSuite([case(method)])); raise SystemExit(0 if result.wasSuccessful() else 1)'
```

Exit 1; 1 failure: `AssertionError: GateError not raised`. This preserves the red
requirement without turning the ordinary runner green by weakening an assertion.

## Inspected legacy route

Read-only inspection of `Invoke-NativeAcceptance.ps1` lines 644-682 confirms that
failed owner discovery may provision the reviewed private archive and then set
`BarePythonPath` to that baseline. The current legacy consumer deliberately
accepts `reviewed_private3150_archive`; it remains valid supplemental execution
evidence under the old contract. No claim is made that the new controller gate
has been implemented. T05/T07 must remove this route from the required existing
owner Store scenario and prove normal automatic setup/BAT processing on the exact
candidate ZIP. Missing owner Windows 10 evidence remains NOT_RUN/BLOCKED.
