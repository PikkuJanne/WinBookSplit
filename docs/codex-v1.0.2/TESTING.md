# Layered automation and evidence

Reuse existing unittest/Pester/runtime/native/packaging tests. Every case in ACCEPTANCE_CASES.json starts NOT_RUN. Synthetic cases in RUNTIME_POLICY_CASES.json and INCIDENT_CASES.json are inputs/expectations, not passed application tests.

## Source tests

Run policy/identity tests under each advertised Python branch with both PS5.1 and the currently reviewed supported PS7 where appropriate. Use the existing single policy, not an independent duplicate production implementation. Execute both branch boundaries and synthetic later patches, unknown minors, malformed records, wrong architecture/GIL, broken launcher listing, package import origin, stale venv and explicit override. A synthetic 3.15.99 proves numeric handling, not that such an upstream release exists.

Run the existing full automated suite on the integration branch, focused setup/BAT/diagnostic tests and security regression cases. Update expectations tied to old branch preference, exact test versions, package members and output wording deliberately. Do not mass-replace historical versions or remove failing guard tests to turn CI green. Keep runtime-only and developer fixture/renderer dependencies separate.

CI: extend the current workflow using verified pinned actions, least-privilege contents:read for test jobs, bounded execution and redacted authored reports. The matrix must include actual 3.15.0. Do not skip it with continue-on-error, conditional missing-dependency success or fallback to 3.14. Record exact resolved runtimes, actions and wheel hashes; source tests on a Windows Server runner are supplemental, not Windows 10 desktop certification. Reuse generation of trusted fixtures instead of forcing unavailable native developer wheels into the runtime.

## Native final-archive matrix

All native success rows use the same frozen ZIP Z extracted to a new ordinary per-user path, not a live checkout. Verify each package member against the release manifest before execution; adding local `.venv` does not change application member bytes.

| Environment | Required execution |
| --- | --- |
| Windows 10 22H2 x64 / inbox PS5.1 / regular 3.15.0 | New setup BAT; emitted remedy round-trip; fresh `.venv`; automatic/manual PDF plans and writes; offline/check/repair; MOTW under RemoteSigned; failed optional launcher listing; BAT exit/args; original source hashes. |
| Windows 11 x64 / inbox PS5.1 / regular 3.15.0 | Same new setup path and PDF smoke; MOTW and emitted remedy; actual EPUB and AZW3 conversion under 3.15. |
| Reviewed PS7 / regular 3.15.0 on a supported Windows workstation | Invocation from PS7 of new BAT and explicit host remedy; direct PS7 script path under process-only policy; no misleading process-host version claims. |
| Existing accepted 3.14 and 3.13 branches | Automated source regression plus a focused final-ZIP setup/check/PDF smoke for each on an available Windows workstation. Do not require every old OS×minor cross product or repeat manual viewing. |

At least one native 3.15 conversion run per ebook format is required; unchanged converter behavior on Windows 10 may reference earlier evidence only as historical context, not as v1.0.2/3.15 acceptance. A regression in conversion triggers focused additional OS cases. Log exact tested combinations rather than general 'all versions pass'.

## Policy tests without persistent changes

Use process-only RemoteSigned/AllSigned test contexts in owned local NTFS directories. Negative control must prove direct Internet-marked unsigned `.ps1` is blocked. The documented BAT/host entry then succeeds without removing ADS or changing persistent scopes. Assert CurrentUser/LocalMachine/MachinePolicy/UserPolicy and parent process policy unchanged. Child Process policy differs intentionally.

Test organization-policy diagnostic paths with a mocked policy query/code-path check. If an already authorized isolated GPO fixture exists, a native denial test is useful, but do not demand it or edit real GPO/registry. Label simulation accurately. Static review verifies no alternate-host/encoded-code policy evasion. A mocked Group Policy result is not native organization enforcement evidence.

## Evidence

Record source commit/tree or runtime-member digest, archive SHA256, test harness revision, OS caption/build/architecture, shell executable/version, Python real executable/version/GIL, pypdf/Calibre versions, selected source (explicit/venv/PATH/launcher), fixture hashes, commands, native exit, expected/actual, timing, failures/skips and tested scopes. Evidence must refer to owned authored documents and no secrets. Hashes do not make a test true by themselves; inspect results. Retain full local logs, share redacted compact JSON.

Do not mark PASS based on a template, setup banner, version policy alone, simulation of Windows, prior v1.0.1 source, or output existence without its verified result. Missing mandatory native cases stay BLOCKED/NOT_RUN. Failed cases require a defect and rerun on the corrected exact artifact.
