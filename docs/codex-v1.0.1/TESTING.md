# Automated verification strategy

## Reuse and proportionate coverage
Reuse `tests/run_tests.py`, current Python/unit/Pester/fixture suites and current package auditor. Extend them with compatibility/setup cases rather than building another general test framework. Existing regression results from v1.0.0 are historical evidence; rerun the relevant automated tests on v1.0.1. Do not replay the human acceptance programme.

| Layer | Required scope |
| --- | --- |
| Policy unit tests | All advertised minors; below/at/above floor; numeric comparison; prerelease/free-threaded/architecture; malformed policy; dual-layer parity. |
| Discovery/setup | Fake bounded probes for launcher failures, direct PATH, explicit/.venv authority, decoys, no implicit installs; real setup into clean and existing environments. |
| Python matrix | Actual stable x64 interpreter in every advertised minor, real engine execution and important regression fixtures; exact used versions recorded. |
| Reference regression | Full existing Python and PS5.1/PS7 test suites on the designated reviewed toolchain; safety tests may not be removed to obtain green. |
| Native Windows 11 | Final ZIP; PS5.1 BAT/entry point plus PS7 direct path, setup and core smoke; exercise at least two supported Python branches here or on another native workstation. |
| Native Windows 10 | Final ZIP; actual 22H2 x64/PS5.1; one representative policy-supported interpreter is enough for this native platform smoke. Minor-branch coverage is supplied separately by the matrix; report that distinction. |
| Native ebooks | Real self-authored EPUB and genuine AZW3/KF8 conversion on both Windows families with the reviewed converter, followed by output validation. |
| Package/release | Source-to-member allowlist/provenance, secret exclusion, hashes, public read-back and tag identity. |

Do not require every OS × shell × Python × patch × converter Cartesian combination or repeat visual fidelity inspection on unchanged writer code. Review risk and use the targeted matrix above. Later stable-patch acceptance tests can use synthetic probe records; they are never claimed as runs of a not-yet-installed/released interpreter.

## Developer dependencies versus runtime dependencies
Existing `requirements-dev.txt`/tool checks may contain native-wheel hashes or exact Python assertions for the reference environment. Audit and decouple them. Keep the heavy fixture generator/renderers on one reviewed development toolchain where practical. Generate/reuse deterministic licensed synthetic fixtures once, then run the application under each real target interpreter with its runtime-only dependency lock. Pure-Python regression tests should run in each compatible environment when feasible. Do not skip engine tests because a renderer/dev wheel only exists for another minor; solve the fixture/tool separation instead. Never install cp314-only wheels under cp311.

Commit matrix selections/resolved versions and verified hashes for reproducibility. Permit current maintenance updates in application policy independently. Standard GitHub x64 Windows runners are Windows Server, not native Windows 10; label reports honestly. Keep CI permissions contents:read and action revisions pinned after verification. Publishing belongs to a deliberate separate release step. Do not expose a personal Win10 self-hosted runner to untrusted public pull requests or `pull_request_target` checkout code.

## Native smoke harness to implement in V101-08
Create/extend a PowerShell 5.1-compatible harness under `tools/compatibility/` (not inside the runtime path) that accepts an exact ZIP, expected ZIP SHA-256, manifest, output report location and optional approved runtime/converter paths. It must verify bytes before extraction, reject zip traversal/duplicate case-insensitive paths, use a new owned directory, preserve existing environments and never use private input.

Capture real OS caption/build/architecture/ProductType, PowerShell/.NET version, interpreter executable/version/flags, package versions, exact archive SHA-256, release source R and policy hash. Test setup's interpreter-only discovery, install/check/no-op behavior, version query, manual `1,4,7`, first-page inclusion, Level1/Level2 parent boundaries, Unicode and literal paths, repeat-run non-overwrite, malformed input, no-Python and wrong-package messages, failed `py -0p` fallback, timeout/cancellation and cleanup. Run real EPUB/AZW3 conversion and reopen/page-count/coverage/source-hash checks. Use existing fixtures/tests where appropriate.

Use an actual child process and actual BAT for launcher forwarding. Feed prompts in a bounded way or use appropriate pseudo-console/UI automation without adding insecure testing switches to production. Noninteractive paths must never block awaiting input. Don't equate a direct engine call with a BAT test.

A machine with only one interpreter can still supply native OS evidence. The all-minor matrix is collected elsewhere. Kit runtime provisioning is an explicit isolated developer/test action; it does not change the end-user processing rule. It must not require installing every minor on the owner's Windows 10 machine.

## Evidence rules
Every result has test/case IDs, status, command, actual versions, source/package identity, timestamp and relevant stdout/exit status. Use PASS, FAIL, NOT_RUN, BLOCKED; a not-applicable case has a documented reason and may not erase a required Win10/Win11 gate. Distinguish MOCKED_POLICY from NATIVE_AUTOMATED, USER_LAUNCHED automation, and any HUMAN_OBSERVED check. No fabricated screenshots or copied old PASS labels.

Redact user/profile paths, host names, network identifiers, tokens and environment values before public commit. Include only the fixture names and necessary OS/build/tool facts. The raw report stays local. The redacted evidence names the hashes of underlying results but must not imply that hashing authenticates an observer or protects a compromised host.

Missing native Windows 10 access triggers the one-session procedure, not a server substitute or a broad manual checklist. A failure blocks only dependent work; continue independent work with an exact handoff.
