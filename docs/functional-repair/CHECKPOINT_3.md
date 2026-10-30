# Checkpoint 3 — consent-based local setup and request continuation

30 October 2026. Implementation and component verification are complete. The
ordinary supported-Python cold-start acceptance remains NOT_RUN, not a mocked
success. Checkpoint 4 is a separate chat; no publication is authorized.

## Authority and source boundary

Read root AGENTS.md, STATUS_NEXT.md, CHECKPOINT_1.md, CHECKPOINT_2.md, the
historical closure/testing references and the repair pack's CODEX_TASK.md
checkpoint 3 and ACCEPTANCE.md. The Downloads instruction ZIP has SHA256
`f33ab17eab774f4a65cf6e008949012e810cb0c10ce4c9bb9796f0d6524dc70d`.

Started with a clean `codex/functional-repair` at
`47d054728b009caab0ec35fd7b6ce908c5ec9985`, tree
`022b54e6542567cf4a09db2c44b2f1379314f073`, equal to freshly queried live origin.
Origin remains `https://github.com/PikkuJanne/WinBookSplit.git`; one worktree.
CHECKPOINT_3.json records the tested changed-file byte map and local receipt
hashes. Documentation is excluded from that map to avoid self-reference.
Its canonical changed-file map SHA256 is
`a3c1114abeaacd7d143d64788c41a0701995f032c414d073657f8f921fb12734`.

## Implemented behavior

- Verified supported interpreter/package failures now report `setup_required`
  / exit 3. Missing, wrong-version and untrusted-origin pypdf remain distinct
  package reasons in attempts and human summaries. Unsupported/unverified
  interpreters and interpreter-only checks retain their existing authority.
  Explicit-base remedies preserve their literal `-PythonPath`.
- The interactive preflight offers setup once. Only Y/Yes proceeds; No, blank,
  EOF and C cancel without calling setup. An unhealthy authoritative `.venv`
  gets an explicit repair-and-backup choice. Preview and NonInteractive bypass
  consent and setup. Explicit readiness Check remains read-only.
- Consent calls existing `Invoke-WinBookSplitSetup` with structured arguments
  and `-NonInteractive` to avoid a second prompt. There is no new installer,
  shell evaluation or implicit selected-base injection. The original document
  path also reaches setup discovery, preserving document-directory exclusions.
- Setup's final protocol, status, code and exit must agree. Progress, actual
  failure text and the structured setup result are displayed and retained in
  local diagnostics. Cancellation, timeout, setup/rollback failure and malformed
  results stop before engine execution. Existing backup/restore/incomplete-artifact
  behavior is reused without deleting old environments.
- Successful setup triggers one fresh runtime resolution requiring this copy's
  `application_venv`; there is no second setup attempt or global-runtime fallback
  if that environment disappears. A caller's explicit base is used by setup,
  then the newly prepared app environment is selected after consent. Healthy
  initial explicit-runtime behavior is unchanged.
- The caller retains its original literal source, output directory, mode,
  starts/level, converter options and deadlines. Only the runtime is returned
  from the coordinator; no request is rebuilt or redropped. Consent/setup runs
  before console-record/engine resources are held. Component tests exercise
  Unicode, spaces, brackets and ampersands; complete BAT install/resume with
  those paths is still a native acceptance requirement.

The PowerShell/Python/pypdf/Calibre architecture, checkpoint-2 PDF implementation,
Store identities/listing guards, dependency locks and security floors are
preserved. There is no persistent policy/PATH/alias/global-package change,
desktop runtime acquisition, private book, release build, tag or publication.

## Test evidence

Existing developer tools only: Windows 11 build 26300, regular Windows x64
CPython 3.14.8, locked pypdf 6.21.3, Pester 6.2.0 and analyzer 1.25.0. The
developer Python is a test driver, never injected into ordinary discovery.
Existing explicit offline setup tests create their own labelled disposable
fixtures using the already available reviewed wheel; they do not pre-create
the ordinary BAT test application's environment.

| Check | Actual result and limits |
| --- | --- |
| Baseline missing-package regression | FAIL on actual PS5.1.26100.9711 and bundled PS7.6.5: one expected `setup_required` assertion received `runtime_not_found` on each host. Retained as red evidence. |
| Final FirstRun components | PASS, 38 cases on each host, zero failures/skips. Real coordinator/resolver/setup-Check functions with explicitly synthetic discovery/setup/input seams: one consent cycle, fresh app environment, literals, explicit-base remedy/repair, healthy no-op, decline/EOF/cancel, unsafe/malformed/unsupported negatives, failed setup and read-only exclusions. These mocks do not prove native installation/resume. |
| Wider Runtime/Discovery/Manager/Setup/Outcome/Support | Original run retained as FAIL on each actual host (PS5.1.26100.9711 / PS7.6.6): 177 cases, 176 PASS and one obsolete package-free `.venv` assertion expecting `runtime_invalid` instead of `setup_required`. Interpreter authority/no-competing-manager assertions are unchanged. The other cases passed, including actual offline local setup, healthy no-op, backup-preserving repair, native failed-install/cancellation rollback, bounded probes, outcome and redacted-support safety. No production change followed this run; only that expectation was corrected. Final Discovery reruns: 40/40 PASS on each host, zero skips (31.91s / 24.67s); this is a focused rerun, not a relabelled 177-case pass. |
| Fixed PDF/application roster | PASS, 171 methods / ten required modules in 117.164 seconds, zero failures/errors/skips. All five engine Python files were byte-identical before/after. This run loaded the starting outcome map before concurrent PS/contract edits; the final changed outcome contract was tested separately below. |
| Final Python outcome contract | PASS, 11 methods in 2.590 seconds; final map/test bytes unchanged during execution. This overlaps the roster and is not 11 additional unique application tests. |
| Active runtime/native receipt contracts | PASS, 46 methods in 4.157 seconds. Updated current negative producers/validators require `setup_required` for missing/wrong pypdf and retain interpreter failure categories; swapped/obsolete codes are rejected. Synthetic receipts/oracles, not native acceptance. Historical report bytes are unchanged. |
| Package-input/support/runner/CI-report checks | PASS, 94 methods in 258.305 seconds, no skips. The committed-input dependency scanner accepts the added setup helper edge while retaining private/missing helper rejection. |
| Actual healthy PS5.1 application, repeated | PASS, two explicit pre-existing developer-runtime NonInteractive launches from unrelated CWD, with Unicode/bracket/ampersand paths and the combined authored RC4-40/local-GoTo fixture. Each returns exit 0 / `split_complete`, three chapters and all six page-content markers in physical order, expected manual ranges, one ignored-view warning, unencrypted outputs without OpenAction. No setup frame/prompt or `.venv` creation; distinct output directories and preserved source/prior output/first-run chapters. This is labelled healthy NO_INSTALL / NOT_COLD_START, not automatic discovery or consent/install/resume. |
| Final ordinary BAT / natural discovery | One legitimate prerequisite skip, zero test failures. The actual unchanged launcher/application exits 3 / `runtime_not_found`, with `python_security_floor` for existing Python 3.14.7. Final app SHA256 `e655c3e649677464c4eb478f3b02ea17def632cb81869d00a007a781c1d728ee` matches copied bytes. Input/app bytes unchanged, no `.venv` or chapters. Missing-package consent/install/resume remains NOT_RUN. |
| Independent document exclusion | PASS, actual setup Check and Repair with a synthetic inert book-directory decoy: rejected before any interpreter/setup process, zero operations/changes, authored old environment marker preserved. |
| Configured analysis and source review | PASS, zero analyzer findings across seven changed PS files with stable source hashes; independent review found no actionable defect. The final corrected Discovery test also receives a separate zero-finding check. Existing CRLF style retained; CRLF-aware whitespace check passes. |

Native Windows 10, supported existing Store first-run setup/resume, second-run
package-network observation, exact distribution ZIP and private-book acceptance
remain NOT_RUN here. The private book remains unavailable. No helper self-test,
synthetic manager or explicit offline fixture is promoted to ordinary acceptance.

## Commands and receipts

`$Repo` is this checkout. `$DevPython` is the pre-existing
`%TEMP%\WBS-623-finish-q6dwc7hk\dev14\Scripts\python.exe`; tools are under
`%TEMP%\WBS-pypdf-repair-0865p3vw\tools`. The existing reviewed wheel is in
`%TEMP%\WBS-623-finish-q6dwc7hk\wheelhouse`.

```powershell
& $DevPython -I -B "$Repo\tests\run_tests.py" --_application-units-child
& $DevPython -I -B "$Repo\tests\python\test_outcomes.py"
& $DevPython -I -B "$Repo\tests\python\test_first_run.py"
git -c core.whitespace=blank-at-eol,blank-at-eof,space-before-tab,cr-at-eol diff --check
```

The retained first-run driver imports pinned Pester and runs
`FirstRun.Tests.ps1` with repository/work parameters using New-PesterContainer.
Its directory is `%TEMP%\WBS-functional-cp3-first-run-components-20261030`.
The wider driver is `%TEMP%\WBS-functional-cp3-suites-20261030\run.ps1`.
The first-run driver uses actual PS5.1 / existing bundled 7.6.5; the wider driver
uses PS5.1 / existing 7.6.6. They use `-NoProfile -ExecutionPolicy Bypass -File`.
This is process-only policy handling, not an organization-policy or ordinary BAT
success claim.

The 94-method command uses `-I -B`, trusted test-directory imports and explicit
unittest modules `test_ci_package_inputs`, `test_support_receipt`, `test_runner`
and `test_ci_reports`. The 46-method command loads `test_runtime_receipt` and
`test_native_smoke` in the same way; its report/raw output are under
`%TEMP%\WBS-functional-cp3-runtime-propagation-fb2b52aaf55d4b8a8e5ae6ae850f3fe6`.
Healthy application commands use an argv array of inbox PowerShell,
`-NoProfile -ExecutionPolicy Bypass -File $Repo\WinBookSplit.ps1`, the literal
authored input/output paths, `-Mode Manual -StartPages 1,3,5 -NonInteractive
-NoPause -PythonPath $DevPython`. The exact driver/argv/output validation is in
`%TEMP%\WBS-functional-cp3-healthy-20261030-yah5p9xc`; its evidence.json SHA256 is
`bb4f2849f5539c4d30dd463ffde0cdf20b7196f94c07acadecda2f21a8998024`.

PDF roster output and before/after hashes are under
`%TEMP%\WBS-functional-cp3-pdf-20261030-dbd30f46`; final outcome output is under
`%TEMP%\WBS-functional-cp3-final-python-20261030-05e2d368`; final native driver
output is under `%TEMP%\WBS-functional-cp3-native-final-20261030-84868181`.
The actual native receipt is `%TEMP%\WBS-first-run-t7ap4c72\evidence.json`,
SHA256 `6d3a317cfdae4bda54add8851c4f3830c28baba7f91c6c6c1117693d3c9c90a4`.

Final configured analysis/review is retained under
`%TEMP%\WBS-functional-cp3-final-analysis-120198eb64eb46d1af0331920ebbb4ce`.
Earlier one-warning analysis of the pre-existing empty summary catch, the
subsequent source-unstable analysis and intermediate component failure are
retained separately. The empty catch now states its diagnostic intent through
Write-Verbose; final configured analysis is the authoritative gate.

Commit intended files, push normally and verify clean local HEAD against the
freshly queried live branch. Report the post-commit synchronization receipt in
the handover rather than putting a self-referential hash here, then stop.
