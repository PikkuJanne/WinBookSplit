# Checkpoint 1 — failing application regressions

Checkpoint recorded with intentional failing regressions. The native
missing-pypdf scenario remains NOT_RUN for the prerequisite stated below.

The owner requested checkpoint 1 only in this chat. The supplied ZIP's
CODEX_TASK.md, CONFIRMED_FINDINGS.md, ACCEPTANCE.md and verified_diagnosis.json
were read as supporting material. The request controls scope and authority;
later checkpoint descriptions do not authorize publication or continued work
beyond this checkpoint.

Instruction ZIP SHA256:
`f33ab17eab774f4a65cf6e008949012e810cb0c10ce4c9bb9796f0d6524dc70d`.
Its authored generator was run with the existing reviewed developer environment;
fixture-generation/self-test results are not application repair evidence.

Starting checkout was clean main at
`bafa1868da10ab6ba65f112fe29b4bf4deb909f8`, tree
`c2ffea4a2f6b518d25589e2bb95266d087063849`, matching freshly queried live main.
Origin is `https://github.com/PikkuJanne/WinBookSplit.git`; only this worktree was
present. Work proceeds on `codex/functional-repair`. Live v1.0.0–v1.0.3 tags and
public releases were inspected. v1.0.3 still resolves to tested source
`7c5f7afeef9b09733ee8853d6df5d86c42d9c726`; its ZIP remains 938,143 bytes,
SHA256 `ba80fc35e0262f7f422dc1596d887e4a8207ec4ede94b754800b994084da4d20`.
No production files, release assets or historic evidence are changed here.
The new tests are registered in the existing Python/shell runners and fixed
application-runtime module roster; roster/count fixtures were synchronized.

## Production rejection sites

- `engine/winbooksplit_conversion.py`: `PolicyPdfReader._handle_encryption`
  rejects all encrypted files before ordinary library authentication.
- The same module's `_preflight_pdf_features` rejects every catalog OpenAction;
  its later object traversal independently rejects action dictionaries. Merely
  removing the catalog condition would leave valid local GoTo rejected.
- `engine/winbooksplit_engine.py`: `prepare_split(_captured_reader=...)` checks
  `is_encrypted`, which remains true after successful authentication. The
  interactive planner reuses its captured reader through this path.
- Conversion output validation uses the same PolicyPdfReader. `run_split`
  prepare/preview/execute and the interactive confirmation path must agree.
- `WinBookSplit.ps1` preflight catches missing pypdf, reports outer
  `runtime_not_found`, prints a separate setup action and exits. The ordinary
  BAT does not offer preparation/resume.

## Evidence boundaries

The existing developer environment is regular Windows x64 CPython 3.14.8 with
project-pinned pypdf 6.21.3 on Windows 11 build 26300. It is a test driver, never
injected into ordinary BAT discovery. No new Python or application environment
is provisioned for this checkpoint.

Actual automatic discovery on this workstation finds the existing python.org
Manager's Python 3.14.7 and correctly rejects it below the unchanged 3.14.8
floor. The native supported-base/missing-pypdf first-run scenario therefore
cannot be claimed as reproduced here. Its ordinary BAT invocation and exact
prerequisite result will be recorded; controlled summary tests are separate.
No Store-managed acceptance is claimed.

The private source PDF was not found in the checked local Downloads, Desktop
or Documents directories. Actual-book acceptance is NOT_RUN. Only authored
synthetic fixtures enter tests; no book bytes, text, images or chapters are
stored in Git or uploaded.

## Executed results

| Check | Actual result |
| --- | --- |
| New functional PDF regressions | **FAIL, expected before repair**. 11 methods: 2 fully pass, 9 contain failures/errors. unittest reports 28 failed assertions and 8 production policy errors. These include subtests, not 36 independent test methods. |
| Plain control | Manual, Auto1 and parent-aware Auto2 preview/execution pass with independently expected ranges, six physical pages in order, reopened outputs and unchanged source. Plain interactive confirmation and converted-PDF validation also pass. |
| Empty-password RC4-40 | All three modes' preview/execution fail at the blanket encryption hook. |
| Valid local GoTo and destination array | Both independently fail at the catalog OpenAction rejection, for preview and execution in all three modes. |
| Combined RC4-40 + local GoTo | Fails at encryption. A separately publicly authenticated reader reaches and fails the catalog rule, exposing the second blocker without modifying the source. |
| Authenticated captured reader | Public decrypt succeeds and `is_encrypted` stays true; all three modes fail at the separate captured-reader rejection. |
| Shared callers | Empty-password/local-view interactive paths and converted-PDF validation fail through the same production policy. |
| Genuine nonempty password | Fails the desired classification: `unsupported_document` instead of `password_required`; no outputs. |
| Active/malformed controls | JavaScript, chained GoTo + JavaScript, remote GoTo and malformed local destination remain rejected. No action is executed. |
| Missing-package classification | **FAIL, expected before repair**, one Pester test on each actual host: Windows PowerShell 5.1.26100.9711 and PowerShell 7.6.6. Real resolver/summary with explicitly synthetic discovery input returns `runtime_not_found` instead of `setup_required`. No interpreter or installer is run by this component test. |
| Ordinary unchanged BAT | Actual Windows source-copy BAT invocation runs with inherited discovery/environment and returns exit 3 / `python_security_floor` for Python 3.14.7. No venv, chapter output or source/application-byte change. The desired missing-package/setup-offer case is **SKIPPED / NOT_RUN**, not a setup pass. |
| Existing PDF policy suite | **PASS**, 22 methods. This still includes the old blanket-ban regression because production is intentionally unrepaired at checkpoint 1; checkpoint 2 must replace it. |
| Runner/report integration | **PASS**, 15 relevant harness methods; fixed application module roster is 10. This is harness evidence, not application repair evidence. |
| New PowerShell regression static analysis | PSScriptAnalyzer 1.25.0: zero findings. |
| Diff review | No production behavior/dependency/discovery changes; whitespace check passes. Independent review checked production call paths and evidence limits. |

The 11-method PDF suite contains 46 subtests: 27 assertion failures and seven
policy errors, plus one standalone assertion failure and one standalone policy
error. Source, neighboring prior output and absence of false publication are
checked even when positive cases fail. Successful plain output checks compare
decoded page-content bytes in physical order and confirm unencrypted outputs
without copied OpenAction. The new contract requires exactly one
`opening_destination_ignored` warning for an accepted initial view; unrelated
warnings cannot satisfy it. This warning is not implemented yet.

No full-suite/Calibre/runtime-matrix/exact-package pass is claimed here. No
helper self-test is counted as application evidence. No release/package is
built, tagged or published.

## Commands and retained receipts

Existing developer interpreter used (no installation/preparation):
`%TEMP%\WBS-623-finish-q6dwc7hk\dev14\Scripts\python.exe`.
Existing isolated Pester 6.2.0 / analyzer tools:
`%TEMP%\WBS-pypdf-repair-0865p3vw\tools`.
Raw logs and component driver are retained outside Git at
`%TEMP%\WBS-functional-repair-20261030`.

Actual focused Python commands, with `$Repo` set to this checkout's absolute
path and `$DevPython` to the existing interpreter above:

```powershell
& $DevPython -I -B "$Repo\tests\python\test_functional_pdf_repair.py"
& $DevPython -I -B -m unittest discover -s "$Repo\tests\python" -p test_first_run.py -v
& $DevPython -I -B "$Repo\tests\python\test_pdf_policy.py"
& $DevPython -I -B -c "import sys,unittest; sys.path.insert(0,r'$Repo\tests\python'); suite=unittest.TestSuite([unittest.defaultTestLoader.loadTestsFromName('test_runner.ApplicationRouteTests'),unittest.defaultTestLoader.loadTestsFromName('test_ci_reports')]); result=unittest.TextTestRunner(verbosity=2).run(suite); sys.exit(not result.wasSuccessful())"
```

The focused Pester driver imports the explicit pinned module and runs only
`tests/powershell/FirstRun.Tests.ps1` through `New-PesterContainer`, passing the
repository and external owned work directory. It was executed under inbox
PowerShell 5.1 and the existing 7.6.6 host with `-NoProfile -ExecutionPolicy
Bypass -File`; exit 1 on both hosts. This is process-only policy handling, not
a persistent policy change or company-policy acceptance. The test is also
registered in the ordinary shell runner.

Selected raw receipt SHA256 values (raw files remain private/local):

| Receipt | SHA256 |
| --- | --- |
| `pdf-focused-red-final2.txt` | `d9ef61be1fd90af490bfb8eb09a0a148a2e2c1f74e560ee8ab07772b35c6d26e` |
| `pdf-roster-checks-final.txt` | `36d4b178e3612b4fba8579e5bb0dfcf62ef0108c09765b9aa9350544d09b1e10` |
| `baseline-pdf-policy.txt` | `ce43fa47a3427560b272422a93700026780aadc0267da6e43fb1e34b7a323762` |
| `first-run-ps51.json` | `a1a1a7f5c1e53c9593f12df85ef350f7d10479595daf9ebfdcab63ba0421f307` |
| `first-run-ps7.json` | `47fdff4ae042dd98e23786acdfee952c1df1f2011fb39e0e3a84b0bb22601215` |
| `first-run-native-final.json` | `f266f2aa87578d17cdefe8ed4cc77e85cf03519d56a7894572518902cace5de3` |

The final native-test rerun contains one legitimate prerequisite skip and no
test failures. Its preflight validates structured failure codes/attempts and
does not convert unknown resolver errors or unexpected package failures into
skips. BAT and its outcome record must agree on exit code. The original native
work directory and the copied final receipt remain outside the repository.

Tested production bytes equal the starting source above. Final checkpoint commit
and clean live-branch synchronization are verified after commit/push and reported
in the handover, avoiding a self-referential commit hash in this document.
