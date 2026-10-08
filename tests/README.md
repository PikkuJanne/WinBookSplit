# Local verification

Use the explicit fresh developer-venv interpreter from
`docs/codex-v1.0.0/SUPPORT_AND_SETUP.md`. Run the same entry point from any
working directory; it resolves the trusted repository from its own location.
Do not activate a venv or install dependencies globally.

```powershell
$TestPython = 'C:\Tools\WinBookSplit\.venv-dev\Scripts\python.exe'
$Runner = 'C:\Tools\WinBookSplit\tests\run_tests.py'
$ToolRoot = 'C:\Tools\WinBookSplit-test-tools'
$PS51 = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'
$PS7 = 'C:\Program Files\PowerShell\7\pwsh.exe'
$Report = Join-Path $env:TEMP ('WinBookSplit-tests-' + [guid]::NewGuid().ToString('N') + '.json')
& $TestPython -I -B $Runner --layer full --tool-root $ToolRoot --shell-path $PS51 --shell-path $PS7 --report $Report
if ($LASTEXITCODE -ne 0) { throw 'Local verification failed; inspect the report' }
```

Substitute actual absolute interpreter/tool/host paths. Every report must be a
new absolute path outside the checkout. Existing reports are refused, and
reports containing commands or local paths should stay outside Git. Commit only
the reviewed summary with hashes in `docs/codex-v1.0.0/evidence/`.

Use `--layer python`, `--layer shell` or `--layer baseline` for targeted checks.
The shell layer requires explicit hosts and the isolated exact-version modules;
the full layer must include both supported hosts for the milestone gate.
`--failure-probe native`, `--failure-probe python` and
`--failure-probe pester` deliberately fail and must return nonzero even after
the runner's later successful command/report step. These are runner acceptance
probes, not application defects. The Pester probe requires the shell arguments.

The shell bootstrap executes only the authored repository test script in a
fresh `-NoProfile` host, using an encoded command and trusted script block.
The child requests process-only `RemoteSigned` so Windows PowerShell 5.1 can
import the isolated tools when its default is `Restricted`. Machine/user policy
is unchanged; enterprise policies take precedence and a blocked import fails.
This developer invocation does not certify the application's ordinary launcher
under the workstation's policy.
Each child uses its own host's built-in module directory and the absolute
external pins, so an inherited PowerShell 7 module path cannot supply tools to
Windows PowerShell 5.1.

Fixture PDFs, slices, Pester TestDrive files and child working directories live
inside unique temporary directories. Cleanup removes only each run's owned
directory. No user document is a test input. Source documents, environments,
generated outputs and raw evidence are ignored; ignore rules do not replace
reviewing the exact staged paths.

The baseline layer executes the unchanged engine and verifies its deliberately
known-bad expectations. A pass means reproduction, not repaired splitting.
The shell layer gates syntax and the new scaffold's selected static checks;
unchanged application analyzer findings are reported as legacy observations.
Real splitting under both shells, Explorer, Calibre and the extracted release
package remain separate mandatory later checks. CI installation and workflow
execution belong to M5-T01; this local scaffold does not claim a CI run.
