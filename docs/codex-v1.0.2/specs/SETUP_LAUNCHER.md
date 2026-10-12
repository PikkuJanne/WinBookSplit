# Reliable unsigned setup entry point

Add **Setup-WinBookSplit.bat**, a small wrapper around the existing Setup-WinBookSplit.ps1. No rewrite of the setup engine.

## Behavior

Use the trusted Windows PowerShell 5.1 executable from the actual Windows system directory, not a `powershell.exe` in the document/CWD/PATH. Account for 32-bit host redirection if needed without expanding Python architecture support. Launch with `-NoProfile`, a process-only execution policy and `-File` pointing at the sibling setup script. Resolve its directory from the launcher location, never the current directory. Use existing proven BAT quoting/argument conventions; keep delayed expansion disabled when required. Do not `CALL`/reparse an argument string, eval values, or embed user paths in a PowerShell `-Command` script.

The documented entry points are double-click interactive setup and, from PowerShell, a quoted absolute BAT path with setup switches. Supported forwarding includes `-NonInteractive`, `-Check`, `-Repair`, `-PythonPath` and `-Wheelhouse`. Do not add arbitrary process-host flags after `-File` or accidentally consume the setup `-NonInteractive` as only a host switch. Keep explicit repair authority. An unattended invocation cannot pause/read a key/open a window waiting for input; an interactive launch must leave useful status visible through an existing/simple pause mechanism without making CI hang. Preserve exit 0/3/6/130, and host-launch/binding failures as nonzero; do not convert cancellation or a signature block into success. Store the child exit before any echo/pause changes it.

The actual setup script still returns a [SETUP] record after it runs. A host rejection before script entry has no such record. Preserve stderr and nonzero exit; offer a short accurate policy hint without asserting that Python ran. Do not fabricate a setup result for a script that never started. No automatic retry under another security policy to evade an organization rule.

## Policy boundaries [S3, S4]

A process-scoped Bypass invocation is the existing product model for authorized unsigned scripts. It does not persist CurrentUser/LocalMachine changes and cannot override enforced Group Policy. No `Set-ExecutionPolicy` at those persistent scopes, no registry edits, no recursive Unblock-File, no elevation or bypass of enterprise application controls. Files need not be unblocked before this process-scoped launch on an ordinary non-GPO machine; keep their Zone.Identifier streams intact.

Do not claim that a blocked direct `.ps1` can repair its own load-time policy error. Setup cannot act before PowerShell allows it to run; this is why the BAT/explicit host invocation is required. `AllSigned` enforced by organization policy needs an approved signed/authorized route; v1.0.2 must fail clearly there instead of promising universal execution.

## Tests must use the artifact and exact shell path

Create `Zone.Identifier` with ZoneId=3 on the extracted setup .ps1 and shipped PowerShell helpers on an owned local NTFS test directory. Confirm the ADS exists. In a fresh process-scoped RemoteSigned child, direct unsigned script invocation must fail before setup (negative control). In that same parent context, invoke the actual BAT without first unblocking the files; setup and a subsequent PDF split must pass with the selected 3.15 runtime. Capture parent/child policy scopes, return codes, marker presence, actual runtime, ADS and before/after source/settings hashes.

Also test the full process-scoped host remedy printed to a PowerShell user; no byte-sanitized copy or pre-unblocked install may substitute. A fresh developer checkout or curl-only download that never had Internet metadata is not this test. Simulated ADS assignment is an automated Internet-marked-file regression, not evidence of a browser/Explorer download performed by a human.

Exercise invocation from System32/unrelated CWD, from PowerShell 5.1 and 7, as well as direct CMD/BAT invocation. Include normal Windows paths with spaces, apostrophes, brackets, ampersands, parentheses, Unicode, percent/exclamation and leading-hyphen filenames where representable. Pin failures in a supported path domain to explicit safe rejection rather than command execution; do not silently shrink the old accepted path domain. A special-character wrapper bug may be handled with the documented explicit PowerShell-host remedy when correct shell-safe quoting is tested.

Check missing setup file, no compatible Python, malformed switches, successful/check/no-op/repair states and cancellation. Do not assume a bare `%*` wrapper is correct without native argument/exit tests.
