# Actionable and executable diagnostics

Keep technical records in existing local logs. Show the actual selected/authoritative executable path, observed Python version and specific reason in a concise console summary. Do not dump all probe JSON by default. Redact account/document paths from shareable reports; never upload raw logs automatically.

## Required distinctions

| Situation | Meaning and remedy |
| --- | --- |
| Supported 3.15 base, pypdf absent | Setup required, not unsupported Python. Run setup to create the local environment. |
| 3.14.2 | Below reviewed branch security floor. State observed version and current floor; do not say no Python was found. |
| 3.16+ unknown minor | Not yet validated for this release. Do not call it unsafe or claim the splitter failed. |
| x86/ARM/free-threaded/PyPy | Explain the build/architecture requirement, not a security-floor mismatch. |
| Missing/bad pypdf | Distinguish missing package, wrong version and untrusted origin. |
| Broken or below-floor `.venv` | Name it as authoritative; preserve it; propose explicit setup repair with a valid base. PATH version is not the selected venv version. |
| Failed optional `py -0p` but good direct Python | Continue discovery; do not let an optional listing dominate the successful result. |
| PowerShell blocks setup before entry | Execution-policy/host failure; cannot be reported as a tested Python failure. |

Prefer structured existing reason codes and additive fields instead of gratuitously changing the outcome protocol. If a protocol shape must change, version it and update every consumer/test. Unknown malformed input remains a genuine error. Version text in console/JSON must reflect the interpreter actually probed.

## One maintained remedy formatter

Build current setup suggestions from one tested helper/template and the real application path. Use either the new setup BAT or a complete explicit trusted PowerShell-host `-NoProfile -ExecutionPolicy Bypass -File ...` command. The latter is a useful PowerShell-safe route for characters which CMD would reparse. Bare `& '<app>\Setup-WinBookSplit.ps1'` is not an acceptable active default remedy. Label the target shell and quote for it, including apostrophes in PowerShell literals. Paths/options are data, not executable code. Do not include secrets or automatically execute a suggested command in production.

Examples in help, README, docs/SETUP.md, diagnostics and release notes must agree. Preserve `-Check`, correct `-Repair` selection and an appropriate explicit base `-PythonPath` where discovery cannot decide; never suggest the invalid venv interpreter as its own repair base. A placeholder-only 'pick a compatible Python' action is not sufficient when a suitable 3.15 interpreter is already discoverable.

## Execute the emitted action in regression tests

Use an authored finite test fixture to capture the exact Action emitted by a dependency failure, not a manually reconstructed equivalent. For a compatible 3.15-only installation with missing pypdf, run that action unchanged in a fresh RemoteSigned parent against Internet-marked scripts. Then run the original operation; it must succeed. For a stale `.venv`, run the emitted explicit repair action with the discovered valid base and verify backup retention. Include the System32 CWD case and a path with spaces/apostrophe. Capture exit/result and before/after state. Do not evaluate arbitrary user log text outside this controlled synthetic test.
