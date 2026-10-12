# Minimal owner involvement

**Default: zero manual functional checks.** The owner has already done too much compatibility troubleshooting. Do not recreate the first-release manual viewing/GUI programme. Run tests yourself in authorized local/CI/native environments, reuse the existing kit, and make the implementation robust before asking the owner to run anything.

## Only when native Windows 10 access is absent

Finish source tests, launcher/MOTW tests on available Windows, docs/packaging and internal kit validation first. Prepare ONE consolidated per-user verification kit, extending existing `tools/compatibility/kit_builder.py`, `Invoke-NativeAcceptance.ps1` and `native_smoke.py`. Supply one exact command that starts the kit reliably despite Internet marks. No preliminary six-command dependency install instructions or repeated screenshots.

The kit uses the exact candidate ZIP and authored fixtures; carries/locates a validated runtime-only test environment through an explicit documented setup step; selects the owner's regular 3.15 when usable; and never asks them to downgrade to 3.14. It checks actual executable/build and packages, performs setup/check/repair, emits/captures actions, sets Zone.Identifier only on its owned test files, runs targeted splitting and reports. It does not modify the user's installed application, working `.venv`, original books, global Python, PATH or persistent security settings. Any separate runtime/package bytes must have reviewed source/hash/license and remain a private testing resource, not an unnoticed addition to the public runtime ZIP.

The owner action is run the supplied launcher/command once and return the produced redacted report. Tag this USER_LAUNCHED_AUTOMATED, not manually inspected or agent-launched. Prepare all relevant cases for one session; do not request separate Python-branch/GUI/fidelity checks. A real defect can justify a focused rerun after correction, not a routine repeated acceptance ceremony.

## What is not required

No manual chapter inspection, no compulsory dragging through Explorer again, no approving every task/commit/release, no screenshots of dependency versions, no paid VM or signing certificate. BAT invocation/argument forwarding is tested automatically. A changed setup BAT can be automated; do not impose a GUI waiver ceremony. Note exactly that Explorer was not manually exercised if it was not.

## Evidence boundary

Actual Windows 10 execution with 3.15 on the release package remains mandatory; old 3.14/Windows 10 evidence or Server CI is not a substitute. When access is blocked, complete independent work, provide the validated kit and record the remaining exact gate honestly. Do not publish falsely, quietly remove 3.15/Windows 10 from scope, or promise unattended background work.
