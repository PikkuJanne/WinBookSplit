# V102-05 reproduced CMD setup output encoding defect

Observed 12 October 2026 during actual setup launcher testing. Source baseline was `22ee5fbace6ff2f80ea38e2876a66f5be096a3d7`, with the task-5 changes in the working tree. Testing used owned authored application copies and an actual regular Python 3.15.0 runtime; no owner source-ZIP installation was changed.

## Reproduction and outcome

The direct CMD route created a valid local 3.15 environment and returned native exit `0`, but the setup script wrote its structured result using OEM encoding at a Unicode application path. The existing strict UTF-8 native transport reported:

`stdout: Unable to translate bytes [8F] at index 114 from specified code page to Unicode.`

The receipt recorded `ParentStopped=true`, `DescendantsStopped=true`, and `StreamsComplete=false`. Its `[SETUP]` output was truncated within `environment_path`. This route is a reproduced failure, despite the child exit and completed environment; neither transport success nor complete structured-result acceptance is claimed.

Raw private receipt identifier: `task05-Actual-86256832e9d846b6b1bbe1d7cd3b6a8a/CMD/fresh-process.json`. Actual SHA-256 `77a1bfb1faf5e1b9dbae864493d01c7c161883e1afc1bde867a032739b294819`, size `1383` bytes. Raw account paths and process output remain outside the repository.

The setup script before the fix had SHA-256 `4cb6b2242604a0b4c6e3708dcf51c9237e712a3169ea5086edcddf8c7f3007e5`. The BAT SHA-256 was, and remains, `b756dffa3e4eb1b5adb04d4bc06ae1bd87718b42ca42183572dc6ef820fa6500`.

## Minimal fix and checks

Setup now initializes `[Console]::OutputEncoding` and `[Console]::InputEncoding` with `System.Text.UTF8Encoding($false)` immediately after parameter binding. This reuses the existing WinBookSplit entry point's encoding behavior. It makes no persistent policy, registry, Python, PATH, dependency or source/output-safety change.

Fixed setup SHA-256: `d5e5f87705d176e87189141f8840e42ffc4b66a2fb7363400102dfeb17bee096`. The fixed script passed actual parser and PSScriptAnalyzer `1.25.0` checks under inbox PowerShell `5.1.26100.9444` and reviewed PowerShell `7.6.6`, with zero findings/errors and unchanged source hashes during both checks.

The external static-check driver uses `-NoProfile -ExecutionPolicy Bypass -File Check-SetupSource.ps1 -RepositoryRoot <checkout> -AnalyzerManifest <reviewed-module>` with literal native arguments. Driver SHA-256: `620f93cac2802bff99617064424cc698fb9e64b2663a596ed1c26eb22e6857a0`. Initial and fixed-source static receipts are retained separately; checks on the earlier script are not relabeled as fixed-source checks.

Subsequent native testing reran all 18 actual setup stages across inbox PowerShell, reviewed PowerShell 7.6.6 and direct CMD on fresh fixed-source copies; all passed, including strict UTF-8 result decoding on the reproduced CMD route. The actual hidden-console zero-argument setup/result/Enter wait also passed on the fixed source. Their raw/source identities and scope are retained in `V102-05-native.json` and `V102-05-native.md`.

The earlier 30 controlled BAT argument/exit tests used the unchanged BAT bytes and retain their original scope. Earlier actual successes and the OEM decoder failure remain history; they are not relabeled as fixed-source reruns. This note is not final-package or Windows 10 acceptance.
