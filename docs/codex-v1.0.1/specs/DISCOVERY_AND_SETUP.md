# Discovery and explicit setup contract

## Discovery
Keep explicit `-PythonPath` and an existing application .venv authoritative. Then use tested **read-only** legacy-launcher/manager listing and independent trusted direct PATH discovery. A bounded PEP 514 registry-location fallback is optional where useful; do not scan whole disks.

The current code already falls through after `py -0p` fails. Retain that and add regression coverage for the supplied `Unknown option: -0` case. Do not assume every executable called py.exe is a particular launcher. Detect/query the appropriate installed-only listing method; the official manager supports `py list --format=exe` and retains legacy `-0p` [S1]. Prefer unambiguous manager invocation when available and validated. A failed optional listing must not discard valid PATH candidates or generate a frightening failure dump after successful fallback.

Never call `py exec`, `py install`, an online listing, a zero-argument launcher, or an invocation that may auto-install a runtime during discovery/processing. Apply child-local automatic-install disablement and bounded process controls; retain protection against document-directory/current-directory decoys, environment-module shadowing, alias tricks and unsafe paths. No wholesale exemption for WindowsApps/reparse points: where manager aliases cannot meet existing trust rules, use a narrowly audited route to actual runtime paths or documented direct-path fallback. No elevation, registry edits, PATH changes, global environment changes or remote scan.

## Setup-WinBookSplit.ps1
An explicit setup entry point, not an installer framework. Proposed stable interface:

- `-PythonPath`: optional trusted base interpreter.
- `-Check`: read-only readiness diagnosis; never install or create a venv.
- `-NonInteractive`: explicit invocation authorizes this script's normal local setup/download action and disables all prompts; return clear failure codes.
- `-Wheelhouse`: optional approved local offline directory; use only the hash-locked artifacts and no index/network fallback.
- `-Repair`: explicit replacement of an unhealthy application .venv with safe rollback; never implied by an ordinary run.

Plain interactive setup explains the selected interpreter, destination and package source once and requests consent for changes. Validate the interpreter **without requiring pypdf**. A healthy existing .venv is a validated no-op. An incompatible/damaged existing .venv is not deleted or bypassed implicitly. Use `-Repair` or actionable instructions. Avoid adding `-Force`/unsafe trust overrides.

Create the environment using that concrete interpreter, then invoke its concrete `Scripts\python.exe -I -m pip`, not bare pip. Use isolated/no-input hash-required binary-only installation from the reviewed source or supplied wheelhouse and verify dependency consistency/import origin afterwards [S8]. Preserve TLS validation. Do not upgrade unrelated packages, use global/site-user installs, or execute downloaded shell content. Runtime dependencies and developer dependencies remain separate. Record actual package versions and permitted downloads.

### Avoid two common implementation traps
**Bootstrap catch-22:** setup must discover a compatible Python with no pypdf installed. Reusing the full processing probe without an interpreter-only mode defeats setup.

**Venv relocation:** do not construct a complete venv under a temporary name and blindly rename it into .venv; virtual environments are not generally relocatable [S6]. Create the venv at its final validated location under an exclusive setup lock. Track ownership so an interrupted first install can remove only its own incomplete files. For explicit repair, preserve the old environment under a backup identity, create the new environment at the original final path, and restore the old environment to its original path on failure. Do not use a moved backup as a working environment. Test rollback, held files and interruption. If ownership or path safety cannot be established, preserve files and fail with a clear next step instead of recursive deletion.

Refuse setup in unsafe/reparse destinations and serialize two simultaneous setup attempts. Temporary lock/marker/log writes must be bounded and owned. Unwritable app roots get a clear per-user extraction recommendation, not UAC elevation. Moving the application later requires environment revalidation/recreation, not a promise that .venv is portable.

No Python or Calibre download/install during ordinary application setup in this patch. Missing Python gets compatible-branch instructions without demanding one exact patch. Separate developer acceptance-kit provisioning may fetch approved isolated official runtimes as explicitly described in TESTING.md.

## User-facing behavior
Keep drag-and-drop unchanged after setup. Ordinary failures show a brief reason and one actionable setup command, with structured technical details available in existing diagnostics. No global policy changes and no automatic Unblock-File across arbitrary trees. Respect organization policy and explain unsigned scripts honestly.
