# Scope and authority

## Required outcome

Repair automatic discovery of an existing Microsoft Store Python Install Manager
installation and its regular x64 Python runtime on Windows 10. Preserve Win11,
regular supported 3.13/3.14/3.15 runtimes, the BAT shortcuts, PS5.1 console entry,
PS7 entry, existing setup, PDF engine and optional Calibre conversion. Package and
publish the intended next patch v1.0.3 after native Store-route verification.

This continues the owner's existing implementation/synchronization/publication
workflow. Ordinary scoped commits/pushes, PRs, passing reviewed merges, explicit
application-local package setup and this release are within the requested work.
Tool approvals and organization controls still apply. A stale AGENTS restriction
to v1.0.2 must be reconciled visibly, not silently ignored.

## Do not change to hide the failure

- Do not install/reinstall/update Python or Python Manager on the target desktop,
  modify its PATH or alias settings, or ask the user to switch to python.org.
- Do not pre-create a venv or pass a test-selected runtime to the application and
  call that automatic discovery. Do not copy a private runtime into the test kit.
- Do not remove trust checks globally, lower Python security floors, broaden
  pypdf versions blindly, or suppress failed probes.
- Do not hard-code a username, desktop path, Python folder spelling, manager
  package version or the four hashes from the captured installation into runtime
  acceptance. Those values identify a regression, not a product allowlist.
- Do not rewrite the splitter, replace the UI, build a website, add telemetry,
  upload documents, bundle a Python distribution, or buy signing/services.
- Do not retag/reupload/edit published v1.0.0-v1.0.2 assets, reset history,
  force-push, delete unrelated files, auto-stash or change repository visibility.

Keep ordinary venv creation, pypdf installation, repair and authoritative explicit
runtime behavior. A compatible base without pypdf is setup-required, not absent.
Readiness checks must not install; normal processing must not install. No
permanent execution-policy changes, elevation or bypass of enforced controls.

## Work discipline

Eight tasks, one completed task per Codex thread. Provide partial progress and
specific blockers. Pause after each task, not each tiny edit or test invocation.
Do not ask the user to repeat decisions already recorded. No success claims from
static/Windows Server/mocked results for the actual desktop. A missing native
gate blocks publication, not all remaining independent work.

## Dependency/security scope

Retain the current runtime policy and pypdf/Calibre lock unless a newly verified,
release-relevant issue requires a separate small remedial change with evidence.
Record current advisory review and scope honestly. Do not turn an unrelated
advisory into a reason to redesign discovery or claim the app is a sandbox.

## Deferred native checks without repeated owner requests

Task dependencies order implementation checkpoints, not repeated owner sessions.
If code/tests can advance but a native case needs the owner's desktop, record
IMPLEMENTED_PENDING_NATIVE and the exact deferred acceptance ID, then proceed
with independent next work. Do not mark the case PASS or the task fully COMPLETE.
Consolidate owner-only cases into T07's single kit. All required cases must pass
before T08 publication. Use an available authorized agent Windows environment
for early alias/supervisor tests whenever possible; do not wait to discover
obvious launch errors on the owner desktop.
