# V101-04 independent read review

Review date: 2026-10-11. Repository: <checkout>. No repository source, tests, status, commit, index, GitHub settings or release files were edited by this review.

Final reviewed engine/WinBookSplit.Runtime.ps1 SHA-256: 611e1b633aea7ad9458922443209b7f67f10ef7ba7ce4129e7604323e7b18211. Review is source inspection plus the limited native command below; it is not a final-ZIP acceptance report and does not establish Windows 10 compatibility.

## Source review result

No blocking defect found after the final narrow optional timeout change.

- Explicit Python and an existing application .venv remain authoritative. Presence/type/permission/reparse errors are failures rather than permission to switch runtimes.
- Independent direct PATH candidates are guarded/probed before optional listing. Stable selection follows policy branch preference, numeric patch ordering and ordinal canonical path ties. Candidates are deduplicated; bounded candidate/list/launcher counts and an overall process-discovery budget remain explicit.
- Unknown -0 exit-2, malformed/empty listings and safely stopped optional listing timeouts remain diagnostic failures that permit valid PATH fallback. Optional exhausted budget skips later optional work. Only the optional listed-interpreter loop catches processor_timeout when the probe is timed out, not cancelled, not FatalDiscovery, both parent and descendants are proved stopped, and no StopError exists. All other errors rethrow; direct PATH, explicit and .venv probes retain fatal timeout handling.
- Arbitrary PATH pymanager.exe is never executed. Only independently verified current-user PSF python.org package registration authorizes physical manager enumeration. The fixed query uses the OS-known inbox Windows PowerShell path, a static encoded command, explicit inbox Utility then Appx manifests and disabled module autoload; no discovered path or identity enters child code. The exact OS-known inbox shell retains ordinary/reparse/canonical checks but allows its observed OS servicing hardlinks. All discovered managers, legacy launchers and interpreters retain Automatic alias/hardlink guards.
- Parent registration validation requires exact Name, signing Publisher, PublisherId, FamilyName, full-name/version/architecture relation, SignatureKind Developer, Status Ok, four false incomplete/development/framework/resource flags and a canonical physical path at the OS-known ProgramFiles WindowsApps location. This is a narrow tested python.org identity, not a universal Store or activation guarantee.
- Manager and legacy listing use only option-only -0p. The final parser accepts the native bracketed optional architecture tag, absolute literal executable paths and existing punctuation without shell evaluation. No installer, online list, registry edit, PATH/global policy/environment mutation or zero-argument launcher call was introduced.
- The shipped policy and dependency origin guards continue to run after concrete executable selection. No separate runtime policy or package framework was introduced.

## Why registered manager -0p was selected

An arbitrary executable merely named pymanager.exe may be renamed CPython, where 'list' is a script operand. It cannot authorize a modern 'list --format=exe' call. The modern manager ListCommand also creates a temporary log and may update a welcome marker. The installed manager's package configuration explicitly has welcome_on_update=true.

Pinned upstream source release26.4, commit3f6af8deae987ceab472c68bfaa864e911c86dc0, published2026-10-09, confirms ListLegacy0pCommand inherits _create_log_file=False and a no-op show_welcome. Logger starts with file=None. Inherited PYMANAGER_DEBUG/VERBOSE affect console threshold; they cannot enable the legacy command's file logging or welcome writes. PYMANAGER_IMPORT_TIME can add diagnostics to stderr, which are bounded/rejected rather than executed. Modern listing is absent from the final discovery execution route.

Primary sources reviewed:
- https://docs.python.org/3.14/using/windows.html#listing-runtimes
- https://docs.python.org/3.14/using/windows.html#advanced-installation (official package identities and direct ProgramFiles activation limitation)
- https://github.com/python/pymanager/blob/26.4/src/manage/commands.py
- https://github.com/python/pymanager/blob/26.4/src/manage/logging.py
- https://github.com/python/pymanager/blob/26.4/src/manage/__init__.py
- https://github.com/python/pymanager/blob/26.4/src/pymanager/main.cpp
- https://learn.microsoft.com/en-us/powershell/module/appx/get-appxpackage?view=windowsserver2025-ps

## Actual limited native observation

Host: Windows11 version10.0.26300.0; reviewer shell7.6.5. This host is not Windows10 22H2.

The actual current-user installed physical manager was PythonSoftwareFoundation.PythonManager_26.4.240.0_x64__3847v3x7pw1km\pymanager.exe at the registered ProgramFiles WindowsApps location. Existing ordinary Automatic guards accepted it. The reviewer dot-sourced the current Diagnostics/Runtime and used existing Invoke-WinBookSplitRuntimeProbe with arguments @('-0p'), ApplicationRoot <checkout> and TimeoutSeconds10.

Actual result: exit0; elapsed0.1593047s; parent+descendants stopped, JobAssigned=true, StreamsComplete=true; no stderr/start/stop/stream error, timeout, cancellation or truncation. One concrete legacy row '-V:3.14[-64] * %LOCALAPPDATA%\Python\pythoncore-3.14-64\python.exe'. The welcome marker LastWriteTimeUtc ticks were639272562163944895 before and after. This command exercised installed-manager listing and the unchanged owned transport; it did not rerun the entire final resolver, install pypdf, process a document or test the final ZIP. Final focused suite results are owned by the implementation/native-runner agents.

The earlier modern-list research receipt demonstrates installed-only enumeration and no installation/global-setting action, but must not be interpreted as proof that modern listing never creates incidental log/welcome files. The final route avoids that behavior with -0p.

## Related existing version fixture

The broad task03 Python run executed559 tests and reported two failures: the old28-versus31 payload assertion and test_version.py expecting dependency_missing from a fixture containing only engine+outcomes. The latter now reaches the newly required missing compatibility helper first and correctly reports runtime_invalid. Parent added the real helper+policy copies without removing assertions. Read review confirms VERSION remains validated before loading those helpers; valid VERSION with -S can then test missing pypdf intentionally. Parent reported actual isolated dev14 test_version.py5/5 pass in0.784s; this was not a reviewer execution.

Checkpoint: independent source review complete, no blocking source defect. Parent owns final native suite results, evidence synchronization and subsequent setup task. Native Windows10 acceptance remains outstanding.
