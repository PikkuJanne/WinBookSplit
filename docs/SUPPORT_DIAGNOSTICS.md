# Export a local support summary

After processing setup, an application run keeps a bounded UTF-8 `console.log`
in its unique console folder. Successful diagnostic finalization publishes its
neighbor `WinBookSplit_Run.json`, including the actual failed or cancelled
outcome when processing did not succeed. These original files can contain book
paths, chapter titles, error messages, and other local details. Keep them under
your control.

The log reserves 32 MiB for ordinary records and 16 MiB for the final outcome,
for a total cap of 48 MiB. A record that exceeds its remaining space is refused;
it is never partly written. The original process streams are still displayed.
A diagnostic write failure preserves a primary nonzero processing or cancellation
exit. If processing succeeded but diagnostics cannot finalize, the application
reports `console_finalize_failed` with exit `6`, retains completed engine output,
and omits `Done.` A corrective outcome is appended when the held log permits it.

Manifest publication uses an exclusive pending file and occurs only after UTF-8
write, flush, and close. On publication failure the application attempts to mark
the pending record with `diagnostics_finalized: false`. If that corrective write
also fails, the retained file remains unpublished. A pending filename is always
rejected by the export command. Existing files are never overwritten.

Preview and early validation or dependency preflight create no disk diagnostics.
When cancellation occurs before the engine captures a plan, the local source
identity is explicitly `metadata_only` with no invented hash. Once captured,
the local plan records the processed source identity and the exact page ranges.
Full local runtime versions remain in the original manifest; the redacted
summary uses the fixed PowerShell buckets described below.

To create a separate redacted summary, choose the finalized
`WinBookSplit_Run.json` and a new local destination file. Run the shipped command
from Windows PowerShell 5.1 or PowerShell 7, using absolute paths:

```powershell
& 'C:\Tools\WinBookSplit\Export-WinBookSplitDiagnostics.ps1' `
    -ManifestPath 'C:\LocalRuns\ChosenRun\WinBookSplit_Run.json' `
    -OutputPath 'C:\LocalRuns\ChosenSupportSummary.json'
```

These paths are examples. The destination directory must already exist. The
command creates the destination exclusively and never overwrites an existing
file. Success returns native exit code `0`; a rejected export returns `2` and
a fixed error message. No export is automatic, and nothing is uploaded or sent
over the network.

The summary contains the application version, recognized runtime versions,
split mode and input kind, Boolean options and bounded timeout values, physical
page ranges and coverage counts, the final outcome code and file count, known
warning categories and counts, and bounded log size counters. Page ranges use
zero-based starts and exclusive ends. The full partition must cover every
physical page once.

It excludes paths, titles, filenames, timestamps, run IDs, source hashes and
identities, command arguments, environment values, messages, and arbitrary
free-text diagnostics. Parser warning messages remain only in the original
local record; the summary retains known warning or error codes and counts.
PowerShell versions are validated as bounded supported numeric versions and
reduced to the fixed `5.1` or `7` bucket; build text is omitted. Python, pypdf,
and Calibre version text must match the shipped trusted version allowlist.
Unknown versions are rejected instead of copied into the summary.

Only version `1` of the finalized `winbooksplit.run` schema is accepted, with
`diagnostics_finalized` set to Boolean `true`. Pending records are rejected.
Input must be strict UTF-8 JSON, at most 32 MiB, with no duplicate or ambiguous
case keys. Unknown root fields and invalid selected fields are rejected. The
summary is capped at 1 MiB. Source and destination must be ordinary local files
with no reparse point in either ancestor path. Relative, device, network, and
alternate-data-stream paths are rejected. The original record is held read-only
while the summary is validated and written.

The summary is intentionally limited. Review it before sharing it yourself. A
failed or incomplete run remains failed or incomplete; the command does not
turn a retained partial result into a success claim. If writing a new summary
fails during I/O, a partial new destination may remain and the command reports
failure. It does not remove that file or touch the original diagnostics.
