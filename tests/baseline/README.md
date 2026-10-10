# Original behavior characterization

These checks execute the unchanged Python here-string from `WinBookSplit.ps1`
in an isolated temporary directory, using real synthetic PDFs and reopening
every written slice to read its original page IDs. They do not change or import
the application at runtime. The original PowerShell and batch files are pinned
by SHA-256; a changed launcher makes this harness fail rather than apply its
known-bad expectations to repaired code.

After M1 extraction, invoking this historical command deliberately fails on the
changed PowerShell hash. Use `tests/extraction/characterize_extraction.py` or the
central `--layer extraction` for explicit immutable-original/extracted-source
comparison. Neither the original hash guards nor the expectations are changed.

`expected_original.json` holds explicitly original results, including success
with zero files and omitted pages. `docs/codex-v1.0.0/PLAN_ORACLES.json` remains
the independent corrected-behavior oracle. Each report includes both results.
A successful characterization means the original behavior was reproduced; it
does **not** mean those defects pass fixed-version acceptance. M1 connects the
target oracles to the extracted/repaired engine separately.

Use an explicit Python with pypdf and ReportLab available. M0-T02 used the
bundled isolated Codex runtime, not the default Python missing pypdf; this does
not select the supported application dependency matrix.

```powershell
$FixturePython = '<absolute bundled Python path>'
$ReportPath = Join-Path $env:TEMP ('WinBookSplit-baseline-' + [guid]::NewGuid().ToString('N') + '.json')
& $FixturePython -I .\tests\baseline\characterize_original.py --launcher-probes --report $ReportPath
if ($LASTEXITCODE -ne 0) { throw 'Baseline characterization failed' }
```

Use a new report path; existing reports are refused. On a non-Windows developer
host omit `--launcher-probes` and explicitly record that launcher probes did not
run. The entry point works from another working directory when its absolute
script path is supplied. Its children use the same explicit isolated Python,
argument arrays, concurrent stdout/stderr collection and a 30-second timeout.

The harness checks 14 original engine cases: manual first-page loss and invalid
token filtering, Level 1 front matter, and nested Level 2 parent crossing.
Fixture bytes and provenance must match an independent second generation.
Every case checks exact expected original ranges and unchanged source/input/
neighbor sentinel hashes. It also reports complete coverage and fixed-target
matches separately. A matching total page count alone cannot pass these checks.

The two optional launcher probes use the unmodified files only on their early
rejection paths: a nonexistent PDF passed directly to Windows PowerShell 5.1,
and no arguments passed to the batch file. They stop before the application
creates output or its shared temporary engine. These probes demonstrate zero
exit codes for those errors; they do not certify full splitting, Explorer drag/
drop, PowerShell 7 application execution, Calibre, or the release package.

All generated PDFs, the extracted test-only engine, slices and sentinels belong
to one `TemporaryDirectory` and are cleaned when the run finishes. The generator
module import suppresses bytecode writes into the checkout. The sole persistent
run output is the explicitly requested JSON report. Commit only reviewed safe
evidence summaries; keep generated PDFs and render images outside Git.
