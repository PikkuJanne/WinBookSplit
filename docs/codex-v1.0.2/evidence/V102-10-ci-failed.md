# V102-10 exact-R required CI: both reference failures preserved

Required [CI on source R](https://github.com/PikkuJanne/WinBookSplit/actions/runs/38303397809) remains **FAIL** at `001669a051077536864b3f787ca0ec111fe06013`. Three actual application roles from attempt 1 passed strict validation with Python 3.13.16, 3.14.8 and 3.15.0. The original reference and its one unchanged failed-job retry both failed strict validation. Application jobs were not rerun.

Both inbox children reached the existing 300-second bound without a completed Pester child report. Exit 124 is synthesized by the maintained runner on timeout; native exit and assertion result are unavailable. Retained stderr contained 921600 bytes / 4014 first-use progress records in attempt 1 and 847870 bytes / 3693 records in attempt 2. No retained error/assertion marker established a cause. Source/input remained unchanged and work was retained. Each separate Python and PowerShell 7 report passed; these results do not satisfy the missing inbox role.

Five downloaded API ZIP hashes/sizes, all 21 member identities and each 207-path source map were verified against exact R Git blobs. The companion JSON preserves actual role attempts, both failures, stream hashes, successful component counts and the corrected earlier progress observation. Inspector SHA: `b542efd708518be046d038de31bc620cbe4f6b609b477a178e413c2178f1ebe9`.

These are Windows Server 2025 results. Full local regression, diagnostic-only shadow instrumentation, final ZIP native Windows 10/11 and public release gates remain separate. No source or timeout change, absent-report success or second retry was used.
