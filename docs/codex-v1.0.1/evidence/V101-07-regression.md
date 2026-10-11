# V101-07 — actual reference and two-minor application checks

PASS in the available Windows11 source-checkout scope. Full Python576/576
passed in688.968s, and inbox5.1.26100.9444/isolated7.6.6 each passed182/182
Pester tests with zero skips, not-run cases, failed containers, syntax errors
or scoped analyzer findings. Each separately provisioned application environment,
regular x64 CPython3.13.16 and3.14.8 with pypdf6.20.0 and optional pip only,
passed147 methods across the fixed nine-module subset and ten native cases.
The driver remained dev14 with the reviewed developer wheels. No runtime
dependency was installed by these checks.

All four final reports bind prior HEAD21408fde0fc0bfb4eff547d8cff23a1375f1fada,
tree d811ee1a4011f320c02d062696a055a2f7d77d83 and the same actual checked-out
source digest d4f526aac1d078c57a13846ee53b52c02da8acbc32f7fdecf69ebe1db1ec9d96.
The paired JSON includes the complete tested path/raw-SHA map and actual
versions; HEAD alone is not a claim that these uncommitted changes existed in
the preceding commit. Each source map was independently compared with the
current file bytes after completion. All four report unchanged sources/inputs,
proved cleanup and removed owned work, with no timeout or omitted suite.

Actual commands, from the existing checkout, using explicit isolated per-user
executables and new external reports:

```text
<isolated-work>\dev14\Scripts\python.exe -I -B tests/run_tests.py --layer python --report <isolated-work>\regression-07-final-02\reference-python.json
WBS_TEST_WHEELHOUSE=<isolated-work>\setup-tests-05-01\wheelhouse
<isolated-work>\dev14\Scripts\python.exe -I -B tests/run_tests.py --layer shell --shell-path C:\WINDOWS\System32\WindowsPowerShell\v1.0\powershell.exe --shell-path <isolated-work>\runtimes\powershell-7.6.6-win-x64\pwsh.exe --tool-root <isolated-work>\shell-tools --report <isolated-work>\regression-07-final-02\reference-shell.json
<isolated-work>\dev14\Scripts\python.exe -I -B tests/run_tests.py --layer application --application-python-path <isolated-work>\runtime13\Scripts\python.exe --shell-path C:\WINDOWS\System32\WindowsPowerShell\v1.0\powershell.exe --shell-path <isolated-work>\runtimes\powershell-7.6.6-win-x64\pwsh.exe --calibre-path <profile>\AppData\Local\Programs\Calibre\ebook-convert.exe --report <isolated-work>\regression-07-final-02\application-13.json
<isolated-work>\dev14\Scripts\python.exe -I -B tests/run_tests.py --layer application --application-python-path <isolated-work>\runtime14\Scripts\python.exe --shell-path C:\WINDOWS\System32\WindowsPowerShell\v1.0\powershell.exe --shell-path <isolated-work>\runtimes\powershell-7.6.6-win-x64\pwsh.exe --calibre-path <profile>\AppData\Local\Programs\Calibre\ebook-convert.exe --report <isolated-work>\regression-07-final-02\application-14.json
```

| Final report | Actual result | Raw SHA256 |
| --- | --- | --- |
| reference-python | PASS576/576, zero skips | b5fc116d120748327ac63e0ab0778132f0b4e2acf33ee9a898c639f55518bffc |
| reference-shell | PASS182/182 per host | 1381576429b59a7900098768da1f0267e2e96bc07104eb642dfd0ec4461615a9 |
| application-13 | PASS147 units and ten native cases | 2178c8cbe049edffaf66ac69a204a91b21f4336f6a6b6fa6c5dbb5965bc6372f |
| application-14 | PASS147 units and ten native cases | bf2bad9ebd0cb40a61ae98f322354c1fa36e33fbb0e8146cf85b4743d082b64c |

Each matrix completed21 real owned native probes on Microsoft Windows11 Pro,
CIM version10.0.26300/build26300/ProductType1/64-bit. Cases were manual front
matter, Auto1, Auto2, late-invalid token, no bookmarks, zero pages, both native
PowerShell wrappers, actual EPUB and genuine original AZW3. The authenticated
wrapper log now binds the exact explicit candidate executable, its full native
probe and policy SHA. Page identities/complete ordered ranges, metadata policy,
owned manifest/members, retained/discarded converted PDF, both original ebooks
and their shared same-name neighbor were checked. Before/after runtime-only
inventory, interpreter/pypdf bytes, copied21-member application and every source
input/neighbor remained unchanged. Calibre9.15 actual conversion is distinct
from synthetic converter controls.

[The review](V101-07-review.md) preserves initial full-reference passes and the
two real matrix FAIL reports caused by duplicate exclusive creation of an owned
same-name fixture neighbor. Their unsafe-cleanup work was retained. Only that
fixture creation and the exact wrapper-runtime assertion changed afterward;
these final four complete reruns certify the corrected source. No safety test
was removed to obtain green. The reference count increased from575 to576 due
to the meaningful same-version/different-interpreter contradiction control.

Numeric later-patch acceptance and excluded/floor records are SYNTHETIC policy
controls exercised by the reference and147-method subsets, not claims that a
future patch ran. The3.11/3.12 distribution/security exclusions retain task02's
specific evidence. Developer pins remain separate from application acceptance.
This checkpoint does not pass final-ZIP/nativeWindows10, live Explorer or public
release gates. Task06's independently downloaded matching-source Server CI
receipt is recorded separately in [V101-06-ci.md](V101-06-ci.md).
