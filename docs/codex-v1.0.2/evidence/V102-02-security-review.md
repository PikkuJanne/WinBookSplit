# V102-02 primary-source security and artifact review

Reviewed 12 October 2026. This is upstream source/metadata review, not a native-runtime, application, installer-signature, or downloaded-binary pass. The accompanying JSON records access times, decoded UTF-8 response-content hashes, selected metadata, all published pypdf advisory ranges, and the local source identity at collection. No private documents were used.

## Conclusion

The reviewed sources support adding regular stable Windows x64 CPython 3.15 with initial compatibility/security/recommended patch `0`, retaining the existing 3.13.16 and 3.14.8 floors and exact pypdf 6.20.0 wheel lock. No concrete upstream blocker requiring exclusion of 3.15.0 or a dependency update was found in this review. Actual shared-evaluator tests and real 3.15 application/native acceptance remain separate required evidence.

## Authentic published distributions

[Python.org's release page](https://www.python.org/downloads/release/python-3150/) identifies 3.15.0 as stable, released 9 October 2026. Its recommended regular Windows x64 installer is `python-3.15.0-amd64.exe`, published SHA-256 `67d9a63eb145a34aa939f02eadbf8c57f21fcdd6decb901ccac578d2d1f01d6f`.

Both the [release manifest](https://www.python.org/ftp/python/3.15.0/windows-3.15.0.json) and [current PSF catalog](https://www.python.org/ftp/python/index-windows.json) contain the final `pythoncore-3.15-64`, `sort-version: 3.15.0` record. The regular runtime archive is [python-3.15.0-amd64.zip](https://www.python.org/ftp/python/3.15.0/python-3.15.0-amd64.zip), published SHA-256 `6b1a3f885ce416cf3a68af203e62203664eb5224d8d0c1d2a1945fddec5c500b`. Free-threaded `3.15t` and embeddable records are distinct. These are publisher metadata observations; acquisition must hash the actual bytes and probe the actual selected executable.

## Preserved security rationale

The Windows source-landmark fix for CVE-2026-12003 is documented in [upstream issue 151544](https://github.com/python/cpython/issues/151544) and [3.15.0b3 NEWS carried in the final tag](https://raw.githubusercontent.com/python/cpython/v3.15.0/Misc/NEWS.d/3.15.0b3.rst). [Final tagged getpath.py](https://raw.githubusercontent.com/python/cpython/v3.15.0/Modules/getpath.py) requires the build-directory marker and no longer uses the alternate `Modules/Setup.local` landmark. This justifies using the existing `CVE12003` rationale for 3.15 rather than treating an unknown minor as insecure.

[Final-tag rc3 NEWS](https://raw.githubusercontent.com/python/cpython/v3.15.0/Misc/NEWS.d/3.15.0rc3.rst), [Windows build properties](https://raw.githubusercontent.com/python/cpython/v3.15.0/PCbuild/python.props), and the [official installer SBOM](https://www.python.org/ftp/python/3.15.0/python-3.15.0-amd64.exe.spdx.json) independently identify Windows OpenSSL 3.5.9. The [29 September OpenSSL advisory](https://openssl-library.org/news/secadv/20260929.txt) fixes CVE-2026-35189 in 3.5.9; its malicious-certificate memory pressure during ordinary TLS is relevant to explicit online pip setup. The [current advisory index](https://openssl-library.org/news/secadv/) had no later dated advisory.

The tagged [3.14.8 NEWS](https://raw.githubusercontent.com/python/cpython/v3.14.8/Misc/NEWS.d/3.14.8.rst) and [3.13.16 NEWS](https://raw.githubusercontent.com/python/cpython/v3.13.16/Misc/NEWS.d/3.13.16.rst) also confirm Windows OpenSSL 3.5.9. The [3.14.8 release](https://www.python.org/downloads/release/python-3148/) and [3.13.16 release](https://www.python.org/downloads/release/python-31316/) document the serviced releases. Existing floors remain justified; 3.14.2 remains below its reviewed floor.

The CPython API query for `type-security` issues updated since 8 October returned five records. Open [issue 157579](https://github.com/python/cpython/issues/157579), concerning `TemporaryDirectory` cleanup races, is already fixed in final 3.15 source by [merged backport 158431](https://github.com/python/cpython/pull/158431). Its fix was not found in the tagged 3.13.16/3.14.8 NEWS. The application engine, setup and tools search did not find use of `TemporaryDirectory`, `tempfile` or `rmtree`; this review did not reproduce a corresponding Windows project failure or establish a new required legacy floor. Preserve the floors and cleanup guards, and recheck relevant advisories at release. An issue's open/closed state alone is not a version-vulnerability verdict.

## pypdf origin and advisory boundary

[PyPI JSON](https://pypi.org/pypi/pypdf/6.20.0/json) reports the non-yanked universal `pypdf-6.20.0-py3-none-any.whl`, Python `>=3.9`, SHA-256 `f003fc2014814d264fe7dd3f9d435c158e23e1a85a2233f87a0a2d6d21c914ad`, matching the shipped lock. Metadata alone does not prove Python 3.15 application compatibility or import origin.

The [publisher advisory API](https://api.github.com/repos/py-pdf/pypdf/security-advisories?per_page=100) returned 56 advisories and 53 vulnerability ranges for package `pypdf`; all ranges were evaluated numerically and none included 6.20.0. Separate `PyPDF2` records were not attributed to `pypdf`. This includes [multiple-filter memory use](https://github.com/py-pdf/pypdf/security/advisories/GHSA-7mh2-4gg2-979f) and the newly published [EOF/xref infinite loop](https://github.com/py-pdf/pypdf/security/advisories/GHSA-39mj-jf6j-2qj3), both patched in 6.20.0. Retain origin validation, document guards and finite process limits.

## 3.15 changes requiring actual regressions

[Windows documentation](https://docs.python.org/3.15/using/windows.html) states Windows 10 and newer support. [What's new](https://docs.python.org/3.15/whatsnew/3.15.html) records default UTF-8, new package `.start` startup files, and ctypes changes. Preserve isolated startup and bytecode suppression, and exercise malicious working-directory/book-directory shadows, pip/venv, Unicode CLI paths and real ctypes/process paths. No workaround importing arbitrary startup files is justified by this review.

## Commands and limits

Sources were opened with web browsing and retrieved with `Invoke-WebRequest -UseBasicParsing` / `Invoke-RestMethod`; API ranges were classified using numeric PowerShell `System.Version` comparisons. Local instructions/policy were read; `rg` audited the relevant application paths; `git rev-parse HEAD`, `git rev-parse HEAD^{tree}` and `Get-FileHash` recorded source identity. The large complete 3.15 changelog exceeded the web reader limit, so final-tag NEWS files and tagged source were inspected directly. No production policy or historical evidence was edited by this reviewer.

Recommended new policy references: `Release3150`, `Changelog315`, `OpenSSL3150`, `Python315Sbom`; use existing `CVE12003`, `OpenSSL20260929`, and `WindowsCatalog` references as appropriate. Their exact URLs are above and in the JSON snapshot.

## Independent implementation review

Reviewed the task-2 diff against checkpoint `41b281fcaa4302ed9d362c62ee39f64bad7e14f0`: production changes are confined to the shared JSON policy; active Python/PowerShell policy, discovery and outcome tests are updated. Reviewed policy file SHA-256: `57811a77c837bd891750ea126d19c876e0ad7099793b490445d126720baf4582`.

Programmatic before/after comparison confirmed both legacy branch objects, pypdf lock/hash, Calibre lock and excluded-branch records are exactly semantically unchanged. The new branch references the verified upstream sources and stays below the policy's size/source-count limits. Existing evaluator, strict duplicate-safe parsing, policy-byte identity binding, runtime identity/flavor/architecture checks and package-origin checks are unchanged. The unchanged selector ranks policy branch order first, numeric patch second, and ordinal path ties last, so prepending 3.15 produces the required preference.

The active tests add actual shared-evaluator fixture/parity calls, retain invalid-policy/source/identity cases, preserve legacy-floor rejection, and move the synthetic unknown-minor case to 3.16. Synthetic patch versions are explicitly identified; no historical v1.0.0/v1.0.1 report was changed. No actionable correctness/security finding was identified in this diff. This independent review did not run an additional test suite; test outcomes belong to the task runner's separate evidence.
