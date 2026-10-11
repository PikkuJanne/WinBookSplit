# V101-02 security and compatibility design

Review date: 2026-10-11. Baseline source commit:
`45507d427b4238425fa0ffe91982e53d4499385c`; its Git tree:
`7114363b9e4abc195e5439c9e7d2af4ebaab2a8a`. The design adds
`engine/WinBookSplit.Compatibility.json` and a separate exact test selection in
`tests/compatibility/runtime-lock.json`. Historical v1.0.0 evidence is unchanged.
This record is a source/security review and syntax audit, not application
acceptance or a claim that a selected interpreter has run the regression matrix.

## Decision and candidate evaluation

| Candidate | Design decision | Concrete reason |
| --- | --- | --- |
| Regular stable x64 CPython 3.11 | Excluded | The material Windows import-search-path fix is in source-only 3.11.16; the current PSF Windows distribution catalog ends at 3.11.9. Requiring an end user to compile Python is outside scope. |
| Regular stable x64 CPython 3.12 | Excluded | The same fix is in source-only 3.12.14; the current PSF Windows catalog ends at 3.12.10. The exclusion concerns patched official Windows distributions, not syntax inconvenience. |
| Regular stable x64 CPython 3.13 | Candidate policy branch, minimum 3.13.15 | Official patched Windows binaries exist. Recommended and selected exact test target is 3.13.16. Actual engine matrix remains NOT_RUN in this review. |
| Regular stable x64 CPython 3.14 | Candidate policy branch, minimum 3.14.7 | Official patched Windows binaries exist. Recommended and selected exact test target is 3.14.8. Actual engine matrix remains NOT_RUN in this review. |

The numeric compatibility minimum is patch zero because this static source/API
audit found no maintenance-patch API requirement. The effective acceptance floor
is the higher security minimum. Branches are ordered 3.14 then 3.13 for a
deterministic preference. An accepted later stable patch is not an exact-version
test claim. The policy excludes other implementations, 32-bit and ARM processes,
free-threaded builds, prereleases and unlisted minors, including 3.15. The current
latest feature release does not expand this bounded patch target.

### Applicable hard floor

[CVE-2026-12003 / gh-151544](https://github.com/python/cpython/issues/151544)
describes a Windows import-search-path weakness in legacy all-user installations:
a writable location above the interpreter can supply the old in-tree landmark
and an alternative standard-library directory. Import-origin checks performed
after Python starts cannot undo code that was loaded while Python initialized.
The application discovers all-user as well as per-user installations, so a
single security floor avoids a location-dependent ancestor-permission exemption.
The fallback is removed in [3.13.15](https://docs.python.org/3.13/whatsnew/changelog.html#python-3-13-15-final)
and [3.14.7](https://docs.python.org/3.14/whatsnew/changelog.html#python-3-14-7-final).

Upstream explicitly says protected per-user installs can be unaffected. The
user's recorded 3.14.2 path is a per-user location; this review does not establish
that their individual installation is exploitable. Its proposed result is
`below_security_floor`, with an explained reviewed-baseline requirement, never
the old exact-3.14.8-equality failure. The log did not reach the pypdf import and
does not prove that pypdf was absent.

### Advisory versus minimum

[3.13.16](https://www.python.org/downloads/release/python-31316/) and
[3.14.8](https://www.python.org/downloads/release/python-3148/) include newer TLS,
archive, ZIP decompression and XML fixes. They remain recommended values and
exact test selections, not equality checks or automatically substituted floors.
Normal engine processing has no zipfile/tarfile/SSL/XML input-parser import; ebook
conversion is delegated to the separately reviewed Calibre process. Explicit
setup installs only the permitted hash-locked wheel. This limited applicability
assessment is not a universal security guarantee for the interpreter. A pending
[TemporaryDirectory cleanup race](https://github.com/python/cpython/issues/157579)
does not identify a runtime cleanup path here: existing output/conversion cleanup
uses owned ledgers and held Windows handles. Recheck upstream at release freeze.

## Syntax, standard library and native API audit

The four existing shipped Python files were read without execution. The source
uses ordinary imports, dataclasses, MappingProxyType, pathlib, hashlib, threading,
explicit ctypes/wintypes declarations and subprocess. `BaseException.add_note`
requires Python 3.11; the preflight's `Path.is_relative_to` is available since
3.9. No application syntax or standard-library API exclusion for 3.11/3.12 was
identified. Their exclusion instead follows the material distribution/security
decision above. [Python 3.11 pathlib](https://docs.python.org/3.11/library/pathlib.html#pathlib.PurePath.is_relative_to),
[Python 3.11 ctypes](https://docs.python.org/3.11/library/ctypes.html#ctypes.WinDLL).

The application declares native argument and result types rather than relying
on implicit integer handle widths. Directory identity/publication calls and
job-object process ownership do not intrinsically require Windows 11.
Microsoft documents [GetFileInformationByHandleEx](https://learn.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-getfileinformationbyhandleex)
and [CreateJobObjectW](https://learn.microsoft.com/en-us/windows/win32/api/jobapi2/nf-jobapi2-createjobobjectw)
before Windows 10; actual filesystem behavior, ctypes layout, job assignment and
process cleanup remain native-test obligations. No weaker sharing, cleanup or
process-safety fallback is authorized by this audit.

One actual compiler run used official regular Windows x64 CPython **3.14.8**:

```powershell
& '<isolated-runtime>\python-3.14.8-amd64\python.exe' -I -B -c "import ast,pathlib,sys; files=sorted(pathlib.Path('engine').glob('*.py')); print('compiler',sys.version.split()[0]); print('files',len(files)); [(ast.parse(p.read_text(encoding='utf-8-sig'),filename=str(p),feature_version=(3,n)),print('SYNTAX_ONLY',str(p),'3.'+str(n))) for n in (11,12,13,14) for p in files]"
```

Exit 0; compiler reported 3.14.8; all **16** parses (four files times four
requested grammar versions) succeeded. These are `SYNTAX_ONLY` checks under a
single compiler, not executions of 3.11, 3.12 or 3.13 and not native application
or Windows 10 passes. `ast.parse(feature_version=...)` is a best-effort grammar
check. The actual multi-minor matrix must exercise the engine and safety fixtures.

## Separate pypdf and Calibre review

[PyPI 6.20.0 JSON](https://pypi.org/pypi/pypdf/6.20.0/json) identifies a non-yanked
pure-Python wheel, Python >=3.9 metadata, and the retained SHA-256
`f003fc2014814d264fe7dd3f9d435c158e23e1a85a2233f87a0a2d6d21c914ad`.
The metadata makes all four candidate minors feasible; it is not proof of
WinBookSplit engine compatibility. The
[maintainer advisory API](https://api.github.com/repos/py-pdf/pypdf/security-advisories?per_page=100)
returned 55 records. Current pypdf affected ranges inspected exclude 6.20.0;
old PyPDF2 package records were kept distinct. The newest
[GHSA-7mh2-4gg2-979f](https://github.com/py-pdf/pypdf/security/advisories/GHSA-7mh2-4gg2-979f)
was published October 9 and fixes memory consumption through multiple stream
filters in 6.20.0. Keep this exact package lock, encryption rejection hook,
`root_object_recovery_limit` and current origin/resource/process safeguards.

Calibre remains exactly **9.15.0**. Its
[exact installer source](https://raw.githubusercontent.com/kovidgoyal/calibre/v9.15.0/bypy/windows/wix-template.xml)
permits Windows 10 builds; the [current download page](https://calibre-ebook.com/download_windows)
requires Windows 10 1809 or newer. These are prerequisites, not WinBookSplit
acceptance. Latest Calibre observed is 9.16.0. The
[maintainer advisory API](https://api.github.com/repos/kovidgoyal/calibre/security-advisories?per_page=100)
returned 21 records whose affected ranges inspected exclude 9.15.0. The newest
[conversion-options content-server advisory](https://github.com/kovidgoyal/calibre/security/advisories/GHSA-p6rh-fpfg-37w2)
is patched in 9.15; the earlier
[PalmDOC heap overflow](https://github.com/kovidgoyal/calibre/security/advisories/GHSA-8rcp-ffwf-rf64)
is patched in 9.13. No published advisory identified requires broadening this
converter pin. Actual Win10 and Win11 EPUB/AZW3 conversion remains NOT_RUN here.

## Primary observations and reproducibility

All cited sources above were fetched on 2026-10-11, using primary PSF/Python,
PyPI, maintainer GitHub and Microsoft documentation. Additional refreshed
sources were [Python lifecycle](https://devguide.python.org/versions/),
[3.11.16](https://www.python.org/downloads/release/python-31116/),
[3.12.14](https://www.python.org/downloads/release/python-31214/),
[3.11.17](https://www.python.org/downloads/release/python-31117/),
[3.12.15](https://www.python.org/downloads/release/python-31215/) and
[Python Windows documentation](https://docs.python.org/3.14/using/windows.html).
Python's own repository advisory API returned zero records; that absence was not
treated as an absence of CPython vulnerabilities. Release changelogs and upstream
issues supplied the relevant CPython security evidence.

Read commands included `rg --files -g AGENTS.md`, `rg -n` over engine imports and
API use, the selected task/specification reads, `git status --short`,
`git remote -v`, `git branch --show-current`, `git log -1`,
`git rev-parse 'HEAD^{tree}'`, and `git ls-remote origin` for main and release
tags. Git showed the existing scoped work branch at the baseline above, live
main at `be1fa1e271aa2bd031d907e9dbc9e7bfc32bbb62`, v1.0.0 tag at
`97df3d364e9bde0b2c2c6c3655d073b7b56b5dd2`, and no v1.0.1 tag. No Git writes
were performed by this review worker.

Web observations used `web__run` plus read-only `Invoke-WebRequest` for release
pages/changelogs/source and `Invoke-RestMethod` for PyPI/advisory/PSF catalog JSON.
The PSF index was followed through all three linked pages to the terminal null
`next`; newest regular AMD64 entries were 3.11.9, 3.12.10, 3.13.16 and 3.14.8.
Published selected full-package hashes were:

| Exact selected target | Official URL | Published SHA-256 |
| --- | --- | --- |
| 3.13.16 | https://www.python.org/ftp/python/3.13.16/python-3.13.16-amd64.zip | `bbf675bb5e763c1efbb09a3a461b259d81598a63c30c4b0d7ea11b9f063df159` |
| 3.14.8 | https://www.python.org/ftp/python/3.14.8/python-3.14.8-amd64.zip | `4873947a8afc037846b180312b83c744a4146a851cfd316a75c3125a4d8299da` |

Those values are published source observations, not independent byte verification
by this review worker. Isolated acquisition/probe records are recorded separately.
After authoring, CPython 3.14.8 loaded both new JSON files with `json.loads` and
reported schema version 1, exit 0. `git diff --check` exited 0. These are JSON
syntax/format checks, not the strict policy-validator or application matrix.
The authored policy SHA-256 was
`b2776ea7e75629a45ad0cf664e1c9002cdf61a0fc8a8815bf879fd34377c9c62`;
the exact-test selection SHA-256 was
`48d19fc69d2c9fb8831448b1fbd20a71ab7582a8e491c10494d8e4e21f7f3f3a`.
One initial audit command tried a known absolute Python path from the user's old
log; it was absent on this host, so no invocation occurred. A PATH inspection
showed a WindowsApps alias, which was not launched. The audit then used the
explicit provisioned official 3.14.8 executable. Two initial `rg` lookups used a
nonexistent optional setup path or a Windows wildcard form; they were corrected
to `-g '*.py' engine`. No missing-path result was represented as a test pass.

## Acceptance boundaries and next obligation

- V101-AC-006: design evaluation complete for all four candidates with concrete
  security/distribution reasons.
- V101-AC-007: one strict schema and separate exact test selection authored;
  validator tests and consumption by every application entry point are separate
  implementation evidence, not established by this document.
- V101-AC-008: policy design permits numeric later stable patches; actual
  validation/parity tests remain the implementation obligation.
- V101-AC-009: floor has applicable Windows discovery evidence; advisory and
  recommended values remain separate. No exact-equality floor is intended.
- V101-AC-010: only candidate policy branches are selected here. Public
  advertisement remains gated on real multi-minor engine and native package
  acceptance; matrix and native Windows 10/11 checks are NOT_RUN in this review.

No private document, network identifier, host name or credential is included.
No manual acceptance, native Windows 10, conversion, full regression matrix or
release-package PASS is claimed. Next: validate and integrate this shared policy,
then collect actual branch matrix and native final-ZIP evidence before freeze.
