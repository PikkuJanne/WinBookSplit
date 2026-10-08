# Sources and verification boundaries

Repository baseline and primary technical documentation checked during the review/bundle preparation on 7 October 2026. Live documents/versions may change: recheck before choosing pins or publishing. These references explain technical behavior; the milestone design and concrete safety policy are project decisions. No third-party review/news article is a technical authority here.

## [R1] Baseline PowerShell source

`https://github.com/PikkuJanne/WinBookSplit/blob/6edbed7c1a0c94968999882c5a46d90492d3c327/WinBookSplit.ps1`

Existing implementation, defects and entry points; inspected via GitHub connector.

## [R2] Baseline launcher

`https://github.com/PikkuJanne/WinBookSplit/blob/6edbed7c1a0c94968999882c5a46d90492d3c327/WinBookSplit.bat`

Launcher behavior; inspected in preceding review.

## [R3] Baseline README

`https://github.com/PikkuJanne/WinBookSplit/blob/6edbed7c1a0c94968999882c5a46d90492d3c327/README.md`

Original documented scope and mismatches; preceding repository review.

## [R4] Baseline license

`https://github.com/PikkuJanne/WinBookSplit/blob/6edbed7c1a0c94968999882c5a46d90492d3c327/LICENSE`

Existing MIT license; preceding repository review.

## [R5] Live GitHub repository/release metadata

`https://api.github.com/repos/PikkuJanne/WinBookSplit/releases`

No releases observed 2026-10-07; default main SHA/tree also rechecked via connector. Recheck live in M0/M6.

## [S1] OpenAI Codex AGENTS.md guidance

`https://developers.openai.com/codex/guides/agents-md`

Codex loads repository guidance; keep root guidance compact and detail in task files. Official URL redirects to learn.chatgpt.com.

## [S2] GitHub CLI official manual

`https://cli.github.com/manual/`

Use installed-version help for exact syntax and authentication; no API credentials belong in repository files.

## [S3] pypdf installation

`https://pypdf.readthedocs.io/en/stable/user/installation.html`

Package/runtime requirements and optional dependencies; select/test actual versions rather than claim all Python3.

## [S4] Calibre ebook-convert

`https://manual.calibre-ebook.com/generated/en/ebook-convert.html`

Format-dependent conversion options; test actual EPUB/AZW3 -> PDF settings.

## [S5] Microsoft Process.StandardError

`https://learn.microsoft.com/en-us/dotnet/api/system.diagnostics.process.standarderror`

Redirected-stream reads can deadlock; use compatible concurrently drained process streams.

## [S6] pypdf encryption/decryption

`https://pypdf.readthedocs.io/en/stable/user/encryption-decryption.html`

Encryption APIs/dependencies; v1.0.0 deliberately rejects encrypted input instead of claiming password support.

## [S7] Git ls-remote

`https://git-scm.com/docs/git-ls-remote`

Fresh remote ref/object IDs, exit status and annotated tag peeled refs.

## [S8] GitHub CLI release create

`https://cli.github.com/manual/gh_release_create`

Existing-tag verification, draft assembly, assets and immutability considerations.

## [S9] GitHub CLI release edit

`https://cli.github.com/manual/gh_release_edit`

Publishing draft and explicit release fields; verify public outcome afterward.

## This handoff does not certify the application

The repository was read via the GitHub connector. A local baseline clone was unavailable in the bundle-building container because its network name resolution failed; no Windows/PowerShell/Calibre application execution is claimed here. Portable bundle-helper tests use temporary local repositories and mocked remote network results. M0 must reproduce runtime defects and M6 must obtain real Windows/release evidence.

## M0-T03 primary-source refresh — 8 October 2026

The following sources were read for the frozen `SUPPORT_AND_SETUP.md` target.
Package metadata and vendor/advisory reads are dependency selection evidence,
not a scanner or application acceptance pass. Historical baseline versions and
the above bundle-preparation limits remain intact.

### [S10] CPython security release and isolated Windows setup

- [Python 3.14.8 release, 30 September 2026](https://www.python.org/downloads/release/python-3148/): expedited security release, regular Windows x64 artifacts.
- [Python 3.12.15](https://www.python.org/downloads/release/python-31215/): newer source-only security release than the historical bundled 3.12.14.
- [Windows Python manager documentation](https://docs.python.org/3.14/using/windows.html): explicit `--target` extraction and interpreter selection. Installed manager 26.3 help independently confirms extraction without registration.

### [S11] Exact pypdf requirements and published advisories

- [pypdf 6.19.0 package](https://pypi.org/project/pypdf/6.19.0/), [metadata](https://pypi.org/pypi/pypdf/6.19.0/json), [versioned installation docs](https://pypdf.readthedocs.io/en/6.19.0/user/installation.html): Python >=3.9, plain package has no mandatory transitive package on Python 3.14.8; BSD-3-Clause.
- [Public advisory API](https://api.github.com/repos/py-pdf/pypdf/security-advisories?per_page=100): 54 records, no next page at observation. Manual declared-range review found none including 6.19.0 and 31 including historical 6.10.0.
- [Outline runtime/memory advisory](https://github.com/py-pdf/pypdf/security/advisories/GHSA-23w6-3w8w-8484), [appearance stream](https://github.com/py-pdf/pypdf/security/advisories/GHSA-php9-fj8v-98fj), [page labels](https://github.com/py-pdf/pypdf/security/advisories/GHSA-w23x-9jrw-r45c), [embedded files](https://github.com/py-pdf/pypdf/security/advisories/GHSA-v247-6f48-mgcj): inspected recent patch boundaries through 6.19.0. No exploit or hostile-document test was run.

### [S12] Calibre target and conversion boundary

- [Official Windows download](https://calibre-ebook.com/download_windows), [release notes](https://calibre-ebook.com/whats-new), [ebook-convert manual](https://manual.calibre-ebook.com/generated/en/ebook-convert.html): current 9.15.0 target, upstream Windows10 >=1809 requirement and format-dependent conversion options. This does not establish WinBookSplit Windows10 or conversion support.
- [Public advisory API](https://api.github.com/repos/kovidgoyal/calibre/security-advisories?per_page=100): 21 records, no next page; no returned declared affected range includes 9.15.0.
- [Content-server conversion options](https://github.com/kovidgoyal/calibre/security/advisories/GHSA-p6rh-fpfg-37w2), [palmdoc heap overflow](https://github.com/kovidgoyal/calibre/security/advisories/GHSA-8rcp-ffwf-rf64), [metadata template](https://github.com/kovidgoyal/calibre/security/advisories/GHSA-2j4m-2q7x-2c47) were inspected. [Older incomplete patch metadata](https://github.com/kovidgoyal/calibre/security/advisories/GHSA-8r26-m7j5-hm29) prevents a claim that every entry has complete patch information. No converter or bundled-component audit was performed.

### [S13] ReportLab developer pin

- [ReportLab 5.0.1 metadata](https://pypi.org/pypi/reportlab/5.0.1/json): Python >=3.9,<4; mandatory Pillow and charset-normalizer, no extras needed for the generator.
- [4.5 release notes](https://docs.reportlab.com/releases/notes/whats-new-45/), [4.4.10 security patch](https://hg.reportlab.com/hg-public/reportlab/rev/9c95ffc4c865), [5.0 security settings](https://docs.reportlab.com/releases/notes/whats-new-50/): vendor fix after historical 4.4.9 and safer URL-fetch defaults. Empty PyPI vulnerability metadata is insufficient by itself.

### [S14] Developer transitive pins

- [Pillow 12.3.0](https://pypi.org/project/pillow/12.3.0/), [release notes](https://pillow.readthedocs.io/en/stable/releasenotes/12.3.0.html), [PDF decompression advisory](https://github.com/python-pillow/Pillow/security/advisories/GHSA-jjj6-mw9f-p565), [Windows viewer advisory](https://github.com/python-pillow/Pillow/security/advisories/GHSA-4x4j-2g7c-83w6): Python >=3.10 and inspected fixes in 12.3.0.
- [charset-normalizer 3.5.2 metadata](https://pypi.org/pypi/charset-normalizer/3.5.2/json), [maintainer advisories](https://github.com/jawah/charset_normalizer/security/advisories): Python >=3.7, no mandatory dependencies, no published advisory observed. No scanner or unpublished-defect claim.

### [S15] Pester tooling target

[Gallery 6.2.0](https://www.powershellgallery.com/packages/Pester/6.2.0), [official release](https://github.com/Pester/Pester/releases/tag/6.2.0), [prerequisites](https://pester.dev/tutorial/introduction/prerequisites): Windows PowerShell 5.1 / PowerShell7.4+ compatible target; no module dependencies. Selection only; module has not been imported or tested here.

### [S16] PSScriptAnalyzer tooling target

[Gallery 1.25.0](https://www.powershellgallery.com/packages/PSScriptAnalyzer/1.25.0), [official release](https://github.com/PowerShell/PSScriptAnalyzer/releases/tag/1.25.0): minimum PowerShell5.1; selected analyzer version, not an executed analyzer pass.

### [S17] Poppler renderer boundary

[Official release history](https://poppler.freedesktop.org/releases.html): 26.10.0 is newer than the observed bundled 26.07.0, with later malformed-document fixes. Select/review the Windows binary before later fidelity testing; the historical renderer observation is not a current security certification.

### [S18] pip hash-verified installs

[Official secure-install documentation](https://pip.pypa.io/en/stable/topics/secure-installs/): explicit pinned/hash-verified wheels with `--require-hashes --only-binary=:all:`. M0-T03 actually used these flags in separate fresh runtime/dev venvs and recorded the downloaded wheel identities.
