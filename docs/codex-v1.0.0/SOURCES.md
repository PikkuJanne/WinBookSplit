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
