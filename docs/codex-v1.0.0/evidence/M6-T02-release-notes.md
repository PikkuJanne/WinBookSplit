# WinBookSplit v1.0.0

These notes are prepared from frozen source `73c90d10dc16b1c321737afbdfbe959d61a24573`. Final ZIP acceptance and public release verification are subsequent steps; no published tag, release URL or downloaded-asset result is asserted here.

WinBookSplit splits a local PDF into chapter PDFs using existing bookmarks or one-based physical page starts. EPUB and genuine, unencrypted AZW3/KF8 inputs are converted to PDF through separately installed Calibre, then use the same splitter. Processing keeps the source file unchanged and does not upload documents or add telemetry.

## Included behavior

- The existing BAT/PowerShell console workflow supports one-file drag-and-drop and no-file path selection. Multiple dropped files are rejected.
- Manual starts and Level 1/Level 2 bookmarks share a validated plan covering every physical page exactly once. Front matter and parent-opening pages are retained; Level 2 sections stay within their Level 1 parent.
- Interactive runs show the complete plan before chapter creation. Scripted parameters support explicit choices, noninteractive execution and plan-only preview.
- Chapters are checked before publication into a unique new output directory. Existing outputs and neighboring files are preserved. Ebook runs can optionally retain the complete converted PDF.
- Bounded process handling preserves conversion errors, cancellation and timeout results. Local logs and manifests support troubleshooting; a redacted support export is an explicit local action.
- PDF output retains the documented page content/resources, limited metadata, chapter-start bookmarks, tested static annotations and supported links within a chapter. Navigation exclusions produce categorized warnings.

## Supported setup

| Component | v1.0.0 scope |
| --- | --- |
| Windows | Windows 11 x64; ordinary local filesystem paths |
| PowerShell | Windows PowerShell 5.1 and PowerShell 7.6.6 |
| Python | Regular GIL CPython 3.14.8 x64 |
| PDF library | Plain, hash-pinned pypdf 6.20.0 |
| Ebook converter | Windows x64 Calibre 9.15.0, required only for EPUB/AZW3 |

Keep the entry points, `VERSION`, requirements and `engine` folder together. Use the included setup guide to create a fresh application `.venv` with the exact Python/pypdf pins. Install Calibre separately for ebooks. Python, pypdf, PowerShell and Calibre binaries, test packages and renderers are not bundled. Processing does not install dependencies, request elevation or change global execution policy.

Drop one supported file onto `WinBookSplit.bat`, or double-click it and enter a literal path. Select Level 1, Level 2 or manual mode and review the plan. Manual starts such as `1,4,7` are physical PDF page numbers, not printed page labels or ranges. For an ebook, use the converted PDF's physical pages; conversion can change layout and introduce a table-of-contents page. The PowerShell entry point also provides `-Version`, full help and scripted parameters.

## Limits and privacy

Encrypted PDFs are rejected, including owner-password cases with an empty opening password. Interactive forms, signatures, attachments, portfolios and detected unsupported active features are rejected according to the PDF policy. Other narrowly defined annotation/navigation relationships are omitted with warnings. Complete metadata, original outlines, signatures and links across chapters are not generally preserved. Read the included PDF policy and fidelity guides before relying on specialized document features.

Automatic splitting requires actual outline destinations. Image-only scans can use manual starts; there is no OCR or inferred chapter detection. Output is PDF only. Password processing, DRM removal, native ebook output and batch jobs are outside scope.

Windows 10, ARM, other runtime versions, Linux/macOS, network paths and arbitrary long paths are outside the support claim. Recorded Windows checks used isolated extractions on an existing workstation, not a clean operating system. Document fixtures and bounded resource checks do not establish compatibility with every PDF or ebook. WinBookSplit is not a document sanitizer or security sandbox; Calibre/plugins and external ebook resources are not isolated by an OS network-denial mechanism. Prefer trusted dependencies and self-contained ebooks.

Local logs may contain sensitive paths, titles, messages and source identities. Keep raw diagnostics private and review a redacted support summary before sharing it. Follow the included security and contributing guides; use an authored minimal reproducer rather than uploading private documents.

## Assets and checksum verification

The intended release assets are `WinBookSplit-v1.0.0.zip`, `release-manifest.json` and `SHA256SUMS.txt`. The manifest identifies source R, package members, dependencies and per-file hashes. `SHA256SUMS.txt` lists the ZIP and manifest; it does not contain its own hash. Final asset sizes and hashes belong to the inspected release record and checksum file.

The package files are frozen for the subsequent exact-ZIP Windows acceptance. Their recorded sizes and SHA-256 values are:

| Asset | Bytes | SHA-256 |
| --- | ---: | --- |
| WinBookSplit-v1.0.0.zip | 414497 | e78a1d0b976223832fcff1c7f17d0d53473112324b05eeec3ec7f74e575ccdc3 |
| release-manifest.json | 11547 | ea5c29faaaeda9132ae75546e1fe0751d8f6417b109b23047c7590875a2f182f |
| SHA256SUMS.txt | 178 | 1bdff00a4589049e41901417784d58b43b31371ca2b74cd98f14a5c4f535548b |

Application scripts are unsigned. After obtaining all three assets from the intended publisher, compare the complete SHA-256 values before extraction. From their download directory, run:

```powershell
Get-FileHash -LiteralPath '.\WinBookSplit-v1.0.0.zip' -Algorithm SHA256
Get-FileHash -LiteralPath '.\release-manifest.json' -Algorithm SHA256
Get-Content -LiteralPath '.\SHA256SUMS.txt'
```

Match all 64 hexadecimal digits for each named file against the checksum file and trusted expected release values. Stop on a mismatch. Matching hashes check bytes; they do not authenticate the publisher, sign scripts or guarantee SmartScreen acceptance. Follow your organization's script policy.

Project code and artwork use the MIT license and are supplied without warranty. Separately installed dependencies retain their upstream licenses; see the included `LICENSE` and `THIRD_PARTY_NOTICES.md`.

The frozen [README](https://github.com/PikkuJanne/WinBookSplit/blob/73c90d10dc16b1c321737afbdfbe959d61a24573/README.md), [setup guide](https://github.com/PikkuJanne/WinBookSplit/blob/73c90d10dc16b1c321737afbdfbe959d61a24573/docs/SETUP.md), [PDF policy](https://github.com/PikkuJanne/WinBookSplit/blob/73c90d10dc16b1c321737afbdfbe959d61a24573/docs/PDF_POLICY.md), [fidelity guide](https://github.com/PikkuJanne/WinBookSplit/blob/73c90d10dc16b1c321737afbdfbe959d61a24573/docs/PDF_FIDELITY.md) and [ebook guide](https://github.com/PikkuJanne/WinBookSplit/blob/73c90d10dc16b1c321737afbdfbe959d61a24573/docs/EBOOK_SUPPORT.md) describe the same source. Their candidate labels record the state at source freeze and do not establish later final-asset acceptance or publication.
