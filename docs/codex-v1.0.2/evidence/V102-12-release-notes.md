WinBookSplit v1.0.2 adds regular x64 Python 3.15 support and fixes setup commands that could be blocked before unsigned setup started. It preserves the BAT/PowerShell console, drag-and-drop workflow, local PDF splitting and optional Calibre EPUB/AZW3-to-PDF conversion.

- Accepts regular GIL Windows x64 CPython **3.15.0 or later stable 3.15 patches**, **3.14.8 or later stable 3.14 patches**, and **3.13.16 or later stable 3.13 patches**. These are numeric branch floors, not exact-patch requirements. The executed versions are 3.15.0, 3.14.8 and 3.13.16; every later patch has not been executed. Unlisted minors are unsupported/unvalidated, rather than automatically labeled insecure. Prerelease, x86, ARM, free-threaded Python and PyPy remain outside this policy.
- Adds `Setup-WinBookSplit.bat` and executable setup remedies using a trusted PowerShell host with a process-scoped policy. Internet marks remain intact, organization-enforced controls still apply, and no persistent execution policy or global PATH change is made.
- Reports the selected runtime and reason accurately. Processing does not install packages or silently fall back from an invalid application `.venv`. Setup and explicit repair remain separate operations.
- Retains the exact **pypdf 6.20.0** wheel lock, isolated imports, dependency-origin validation, source preservation, no-overwrite output checks, encryption policy and finite process supervision. Optional conversion uses separately installed **Calibre 9.15.0**.

Extract the complete distribution ZIP to an ordinary writable per-user folder. With a supported Python installed, run once from that folder:

```powershell
.\Setup-WinBookSplit.bat -NonInteractive
```

Then use `WinBookSplit.bat` as before. See the included [setup guide](https://github.com/PikkuJanne/WinBookSplit/blob/v1.0.2/docs/SETUP.md) for explicit runtime selection, read-only checks, reviewed offline wheels and repair. PDF input does not require Calibre.

The **same frozen distribution ZIP** was tested in these exact combinations:

| Native environment | Runtime and automation |
| --- | --- |
| Windows 10 Home 22H2 x64, build 19045.7417, inbox PowerShell 5.1.19041.7417 | Regular Python 3.15.0 and pypdf 6.20.0; owner-launched automated kit, all 16 stages and 19 cases, including authored PDF and genuine Calibre EPUB/AZW3 checks. The default marked-script matrix ran 23 steps. No usable owner 3.15 was discovered on this host, so the explicitly provisioned reviewed private 3.15.0 archive was selected; no global owner installation is claimed. |
| Windows 11 Pro 26H2 x64, build 26300.9457, inbox PowerShell 5.1.26100.9444 | Regular Python 3.15.0 and pypdf 6.20.0; agent-launched automated acceptance, all 16 stages and 19 cases, with an authored private child-PATH runtime fixture and 39 marked-script/control steps. |
| Windows 11 x64, PowerShell 7.6.6 | Four same-ZIP Python 3.15.0 cases: version, setup check, manual split and Auto2. Separately, a genuine inbox PowerShell 5.1-hosted unredirected zero-argument setup BAT pause/Enter case was automated. |
| Windows 11 x64, inbox PowerShell 5.1 | Separate same-ZIP Python 3.13.16 and 3.14.8 setup/check/PDF checks, with independently validated page coverage and preserved sources. |

The full source regression passed 23 stages, 670 Python units and both 204-case Pester routes. Required source CI passed all four jobs. Windows Server CI is supplemental source evidence, not workstation acceptance. Organization-denial controls are simulations; Explorer was not manually exercised, and no repeated chapter-viewing programme was imposed. The returned Windows 10 report was validated against the unchanged maintained report contract; its raw local artifacts were not returned for independent reopening. Earlier external controller and validation-driver failures remain disclosed in the [release evidence](https://github.com/PikkuJanne/WinBookSplit/tree/main/docs/codex-v1.0.2/evidence).

Windows 10 Home/Pro 22H2 standard servicing ended on 14 October 2025. OS security servicing is separate from this application compatibility result; the tool does not enroll a machine in servicing or require a purchase.

The retained pypdf 6.20.0 dependency is generally affected by newer outline/tree advisories. The required application paths were reviewed against those operations; this does not establish that the dependency is generally unaffected or every possible document is safe. Calibre 9.15.0 also retains older embedded libraries. Process deadlines are not an OS/network sandbox. See [SECURITY.md](https://github.com/PikkuJanne/WinBookSplit/blob/v1.0.2/SECURITY.md) and the scoped security evidence for the exact limits. There is no new OCR, DRM removal, uploader, telemetry, installer, updater or native ebook-output feature.

The scripts and ZIP are **unsigned**. SHA-256 values detect byte changes against the published values; they are not a publisher signature.

Release source: `483baf858569bc838ac8a809d329b2595cdacf57`, tree `818de909908c95395691dc962395f270c1cd5735`.

| Accepted asset | SHA-256 |
| --- | --- |
| `WinBookSplit-v1.0.2.zip` | `37d017e2a7bf704676b637f89239ec2cbaaa67e76f190f3dbbd17fec87e5092e` |
| `release-manifest.json` | `a8f8109bc8a49dadf890a10989a92accf6c0f6a038f29aac4ea51474f0792c03` |
| `SHA256SUMS.txt` | `5aea43906c7cd55b841fad51e4d1e36b3c2425edd162f8b0cc388440f979a00e` |

The published v1.0.0 and v1.0.1 tags, release text and assets are preserved.
