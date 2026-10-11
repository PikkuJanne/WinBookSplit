# Security and reporting

WinBookSplit, Python/pypdf and optional Calibre run with the current user's
privileges. WinBookSplit is not a security sandbox or a general PDF sanitizer.
Its tested conservative policy rejects encrypted and unsupported documents and
omits narrowly documented relationships/actions with warnings. That does not
prove every hostile file is safe. Use trusted dependency installations and
self-contained ebooks; Calibre/plugins and ebook resources are not isolated by
an OS network-denial mechanism.

The application does not upload documents, add telemetry, install dependencies
while processing, request elevation or change global execution policy. Sources
are immutable inputs; new output directories and cleanup use run ownership.
Local diagnostics may contain sensitive paths, titles, messages and source
identities. Keep them private and review any [redacted support summary](docs/SUPPORT_DIAGNOSTICS.md)
before sharing it.

## Runtime maintenance

Use regular stable Windows x64 CPython accepted by the shipped
[compatibility policy](engine/WinBookSplit.Compatibility.json): Python 3.13.16
or later in 3.13, or 3.14.8 or later in 3.14. The reviewed/tested recommendations
are 3.13.16 and 3.14.8; later stable maintenance patches in these branches remain
accepted. These floors include the Windows startup import-path
[fix](https://github.com/python/cpython/issues/151544) and patched bundled
OpenSSL for [CVE-2026-35189](https://openssl-library.org/news/secadv/20260929.txt),
which can affect TLS clients during explicit online setup. They do not declare
every accepted interpreter or document secure. The patched 3.11/3.12 releases
are source-only and lack a patched official Windows binary in the reviewed
catalog, so those branches are excluded from this application policy.

pypdf remains separately locked to 6.20.0 and its reviewed wheel SHA256,
including the upstream
[multiple-filter fix](https://github.com/py-pdf/pypdf/security/advisories/GHSA-7mh2-4gg2-979f).
Optional ebook conversion remains locked to Calibre 9.15.0. Do not bypass
origin, document-feature, encryption, process or cleanup checks to resolve a
dependency error. Run the explicit [application-local setup](docs/SETUP.md).

Windows 10 22H2 compatibility does not extend Microsoft's servicing lifecycle.
Standard Home/Pro support
[ended October 14, 2025](https://learn.microsoft.com/en-us/lifecycle/products/windows-10-home-and-pro).
Keep OS security servicing current under an applicable Microsoft programme.

## Report a concern

Private GitHub vulnerability reporting is not enabled for this repository.
Open a [GitHub issue](https://github.com/PikkuJanne/WinBookSplit/issues/new) containing
only a non-sensitive description and affected application version. Ask the
maintainer to arrange a private follow-up before sharing exploit details,
private documents, raw logs or identifying paths. Do not post those details
in the initial issue. There is no promised response deadline or supported private
contact address asserted by this policy.

For an ordinary bug, use an authored minimal reproducer and the guidance in
[CONTRIBUTING.md](CONTRIBUTING.md). Do not upload another person's copyrighted or
private document to demonstrate a problem.

The application is unsigned and supplied without warranty under [MIT](LICENSE).
SHA-256 can detect a byte mismatch against a trusted expected value; it does not
authenticate a publisher or guarantee SmartScreen acceptance. See [setup](docs/SETUP.md)
for checksum use, [PDF policy](docs/PDF_POLICY.md) for input limits and
[ebook support](docs/EBOOK_SUPPORT.md) for the bounded network observation scope.
