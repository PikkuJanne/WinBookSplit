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
