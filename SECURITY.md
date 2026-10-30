# Security and reporting

WinBookSplit, Python/pypdf and optional Calibre run with the current user's
privileges. WinBookSplit is not a security sandbox or a general PDF sanitizer.
Its PDF policy permits ordinary empty-password authentication and rejects files
that require a nonempty password, unsupported encryption and other excluded
features with specific outcomes. Bounded validated local opening views are ignored
with a warning, without executing actions; other documented navigation omissions
also produce warnings. That does not
prove every hostile file is safe. Use trusted dependency installations and
self-contained ebooks; Calibre/plugins and ebook resources are not isolated by
an OS network-denial mechanism.

Interactive processing can offer the existing application-local setup or explicit
backup-preserving repair once. Only affirmative user consent permits changes,
followed by a fresh runtime check and continuation of the original request.
Preview, noninteractive processing and read-only checks neither prompt nor install.
Setup uses an existing accepted Python and the hash-locked pypdf package; it does
not install Python or Calibre or modify global packages/PATH. The application does
not upload documents, add telemetry, request elevation or change global execution policy. Sources
are immutable inputs; new output directories and cleanup use run ownership.
Local diagnostics may contain sensitive paths, titles, messages and source
identities. Keep them private and review any [redacted support summary](docs/SUPPORT_DIAGNOSTICS.md)
before sharing it.

## Runtime maintenance

Use regular stable Windows x64 CPython accepted by the shipped
[compatibility policy](engine/WinBookSplit.Compatibility.json): Python 3.15.0
or later in 3.15, 3.14.8 or later in 3.14, or 3.13.16 or later in 3.13. The
exact source-test targets are 3.15.0, 3.14.8 and 3.13.16. Later stable maintenance
patches in those branches are accepted numerically; that is not a claim that
each patch was tested. Unlisted minors, such as 3.16, are not yet validated
for this release and are not labeled insecure. Prereleases, free-threaded,
PyPy and non-x64 builds remain excluded. The reviewed floors include the
Windows startup import-path
[fix](https://github.com/python/cpython/issues/151544) and patched bundled
OpenSSL for [CVE-2026-35189](https://openssl-library.org/news/secadv/20260929.txt),
which can affect TLS clients during explicit online setup. They do not declare
every accepted interpreter or document secure. The patched 3.11/3.12 releases
are source-only and lack a patched official Windows binary in the reviewed
catalog, so those branches are excluded from this application policy.

pypdf is separately locked to 6.21.3 and its reviewed wheel SHA256,
retaining the upstream
[multiple-filter fix](https://github.com/py-pdf/pypdf/security/advisories/GHSA-7mh2-4gg2-979f).
Optional ebook conversion remains locked to Calibre 9.15.0. Do not bypass
origin, document-feature, encryption, process or cleanup checks to resolve a
dependency error. Use the consent offer or explicit [application-local setup](docs/SETUP.md)
through `Setup-WinBookSplit.bat` or the complete process-host command. Empty-password
access is not per-operation PDF-permission enforcement; there is no nonempty-password
entry, guessing or DRM removal. Source files are never externally decrypted or
resaved, and chapter PDFs are unencrypted without the ignored opening view. The
scripts are unsigned; Internet marks remain intact and persistent execution
policy is unchanged. Enforced organization controls still apply.

The earlier 18 October review covered upstream fixes to
excessive memory consumption while parsing
[arrays/dictionaries](https://github.com/py-pdf/pypdf/security/advisories/GHSA-f54x-cmq8-qg57)
and [content streams](https://github.com/py-pdf/pypdf/security/advisories/GHSA-cgf9-327m-v77c)
in versions below 6.21.0, plus the fixes described for
[form-field retrieval](https://github.com/py-pdf/pypdf/security/advisories/GHSA-6vvc-x485-93v7)
and [image-loading](https://github.com/py-pdf/pypdf/security/advisories/GHSA-5hx5-vpf9-5wcv)
advisories. The JPEG advisory names 6.21.1 as fixed. The form-field advisory's
patched-version metadata disagrees with its affected range and prose; that
discrepancy remains recorded in the historical gate-review evidence. Small
developer experiments demonstrate the selected dependency's array, dictionary
and content-stream guards with temporarily reduced limits. They do not measure
memory use or establish full application or native acceptance.

WinBookSplit reads PDF structures and content streams. Its production code does
not call form-field retrieval or image extraction APIs; that limited observation
does not prove hostile documents harmless. Bounded child-process supervision
is not a memory or network sandbox. The 27 October update selects 6.21.3 for
the upstream [cross-reference stream entry budget](https://github.com/py-pdf/pypdf/security/advisories/GHSA-x3rh-gjgm-fmqc)
and the preceding [xref table runtime fix](https://github.com/py-pdf/pypdf/security/advisories/GHSA-gmh7-3g9r-mjh5).
Reader initialization occurs before the application's structural preflight, so
those dependency limits remain necessary. Earlier accepted receipts cover only
their original bytes. The owner closed further manual acceptance for v1.0.3;
fresh final-package Windows10 Store and genuine company-policy testing remain
untested and are disclosed. The owner's subsequent repair order removes all
human testing and owner sign-off from active scope. Unavailable ordinary Store,
Windows 10 and other native scenarios remain NOT_RUN evidence limits with no
owner action required; explicit-runtime or component passes do not prove them.
Version 1.0.4 preserves these dependency locks and security floors. Its
exact-source/package results, publication status and limits are in the
[repair status](https://github.com/PikkuJanne/WinBookSplit/blob/codex/functional-repair/docs/functional-repair/STATUS_NEXT.md).
The reviewed fixes do not establish general security
clearance for the application or documents.

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
