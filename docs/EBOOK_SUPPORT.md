# EPUB and AZW3 conversion

This describes version 1.0.4 and its
[30 October functional repair](https://github.com/PikkuJanne/WinBookSplit/blob/codex/functional-repair/docs/functional-repair/STATUS_NEXT.md).
Converted PDFs use the same access and document policy as direct PDF inputs,
including empty-password authentication and narrowly ignored local opening views.
Published v1.0.3 assets retain their original bytes and behavior. Ebook DRM and
unrelated document/process restrictions remain unchanged.

WinBookSplit converts an EPUB or AZW3 to PDF using the selected installed
Calibre `ebook-convert.exe` (stable 9.17.0 or newer), then splits that generated PDF through the
same validated physical-page plan used for PDF inputs. The original ebook is
an immutable input. The outputs are PDF chapters. Install Calibre separately.
The numeric security minimum permits later stable releases. Exact converter
versions and binary hashes in evidence identify the tested target, rather than
requiring users to install that precise patch.
The interactive application can prepare the locked PDF
package with explicit consent and resume the original request, as described in
[setup](SETUP.md). Preview and NonInteractive runs never prompt or install.
WinBookSplit does not upload documents.

The application invokes the resolved executable with separate literal arguments:

```text
ebook-convert.exe <original-ebook> <owned-new-generated.pdf> --output-profile tablet
```

The existing `tablet` profile stays in place. A profile does not guarantee a
particular PDF layout for every ebook. Calibre reconstructs layout: EPUB and
AZW3 versions of the same authored book can have different physical page counts
and a generated table-of-contents page. Use the generated PDF's physical pages
for manual starts. Outline destinations determine Auto1/Auto2 boundaries; the
shared plan includes front matter and every generated physical page exactly once.
When the requested outline level has no usable plan, select a supported mode
explicitly or provide valid manual starts.

Use `-KeepConvertedPdf` to retain `WinBookSplit_Converted.pdf` inside the new
successful run directory for inspection. It is listed separately from chapters
in the manifest. Default execution removes the owned intermediate after capture.
Plan-only conversion creates temporary PDF data without publishing chapters.
Interactive manual selection opens the generated PDF only when requested; its
preview remains available until that selection ends. Files beside the ebook
and previous run directories are preserved.

A converter failure returns a nonzero result with the conversion diagnostic;
zero converter status alone does not establish success. The generated file must
be new, nonempty, readable under the shared [PDF policy](PDF_POLICY.md), and contain
positive physical pages. Empty-password encrypted output can be accepted;
password-required, unsupported-encryption and missing-backend failures retain
their specific outcomes. New chapter PDFs are unencrypted and omit any ignored
opening view. The retained full converted PDF, when requested, is the converter's
output and is not rewritten by the chapter writer.
The application records exact converter arguments, status, bounded output and
owned cleanup. Read the console/run diagnostics when a malformed or unsupported
ebook fails. DRM removal and native ebook output are outside scope.

Process self-contained ebooks with installed dependencies. The M4-T04 route
uses original EPUB3/NCX and independently generated genuine, unencrypted KF8
AZW3 fixtures; it checks resource closure before conversion. Its separate
loopback image/CSS canary records endpoint liveness, actual converter diagnostics
and whether those exact requests occur. This is a bounded observation of the
tested PDF renderer, not network isolation for every plugin/process or a security
sandbox. Avoid ebooks requiring external resources. Calibre and pypdf run with
the user's privileges.

See [the automated source route](https://github.com/PikkuJanne/WinBookSplit/blob/main/tests/ebooks/README.md),
[source task evidence](https://github.com/PikkuJanne/WinBookSplit/blob/main/docs/codex-v1.0.0/evidence/M4-T04-ebook-conversion.md) and
[Calibre's conversion command documentation](https://manual.calibre-ebook.com/generated/en/ebook-convert.html).
Current online documentation may describe a newer Calibre; the route records
the actual installed version and format-specific local help. Render comparisons,
agent image inspection and historical closed human testing are distinct.
