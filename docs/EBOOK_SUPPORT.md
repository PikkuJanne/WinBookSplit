# EPUB and AZW3 conversion

WinBookSplit converts an EPUB or AZW3 to PDF using the selected installed
Calibre `ebook-convert.exe` (tested with 9.15.0), then splits that generated PDF through the
same validated physical-page plan used for PDF inputs. The original ebook is
an immutable input. The outputs are PDF chapters. Install the documented
dependencies before processing; WinBookSplit does not download dependencies
or upload documents.

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
be new, nonempty, readable, unencrypted and contain positive physical pages.
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
