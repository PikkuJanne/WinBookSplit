# Original conversion fixtures

`generate_ebook_fixtures.py` authors a tiny offline EPUB with three original
chapters, a navigation table and an NCX table. Calibre 9.15.0 then converts that
EPUB to a genuine AZW3. The generated book text is MIT-licensed test material;
no private book or existing document is used. There are no remote resources,
scripts or external links in the EPUB.

The generator requires the recorded absolute converter path and executable hash.
It records the actual version command, AZW3 conversion arguments, native status,
streams and both file hashes in the external output directory. It does not install
Calibre, alter its global settings or bundle its binaries. EPUB ZIP bytes are
reproducible; generated AZW3 bytes are observed and hashed, without claiming that
Calibre's binary output is reproducible. Existing historical PDF fixtures and
oracles remain unchanged.

Run `tests/run_tests.py --layer conversion` with both actual `--shell-path`
values and the absolute pinned `--calibre-path`. The eleven-stage full route
includes this acceptance after paths. Reports must be new absolute files outside
the checkout. PDF-only targeted layers require no converter.

The route runs eight real PS5.1/PS7 EPUB/AZW3 default/`KeepConvertedPdf` cases,
two direct API captured-reader controls, eight zero-exit native converter controls
(no write, empty, corrupt and zero-page PDF), and one actual unchanged BAT
missing-converter failure. It supplies scripted manual starts `2,3`; conversion
uses exactly `--output-profile tablet`. These fixtures demonstrate the reviewed
profile's actual behavior without claiming that it improves every ebook layout.
The BAT failure must contain the actual `converter_not_found` dependency record
and exact-version guidance, return 3, and create no console or engine result.
It preserves the source and neighbor; successful discovery is tested separately.
Only this negative case redirects child `LOCALAPPDATA` to a fresh empty owned
directory, leaving the parent's environment and real per-user Calibre unchanged.
Both fixed system Calibre paths must actually be absent; the test fails clearly
if this missing-dependency premise cannot be established. No installation is
altered and no mandatory assertion is skipped.

The EPUB yields three physical pages. The genuine AZW3 has three chapters plus
an inline table of contents and yields four physical pages. All pages, including
that fourth page, must appear exactly once in the three requested sections.
Three original visible chapter markers must each appear once. Independent real
reference conversions supply page-content hashes. Default launcher cases compare
actual converter metadata with those references and every reopened output slice;
this treats the source hash vector as engine metadata rather than claiming an
independent observation of its deleted workspace bytes. The direct API cases
independently inspect the captured reader, and retention cases independently
hash/reopen the actual full converted PDF inside the completed run.

Original EPUB/AZW3 sources are marked read-only for actual launches, then their
exact identities and attributes are checked before restoration. Genuine unrelated
same-basename PDFs remain beside every source; hashes must remain unchanged.
Manifest membership, owner markers, chapter digest/size/count and optional
`WinBookSplit_Converted.pdf` are exact. Outputs and console records are deleted
only through the existing held, identity-checked cleanup helper. Conversion-only
publication checks accept the explicit retained member; earlier PDF-only helper
guards remain strict. Stored execution policies and immutable historical sources,
fixtures and oracles remain unchanged.

Successful BAT converter discovery is separately covered by the runtime
acceptance layer with the actual portable converter on a controlled trusted
PATH. The batch launcher remains unchanged. These checks do not certify Calibre GUI,
Explorer drag/drop, arbitrary ebook resources, a network sandbox, DRM handling or
release packaging. Fixture generation alone is not application acceptance.
