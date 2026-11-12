# Checkpoint 4 — Exact v1.0.5 package acceptance

The tested source was `56dc118313cd610bde4cbb578f46e9acbb273610`, tree
`60190bd5286cee963cf1707a5990aaa5ca83166b`. Two independent builds from
the committed builder produced byte-identical assets and passed its verifier.
All 37 fresh extraction members matched the Git-bound package manifest.

The earlier hosted Python regression failed. Local reproduction identified
stale unit expectations for the payload count and the ebook Preview converter
argument. Both focused groups passed after correction (23 cases). The next
hosted run passed Python and all three application jobs. Both hosted shells
exposed a registry fixture under the runner's working directory. The fixture now
exercises registry fallback from the application directory while retaining the
working-directory guard. All 38 runtime controls passed on each native shell.

## Reported document or resource action failure

The archived public v1.0.4 package ran through Windows PowerShell 5.1 on a
four-page authored PDF with a resource-only action. It exited 7 with
“Active document or resource actions are not supported.” The authored
fixture SHA-256 was
`15487f5740fb4e92e07fccf15c2ff9bd0078c57ebec8935e67d62bd7a612aee2`.
An earlier v1.0.5 build split that fixture, then refused
an optional action on a drawn XObject in a deeper combined fixture. The final
implementation handles optional action/script entries in graphics streams
and resource dictionaries by context. It preserves used drawing resources,
their names, transparency, Type 3 glyphs, and actual font/image objects even
when a malformed optional reference also points to them.
The final ZIP split the same four-page reproduction into two reopened
chapters with all page markers and an action-omission warning. The input
and a neighboring file retained their original bytes.

The combined native regression includes empty-password encryption, a saved
form value, an image, local navigation, catalog and widget actions,
attachments, a drawn Form XObject with malformed and executable action
branches, and indirect script payloads also reachable through earlier page
resources. It reopens the written chapters, checks exact page order and
rendered output, and inspects the serialized graph for omitted objects.

| Asset | Bytes | SHA-256 |
| --- | ---: | --- |
| WinBookSplit-v1.0.5.zip | 1,133,973 | `756a467136f7c41a70fbdbcfcfcf073a772bc338a10ce2141fe2669f396431c3` |
| release-manifest.json | 17,778 | `a092723d2dfab81f7ffc1a40a4a2fef48427f2fa25ef9a6134da4360a4c46bce` |
| SHA256SUMS.txt | 178 | `22b80f6647bea8234bcbef4ca4f35382d209065c10661a98038ed54eab883625` |

## Windows 11 desktop results

The host was Windows 11 Pro build 26300. Tests used inbox PowerShell
5.1.26100.9711, PowerShell 7.6.9, regular x64 Python 3.14.8 and pypdf 6.21.4.
The form matrix used Poppler 26.07.0 and PDFium 153.0.7999.0.

| Workflow | Observed result |
| --- | --- |
| Exact ZIP form and PDF matrix | 54/54. There were 24 cases per PowerShell host and six BAT cases. Forty-five split runs produced 134 reopened chapters, three previews produced read-only plans, and six malformed or unreproducible cases failed for their expected reasons. |
| Exact ZIP Calibre matrix | 15/15. The normal interactive launcher downloaded and verified the official Portable package, resumed the original illustrated Unicode EPUB and produced three chapters. A second run reused it with downloads unavailable and produced three more. |
| Exact ZIP offline Calibre and older install | The signed 9.15.0 installation remained byte-identical. Its actual rejected path and version were recorded. Explicit offline setup prepared verified 9.18.0 separately, then Manual split covered all three generated pages in two chapters. |
| Existing Calibre and converter errors | Actual PATH and explicit 9.17.0 runs produced three illustrated EPUB chapters each. AZW3 produced three chapters. Invalid explicit selection, decline, unavailable download, invalid offline artifact and corrupt ebook each had its own expected failure. Preview, noninteractive processing and Check left converter preparation untouched. PDF-only split succeeded without Calibre. |
| Exact ZIP cold PDF setup and repeat | With an explicit supported base Python lacking pypdf and a fresh extraction, Preview and NonInteractive returned `setup_required` without setup. Consented PowerShell 5.1 installed locked pypdf and resumed the same filled PDF. PowerShell 7 reused the application environment. Both split runs reopened and verified the complete static output. |
| Native document policy | 75/75 on the package-matched frozen source archive. Seventy-three PowerShell/BAT controls and two redacted exports passed with complete owned cleanup. |
| Supplemental real Calibre 9.17.0 | 10/10 from the immutable source archive with package members byte-matched. Real EPUB and AZW3, corruption and loopback controls preserved inherited converter state, all source inputs and owned cleanup. The loopback controls observed zero conversion requests. |

The shared native shell runner passed on both hosts. Each ran 300 Pester
tests, 36 syntax checks and 35 scoped analysis checks.

The form and PDF driver checks original and neighboring files, ordered page
ranges, physical markers, current text, rendered field regions, navigation,
and every serialized object. The BAT controls retain exact BAT bytes and use
the three documented copied PowerShell parameter defaults. The ebook driver
compares retained converter PDF content, RGB images and Unicode markers with
the output slices and checks actual EPUB navigation and chapter order. It binds
the application files and ZIP hash to the frozen release manifest. Existing
Calibre configuration, cache and plugin data are isolated from converter
probes and processing.

The executable native drivers are
[`characterize_form_support.py`](../../tests/document_policy/characterize_form_support.py),
[`characterize_managed_calibre.py`](../../tests/ebooks/characterize_managed_calibre.py)
and [`characterize_document_policy.py`](../../tests/document_policy/characterize_document_policy.py).
Their invocations are in the respective test READMEs. The Calibre run used
`--online` with both native hosts, the exact package manifest, the original
9.15.0 and 9.17.0 installations, and the verified official 9.18.0 offline
artifact for the separate offline scenario.

| Local raw receipt | SHA-256 |
| --- | --- |
| Exact form and PDF matrix | `7e5e4a66e18595f6d8eee2b72fb99e223c23c64bddd1ada5cf720a32104ac048` |
| Exact managed Calibre matrix | `6f35b5958720ef0bffd5492fc6704fcd8ff4e0bb487d1ea7e3bd444f730e269a` |
| Exact resource-action reproduction | `6a907d8005c0ca133caf491619d348d4ba7a96a44f813c517d6c156e2b3c2f45` |
| Exact cold form setup and repeat | `449e6654864923a3c88a64d1997c270522571fe038fc6fdb3a289bfad34f6a00` |
| Native document policy | `28a8937c242d867ef1329ea25ae230b81935c4542f0f0d1a5ca310250df2569c` |
| Supplemental Calibre 9.17.0 | `0a74b4fdfb7fe4f5f2c1673bf9f617ed2c67cbf9614c0abde330170a2598028e` |
| Native Windows PowerShell 5.1 shell regression | `4519979cdbb6d24f1934799d6444bd97a0aeb8154efb70d22ed058ef39fa2c72` |
| Native PowerShell 7 shell regression | `f245b12e379c3b14de6b632ed75afde8e4954a2558606e037cd5d6b10d389e06` |

## Coverage limits

The reported private PDFs and the owner's ten books were unavailable. The
reported refusal was reproduced with an original authored fixture.

Ordinary Store first-run consent/resume was `NOT_RUN`. The fresh extracted
application attempted natural discovery in a read-only preview. It found
a registered installed Python 3.14.7 below the existing 3.14.8 floor and
returned `runtime_not_found` without creating an app environment. The separate
successful consent run used an explicit compatible base. No natural runtime
was substituted.

Native Windows 10 and genuine XFA rendering were `NOT_RUN`. The authored
XFA detection sentinel returned its specific required-rendering error.
Visible signature artwork was checked as static graphics. Cryptographic
signature verification was outside the tested chapter-copy behavior.
