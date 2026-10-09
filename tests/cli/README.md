# Actual Windows CLI acceptance

Run with the pinned developer Python, both explicit supported hosts, the recorded
real Calibre 9.15.0 executable and a new external report:

```powershell
& $DeveloperPython -I -B tests/run_tests.py --layer cli --shell-path $PS51 --shell-path $PS7 --calibre-path $Calibre --report $NewExternalReport
```

The full route retains all inherited stages and adds CLI acceptance as stage 15.
Every application process uses `subprocess.DEVNULL` stdin and the actual host's
`-NonInteractive`. Application `-NonInteractive` is separately exercised; complete
explicit Manual with only `-NoPause`, and Preview without application
`-NonInteractive`, prove their independent behavior. PDF manual/Level1/Level2
execution reopens every actual chapter for exact physical identity/order/content,
filenames, manifest bytes and complete ranges. PDF previews and actual offline
EPUB/genuine AZW3 previews return the same immutable validated plan with zero
execution/writes. Independent real conversion references bind ebook physical
content, original source identity and absent temporary conversion files.
The proposed EPUB Auto Level1 `-KeepConvertedPdf -NonInteractive` example also
executes under both hosts and checks the separately retained full converted PDF
against actual captured bytes, manifest metadata and independent physical content.

Missing/contradictory choices, invalid tokens, unknown arguments and incompatible
preview retention must reject with 2 before dependency discovery or output/log
creation. Explicit missing Python/Calibre rejects with 3; out-of-range physical
starts reject with 2 through the existing engine. Flat/parent-only bookmark
controls return 5 without automatic fallback. Version runs in a copied application
containing only the PS entrypoint and shared version contract, with no Python or
Calibre on the controlled PATH. Full comment-based help is read under both hosts.

All sources are authored synthetic material. Literal Unicode/special-character
paths, actual read-only sources, same-name ebook PDF neighbors and prior output
files retain exact identity/attributes/bytes. Exact emitted chapter/log/diagnostic
children use authenticated held-object cleanup. Any failed test retains its
authored workspace; a timeout stops only the retained parent handle and preserves
the workspace because descendant stop is unproved. Stored execution policies and
source/application bytes remain unchanged. Strict independent receipt validation
rejects incomplete or contradictory promised evidence; structural validator units
are not Windows acceptance. This layer makes no human Explorer, clean OS,
rendered-fidelity, final-package or public-release claim.
