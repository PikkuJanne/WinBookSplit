# Checkpoint 1 — Reproduction and executable regressions

The owner authorized ordinary AcroForm support and automated release work
through the five checkpoints. The baseline is the clean checkout at
`aef08eba589e82960774168edcd100b302798e6d`, equal to live `main` when work
began. The v1.0.4 tag dereferences to
`22745a9d4361a77817754d28c9499cbf8fb33532`. The committed release payload
matches the tag. The live tag and release queries found no v1.0.5. Check again
before tagging. No old tag, asset, or evidence file changed.

## Native regression

The [fixture generator](../../tests/fixtures/generate_form_fixtures.py) comes
from the handoff and writes original synthetic inputs and independent static
references outside Git. The
[native driver](../../tests/document_policy/characterize_form_support.py) runs
the actual PowerShell and BAT entry points with the pinned runtime. It asserts
the intended successful result. The four form cases intentionally fail the
checkpoint 1 regression because v1.0.4 still refuses them.

On a successful split, the driver checks the physical page partition, the
original source SHA-256 in the plan, source and prior-output preservation,
the expected publication members, and reopened page markers, visible values,
geometry, and serialized objects and decoded streams. It compares Poppler
renders to the independent references at 72 DPI with exact RGB matching for
these unrotated seeds, including per-field regions. It tests those comparisons
against blank regions and opposite checkbox and radio states. Reports include
the native argv, outcomes, fixture/reference hashes, output hashes, and render
receipts. Reports and generated PDFs remain outside Git.

Run from the checkout with an isolated developer interpreter containing the
locked requirements-dev packages, a separate pinned application runtime, and
both native PowerShell hosts:

```powershell
& $DevPython -I -B tests/document_policy/characterize_form_support.py `
  --application-python-path $RuntimePython --shell-path $PS51 --shell-path $PS7 `
  --renderer-path $Poppler --report $NewAbsoluteReport
```

Omit `--app-root` to run the checkout. Use `--app-root $FreshExtraction` to
run a specific package and `--case acroform_empty` for a focused check. Direct
PowerShell cases run the untouched scripts. BAT is byte-identical and calls a
copied PowerShell script with three controlled parameter defaults: output
directory, explicit pinned Python, and no final pause. The form processing
code is unchanged. This is a controlled BAT case with an explicit runtime.

### Observed results

Both runs used Windows 11 Pro build 26300, Windows PowerShell 5.1.26100.9711,
PowerShell 7.6.9, CPython 3.14.8 with pypdf 6.21.4, and Poppler 26.07.0. The
packaged run used a fresh extraction of the published v1.0.4 ZIP with SHA-256
`9d402f72092b20ad4d2a1985d8affef0ccfae525f520817e276185db4920a309`.
The checkout run had the same outcomes and matching output bytes.

| Entry point | Fixture and mode | Actual result in both runs |
|---|---|---|
| PowerShell 5.1 | Plain PDF, Manual | `split_complete`, exit 0, two verified chapters |
| PowerShell 7 | Plain PDF, Auto2 | `split_complete`, exit 0, four verified chapters |
| BAT | Plain PDF, Manual | `split_complete`, exit 0, two verified chapters |
| PowerShell 5.1 | Empty AcroForm dictionary, Auto1 | `unsupported_document`, exit 7, catalog rejection |
| PowerShell 7 | Filled text field, Auto2 | `unsupported_document`, exit 7, catalog rejection |
| BAT | Checkbox, radio, and combo controls, Manual | `unsupported_document`, exit 7, catalog rejection |
| PowerShell 5.1 | Orphan widget, Manual | `unsupported_document`, exit 7, annotation rejection |

The catalog rejection was `Interactive AcroForm/XFA documents are not supported.`
The orphan was refused with
`Interactive widgets, signatures, attachments and multimedia annotations are not supported.`
All four rejections occurred in shared reader preflight before a plan, with
zero chapter files and null execution and engine diagnostic. The driver
returned 1 on each run. Source hashes and earlier output matched before and
after every case. The render oracle detected all five blanked regions and both
opposite button-state pairs. The actual form cases did not reach output rendering.

Local raw report hashes:

- Released ZIP run: `81cffdb76e527fca40acc982726de79cde2171fc590cbb4342e07a7db38e9a1f`
- Checkout run: `42c33cd25142f57bd958c3305c98706f93e3a3d1ade3e1ea5d84f20dcaf2e98d`

Authored input SHA-256 values in both runs:

| Fixture | SHA-256 |
|---|---|
| Plain | `1ff86136bd0760ee1d76e678dc6e1ca4e6c0c89cd89971692fc640ca699bfc86` |
| Empty AcroForm | `178f487d0be622e8ef0461915bd72ec53b24eec1f816bc6f66e834bae047a313` |
| Filled text | `3398410378cd173483f5563ae24cca3b2f700c87eaa926d5c4719b55df24d11f` |
| Controls | `6ae77faed512e8e8233c230ec9d5985463228e72d30bc6a1926ee523397dd112` |
| Orphan widget | `004a8007ee81471f6144e6bd6cb2d1c655333af04149b4af12586f360b2c2c77` |

### Inherited application tests

`$RuntimePython -I -B tests/run_tests.py --_application-units-child` passed
all 172 tests in 148.208 seconds using CPython 3.14.8 and pypdf 6.21.4 in a
fresh developer-only runtime installed with `--require-hashes -r requirements.txt`.
The first run encountered four child CLI failures after the reused temporary
runtime lost package files. The fresh runtime passed the entire roster with
the same source. No application setup or runtime discovery code changed.
Passing log SHA-256:
`57077d7b82cc1d8b53a05a89f7b700f35f2ff394ba7328e01990fba5b256b97b`.

## Diagnosis

- `_preflight_pdf_features` in [winbooksplit_conversion.py](../../engine/winbooksplit_conversion.py#L293)
  refuses the presence of `/AcroForm` before checking the fields, page widgets,
  or XFA payload. Its separate page-annotation scan also refuses `/Widget`.
  The graph checks protect signatures, active content, attachments, and
  document references copied through page resources.
- `PolicyPdfReader` applies this policy for direct inputs and converted PDFs.
  `prepare_split` in [winbooksplit_engine.py](../../engine/winbooksplit_engine.py#L886)
  also checks a captured reader. Interactive retries use that same reader.
  Keep classification idempotent and preserve the original authenticated
  snapshot, hash, and page map for preview and execution.
- [`write_slice`](../../engine/winbooksplit_engine.py#L1795) excludes
  `/Annots`, `/AA`, and `/B`, then
  [`_chapter_annotations`](../../engine/winbooksplit_engine.py#L809) projects
  only its supported local links and static annotations. Accepting a widget
  at the reader alone would lose its visible appearance here. Checkpoint 3
  needs to paint validated content into the page before widget removal and keep
  form trees and other chapters' data outside each output.
- [`_validate_output`](../../engine/winbooksplit_engine.py#L1269) uses the
  shared reader for staged chapters and original converted-PDF copies. It
  reopens and resolves pages, checks the count and hashes the bytes. Once the
  input reader accepts AcroForms, chapter validation needs explicit static
  invariants. Retained original copies have a separate contract.

## Coverage limits

The latest owner PDF was unavailable. Its form type and behavior remain
unknown. Working form output, transformed appearances, actual XFA rendering,
Windows 10, and ordinary Store first-run/resume were not exercised in this
checkpoint. The release remains v1.0.4. The next checkpoint is [classification
and harmless metadata](STATUS_NEXT.md).
