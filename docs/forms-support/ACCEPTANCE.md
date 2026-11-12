# Acceptance: actual split output, not just preflight

## Reuse the existing suite

The fixture generator and self-tests in this handoff validate test inputs only.
Codex must add positive and negative **application** tests using the installed
pinned runtime, PowerShell entry point, and BAT launcher. No replacement demo
splitter can satisfy acceptance.

Check every generated input's hash before/after. Exercise Manual, Auto1 and Auto2
on form-bearing pages. Synthetic fixture bookmarks and expected page markers
make skipped, duplicated, reordered or leaked pages detectable.

## Required case matrix

| Group | Cases and assertions |
|---|---|
| Harmless metadata | Absent/null/empty AcroForm; empty Fields; resource-only container; null/empty XFA markers. Normal output, exact original page content. |
| Text | Filled, blank, read-only, multiline, inherited font/resources, missing AP and stale AP/NeedAppearances. Current values visible, no replacement glyphs/truncation. |
| Buttons | Checkboxes on/off with nonstandard state names; radio group across pages. Render state matches intended value; opposite state cannot pass. |
| Choices | Combo export/display strings and multiselect list. Preserve selected labels/marks, not merely internal export values. |
| Hierarchy | Shared field with two widgets in separate chapters; nested names; same leaf name under different parents; orphan widget. No outside-chapter data serialized. |
| Geometry | CropBox offset, rotation, nonzero appearance BBox/Matrix, appearance resource collision. No double translation/rotation or clipped value. |
| Nonvisual data | Hidden widgets and data-only fields. Do not print secret values or disclose them in diagnostics. Report omitted interactivity/data without values. |
| Scripts/actions and attachments | Cover JavaScript and EmbeddedFiles trees alone, together, empty or cyclic. Exercise catalog, page, annotation and resource actions, associated files, file-attachment icons and custom appearances. Combine saved form values, actions in a drawn Form XObject, attachments and empty-password encryption. Verify retained text, graphics, resource names and local navigation. Inspect all serialized objects for omitted behavior and payloads. |
| Navigation/layers | Preserve supported named destinations, bookmarks, local links and visible optional-content state alongside excluded trees. Exercise preview and captured-reader reuse. |
| XFA | Distinguish empty markers from nonempty XFA. Detection-only sentinel cannot pass as converted. Genuine hybrid/dynamic XFA acceptance requires genuine renderer evidence, not ordinary-page copying. |
| Signatures/corruption | Copy readable page content and disclose loss of signature validity. A signature widget with reproducible appearance retains its visible content. Required malformed/cyclic graphs terminate clearly. Authored signature dictionaries establish structural behavior only. |
| Existing successes | v1.0.4 empty-password PDFs, local opening destinations, plain/scanned PDFs, EPUB/AZW3, setup consent, Store discovery, Unicode/literal paths and no-overwrite stay working. |
| Ebook setup and conversion | On the exact ZIP, exercise existing and explicit converters, rejection without fallback, consented official download and resume, offline preparation beside an older converter, reuse with the network unavailable, cancel, read-only checks, and genuine conversion failure. Verify illustrated Unicode EPUB chapter text, images, outline and complete converted-page coverage. Include actual AZW3 and a PDF with no converter. |

## Visual and semantic checks

For positive form cases, output text must contain the expected synthetic value and
page markers; field values alone in a dictionary do not prove visible output.
Render each field-bearing source/reference and output page at a consistent DPI.
Use an independent renderer from the splitting library. Existing Poppler/native
render tooling can be reused. Use a second renderer for tricky form appearances
when available; disclose missing comparison instead of creating human work.

Compare **field regions**, not just the full-page pixel average: a missing small
checkbox can disappear in a large blank page and pass a weak average threshold.
Test the oracle with deliberately blanked field regions and flipped checkbox/radio
states to ensure it fails. Decide tolerances before observing candidate results;
retain diffs explaining any non-identical accepted rendering. The bundled static
reference for missing AP contains the current value drawn directly, independent
of viewer appearance generation.

Verify no remaining Widget annotations or operative AcroForm/XFA in default static
outputs, no unintended active action/calc metadata, no dangling field relationships,
and no hidden values/other-chapter markers in serialized objects or decompressed
content. Preserve ordinary non-form annotations through the existing policy.
Do not reject harmless PDF **Form XObjects**: these are graphics resources, not
AcroForm fields, and may be needed for the flattened appearance.

A source has fields only on page 2? Page 1 must remain unchanged. A field has
widgets on pages 2 and 3? Each selected chapter contains only its own visible widget.
If a field cannot be rendered reliably, produce no published successful partial
run. Existing staged ownership/no-overwrite checks remain mandatory.

## Runtime and packaging

Use actual application paths in a fresh extraction of the candidate ZIP. Include
setup-to-resume of a form PDF with no app venv, and an already-healthy environment.
Noninteractive/preview must not prompt/install. Full installed-path tests run on
available supported Windows hosts, including inbox PowerShell 5.1. Run the
existing PowerShell 7 regression where available. Windows Server CI is not a
Windows 10 desktop claim. Do not request owner tests or a new sample just to start.

Record command, OS, shell, Python/pypdf version, source commit, ZIP hash, fixture
hash, exit, outputs, original/output hashes, page mapping and render comparison.
Keep evidence succinct and append-only for completed checkpoints. Do not record
private PDF content or field values in public logs.

## Publication checklist

* Required ordinary AcroForm positive cases really split through the application.
* Targeted error/warning cases are precise, not the old combined blanket message.
* Old successful PDF/setup/discovery behaviors are not regressed.
* Documentation says static form output and states XFA/signature limits honestly.
* The accepted user ZIP is from the tagged commit, not a developer-only checkout.
* v1.0.5 public assets match accepted hashes when redownloaded anonymously.
* Prior releases unchanged; final local/main synchronized; any unavailable native
  scenario disclosed as NOT_RUN, never an owner task or an invented pass.
