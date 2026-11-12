# Form-handling contract

## 1. Classification before rejection

Use bounded resolution of actual PDF object types. Catalog-key presence and
`get_fields() is None` are not sufficient classification. Consult both the
AcroForm field tree and every page's Widget annotations.

| Input structure | Required behavior |
|---|---|
| No AcroForm, no widgets | Existing page-copy behavior unchanged. |
| AcroForm null, empty dictionary, or empty/missing Fields with no widgets or XFA payload | Split normally; omit unused form container in output. |
| Resource-only form container and no actual fields/widgets | Preserve visible page content; drop unused form metadata only after inspection. |
| Ordinary non-signature field/widget tree | Normalize visible values to static page content. |
| Orphan widget with locally resolvable field data | Recover only unambiguous field relationships or flatten its validated appearance; do not discard it as an empty form. |
| Hidden/nonvisual data-only fields | Do not expose hidden values on pages. Drop nonvisual form data from static chapters with a count-only notice. |
| Signature or permissions structure | Omit signature validity data with a warning. Preserve a visible signature widget's static appearance when reproducible. |
| XFA null/empty array/empty stream, no dynamic content | Treat as an empty marker after type checks; do not imply genuine XFA was rendered. |
| Nonempty XFA with no verified complete static rendering | Distinct `xfa_rendering_required`/equivalent reason; no success from page-shell copying. |
| Malformed, cyclic or excessive field/widget structures | A specific bounded failure, no crash/hang or published partial results. |

Do not conflate a legitimate repeated reference with a directed field-tree cycle.
Some fields have several widgets on different pages; parent links are normal.
Use identity tracking appropriate to traversal direction and object ownership.
Bound work per document; limits must be justified and generous for real documents.
Never load remote resources or execute field JavaScript to classify a file.

## 2. Default output is static, not an editable form

A chapter splitter is not a form editor. Preserve visible page appearance and
current field values as page content. Declare non-editability in the plan/docs and
one warning. Keep selectable text/vector artwork where the source supports it.
Do not rasterize entire pages as the ordinary fix.

Prefer the existing pypdf dependency and public APIs. Its flattening feature is a
building block, not permission to delete all Widget annotations. A library routine
that updates field values may regenerate appearances; preserve already correct
appearances and original formatting where feasible, and test regeneration cases.

Do not pass every field through a flat text dictionary: names can be hierarchical,
widgets can share fields, values can be names/arrays, radio groups have children,
and choice export/display values differ. Resolve inherited `/FT`, `/Ff`, `/V`,
`/DA`, `/DR`, `/Q`, `/Opt` and related fields as needed, without logging values.

For missing/stale appearances, generate through a supported, tested path using
actual current values and available fonts. Respect `/NeedAppearances`; do not
preserve a known stale visual value as if it were current. Do not silently flatten
an empty rectangle for a visible nonempty value. Unsupported font/layout cases
must explain the concrete inability without deleting values or emitting success.
Missing-appearance support for ordinary text is mandatory, not an optional follow-up.

Button appearance dictionaries require correct `/AS` selection and the logical
field `/V`; state names are not always `/Yes`. Selected/unselected radio widgets
must agree. Preserve field borders, background, check marks and selected choice
labels. Respect annotation visibility flags; hidden/NoView data is not newly painted.
Do not use 'Print' as the sole visibility criterion for screen-reading chapters.

Handle `/Rect`, appearance `/BBox` and `/Matrix`, page rotation, crop/media boxes,
resource names and clipping without double rotation or scaling. Test transformed
and nonzero-origin appearances. Each appearance must reference only resources
needed for that visual content, not the parent form tree or unrelated pages.

## 3. Integration order

1. Open with the existing accessible/empty-password reader behavior.
2. Inspect and classify actual form structures without triggering the old blanket
   rejection. Apply the shared PDF page-copy policy to optional document features.
3. Resolve a single source-bound form plan shared by preview/execution; maintain
   original physical page mapping and original source identity.
4. On a private working representation, preserve/generate visible appearances;
   validate resources and omit form-specific action edges before copying.
5. Paint those appearances into their pages; only then remove widgets/form data.
   Do not delete unrelated links or static annotations.
6. Use the existing chapter partition and safe writer. Ensure later annotation
   projection/reopen validation recognizes the normalized result.
7. Reopen outputs, verify page coverage and static-content expectations, and only
   publish the completed new run directory when all normal checks pass.

Idempotence matters: a captured/prepared reader must not flatten a second time or
reset authentication. Bookmarks, metadata policy and local opening destinations
must survive the planned treatment. No early `get_fields()` success should skip
later necessary appearance checks. No second blanket ban should undo support.

## 4. Actions and signatures

Treat the saved field values and their visible appearances as the source of
truth. Omit form calculations, JavaScript, submit/reset/import actions, launch
actions and calculation-order links. The application performs no document action
or network request. An appearance that depends on unavailable recalculation
remains a processing failure. Chapter copies omit signature data and warn that
the source signature's validity does not transfer to them.

Only the validated static representation should reach the general page-copy
traversal. Classify `/AcroForm` and page widgets before omitting form trees.
Do not recursively copy whole
field trees/resources into each chapter. Hidden field values and content of
other chapters must not leak into serialized output objects.

## 5. XFA is not a synonym for AcroForm

The new error must not say 'AcroForm/XFA unsupported' for an empty AcroForm.
Nonempty XFA may override visible PDF pages. A nonempty page tree, a successful
ordinary renderer, or lack of AcroForm fields does not prove that XFA has a complete
static counterpart. A 'Please wait' page is not the form.

For this patch, mandatory support is ordinary AcroForms and harmless empty markers.
A real XFA input may only be accepted when a maintained local rendering/conversion
path demonstrably produces complete static pages with current data and layout.
Static-hybrid XFA must also be validated, not guessed from one flag or XML string.
If such a path is absent, return a specific reason explaining that XFA rendering
is required; do not represent this as full XFA support. Do not parse untrusted XML
with entity expansion/external resolution or enable V8/scripts just to make a
smoke test pass. Do not install a large renderer suite as a default dependency for
all PDF users. Any actual XFA renderer extension needs genuine regression inputs,
source/output visual comparisons, package/security review and explicit scope evidence.

The supplied XFA sample is deliberately a **detection-only sentinel**, not a real
XFA-renderer acceptance document. It must never be used to claim XFA conversion.
