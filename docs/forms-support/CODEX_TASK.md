# Five-checkpoint functional repair: form-bearing PDFs

## Outcome

A user can split an ordinary PDF that happens to contain empty AcroForm metadata,
or real standard form fields, without editing the file externally, printing it
to PDF, upgrading Python manually, or toggling a hidden unsafe flag. Static chapter
outputs retain the visible form content. Empty-password and local opening-view
support from v1.0.4 continues working. Chapter output keeps readable physical pages
when the input has optional scripts, actions, attachments or portfolio metadata.
It retains supported navigation and static appearances, omits executable and
attachment behavior, and discloses signature limitations. Unverified XFA and
other unavailable required visual content cannot produce successful chapters.

This is an incremental document-handling feature. Keep PowerShell/BAT, the Python
engine, the shared policy/setup/process/output infrastructure, and optional Calibre.
Do not redesign the console or add a new application framework.

## 1. Establish the failure and add executable regressions

Read active instructions, Git state, origin, current branch, VERSION and live
releases. Preserve unrelated edits. Use a scoped branch such as
`codex/forms-v1.0.5` without resetting to the recorded baseline. If a release
already exists, inspect it and reconcile the request; never retag it.

Import only relevant handoff files into a new docs/forms-support directory after
checking conflicts, and incorporate the fixture generator into existing test
facilities if useful. Update active root routing with the owner's new request;
do not replace root AGENTS.md wholesale or rewrite old evidence.

Run the unmodified application against at least an empty form dictionary, a filled
text field, and a checkbox/radio form. Record the actual rejected stage/code.
Add failing automated expectations for the intended behavior before implementation.
Use the existing Windows runner when accessible. A component reproduction on
another OS must be labeled component-only.

Inspect both the shared reader/preflight and the output writer, including
`_chapter_annotations`, captured readers and validation of reopened outputs. Do not
fix only the catalog-key test. Keep classification, visual preservation, and security
policy in one shared path instead of divergent ad hoc retries.

Deliverable: small regression changes plus a short diagnosis. Do not stop with a
pure design note when executable reproduction is possible.

## 2. Classify forms and support harmless form metadata

Implement the classification and bounded reference checks in FORM_HANDLING.md.
Treat absent/null/validated empty AcroForm containers as normal PDFs. An empty
`/Fields` array alone is insufficient: scan page widgets and inspect XFA presence.
Accept resource-only harmless containers without losing page resources.

Add structured counts/kind information to the existing diagnostics/manifest.
Do not log field names or values. Use specific errors for malformed structures,
password requirements, unsupported form appearance and XFA rendering needs.
Preserve existing exit-code contracts where possible by using reason/detail codes
inside `unsupported_document` or `invalid_document`; avoid a protocol rewrite.

Reproduce the v1.0.4 name-tree refusal with separate authored JavaScript and
attachment fixtures. Audit the remaining presence-only bans and implement the
page-copy policy for optional document features. Check actual chapter bytes for
retained graphics/navigation and excluded objects, including shared references
and malformed or cyclic omitted branches. Keep widget-bearing inputs gated until
checkpoint 3 can preserve their visible form content.

Deliverable: harmless form metadata and ordinary optional-feature positives pass
through normal preview and splitting, with verified written chapters. The audit
records omissions and genuine failures in the active PDF policy.

## 3. Preserve ordinary AcroForm content in static chapters

Implement the default static normalization in FORM_HANDLING.md. Support text,
read-only text, multiline text, checkboxes, radio groups, choice/list fields,
parent-inherited properties, and fields with multiple widgets across chapters.
Handle common absent/stale appearances using the pinned library's supported
facilities. Do not invent text or drop values when resources are inadequate.

Work on a bounded private in-memory/owned working representation, never mutate
input bytes. Preserve the original page indices/bookmarks/source identity used by
the split plan; do not silently reflow/reorder or replace the original source hash
with the normalized intermediate's hash. The intermediate may have its own hash.

Flatten visible widget appearances into page content before removing interactive
field/widget data. Reuse the existing page-level writer and annotation projection
where possible. Do not clone the entire document/form tree into every chapter.
The existing graph guard must permit legitimate field-parent/widget relationships
only inside the new validated path, not become globally disabled.

Output is non-editable. Present one clear notice in the normal plan and record it
in warnings: `Form fields will be preserved as static page content; chapter copies
will not be fillable.` Noninteractive processing uses the same documented default
without prompting. An optional preserve-editability mode is not required for this
patch and must not delay the default workflow.

Maintain a distinct, truthful XFA branch. Empty XFA markers can be harmless.
Nonempty XFA is not an ordinary AcroForm promise. Do not add a runtime renderer
unless it is justified, packaged, isolated and tested against genuine XFA output.
Unrenderable or unverified XFA must return an XFA-specific reason, not success
with placeholder pages. State this boundary explicitly in release notes.

Run the same chapter-copy path for forms with saved values and optional scripts,
attachments and named navigation, including empty-password inputs. Preserve
visible signature appearances where reproducible and omit signature validity
metadata from chapter copies.

Deliverable: real static chapter PDFs, not merely successful classification.

## 4. Exercise the packaged workflow and prevent silent data loss

Extend existing Python, PowerShell and native tests rather than build a new test
platform. Apply ACCEPTANCE.md. Use the shipped pinned dependencies; local fixture
self-tests in this handoff are not substitutes. Update active assertions that
previously expected every `/AcroForm` or `/Widget` to fail; keep unrelated negative
security cases and historical evidence unchanged.

Run manual, Level 1 and Level 2 splitting with mixed ordinary/form pages,
empty-password encryption plus forms, benign opening destinations plus forms,
and source paths containing spaces/Unicode. Confirm ordinary first-run setup still
works and resumes the original PDF, and noninteractive/preview stays non-installing.
Preserve the functioning Store discovery path; never solve form errors by replacing
Python, injecting PATH, preflattening the test input externally or using a different
engine only in tests.

Render source/reference/output pages automatically with existing available
renderers. Compare widget regions and state markers, not only full-page averages
or page counts. Check chapter isolation across serialized objects, not only the
page tree. Keep renders and synthetic artifacts private to the test output unless
intentionally included in reviewed reports.

No human test programme. No owner screenshots, tests, kit returns, sign-off or
manual PDF preparation. Use authorized agent-accessible Windows 10/11 hosts; any
unavailable combination remains NOT_RUN and is stated in release evidence without
claiming coverage. Do not fabricate a new Windows 10 pass or make owner work a gate.

Deliverable: automated results, known unsupported cases, and a tested actual ZIP.

## 5. Publish the verified patch and synchronize

Target v1.0.5 after confirming the live state. Preserve all older tags/assets.
Use the existing release/package tools and allowlist; include new runtime helpers
where needed, but exclude fixtures, development packages, renders and private data
from the user distribution. Keep licensing/notices accurate; no new runtime
binary is implicitly approved by this handoff.

Update current README, PDF policy/fidelity docs, help, changelog and active status.
Say ordinary AcroForms are converted to static content, distinguish empty metadata,
record genuine XFA limits and retain signed-document limitations. Do not advertise
universal PDF/XFA compatibility. Do not downgrade dependency security fixes.

Freeze a commit, build the actual distribution, execute package acceptance, and
publish the accepted bytes as a non-draft, non-prerelease v1.0.5. Redownload public
assets without authentication; compare hashes and tag commit. Required CI failures
must be fixed, not waived as a way to force publication. Missing credentials/network
are genuine execution limits: report them precisely and keep completed independent
work; do not claim a published release without verification.

Commit/push only intended changes, use normal review/merge and repository protections.
Verify final clean local main against a freshly fetched/live remote. A later
closure-doc commit need not rewrite the tested release tag.

## EPUB and AZW3 correction

Exercise genuine Calibre conversion with an original EPUB containing multiple
chapters, navigation, images and Unicode. Verify existing and custom
installations, consent-based managed preparation and resume, reuse, declined
or unavailable preparation, and PDF-only independence. Bind the supported
official Portable download to reviewed hashes.

## Work cadence

The owner directed completion of checkpoints 3 through 5 and the EPUB correction
in one chat. Keep the evidence for each checkpoint and proceed through
verification and publication without an intermediate handover or approval request.
