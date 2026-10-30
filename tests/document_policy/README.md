# Unsupported-document policy acceptance

The current producer/validator migrates the M4-T02 route to the repaired PDF
contract. Empty-password execution and preview are positive cases; genuine
nonempty-password inputs require `password_required` / 7. Rejected-input support
exports use the password-required fixture. The original matrix and receipts stay
historical; no production ban is restored to satisfy obsolete expectations.
Current execution results belong to the checkpoint evidence linked from
[STATUS_NEXT](../../docs/functional-repair/STATUS_NEXT.md), not this description.
Human/owner tests, sign-off and report requests are out of scope under the latest
owner order. Missing native scenarios remain NOT_RUN, with no owner action required.

`characterize_document_policy.py` uses 25 original generated PDF fixtures. It
starts 30 application controls in each actual supported PowerShell host and
13 through the shipped batch launcher, then exports one rejected-input run
record per host. The 75 controls require explicit host paths and isolated
developer Python; they do not use GUI automation, human input, private books,
password arguments, or Calibre.

The fixtures cover user- and empty-user-password encryption, AcroForm/XFA,
authored signature structures, orphan widgets, catalog actions/JavaScript,
attachments/associated files, portfolios, 3D/media annotations, malformed and
pathological page/outline structures. Signature fixtures prove structure
detection, not cryptographic signature validity. Their actions and payloads are
inert original data and are never opened or executed. Ordinary manual/Level 1/
Level 2 controls and the documented page-action exclusion keep the complete
physical-page plan, reopened content hashes, and first/last-page identities.

The final fixture labels use 8-point Helvetica so the complete first/last-page
marker fits inside the 120-point page. The immutable baseline fixtures retain
their earlier 12-point labels; this generated-fixture QA correction does not
rewrite the baseline reproduction or its provenance.

Unsupported feature inputs must return `unsupported_document`/7 with no extraction,
plan, interaction request, chapter/staging publication, or fallback. Malformed
inputs retain their declared native 6 result. Genuine nonempty-password inputs
return `password_required`/7 with the same no-output requirements. Empty-password
inputs must produce complete reopened output; their Preview must cover every
physical page without writing chapters or disk diagnostics. New chapter PDFs
are unencrypted. Ordinary rejected runs may keep
their authenticated local log and finalized metadata-only run record; rejected
Preview controls must write nothing. Both actual rejected run exports must
contain only the allowlisted redacted summary.

Each receipt binds native argv, UTF-8 streams/hashes, finite deadlines, process
and stream shutdown, actual host/version/settings, copied application bytes,
fixture provenance, immutable read-only source/neighbor/prior identities, and
exact owned cleanup. Batch controls use the unmodified launcher with three
disclosed copied PowerShell parameter defaults for explicit Python/output paths
and pause suppression; the receipt proves the entire modified script against
the original raw script bytes. It never supplies an application NonInteractive
default or alters policy behavior. Failures and unproved shutdown retain their
authored workspaces. Receipt unit tests are synthetic negative controls and
are not native acceptance evidence.

Run from the repository using the pinned developer interpreter:

```text
python -I -B tests/run_tests.py --layer document-policy --report <new-external-report.json> --shell-path <absolute-powershell.exe> --shell-path <absolute-pwsh.exe>
```

The integrated full runner includes this route as stage 20. The report must
pass `validate_document_policy_report` before it can support a pass claim.
