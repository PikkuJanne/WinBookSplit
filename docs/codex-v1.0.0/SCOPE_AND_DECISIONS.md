# Scope and initial decisions

## User-mandated

U01. Improve, do not rewrite. Preserve the existing entry points and backend choices.
U02. Documents are processed locally. The future website presents/distributes tools; it is a separate project and has no tasks here.
U03. The only public GitHub release in this project is `v1.0.0`; do not publish alpha/beta/RC/0.x or invent prior releases.
U04. Keep local work and GitHub synchronized at every completed task checkpoint. Finish only after publication and verification, not at 'ready'.

## Concrete implementation defaults

These resolve details left open by the review. Codex records any justified deviation before implementing it; do not silently remove safety gates.

D01. Keep manual **start-page** input. Accept comma-separated ASCII decimal integers with surrounding whitespace. Reject empty tokens/input, signs, decimals, ranges, letters, and out-of-range values. Dedupe/sort valid starts with a notice; always include physical page 1. `1` means a valid whole-document one-section run with an explicit message.
D02. All v1.0.0 runs preserve every input page exactly once. Page exclusion/range extraction is deferred. Front matter is its own segment. Level 2 splitting is constrained by Level 1 parent intervals, with parent-opening and no-child fallback segments.
D03. Per-run unique destination beneath Documents by default; `-OutputDirectory` changes the base. No overwrite mode in v1.0.0. Conversion and output staging are private per-run; final output is promoted only after validation.
D04. Split PDF output only. EPUB/AZW3 conversion is required support, not a simulated optional feature. Keep a tested compatibility conversion profile initially; add a small optional layout preset only if it is verified without expanding the release.
D05. Encrypted PDFs are **detected and rejected clearly** in v1.0.0, including owner-password/empty-open-password cases. No password parameter, decrypt/re-encrypt path, or crypto dependency is required. A future secure password feature is deferred. DRM-protected ebooks are not supported; do not add removal plugins.
D06. Target Windows 11 x64 with Windows PowerShell 5.1 and a tested PowerShell 7 version. Record exact builds. The historical README's Windows 10 claim is not evidence: label Windows 10 unverified/not part of v1.0.0 support unless it is actually tested. No ARM, Linux/macOS runtime, or live UNC support claim without evidence. Ordinary paths on the local workstation are mandatory. Do not require new hardware for unsupported platforms.
D07. M0 selects exact supported Python/pypdf/Calibre and developer-tool versions from current documentation plus local tests. Prefer a project-local virtual environment and pinned requirements, no bundled Python/Calibre or silent global installs. Separate dev-only dependencies.
D08. The default release may be unsigned, with accurate notices and SHA-256 checksums. No certificate purchase or SmartScreen guarantee. Keep MIT for project code; review notices for anything actually redistributed.
D09. Offer no-file interactive selection; reject more than one dropped input clearly. Batch/multi-file jobs are deferred. Invalid menu answers reprompt; no silent default that changes intent.
D10. Normal commits, pushes, scoped PRs/merges, and final publication after passing gates are delegated by the supplied start prompt. No history rewrite, tag replacement, deletion, security-setting changes, new paid service, or different public version is authorized.
D11. A clean extracted-package Windows test on an isolated current workstation setup is acceptable if explicitly described as such: fresh venv, controlled dependency paths, no imports from the checkout. A Sandbox/VM is preferred where available, not a new hardware prerequisite. Never describe this as a clean OS installation unless it is one. Actual Explorer drag/drop and Calibre runs remain required.
D12. Basic page fidelity, common annotations, source author/chapter title, and a chapter-start bookmark are in scope. Reconstructing arbitrary forms, signatures, attachments, document scripts, or cross-chapter navigation is not. Reject unsupported features that would otherwise be silently broken; warn about the defined link limits. See `specs/PDF_SUPPORT.md`.

## Change record

For a changed default add date, decision ID, evidence, user-impact, acceptance-case changes, and authority. No new public version or website task may be introduced under 'small improvement'. A blocked test is not a reason to silently redefine the finish line.
