# Release-oriented safety review

## Assets and trust boundaries

Protect original documents, neighboring files, prior output runs, credentials/paths in diagnostics, the source repository, and released artifact integrity. Inputs, metadata, bookmarks, filenames, Calibre diagnostic text and PDFs/ebooks are untrusted. The user owns the local account; document parsing is not sandboxed by this tool. Runtime dependencies and executables are a separate trust boundary.

## Required checks

- No shell evaluation of document-controlled strings; inspect Python/PowerShell/batch argument handling and test hostile filenames.
- No source/neighbor overwrite, deletion outside an owned run directory, unsafe symlink/junction cleanup or shared temporary engine file.
- No silent global installation, administrator requirement, global execution-policy weakening, telemetry, document upload or DRM-removal plugin.
- Bounded process output, explicit stderr drain, safe process-tree cancellation, and resource-aware handling of malformed inputs.
- Interpreter/converter path provenance, pinned reviewed dependencies, current advisory review and remediation decisions.
- No secrets or private PDFs in Git history changes, test fixtures, logs, CI artifacts, ZIPs, screenshots or public issue reports.
- CI uses least privilege, pinned reviewed action revisions, no untrusted PR data interpolated into a privileged shell and no automatic release from ordinary pushes. Existing repository protections are respected, not changed to make CI pass.
- Package is allowlisted, tied to a commit, and verified after public download. Checksums do not authenticate an unsigned publisher. Optional signing may be added only without an unapproved purchase and must be validated accurately.

Review concrete diffs and test failures, not only scanner output. If an automated scanner is unavailable, record that; a scoped manual review is not a substitute for an invented scan. Do not declare the tool '100% secure'. Severity is contextual; unresolved data-loss, arbitrary execution, unsafe cleanup, secret-exposure or release-mismatch issues block publication.
