# Seven milestones to public v1.0.0

36 focused tasks. Complete one task per thread; split a large task into recorded coherent substeps without losing its gate. Sequential dependencies keep local work predictable. Do not publish until M6.

## M0 — Baseline, scope and regression harness

Gate: A reconciled live checkout, reproducible baseline failures and an honest supported-tool matrix.

- **M0-T01** Reconcile the live repository and import the handoff.
- **M0-T02** Capture baseline behavior with safe fixtures.
- **M0-T03** Freeze support and dependency decisions.
- **M0-T04** Add local test and evidence scaffolding.

## M1 — Split correctness and shared plans

Gate: Manual and bookmark plans cover every physical page once, with parent-aware Level 2 boundaries.

- **M1-T01** Extract the existing Python engine narrowly.
- **M1-T02** Fix manual start-page parsing.
- **M1-T03** Normalize bookmarks and fix Level 1 coverage.
- **M1-T04** Make Level 2 parent-aware.
- **M1-T05** Unify planning, preview data and execution.
- **M1-T06** Distinguish no-outline and invalid-document results.

## M2 — Safe files, conversion and child processes

Gate: Originals/previous runs cannot be overwritten; dependencies and process streams are supervised.

- **M2-T01** Isolate runs and stage validated output.
- **M2-T02** Harden literal paths and output filenames.
- **M2-T03** Move ebook conversion into owned workspace.
- **M2-T04** Implement interpreter and converter preflight.
- **M2-T05** Supervise subprocesses and UTF-8 streams.

## M3 — Reliable CLI, launcher and diagnostics

Gate: Interactive and non-interactive entry points share tested behavior, exit codes and useful local diagnostics.

- **M3-T01** Propagate outcomes and cancel safely.
- **M3-T02** Add the small non-interactive CLI.
- **M3-T03** Polish launcher and menu behavior.
- **M3-T04** Show plans and converted-PDF page guidance.
- **M3-T05** Produce useful local manifests and redacted diagnostics.

## M4 — PDF fidelity, integration and CI

Gate: Supported documents preserve tested page behavior; rejected features and real conversions are covered.

- **M4-T01** Preserve tested PDF page fidelity and metadata.
- **M4-T02** Enforce unsupported-document policy.
- **M4-T03** Expand regression and fault-injection coverage.
- **M4-T04** Test real EPUB and AZW3 conversion.
- **M4-T05** Add least-privilege Windows CI.

## M5 — Release readiness on Windows

Gate: User documentation, packaging, security review and extracted-package Windows checks are complete.

- **M5-T01** Correct user docs, licensing and versioning.
- **M5-T02** Build an allowlisted release package.
- **M5-T03** Smoke-test the extracted candidate on Windows.
- **M5-T04** Perform the release security and dependency review.
- **M5-T05** Close acceptance gaps and merge final release scope.

## M6 — Publish and verify v1.0.0

Gate: A tested exact-commit package is publicly downloadable and the final main checkpoint is synchronized.

- **M6-T01** Freeze and verify release source commit R.
- **M6-T02** Build and hash the final assets from R.
- **M6-T03** Accept the exact final ZIP on Windows.
- **M6-T04** Create the fixed tag and inspect the v1.0.0 draft.
- **M6-T05** Publish and verify anonymous downloads.
- **M6-T06** Commit closure and verify final local/main synchronization.
