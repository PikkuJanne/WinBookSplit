# Runtime policy and version validation

## One shipped contract
Introduce one small JSON policy, for example `engine/WinBookSplit.Compatibility.json`, with a validated schema version. Setup, PowerShell preflight, embedded/direct Python validation, diagnostic formatting, documentation tests and matrix generation consume it. No separate hard-coded lists that can drift. The example under templates is **unresolved planning data**, not production policy.

Represent supported Python branches as numeric major/minor plus minimum patch, optional precise exclusions, final/stable release requirement and explanatory source references. Keep the CPython implementation, Windows x64/64-bit process and regular-GIL checks. Do not compare version strings lexicographically; reject booleans/malformed numeric fields and distinguish prerelease/dev builds. `3.14.10` must sort after `3.14.8`. A later stable patch in an approved branch must not fail merely because it differs from a build pin. Future-patch records in tests are synthetic policy checks, never an installation/test claim.

Separate fields for compatibility floor, advisory/recommended patch, hard security exclusions and versions actually exercised. Determine security floors from applicable upstream fixes, not blanket fear or the developer's current version. Keep references and a review date; no network or auto-update policy at normal launch. Test the user's 3.14.2 record: the result must be either accepted under a reviewed policy or a specific explained security/compatibility floor, never the old equality error. Do not claim pypdf is missing from this log because that import did not occur.

## Verification sequence
1. Trusted absolute executable selection and bounded process transport remain intact.
2. Probe interpreter identity/architecture/version/stability and supported features, without importing pypdf for **setup discovery**.
3. For **processing**, require the locked pypdf import and distribution metadata from this interpreter's expected package roots, retaining all origin and isolated-mode checks.
4. Parse exactly one schema-valid result and revalidate compatibility in PowerShell against the same policy. Do not rely solely on `ok: true` supplied by a candidate.
5. Log the selected interpreter, policy version and actual package version; do not label an accepted later patch as specifically tested.

An explicit invalid `-PythonPath` or broken authoritative application .venv remains a clear failure; do not silently switch to a different interpreter. If normal automatic discovery finds several valid candidates, select deterministically using policy preference, recording selection. Do not require pypdf in a global interpreter just to set up a fresh venv.

## Separation from pypdf and build locks
The runtime pypdf lock starts with reviewed 6.20.0 and exact hashes; retest any necessary security update. Leave the encryption/private-reader guards intact. The lock is installed by setup and can be invisible to casual users without being made broad or unpinned. Put exact release/test tool versions into a development/release lock or record rather than Python runtime acceptance.

Current pypdf package metadata advertises a lower Python floor than the proposed application range; this is only feasibility evidence. Audit app syntax, pathlib methods, ctypes declarations and developer tooling on each branch. A successful `import pypdf` is not full PDF-engine compatibility.

## Fail safely but informatively
Differentiate no Python, unsupported Python, security-floor failure, unsupported architecture/free-threaded build, missing locked package, wrong package origin, damaged policy and incomplete probe. Preserve existing application outcome/exit contracts; add a subordinate diagnostic reason rather than needlessly breaking consumers. Missing/malformed policy must never fall back to accept-anything.
