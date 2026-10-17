# Offline handoff helpers

These are developer handoff tools, not WinBookSplit runtime files. They read local
files only and do not invoke an interpreter discovery, install anything, launch
Windows processes, query GitHub, change the repository, or publish a release.
Use an existing developer Python (3.11+) outside the owner test scenario.

From the handoff/repository root:

```text
python -B tools/codex-v1.0.3/check_handoff.py
python -B -m unittest discover -s tools/codex-v1.0.3/tests -v
python -B tools/codex-v1.0.3/check_native_store_report.py report.json --expected-commit RELEASE_COMMIT --expected-zip-sha256 ACCEPTED_ZIP_SHA256
```

Substitute actual externally verified release commit and accepted archive hash in
the final command. Do not run that command with placeholders and count its error
as evidence of an application failure. `check_handoff.py --mode bundle` validates
the original bundle and its manifest. `--mode checkout` validates imported task
links, fixtures and active routing; it scans only the two imported directories
and makes no pristine-byte claim. The default `auto` prioritizes an existing
handoff manifest (bundle mode), then selects checkout mode when root AGENTS.md
exists. The root handoff manifest and prompt files are not copied into a
changing checkout.

`check_native_store_report.py` returns 0 only for a structurally consistent
existing-owner Windows10 Store report with complete no-substitution observations,
pre-existing runtime correlation and six matching receipt envelopes/files. It
returns 2 for blocked/missing/malformed/mismatched or disqualifying evidence.
It does NOT independently prove that the reported Windows actions happened.
Review native receipts with the maintained application validators before release.

## Native report adapter

The automated native kit should write the report shape in
`docs/codex-v1.0.3/templates/native-store-report.json`. Null fields are intentionally
unknown, not false. Hashes are lowercase SHA256; source commit is a full 40-character
Git SHA. `runtime_ref` is a stable normalized/redacted identity token, not a
runtime path to be supplied to application discovery. The before/after Store
inventory must be unchanged; setup's base and the new app venv's base both match
the discovered member. Configuration observations are made by the real kit. Schema 2 requires the
explicit disabled-before-discovery acquisition policy, false attempted/used
flags, and verification-only inventory that never supplies selection to the app.
A fallback attempt disqualifies this scenario even when no runtime was acquired.
Schema 1 and historical private-runtime reports cannot satisfy this gate.

Each `artifacts` entry names one of `registration`, `discovery`, `setup`, `bat`,
`pdf_validation`, `preservation` and contains `path` (a relative forward-slash
path) and `sha256` for the returned redacted receipt. Each receipt is plain UTF8
JSON with `schema_version: 2`, `kind`, `run_id`, `source_commit`, `zip_sha256`,
`summary`, and a nonempty `records` list of actual sanitized command/output and
observation records. `summary` must match the corresponding report projection;
`expected_summaries` documents this mapping. The maintained native/PDF report can
be included in those records; preserve detailed argument vectors and exit codes.

Application records carry `role: application`, an empty override environment,
and literal process argv plus successful bounded supervision. The checker rejects
contradictory runtime inputs in these records. Native stream totals and
`raw_stream_sha256` describe the local original bytes; `redacted_stream_bytes`
measures the returned sanitized strings separately. Sanitizing paths must not
silently relabel shortened strings as the original process byte totals.

The tests generate synthetic positive/negative contract inputs inside temporary
directories. Even their positive inputs are NOT executed-native reports. No
passing native report is distributed. The shipped template must fail the gate.

The checker defends against accidental missing evidence, report substitution and
unsafe receipt paths; it is not a cryptographic attestation service, an untrusted
filesystem sandbox, or a replacement for code/native review. During review keep
the report directory stable. Do not weaken it merely to accept old private-runtime
results; version an adapter only for actual new evidence fields, with negative tests.
