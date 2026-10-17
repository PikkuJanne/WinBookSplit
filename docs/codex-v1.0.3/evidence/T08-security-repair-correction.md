# T08 - standalone legacy export correction

The first broad implementation run is retained, not relabelled as passing.
Its 708 Python tests had 707 passes/1 error: historical standalone export referenced
RUNTIME_POLICY when it was None. The run also recorded source_unchanged=false
because reviewed characterization/docs corrections occurred after suite start;
current delivery-pin preparation is another later source change. Both shell
children exited 0 with 252/252 Pester tests, 35 syntax/34 scoped-analysis entries and
no skipped/not-run cases. Their aggregate also failed source preservation, so
neither broad aggregate is final-source acceptance.

The package/kit at provisional 9f7e081 are retained external artifacts, not a
release freeze to deliver. They predate this helper correction. Production's
34 payload files are untouched by this correction; a new committed source,
manifest and native kit must bind the correction and final CI honestly.

The narrow fix makes export_report compare each ordinary child package version
to the report header's already-schema-validated pypdf version, while broken
children still require None. This keeps self-contained historical exports valid
without loading current application policy. Actual live preparation/validation
still checks the bound shipped policy, version and package origin. Store schema 2
and its no-substitution gate are unchanged.

Two added regression methods cover coherent 6.20.0 and 6.21.1 exports with policy
None, header/child mismatches and invalid broken-child package claims. The
existing historical test explicitly checks the immutable old report's 6.20.0
header/children. Its private-runtime success still cannot satisfy Store schema 2.
All 47 focused native-smoke/Store-gate tests pass in the new developer environment.
Independent read-only review found no remaining issue in the correction.

Raw broad/focused reports, streams and source hashes stay external and are bound
in JSON. No original report, archive, owner installation, policy, PATH or alias
was changed. Original R243d31e and its accepted T07 remain intact for their own
bytes. STORE-024 and new-candidate native acceptance stay NOT_RUN. No publish,
tag or main merge occurred. Final CI and corrected package/delivery receipts
follow separately; no pending outcome is a pass. The correction commit hash
belongs in a subsequent receipt/handover, without self-reference.
