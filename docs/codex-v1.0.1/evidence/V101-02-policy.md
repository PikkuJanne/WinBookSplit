# V101-02 shared compatibility policy

PASS in policy decision, strict-validator, initial native-engine and provisioning
scope. The shared shipped JSON defines regular stable x64 CPython3.13.15+ and
3.14.7+, separate recommendations13.16/14.8, exact pypdf6.20.0 wheel lock and
retained Calibre9.15. See security-review.md for primary applicability evidence,
all four minor decisions and protected-per-user nuance. Later stable patches
are accepted numerically; latest recommendations are not equality/minimum rules.
Exact selected test runtimes are a separate non-shipped lock. Existing entry
points remain unmodified here; policy integration is V101-03.

Actual commands: runtime13 and runtime14 Python -I -B -m unittest discover
-s <checkout>/tests/python -p test_compatibility.py -v:26/26 per interpreter,
zero skips/native0. These are synthetic policy-boundary tests, including future
patch records, not installed future-version claims. The dependency-free loader
rejects malformed/duplicate policy data, numeric booleans, unsupported identity,
missing source references and damaged policy. A bootstrap3.12.14 run also
passed26/26; its initial test-helper defect/failure is retained separately.

Actual direct-engine smoke executed the unchanged current engine under
CPython3.13.16 and3.14.8 with runtime-only pypdf6.20.0, isolated imports, an
authored original10-page fixture and manual1,4,7. Each returned native0,
three chapters, all ten reopened markers in exact order/ranges and unchanged
source hash. This establishes initial genuine engine exercise on two minors;
it does not certify new discovery/setup, full regression or final ZIP acceptance.
JSON records the earlier source and changed policy/test bytes, external report
hash and explicit limits.

Official full Python ZIPs and PS7.6.6 portable ZIP were primary-hash verified,
bounded/exclusive extracted and rechecked in a new per-user cache root. Separate
runtime13/runtime14 and dev14 venvs used exact hash-required binary-only isolated
pip installs and passed pip check/import-origin probes. Existing shared Python,
Calibre, PATH, registry and execution policy were unchanged. Failed initial
GitHubmetadataHTTP415 receipt remains FAILED. Existing tools/ci/bootstrap.py
also provisioned hash-pinned isolated Pester6.2.0/PSScriptAnalyzer1.25.0
(native0); this is setup evidence, not test acceptance.

Windows10 and both final-ZIP native gates remain NOT_RUN. The production policy
is a reviewed candidate contract; public support still awaits later matrix and
native acceptance. Next: V101-03 integrate the same policy into every entry
point and PowerShell revalidation. This receipt does not claim its own commit.
