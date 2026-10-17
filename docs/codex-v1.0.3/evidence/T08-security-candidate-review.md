# T08 — bounded pypdf candidate review

Status: preparation review complete; **no release clearance**. The owner asked
to go ahead after the gate-only continuation. Independent review identifies a
concrete proposed repair: pypdf 6.21.1 with the exact wheel below. The earlier
gate-only scope remains in effect for production writes while the explicit scope
question is pending. No dependency, application, package, frozen source,
accepted receipt or native acceptance status changed.

The review began clean at `bc3c3fc720ea954083ab6b7e380eb5582d9e3360`, tree
`07fc72ada298d87af1b11d04871c8f2fd1fca003`, on
`codex/v1.0.3-store-discovery`, equal to freshly queried live origin.
The required instructions/evidence remain reviewed in this chat.

## Exact proposed dependency

Current official [parser advisory](https://github.com/py-pdf/pypdf/security/advisories/GHSA-f54x-cmq8-qg57)
and [content-stream advisory](https://github.com/py-pdf/pypdf/security/advisories/GHSA-cgf9-327m-v77c)
name 6.21.0 as fixed. The [6.21.0 release](https://github.com/py-pdf/pypdf/releases/tag/6.21.0)
corroborates parser limits. The [6.21.1 release](https://github.com/py-pdf/pypdf/releases/tag/6.21.1)
and [JPEG advisory](https://github.com/py-pdf/pypdf/security/advisories/GHSA-5hx5-vpf9-5wcv)
identify the additional JPEG fix. The retained form-field advisory metadata/prose
discrepancy remains recorded in T08-gate-review.md; no comprehensive exemption
or silent metadata correction is claimed.

Official [PyPI metadata](https://pypi.org/pypi/pypdf/6.21.1/json) and an actual
bounded wheel download establish:

| Candidate field | Verified value |
| --- | --- |
| Version | 6.21.1, non-yanked |
| Wheel | pypdf-6.21.1-py3-none-any.whl |
| Bytes | 405134 |
| SHA256 | `553e2d42f5bb3bb2a785a7d3de8ef90bb7df74b5413eff5a7b4388007ea9254b` |
| Requires-Python | >=3.9 |
| License | Same 1605-byte LICENSE as verified 6.20.0 wheel |

No extras are proposed. Metadata feasibility and the empty PyPI vulnerability
list do not establish application compatibility or general security. Raw wheel,
metadata and probe outputs remain in a new owned external developer directory.
The companion JSON binds them without publishing private paths.

## Actual small dependency experiments

The official wheels were imported directly in separate bundled developer Python
3.12.14 processes with `-I -B`. No package was installed. No application
origin/runtime check was bypassed: the application was not invoked. Python
3.12.14 is outside the supported application policy. Each child had a 20-second
deadline, separately captured streams, actual exit0 and empty stderr. The public
probe copy is T08-security-candidate-guard.py.

Three-item/token positive controls parse in both versions. The old wheel lacks
the array/dictionary/content parsing configuration fields and accepts four
items/tokens. With 6.21.1's corresponding limits temporarily reduced to three,
each fourth-item/token input raises LimitReachedError. Observed post-context
values are 100000, 100000 and 2000000, matching upstream defaults. The script
records these values without a separate equality assertion.

These are low-limit guard demonstrations, not exploit reproduction, memory
measurement, full-threshold or complete PDF tests. ContentStream.operations is
accessed explicitly by this developer probe; that is not proof of this path in
production. Nested budgets, malformed data, form/image fixes, previous filter
guards, supported-runtime compatibility, setup, BAT and native acceptance were
not tested. An independent read-only audit verified probe/result/wheel bindings
and these limits without rerunning the experiment.

## Concrete implementation scope and remaining gates

The smallest proposed repair changes the authoritative requirements/policy lock,
maintained test-runtime lock, package-input and kit prerequisite pins, hardcoded
current tooling assertions, mutable synthetic expectations and current dependency
documentation. Setup hash validation, isolated import/distribution/origin checks,
Python floors, explicit-runtime authority, process bounds and engine stay intact.
Historical reports, owner contracts and frozen archives must remain untouched.

Implementation needs maintained low-limit guards, previous filter guards,
origin/hash rejection, authored PDF/coverage/fidelity and setup checks, supported
runtimes and both-shell regressions. A changed candidate then needs a new
committed freeze, new exact bytes and new native acceptance. Accepted R results
cannot be transferred to changed bytes.

T07 remains COMPLETE for R `243d31ee03efac35dc6f7e88c0c1dfd1203081ed` and ZIP
`70a76bd04f5645e195972bf075ff53d651c9993696a40065241390805d2ee1de`.
STORE-024 remains NOT_RUN: no new already-enforced target or denial receipt was
supplied. The accepted Store run was not repeated. T04 remains
IMPLEMENTED_PENDING_NATIVE; T08/STORE-043 through STORE-048 remain BLOCKED/NOT_RUN.
Retained 6.20.0's security release disposition remains unresolved. No desktop
policy, PATH, alias, runtime/Manager installation, tag, publication or main merge
occurred.

Only new review evidence and active routing/task links change. Initial harmless
read-only filename/path mistakes, output-truncation parsing and an atomic patch
context failure are recorded in JSON. Completed downloads/probes were retained.
Handoff, unchanged statuses, privacy, evidence pins, intended paths and staged
bytes are checked before ordinary commit/push. Final checkpoint SHA and actual
clean/fresh live equality belong in the external handover. Stop after this
checkpoint; apply the explicit scope answer before product writes.
