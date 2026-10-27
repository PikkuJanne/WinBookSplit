# T08 - hosted CI acquisition failure and same-byte repair

SourceG74efe6b CI40528250903 failed its reference job because the Gallery API
returned HTTP404 for the pinned PSScriptAnalyzer1.25.0 archive. Actual reference
Python713tests and all three applicationjobs passed; hosted PS5.1/PS7 steps were
skipped and the strict reference validator failed for missing reports. This is
retained as a failed whole workflow, not relabelled PASS. Raw API/log/artifact and
validator receipts remain external in ci-G; app coverage remains Server/PDF-only.

The official PowerShell/PSScriptAnalyzer1.25.0 GitHub release asset has exactly
14658674 bytes and the existing locked archive SHA256. Actual bounded downloads
from that official release, Gallery and GalleryCDN compare byte-identical; the
latter routes subsequently returned200 locally, so no permanent disappearance is
claimed. The remedy selects the verified official GitHub release URL for this
same package, with no version/hash/fallback change. Pester remains unchanged.

Only bootstrap acquisition data/selection, its focused tests and acquisition
documentation change. Extractor,
path/size/hash/timeout controls and their AST/bounds remain unchanged. Eight
focused tests pass, including bad-hash/HTTP/read-bound refusal controls. A fresh
isolated actual bootstrap passes and produces the same21Pester/53Analyzer file
hashes as the previously validated localtools; that comparison is not cache
adoption. No globalmodule, policy or application/runtime guard changes occur.

Local G broad tests separately pass713Python tests and252Pester tests in each
inbox5.1/reviewed7.6.6 host, source_unchanged=true. Three G Windows11 application
matrices pass150units/10PDF-Calibrecases/21probes each. The new exact-source hosted
workflow must pass after this acquisition repair; it remains pending at the
correction source checkpoint. Product34payload bytes are unchanged, and new
manifest/freeze bindings will use the corrected commit rather than G.

Ownerclosedmanualscenarios stay NOT_RUN/owner-waived, never PASS. This tool repair
does not waive hostedreferencechecks or alterthenativechecker. Current final
source/CI/package/publication receipts follow separately. Sourcecommitidentity
belongs in the subsequent receipt, avoiding self-reference.
