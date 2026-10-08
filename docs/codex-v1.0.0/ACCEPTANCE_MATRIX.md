# Acceptance catalog

107 mandatory cases for the declared v1.0.0 scope. A defined rejection outcome is a passing test only when actually executed. This catalog is not test results; real results belong in evidence/. Case details are canonical in ACCEPTANCE_CASES.json.

| ID | Task | Execution layer | Scenario |
|---|---|---|---|
| AC-001 | M0-T01 | repository_review | Correct repository and live state |
| AC-002 | M0-T01 | repository_review | Safe handoff import |
| AC-003 | M0-T01 | live_github | Initial synchronized checkpoint |
| AC-004 | M0-T02 | python_characterization | Manual first-page bug baseline |
| AC-005 | M0-T02 | python_characterization | Bookmark baseline defects |
| AC-006 | M0-T02 | review | Fixture provenance |
| AC-007 | M0-T03 | environment | Honest supported matrix |
| AC-008 | M0-T03 | environment | Runtime/dev dependency separation |
| AC-009 | M0-T04 | test_harness | Runner failure propagation |
| AC-010 | M0-T04 | test_harness | Isolated repeatable harness |
| AC-011 | M1-T01 | python_integration | Mechanical extraction preserves characterization |
| AC-012 | M1-T01 | windows_integration | Engine resolved from script root |
| AC-013 | M1-T02 | python_unit | Explicit first page alone |
| AC-014 | M1-T02 | python_unit | Explicit and implicit first page |
| AC-015 | M1-T02 | python_unit | Sort duplicate and leading-zero inputs |
| AC-016 | M1-T02 | python_unit | Empty input and tokens |
| AC-017 | M1-T02 | python_unit | Unsupported syntax and mixed tokens |
| AC-018 | M1-T02 | python_unit | Bounds and final page |
| AC-019 | M1-T03 | python_integration | Level 1 front matter |
| AC-020 | M1-T03 | python_unit | Duplicate/out-of-order outline starts |
| AC-021 | M1-T03 | python_unit | Invalid/external destination |
| AC-022 | M1-T03 | python_unit | Deep/malformed outlines |
| AC-023 | M1-T04 | python_integration | Parent-aware Level 2 exact example |
| AC-024 | M1-T04 | python_unit | No children in one parent |
| AC-025 | M1-T04 | python_unit | Children equal parent start or outside parent |
| AC-026 | M1-T04 | python_unit | No usable selected level |
| AC-027 | M1-T05 | python_unit | Plan structural invariants |
| AC-028 | M1-T05 | python_integration | Preview/execution page identity parity |
| AC-029 | M1-T05 | python_integration | Source changes after preview |
| AC-030 | M1-T06 | python_integration | No bookmarks versus unreadable/zero-page input |
| AC-031 | M1-T06 | powershell_test | No-Level2/no-outline fallback distinction |
| AC-032 | M2-T01 | windows_integration | Repeat and same-basename runs |
| AC-033 | M2-T01 | python_integration | Write failure mid-run |
| AC-034 | M2-T01 | windows_integration | Concurrent unique output allocation |
| AC-035 | M2-T01 | windows_integration | Ownership and reparse cleanup |
| AC-036 | M2-T02 | windows_integration | Literal Windows input paths |
| AC-037 | M2-T02 | python_unit | Filename sanitation |
| AC-038 | M2-T02 | python_unit | More than 99 chapters |
| AC-039 | M2-T02 | windows_integration | Long/unwritable destinations |
| AC-040 | M2-T03 | windows_integration | Existing neighboring PDF |
| AC-041 | M2-T03 | powershell_test | Converter zero success with invalid output |
| AC-042 | M2-T03 | windows_integration | Intermediate retention |
| AC-043 | M2-T04 | powershell_test | Missing and wrong Python/pypdf |
| AC-044 | M2-T04 | powershell_test | PDF without Calibre |
| AC-045 | M2-T04 | powershell_test | Trusted explicit/nonstandard converter |
| AC-046 | M2-T05 | windows_integration | Dual stream flood |
| AC-047 | M2-T05 | windows_integration | Fast exit and output tail |
| AC-048 | M2-T05 | windows_integration | Unicode encoding |
| AC-049 | M2-T05 | windows_integration | Argument injection boundary |
| AC-050 | M3-T01 | windows_integration | Failure and fallback exit propagation |
| AC-051 | M3-T01 | windows_integration | Cancel only owned processes |
| AC-052 | M3-T01 | windows_integration | Cleanup and finalization failures |
| AC-053 | M3-T02 | powershell_test | Noninteractive complete run |
| AC-054 | M3-T02 | powershell_test | Invalid/contradictory parameters |
| AC-055 | M3-T02 | powershell_test | Version without dependencies |
| AC-056 | M3-T02 | windows_integration | Explicit Preview semantics |
| AC-057 | M3-T03 | manual_windows | Explorer file formats |
| AC-058 | M3-T03 | manual_windows | No-input selection and cancellation |
| AC-059 | M3-T03 | manual_windows | Multiple files and invalid menu input |
| AC-060 | M3-T04 | manual_windows | Readable complete preview |
| AC-061 | M3-T04 | manual_windows | Converted-PDF manual guidance |
| AC-062 | M3-T04 | manual_windows | Truthful final summary |
| AC-063 | M3-T05 | python_integration | Manifest/log correctness |
| AC-064 | M3-T05 | review | Redacted support information |
| AC-065 | M3-T05 | powershell_test | No blanket warning suppression |
| AC-066 | M4-T01 | pdf_integration | Text and image page fidelity |
| AC-067 | M4-T01 | pdf_integration | Geometry and common annotations |
| AC-068 | M4-T01 | pdf_integration | Chapter metadata/bookmark |
| AC-069 | M4-T01 | pdf_integration | Links across and within segments |
| AC-070 | M4-T02 | pdf_integration | Encrypted variants rejected |
| AC-071 | M4-T02 | pdf_integration | Forms/signatures and active features |
| AC-072 | M4-T02 | pdf_integration | Corrupt/pathological input |
| AC-073 | M4-T03 | fault_injection | Run isolation under failure |
| AC-074 | M4-T03 | python_property | Randomized partition invariants |
| AC-075 | M4-T03 | windows_integration | Full cross-host error suite |
| AC-076 | M4-T04 | calibre_integration | Real EPUB conversion |
| AC-077 | M4-T04 | calibre_integration | Real AZW3 conversion |
| AC-078 | M4-T04 | calibre_integration | Offline/failure/profile behavior |
| AC-079 | M4-T05 | live_ci | Real Windows CI results |
| AC-080 | M4-T05 | security_review | CI privilege and release separation |
| AC-081 | M4-T05 | live_github | Milestone merge verification |
| AC-082 | M5-T01 | documentation | README accuracy and working commands |
| AC-083 | M5-T01 | documentation | Version/license/signing statements |
| AC-084 | M5-T02 | package_test | Allowlisted clean package |
| AC-085 | M5-T02 | package_test | Provenance/checksum integrity |
| AC-086 | M5-T02 | package_test | Broken package inputs fail |
| AC-087 | M5-T03 | manual_windows | Clean extracted candidate both hosts |
| AC-088 | M5-T03 | manual_windows | Complete Explorer candidate workflow |
| AC-089 | M5-T03 | pdf_integration | Candidate visual/output checks |
| AC-090 | M5-T04 | security_review | Concrete safety review |
| AC-091 | M5-T04 | dependency_review | Current dependency/license audit |
| AC-092 | M5-T04 | test_suite | Full release regression |
| AC-093 | M5-T05 | review | Approved scope traceability |
| AC-094 | M5-T05 | live_github | Final candidate merged/synced |
| AC-095 | M6-T01 | live_github | Release commit source proof |
| AC-096 | M6-T01 | release_review | All prerequisites real |
| AC-097 | M6-T02 | package_test | Exact final ZIP from R |
| AC-098 | M6-T02 | package_test | Final frozen hashes |
| AC-099 | M6-T03 | manual_windows | Final exact-asset Windows acceptance |
| AC-100 | M6-T03 | calibre_integration | Final real ebook and PDF fidelity |
| AC-101 | M6-T04 | live_github | Exact remote tag |
| AC-102 | M6-T04 | live_github | Draft asset inspection |
| AC-103 | M6-T05 | live_github | Actual public release |
| AC-104 | M6-T05 | public_download | Anonymous downloaded asset integrity |
| AC-105 | M6-T05 | manual_windows | Downloaded-package smoke |
| AC-106 | M6-T06 | live_github | Final synchronized main with fixed tag |
| AC-107 | M6-T06 | release_review | Durable final evidence/report |
