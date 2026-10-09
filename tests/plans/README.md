# Shared plan regression

Run the trusted absolute `tests/run_tests.py` with the fresh supported developer
Python using `-I -B --layer plan --report <new external absolute path>`. This is
the seventh full-stage child and receives no shell arguments.

M1-T05 checks AC-027..029 through nine structural aggregates: gaps, overlaps,
empty sections, negative/reversed/overflow bounds, zero pages, noninteger and
incomplete entry contracts, and accepted whole-document plans. Additional
contract cases reject inconsistent declared ranges, nonconsecutive sequences,
unsafe/case-duplicate filenames and unsupported modes. Coverage metadata is
derived from the validated partition. A deterministic seed of `20261009`
produces 300 independent physical partitions, 100 for each mode.

The actual manual, Level 1 and Level 2 prepared previews are immutable and write
no chapter files. Execution uses that same plan and captured reader. Each real
output is reopened; numbered filenames, order, ranges and all synthetic physical
page IDs must equal the preview, cover every input page once, and remain nonempty.
The immutable execution result must preserve every preview entry and page count.
Patches fail if execution attempts to reopen the source or rerun planning.
Mutation attempts cover nested entries, normalized inputs, source identity and
the frozen prepared job. Caller-owned nested metadata is detached at validation;
an unbound mapping cannot be executed and leaves an empty output directory.

Source binding tests deliberately modify copies owned by this run. One replaces
the file bytes in place; another renames the captured source and places a
different PDF at its former path, then removes that replacement before a second
execution. All executions must preserve the original ten-page reader snapshot,
source digest, planned physical page IDs and extracted synthetic page-content
hashes. Replacement and moved source bytes stay unchanged. These intentional
copy changes are separate from the immutable
fixture/input preservation checks.

The report contains each case, actual preview/output records, captured and
replacement source hashes, import observation, runtime versions and preservation/
cleanup results. The central runner rejects absent, duplicated, failed or
incomplete promised evidence. Immutable original fixture/oracle bytes and the
historical baseline refusal are checked through the existing guarded helpers.
All PDFs and deliberate replacements live in a unique owned temporary directory;
neighbors are synthetic sentinels and cleanup removes only this run's files.

This route verifies the callable preview model and writer. Interactive preview
controls, Explorer, further launcher paths, Calibre, rendered fidelity, complete
output transactions and release-package checks remain separate later gates.
