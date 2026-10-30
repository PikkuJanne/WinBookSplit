# Checkpoint 5 — documentation, test distribution and closeout

30 October 2026. Documentation and the new exact test distribution are complete.
Functional acceptance remains incomplete: ordinary supported-Store first-run
consent/setup/resume, repeat/network observation, Windows 10 and the private book
remain NOT_RUN. No release, publication, tag or version change is authorized.

## Authority and committed source

Read AGENTS.md, STATUS_NEXT.md, CHECKPOINT_1 through CHECKPOINT_4 and their JSON
maps, the historical v1.0.3 closure/routing/testing references, and the Downloads
repair pack's CODEX_TASK.md checkpoint 5 and ACCEPTANCE.md. The instruction ZIP
still hashes to `f33ab17eab774f4a65cf6e008949012e810cb0c10ce4c9bb9796f0d6524dc70d`.

Started clean on `codex/functional-repair`, local HEAD and freshly queried live
branch both `d3ee23a41fba26eca07e1dc61b564e93586519c1`. Origin remains
`https://github.com/PikkuJanne/WinBookSplit.git`; there is one worktree.
The fourteen intended guidance files were committed before packaging as
`c3d7078a7e54b0020423b1c93d6aca120abd0eda`, tree
`c5fd2b703d174cf8e7b9d98f69abe405624bb450`. The later checkpoint evidence/status
commit is outside that immutable package binding; its synchronization is checked
after ordinary commit/push and reported in the handover, without self-reference.

## Guidance and known harness limitation

README, setup, PDF policy/fidelity, ebook, security, support and contribution
guidance now describe one explicit consent/setup-or-repair/fresh-runtime/resume
cycle, backup preservation and no prompt/install for preview, noninteractive
processing or readiness checks. They describe ordinary empty-password access,
specific password/handler/backend/malformed outcomes, the narrowly ignored local
catalog opening view and one warning, unencrypted chapters, unchanged source and
retained converter-output limits. They make no AES interoperability, permission
enforcement or successful ordinary first-run claim.

Active AGENTS and test/package/kit instructions distinguish current repair
authority from completed v1.0.3 routing, waivers and publication. The changelog
has an Unreleased entry. Complete source archives retain the same required
entry points/coordinator; archive type alone does not explain a setup failure.
Historical records under docs/codex-v* and published tags/assets are unchanged.
Independent final documentation review found no material mismatch with production.

The audit found an existing **unmigrated native harness**, now explicitly disclosed
in tests/README.md and tests/document_policy/README.md. Its document-policy
producer/validator still expects empty-password rejection and generic
`unsupported_document` for passworded inputs; its refused-input export also uses
the now-readable empty-password fixture. `--layer full` includes this stage.
Neither route was run or claimed as a current pass here. The current PDF unit/
functional contracts and exact-package checks below do not repair that older
harness. A future coordinated harness update must preserve its safety/export
controls and must not restore obsolete application bans to satisfy expectations.

## Exact unpublished distribution

The unchanged package builder/allowlist read only committed source blobs. Its
four exact committed tool inputs were materialized externally because this
Windows checkout normalizes some line endings; no builder guard or Git setting
was weakened. Build, verify and the durable copied-assets verify all passed.

The existing filename remains `WinBookSplit-v1.0.3.zip`, but these are **unpublished
checkpoint test bytes**, not the public v1.0.3 asset:

- ZIP: 971,784 bytes, SHA256
  `48dfcd7f7c73e9855733e98ec26880cbf645d299b8f70f4de6ae6480ad1dc3ab`.
- Manifest: 16,633 bytes, SHA256
  `ebd29aefe8d204c4b05c752d1e76d4007b757782b14bde007685442a1c526019`.
- SHA256SUMS: 178 bytes, SHA256
  `9911f23421fad5b42ffa7ec1f453033ed63c725aa4f7873f725e35f9faf5767f`.

All three are saved in the new Downloads folder
`WinBookSplit-Functional-Checkpoint5-c3d7078`. No existing installation or asset
was overwritten. Independent audits verify all 34 allowlisted members against
raw Git blobs and separate fresh extractions, Git IDs/SHA256/sizes, ZIP recipe,
manifest, checksums, shipped relative links and the authenticated Git-free frozen
verifier. All **22 runtime members** are byte-identical to checkpoint 4's package;
only nine payload documents changed. No environment, wheel, private book, chapter,
credential, personal log or developer cache enters either the payload or Git.

## Executed checks and limits

Pre-existing developer tools only: regular x64 CPython 3.14.8, pypdf 6.21.3,
Windows 11 build 26300, PowerShell 5.1.26100.9711 and 7.6.6, Pester 6.2.0.
The driver is never injected into ordinary discovery. No runtime, application
environment, global package, PATH, alias or persistent policy was provisioned
or changed. No new Calibre/runtime-matrix/full-suite pass is claimed; their
checkpoint-4 evidence retains its original source and limits.

| Check | Actual result and scope |
| --- | --- |
| Exact build/verify, independent fresh extraction and Git-free verification | PASS as above; 34 members. |
| Existing functional suite against fresh extracted engine | PASS, 11 methods, zero failures/errors/skips, 22.141s. The trusted unchanged harness selects the exact extracted engine/policy/ROOT. Plain, RC4-empty, local GoTo/array, combined input, captured/interactive reader, converted validation and password/active-content controls; Manual/Auto1/parent-aware Auto2 preview/execution. |
| Actual extracted healthy application, repeated | PASS, four NonInteractive explicit-developer-runtime launches, twice each on PS5.1 and PS7.6.6, unrelated CWD and Unicode/bracket/ampersand paths. Combined authored RC4-empty/local-GoTo input; exit 0 / split_complete; three chapters, independently expected ranges and six physical page-content markers in order; one ignored-view warning; unencrypted chapters without OpenAction. Source, neighbor, prior output and earlier chapters preserved. No setup frames, prompt or app .venv. Healthy NO_INSTALL coverage, **not ordinary Store cold start**. |
| Full source-code ZIP and fresh extraction | PASS, all 798 tracked files and 34 required payload files exactly match Git. Source ZIP: 9,145,792 bytes, SHA256 `d7137bb53dcba0f6cfa18a19d137cfd470df93d6a4caeb91162c918806b3ef83`. No environment/private-output/cache files. This is a separate source-archive check, not an accepted release distribution. |
| Extracted source FirstRun components and version | PASS, 38 cases each on PS5.1 / PS7.6.6, zero failures/skips/failed containers, 13.497s / 10.781s. Unchanged coordinator/resolver tests use explicitly synthetic discovery/setup/input seams; no setup is run. Both actual extracted -Version invocations pass. Components establish the same coordinator is present, not native installation/resume. |
| Unchanged native first-run test from extracted source | One legitimate prerequisite skip, zero failures. Actual ordinary BAT returns 3 / runtime_not_found for naturally discovered Python 3.14.7 / python_security_floor (required 3.14.8). Source and copied application bytes preserved, no app .venv or chapters. Desired missing-package/setup/resume remains NOT_RUN; no selected/private runtime or PATH injection. |
| Live preservation check | Read-only before/after tag references and four public release/twelve asset metadata records agree, including the public v1.0.3 ZIP's original hash. No asset was uploaded, edited, downloaded anew for byte verification, or relabelled. |

The initial plain `git archive` exact-byte assertion is retained as FAIL:
unchanged system `core.autocrlf=true` converted 774 text files, with every
difference independently confirmed as CRLF-only. Its archive hash is
`77e93fe32388bb2ae6f6889395195a2cf9df6000710c85bcf86a411b05392418`.
A second full source ZIP used command-local `git -c core.autocrlf=false archive`
and passed exactness. Git configuration and application behavior were unchanged;
this was an archive-recipe correction, not a first-run repair or compatibility
failure. Both original and final archives/extractions remain outside Git.

The single pretested checkpoint-4 native kit remains unchanged in Downloads:
`WinBookSplit-Functional-Native-Checkpoint4.zip`, SHA256
`8068d08b0f1207bee9c0cbbee2b644386847601844b2840d01cbb77f70f7758f`.
It retains its checkpoint-4 source/archive binding; it is not relabelled as a
test of this new documentation-only ZIP. All runtime payload bytes are identical.
Its one START-HERE.bat/report rehearsal remains the previously recorded result,
not another native pass here. Reuse it only on an authorized qualifying machine;
do not precreate its environment, select a runtime or provision a private Python.

## Commands and retained receipts

Raw receipts remain outside Git under
`%TEMP%\WBS-functional-cp5-146bd0421af94d34b96b82e8ea86630e`.
The independent package audit is under
`%TEMP%\WBS-functional-cp5-independent-audit-6a3fe0320ad2436cabe423e7d1680561`.
Source-archive commands/results and retained failed exactness comparison are under
`%TEMP%\WBS-functional-cp5-source-review-4f218afd6bfd47458bf8edb611193edd`;
its source-archive-results.json SHA256 is
`d0d5a24afd3931e797e6f6095a062c87f4164d20196539a165c184dc13b0c4ee`.
Actual ordinary BAT receipt is `%TEMP%\WBS-first-run-5wxhkhmq\evidence.json`,
SHA256 `6e92fa1d1ce2cb89cf83fb7ad7e1e4b08cc59450584c8f749d49f68938f9d977`.
CHECKPOINT_5.json records source/asset identities, commands/results and receipt
hashes without committing generated PDFs or personal logs.

`$Repo` is this checkout, `$DevPython` the pre-existing
`%TEMP%\WBS-623-finish-q6dwc7hk\dev14\Scripts\python.exe`, `$CommittedBuilder`
the exact external source-blobs tool and `$Assets` a fresh external destination:

```powershell
& $DevPython -I -B $CommittedBuilder build --repo $Repo --commit c3d7078a7e54b0020423b1c93d6aca120abd0eda --output $Assets
& $DevPython -I -B $CommittedBuilder verify --repo $Repo --commit c3d7078a7e54b0020423b1c93d6aca120abd0eda --output $Assets
git -c core.autocrlf=false -C $Repo archive --format=zip --prefix=WinBookSplit-source/ --output=$NewSourceZip c3d7078a7e54b0020423b1c93d6aca120abd0eda
& $DevPython -I -B "$ExtractedSource\tests\python\test_first_run.py"
```

The retained `prepare-and-package.py`/`audit-and-healthy.py` drivers bind the
unmodified existing functional tests to the extracted payload, then launch each
actual host with `-NoProfile -ExecutionPolicy Bypass -File WinBookSplit.ps1`,
literal source/output paths, `-PythonPath $DevPython -Mode Manual -StartPages
1,3,5 -NonInteractive -NoPause`. This is process-only policy handling and an
explicit healthy runtime, not organization-policy or ordinary setup acceptance.
The source component driver imports each actual host's built-in Utility/Security
modules and the existing pinned Pester 6.2.0, then invokes only extracted
FirstRun.Tests.ps1 through New-PesterContainer. Its retained exact argv and test
results bind each host and source; no setup, package install or affirmative
ordinary setup consent occurs.

Commit only the final checkpoint evidence/status, push normally, verify clean
local HEAD against freshly queried live codex/functional-repair, hand over and
STOP. No additional checkpoint, merge, release or publication follows automatically.
