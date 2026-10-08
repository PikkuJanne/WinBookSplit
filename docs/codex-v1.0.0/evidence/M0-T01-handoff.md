# M0-T01: live reconciliation and safe handoff import

Task: M0-T01, acceptance IDs AC-001, AC-002 and AC-003.
Observed: 8 October 2026, Europe/Berlin; exact UTC query times are in `M0-T01-live-audit.json`.
Executor: Codex for authenticated repository owner PikkuJanne. Independent subagent reviewed the archive and importer read-only.
Base commit: `6edbed7c1a0c94968999882c5a46d90492d3c327`.
Base tree: `2443dfef5e6e2f064369f68b6f7b17d061e4b871`.
Application file-set SHA-256: `5a2161df50a442aa6bcd41b4fb43749ff2958dfe09bf42b76b7428b7197ecca2`; recipe and individual byte/hash/blob identities are in the JSON audit.
Test source: manifest-verified extracted bundle, before the implementation commit. No clean-commit application test is claimed.

## Changes and preservation

The supplied `D:\projects\WinBookSplit-main` directory contained the seven original application files and no `.git`, AGENTS, handoff guidance, or nested checkout. Consequently its initial branch, HEAD, origin URLs, Git dirt and historical local Git state were unavailable, rather than clean or synchronized. No applicable AGENTS existed at the root or checked ancestors.

Fresh authenticated GitHub API queries and `git ls-remote` found public `PikkuJanne/WinBookSplit`, default/sole branch `main` at the base commit above, empty tags/releases/open PRs, no rulesets, and an explicit 404 `Branch not protected`. This expected absence is recorded separately from successful queries. Account PikkuJanne has pull/push permissions. No repository settings were changed.

All seven local raw Git blob hashes matched the entire freshly fetched live tree, which also matched BASELINE.json. To make the existing directory usable without touching its files, an external `git clone --no-checkout --single-branch --branch main` fetched live history; `git read-tree HEAD` populated only that external index; its `.git` was copied into the previously absent local `.git`. Local branch/upstream became `main`/`origin/main`, with a clean status. Every original SHA-256 was checked again. This was metadata creation based on live GitHub, not a reset to the bundle snapshot. Missing prior local Git metadata cannot be reconstructed or audited beyond the fetched public history.

Both origin fetch and push URLs are `https://github.com/PikkuJanne/WinBookSplit.git`. The feature branch is `codex/winbooksplit-v1-m0`; it initially has no upstream until its first push. Local commit identity uses the existing public commit author's name and GitHub noreply address. No global Git configuration changed.

The ZIP was extracted under a dedicated external temporary directory after CRC, path, duplicate and symlink-entry checks. Its SHA-256 is `523e5eca12d1176cf301a4f8d1303632e59fafdf088b2f061d1265f579c6c739`. Its 78 manifest rows and 69-target install map were verified. An initial importer preview safely refused the non-Git directory. After metadata reconciliation, the external preview listed 69 CREATE actions, zero conflicts; apply created exactly those 69 files. All imported bytes then matched the verified payload. There was no pre-existing differing guidance/status to merge. The application README, launcher, PowerShell, license and artwork remain byte-identical; no runtime logic changed.

## Actual checks

| Executed command/check | Actual result | Scope |
|---|---|---|
| Git inspection in the initially supplied directory | `not a git repository` | Honest initial limitation; no Git-state success claimed |
| `gh auth status`, `gh api user`, repository/branches/tags/releases/pulls/rulesets/protection queries | Authenticated PikkuJanne; live state above; protection returned expected 404 | Read-only live inspection; GitHub query and authentication command lists in JSON |
| `git ls-remote --symref ... HEAD refs/heads/* refs/tags/*` | Main SHA matched API and fetched clone | Fresh remote observation, not a cached ref |
| `python validate_bundle.py --bundle .` outside repository | Exit 0, VALID: 78 files, 36 tasks, 107 cases, 30 oracles | Delivered integrity and structural consistency only |
| `python -B -m unittest discover -s bundle_tests -v` outside repository | Exit 0, 57 tests in 38.429 seconds: 56 passed, 1 skipped | Importer, Git identity, plan and sync helper regressions |
| Initial `python import_bundle.py --repo <existing directory>` | Safe refusal; no copy because Git metadata was absent | Expected failed preflight, not successful import |
| External `python import_bundle.py --repo <existing directory>` after reconciliation | Exit 0; PREVIEW ONLY, 69 CREATE and zero CONFLICT | Preview before writes |
| Same command with `--apply` | Exit 0; exactly 69 handoff files created | Allowed guidance paths only |
| Python comparison of 69 imported files and seven original SHA-256 values | All matched expected bytes | App README/runtime preserved |
| `python tools/codex-handoff/validate_plan.py --plan-root docs/codex-v1.0.0` | Exit 0, VALID | Imported plan structural consistency |

Environment: reported Windows `10.0.26300`, host PowerShell 7.6.5 Core, Python 3.14.7, Git 2.56.0.windows.1, gh 2.97.0. This reported Windows version does not certify application OS support. Windows PowerShell executable was found but no application checks were executed with it. pypdf is absent from the current Python; Calibre was not found on PATH or at the application's three standard discovery paths. No dependency was installed during M0-T01.

Skip: `SafePaths.test_symlink_escape_refused` could not create a symlink on this host. No elevation or permission weakening was attempted. Bundle integrity does not prove publisher identity. No application execution, bug reproduction, Explorer/drag-and-drop, Calibre conversion, broad security review, or release-package test was run. Those later gates remain open, and source-review findings in BASELINE_AUDIT.md remain reproduction obligations.

## Acceptance and checkpoint

- AC-001: live identity/state reconciled with explicit missing-metadata limitation; original files retained.
- AC-002: verified external extraction, safe failed preflight, reviewed successful preview, create-only import and byte-preservation checks completed.
- AC-003: passed for implementation commit C `ac09c024589158e1402ae63263818466d3e77d5c`: normal push and exit-0 sync checker confirmed clean local HEAD equals fresh live feature branch at `2026-10-08T16:43:50.035807+00:00` (18:43:50 Europe/Berlin).

Pre-commit independent read-only review found no blocking defect and independently confirmed original file hashes/blobs, the file-set digest, imported helper/AGENTS bytes and the scoped guidance changes. The imported plan validator passed during that review.

M0-T01 is `done` with verified implementation C `ac09c024589158e1402ae63263818466d3e77d5c`; its tree is `2ccc56e161053e2ded60df823929b179f6536992`. On clean C, the imported plan validator passed again and all seven original SHA-256 values remained unchanged. Staged diff checks passed; helpers remain byte-identical to the tested bundle, so their suite was not repeated. Application tests remain NOT RUN.

The actual normal push used the authenticated gh credential helper for this invocation only: `git -c credential.helper= -c "credential.helper=!gh auth git-credential" push -u origin HEAD`. Then `python tools/codex-handoff/check_sync.py --repo .` exited 0 and produced the historical C receipt embedded in the JSON audit. The feature upstream is now `origin/codex/winbooksplit-v1-m0`.

[PR #1](https://github.com/PikkuJanne/WinBookSplit/pull/1) is open in draft against main to carry the M0 milestone. Live PR inspection showed head C, MERGEABLE, empty statusCheckRollup and no submitted GitHub reviews. The Actions query for C returned zero runs; no CI or GitHub approval is claimed. The independent local review is recorded above. M0-T04 owns cumulative review/merge; this thread does not merge incomplete milestone work.

This evidence/status update is checkpoint E. It references C's already observed results; E will be normally pushed and checked separately. E's own final SHA/receipt belongs in the final thread/PR or a later record, not inside E. If that push/check fails, repair the incomplete checkpoint before M0-T02 despite the canonical done status.

## Next thread

Complete M0-T02 only after a fresh live check confirms the final checkpoint. Read its brief, BASELINE_AUDIT.md, PLAN_ORACLES.json and TESTING.md. Generate original page-marked fixtures and reproduce original defects before editing runtime code. pypdf setup remains to be resolved in an isolated local environment; the current Python version is not yet a supported application version. Preserve the unchanged baseline application files for that characterization. No private documents were used or uploaded.
