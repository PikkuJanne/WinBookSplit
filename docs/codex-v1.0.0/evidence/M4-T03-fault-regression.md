# M4-T03 — regression and fault-injection evidence

Task completed on 2026-10-10 through **AC-073/074/075**, with automated local Windows evidence and independent source/raw/public review. Human M3-T01 remains absolutely closed. Exact commands, runtime observations, source maps and immutable raw seals are in [the machine record](M4-T03-fault-regression.json).

## Scope and source

Only tests and their documentation changed. Shipped BAT, PowerShell, engine and dependency inputs are byte-identical across all **15 application paths** to user-merged PR23 main `adfb8bd29353d43c6a124ce2bf5bdafea542f31c`. No framework or application replacement was introduced.

Clean implementation C **d12beb8f4b20a9f62f76c109a2f948629fb7ec81**, tree **ab9a9b06740a08875c126f97a9d8488493bef171**, has **146** raw mapped files equal to Git, digest **805d3373e4dbc525b9feeb8c2b32d6cffce33f2d9bcc74807fcb7dc37d11fec1**. It was normally pushed and independently live **SYNCED at 2026-10-10T12:08:50.186784+00:00**. Evidence E references those already observed C facts; its own push/live receipt follows externally after its actual commit.

## Actual verification

| Command projection / acceptance | Actual source | Result | Sealed external record |
|---|---|---|---|
| Pinned Python `-I -B tests/run_tests.py --layer faults`, both actual hosts | Clean C, exact map above | Native 0, **4/4 stages** | `faults01.json` + `faults-command01.json` |
| Pinned Python `-I -B tests/run_tests.py --layer full`, both hosts and explicit Pester/Calibre/Poppler/PDFium tools | Same clean C | Native 0, **21/21 stages**, **422 Python methods** | `full01.json` + `full-command01.json` |
| Actual PS5.1 / PS7 syntax, scoped static checks and Pester | Same C | **112/112 per host**, no skips/failures | Full shell receipts |
| `validate_plan.py` with Python `-B`; normal scoped stage/commit/push; `check_sync.py --repo .` | Actual C | Structural VALID and clean/live SYNCED | External command receipts |

Exact actual argv, timings and raw streams are retained outside Git; readable command projections replace private workstation paths. Versions are freshly observed: Windows 11 Pro x64 26H2 build 26300.9457; PS5.1.26100.9444 / 7.6.5; Python 3.14.8 / pypdf 6.19.0; ReportLab 5.0.1 / Pillow 12.3.0 / charset-normalizer 3.5.2; Pester 6.2.0 / PSScriptAnalyzer 1.25.0; real Calibre 9.15.0; Poppler 26.07.0 and separate Python 3.12.14 / pypdfium2 5.13.0 / PDFium 153.0.7999.0. Legacy static findings retain their declared informational scope; no blanket application static-clean claim follows from shell success.

AC-073: five actual engine worker roles produce two prior publications, two simultaneous real held stages and a final repeat. One worker writes and fsyncs a registered partial second PDF, then returns native 6. While the other remains held, identity/hash/ledger comparisons prove its stage, source, neighbors and prior runs unchanged. The successful worker subsequently publishes all 10 pages in 3 chapters; all four complete publications reopen their physical pages. Exact owner/file identities authorize cleanup. The venv redirector and actual interpreter are distinguished by boot acknowledgement, retained worker handle and exact owned-job membership; native status equality, empty job and both stream EOF are required before cleanup.

AC-074: seed 20261010074 constructs independent per-page ownership before outline arrangement. **720 accepted plans** (240 Manual/240 Level1/240 Level2), **160 invalid manual cases** and **96 malformed outline rejections** pass the strict oracle. Twelve genuine callable writer samples reopen **86 chapters / 168 physical pages** and preserve source/neighbor/prior identities. Callable and synthetic results are distinct from PS/BAT-native results.

AC-075: every same-source path/process/outcome stage passes under both actual hosts, including literal paths, flood/fast-exit streams, nonzero exits, cancellation, timeouts and stopped owned descendants. Four new declared adapter controls replace/delete only an authored processing copy after capture, then run the saved unchanged planner/writer through PS5.1/PS7. Original markers and captured hash survive; original fixture, neighbor and prior publication remain unchanged. The combined route contains **101 distinct controls** (26 supervisors + 75 apps) per execution; targeted and full repeat those controls and are not counted as 202 unique cases. Full also reruns existing cleanup/junction, stale-source, conversion and fidelity controls.

## Defects, retained failures and limits

Only harness defects were exposed. The first isolation focus failed its launcher/worker PID equality assertion; its raw receipt/workspace remain retained, and original worker shutdown cause/status are **UNKNOWN**. The corrected job-bound focus passed; subsequent guard-only recovery changes are verified on C. The first combined pre-C stage failed its receipt assertion despite actual app 0/two correct original-page outputs: `engine.main()` returned `None`, which Python's `SystemExit(None)` maps to 0, and the original adapter lacked the venv parent link. It remains FAILED/native 1 with retained workspace. A fresh source-bound pre-C stage and C routes prove the corrected explicit API/SystemExit/native distinction and actual parent linkage.

Unexpected driver/validator errors now retain partial sections, current/completed case records and owning ancestors. Synthetic controls cover assertion/key errors, interrupts 130, explicit exits, missing/malformed reports and unexpected cleanup members. External capture/probe failures remain explicitly labeled in the raw seals; they are not native acceptance. M4-T01/T02 failed full reports and unknown historical process causes remain immutable and are not diagnosed by this new passing full run.

During the earlier pre-C isolation focus, the agent rendered 17 authored isolation PNGs, compared seven output pages with their source pixels, and inspected four boundary images; those source-bound receipts remain separate from C's fresh native isolation and full fidelity results. This is automated rendering/agent inspection, not human or GUI testing. Finite authored write faults and seeded samples do not fill a physical disk, test power loss/crash recovery, certify arbitrary PDFs or provide a same-account security sandbox. No private files were processed or uploaded. No CI, new human/Explorer/viewer, clean OS/Windows10/ARM/UNC, package/signing/tag/assets/download or public v1.0.0 completion is claimed.

## Git, review and next task

Branch `codex/winbooksplit-v1-m4`, [scoped draft PR24](https://github.com/PikkuJanne/WinBookSplit/pull/24). Independent implementation, raw acceptance and public privacy reviews are sealed in the machine record. M4-T05 owns cumulative review, actual Windows CI and integration.

Next: **M4-T04 — Test real EPUB and AZW3 conversion**, AC-076/077/078. Read the task/specs and reconcile live state before changing anything. Stop at that boundary. Preserve all branches and historical receipts; no force/reset/stash/broad cleanup/settings/interim release.
