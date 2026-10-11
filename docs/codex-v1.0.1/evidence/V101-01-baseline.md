# V101-01 baseline and authorized scope

PASS in recorded inventory, import, routing and synthetic reproduction scope.
Clean starting main be1fa1e271aa2bd031d907e9dbc9e7bfc32bbb62 was reached by ordinary
fast-forward from a5562c4; no reset. Origin is PikkuJanne/WinBookSplit. Current
artwork equals that commit byte-for-byte. Fresh release API and remote tag query
retain v1.0.0 source/tag and all three asset identities; no release write occurred.
Only root AGENTS and new v1.0.1 handoff paths are changed.

Commands: git status/remote/branch/worktree/log/ls-remote; gh release list and
gh api releases/tags/v1.0.0; isolated Python import_bundle.py --repo <checkout>
(preview), then --apply (47 exclusive creates); validate_bundle.py; imported
validate_plan.py. The latter initially failed on a sibling import under -I;
explicit trusted file-directory imports now pass. This is helper validation,
not application acceptance.

Original resolver/probe/supervisor controlled reproduction ran under native
PS5.1.26100.9444 and bootstrap Python3.12.14, with mocked failed -0p discovery
and explicitly synthetic version3.14.2. Sixteen assertions pass: rejected
launcher listing followed by PATH candidate, exact3.14.8 rejection,
runtime_not_found, zero pypdf imports and safely completed candidate process.
No authentic Python3.14.2, pypdf installation state or Windows10 pass is claimed.
External report SHA256 is in the JSON; raw reports/scripts and three failed
preparation attempts remain outside Git. MSIX path virtualization was resolved
by using the actual canonical path, without weakening trust guards.

Actual native inventory is Windows11Pro x64 build26300/26H2/UBR9457. Stale
registry Windows10ProductName is not Windows10 evidence. Read-only Get-VM,
Hyper-V module/service/CIM and standard VBox/VMware CLI/installed-program
checks found no usable authorized VM; Hyper-V namespace returned0x8004100e.
Available tools expose no authorized remote execution route. No owner test
requested: implement/verify the one-command kit first, with Windows10 native
acceptance pending. Existing CPython3.14.7/Calibre9.15, bootstrap3.12.14,
PS5.1 and bundledPS7.6.5 are inventory observations, not renewed acceptance.

Next task: V101-02 shared policy/security decision. This document identifies
the earlier source it observed; its own commit/push/sync is verified afterward.
