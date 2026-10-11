# V101-10 — regressions and security

Task 10 is DONE in the separated source scopes below. Current implementation
**16b2180f917485c584a5511ac7b84d92d6482ba3**, tree
**4d65f34010d7faba6ac03d78a2bea436e782a200**, passed fresh native conversion and
complete [hosted CI](https://github.com/PikkuJanne/WinBookSplit/actions/runs/38165391191).
The [regression receipt](V101-10-regressions.json), [CI receipt](V101-10-ci.json)
and [primary security review](V101-10-security-review.md) preserve commands,
versions, exact source/artifact hashes, rosters and earlier failures.

The native Windows 11 conversion layer passed **19 cases** in 182.030 seconds:
eight real EPUB/AZW3 retention cases, two direct API cases, eight disclosed
invalid-converter controls and one actual BAT missing-converter case. Actual
PowerShell 5.1.26100.9444 and 7.6.6 used reviewed Calibre 9.15.0. The BAT returned
authenticated converter_not_found/3 before engine execution or publication.
The owned job was assigned before resume; the tree and streams stopped without
truncation. Sources, inputs, neighbors and settings remained unchanged, and
owned work was removed. No Explorer gesture or Windows 10 run is claimed.

Current hosted Windows Server 2025 CI passed **640 Python tests**, **192 Pester
tests per shell**, and **149 runtime units plus eight PDF cases per Python
minor**, with zero skips. All three artifact archive hashes match the live API;
all 12 reports pass the existing gates. Every recorded tested path matches its
exact Git blob. Native worktree hashes and Git newline normalization are kept
separate. Server CI does not supply workstation acceptance.

The earlier e21 native source passed eight fidelity cases, 75 document-policy
cases, ten separate genuine ebook cases and all four fault stages. Its aggregate
ebooks run remains **FAIL**: an obsolete BAT assertion failed before writing
the conversion child receipt. Comparing the 200 native tested paths finds ten
changes and 190 identical paths. Planner, writer, conversion/Windows ownership
helpers, supervision, Main and BAT retain identical raw hashes. This composite
audit connects the unchanged regressions with fresh conversion, policy and
shutdown checks; it is not a full native rerun of a later commit.

Policy 2 accepts regular stable x64 CPython **3.13.16+ and 3.14.8+** numerically.
The floors address the reviewed OpenSSL normal TLS handshake issue; later stable
patches remain accepted. Four native shell/runtime combinations each passed
37 Pester, 45 compatibility and eight receipt checks. Converter shutdown controls
passed 29 Runtime tests per shell, including ten synthetic controls; runner
discovery passed 51 units per minor. Exact-byte independent reviews cleared the
changes. pypdf 6.20.0 and optional Calibre 9.15.0 remain separately controlled.

Fresh primary sources support the floors, official Windows binary exclusions
for Python 3.11/3.12, unchanged artifact hashes and qualified converter route.
No resource/security sandbox or universal security claim is made. The final
v1.0.1 ZIP, native Windows 10 acceptance and publication remain NOT_RUN.
