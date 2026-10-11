# Sources and evidence boundaries

Inspected 11 October 2026. Refresh material version/security/release facts before changing the application and before publishing. Bundle design decisions are proposed requirements, not findings in these sources.

**[R1] Observed main**

https://api.github.com/repos/PikkuJanne/WinBookSplit/branches/main

main be1fa1e…; current artwork must be preserved. Access: GitHub connector.

**[R2] Observed published releases**

https://api.github.com/repos/PikkuJanne/WinBookSplit/releases

v1.0.0 published; fixed source and asset hashes; no v1.0.1 in observed list. Access: GitHub connector.

**[R3] Root working agreement**

https://github.com/PikkuJanne/WinBookSplit/blob/be1fa1e271aa2bd031d907e9dbc9e7bfc32bbb62/AGENTS.md

Old v1.0.0 routing and delegated scope. Access: GitHub connector.

**[R4] Runtime preflight**

https://github.com/PikkuJanne/WinBookSplit/blob/be1fa1e271aa2bd031d907e9dbc9e7bfc32bbb62/engine/WinBookSplit.Runtime.ps1

Exact Python checks at both layers, discovery and converter pin. Access: GitHub connector/current conversation code review.

**[R5] Current README**

https://github.com/PikkuJanne/WinBookSplit/blob/be1fa1e271aa2bd031d907e9dbc9e7bfc32bbb62/README.md

Narrow current support matrix. Access: GitHub connector/current conversation code review.

**[R6] PDF reader guard**

https://github.com/PikkuJanne/WinBookSplit/blob/be1fa1e271aa2bd031d907e9dbc9e7bfc32bbb62/engine/winbooksplit_conversion.py

Private encryption hook and reader keyword dependency. Access: GitHub connector/current conversation code review.

**[R7] Current CI**

https://github.com/PikkuJanne/WinBookSplit/blob/be1fa1e271aa2bd031d907e9dbc9e7bfc32bbb62/.github/workflows/windows-ci.yml

Single Python pin and Windows Server automation. Access: GitHub connector.

**[R8] Existing test harness**

https://github.com/PikkuJanne/WinBookSplit/blob/be1fa1e271aa2bd031d907e9dbc9e7bfc32bbb62/tests/run_tests.py

Reuse regression/evidence machinery. Access: GitHub connector.

**[R9] Existing release tool directory**

https://api.github.com/repos/PikkuJanne/WinBookSplit/contents/tools/release?ref=be1fa1e271aa2bd031d907e9dbc9e7bfc32bbb62

Build_package.py and payload.json discovered. Access: GitHub connector.

**[S1] Python Windows documentation**

https://docs.python.org/3.14/using/windows.html

Legacy -0p remains supported; manager listing exists; avoid automatic runtime installation. Access: web.

**[S2] Python 3.14.8 release**

https://www.python.org/downloads/release/python-3148/

Security maintenance release; examine relevant fixes before setting lower floors. Access: web.

**[S3] pypdf 6.20.0 metadata**

https://pypi.org/project/pypdf/6.20.0/

Python >=3.9 metadata is not proof of application compatibility. Access: web.

**[S4] GitHub hosted runners**

https://docs.github.com/en/actions/reference/runners/github-hosted-runners

Standard x64 Windows CI is not Windows 10 desktop evidence. Access: web.

**[S5] Python branch status**

https://devguide.python.org/versions/

Recheck lifecycle and Windows binary availability at execution. Access: web.

**[S6] Virtual environments**

https://docs.python.org/3/library/venv.html

Environment isolation; environments are not generally relocatable. Access: web.

**[S7] Calibre Windows download**

https://calibre-ebook.com/download_windows

Upstream Windows support is a prerequisite, not application acceptance. Access: web.

**[S8] pip secure installs**

https://pip.pypa.io/en/stable/topics/secure-installs/

Hash-required locked installs need exact permitted artifacts. Access: web.

**[S9] Windows 10 lifecycle**

https://learn.microsoft.com/en-us/lifecycle/products/windows-10-home-and-pro

Compatibility does not extend Microsoft security servicing. Access: web.

**[S10] GitHub CLI release creation**

https://cli.github.com/manual/gh_release_create

Draft creation and explicit tag verification. Access: web.

**[S11] GitHub CLI release editing**

https://cli.github.com/manual/gh_release_edit

Publish draft using explicit flags. Access: web.
