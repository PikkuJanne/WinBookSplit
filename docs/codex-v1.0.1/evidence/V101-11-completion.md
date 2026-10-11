# V101-11 — documentation, version and package contracts

Task 11 is DONE at **a1d17c3b48cdc6e6aedc7752954ebf84714cb74d**. The
[complete CI](https://github.com/PikkuJanne/WinBookSplit/actions/runs/38168337363)
passed all three jobs. The [source-bound receipt](V101-11-ci.json) records the
actual rosters, tools, artifact hashes and independently compared source bytes.
This is hosted Windows Server evidence, separate from native final-ZIP testing.

The completion record distinguishes the actual collected report's raw hash from
the linked committed JSON's hash. Review caught Git's CRLF-to-LF transformation;
936 CR bytes differ, with exactly equal parsed values. Both identities remain
recorded, and the earlier checkpoint is preserved in Git history.

The [completion record](V101-11-completion.json) binds the documentation/version
checks and reviews. The 33 package members match the earlier development package
source exactly. Version is 1.0.1; regular x64 CPython 3.13.16+ and 3.14.8+ use
numeric maintenance floors, with pypdf 6.20.0 and optional Calibre 9.15.0
controlled separately. Setup, unsigned status, servicing and tested-versus-
accepted limits are documented. Historical v1.0.0 records remain historical.

The actual 087 failures and qualified focused corrections remain in
[their receipt](V101-11-fixture-correction.json). Review used a separate context,
whose author contributed some earlier code; it is not an independent audit.
Task 12 still must integrate reviewed source, run matching main checks and
freeze R/assets. Native Windows 10/11 final-ZIP acceptance and public publication
remain NOT_RUN. No owner test session, tag or draft has been requested or created.
