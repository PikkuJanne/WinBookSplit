# Status - v1.0.1

V101-01 through V101-12 retain their recorded DONE scopes. V101-13 remains
IN_PROGRESS. Frozen release R and all33 shipped raw bytes/modes/blobs remain
unchanged. The delivered COMMITTED kit is still sourced from K8611; it needs no
rebuild for this nonshipped PS7 supplement recipe correction.

The actual final R ZIP Windows11 kit passed its19 cases under Python3.13.16 and
3.14.8. The separate PS7 supplement has not passed: attempt02 failed before the
adapter because lazy StrictJson initialization was missing; attempt03 reached
Auto2 with invalid `-Mode 2` and missing BookmarkLevel; the application correctly
rejected the mode first with exit2 and message 'Mode must be Auto or Manual.'.
Both failures and safe owned-process receipts are retained. The corrected
recipe supplies `-Mode Auto -BookmarkLevel 2`; the meaningful regression and fixed15 unit
roster passed on both actual Python branches, zero failures/skips. These checks
bind K plus exactly three uncommitted tooling/test/README corrections, not a
clean current-commit or application acceptance pass. PowerShell static checks
were NOT_RERUN for the units-only invocation.

Historical matching K CI38172686734 passed660 Python tests,192 Pester cases on
each shell,28 syntax and27 analyzer rows per shell, and149 runtime units plus8
PDF cases per minor on Windows Server. That pass certifies the old K source;
the next correction commit requires its own complete CI and fresh PS7 run.
See [correction evidence](evidence/V101-13-ps7-recipe-correction.md).

Native Windows10 final ZIP acceptance and public publication remain NOT_RUN.
No human sessions have been requested. Existing v1.0.0 assets/tag are unchanged.
