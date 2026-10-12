# v1.0.2 status

V102-07 is complete and V102-08 source integration is in progress; public v1.0.2 is not yet published.

All four required hosted CI jobs passed on committed C1 `420ae4f9b57acb4034d7336c922b32e32a08f90a`, including actual 3.15.0 beside serviced 3.13.16 and 3.14.8. Both actual private native-kit routes and the four-case PS7 supplement passed independent proof review. Their transitional VERSION 1.0.1 package is distinct from the old published release and supplies no final v1.0.2 or Windows 10 acceptance.

Full local regression finished FAIL with 19 of 23 declared stages passing and unchanged 207-path C1 bytes. Its runtime layer lacked the required existing excluded-interpreter fixture. The separate fixture-corrected runtime run reached a real isolation guard failure: an owned LOCALAPPDATA directory acquired only three empty Microsoft/Windows/Caches directories; its parent environment was unchanged. Outcomes, launcher and document-policy fixtures also refused ambiguous whole-source parameter replacements after a task6 helper added another PythonPath declaration; the outer script declaration was unchanged. Original work and all failures are retained.

Nine focused test-only corrections are implemented and independently reviewed: reuse the conversion fixture's exact authored empty cache layout and held identities, and target only the unique top-level script parameter block in owned copies with strict declaration/count/type checks. Meaningful positive/adversarial controls passed. All four affected native layers passed against the same unchanged 207-path source digest `d0a004cfb9eeb097251947e740ed8d762205ffaecc83168737984f58e2a83809`: runtime 33, outcomes 51, launcher 70 and document-policy 75 cases, with source/input preservation and safe owned cleanup. Source holds are released. Complete receipt review, commit/push C2 and verify fresh hosted CI before marking task 8 complete. Production, package and kit helper bytes are unchanged. Preserve the original full FAIL and distinct corrected results.

Current task: V102-08. Checkpoint branch: `codex/v1.0.2-setup-python315`. Source/evidence references are in TASKS.json. Each commit is followed by normal push and fresh clean/live equality verification outside its own evidence commit.

Required native Windows 10/11 final-ZIP Python 3.15 gates remain NOT_RUN until actual execution. No historical result substitutes for these gates.
