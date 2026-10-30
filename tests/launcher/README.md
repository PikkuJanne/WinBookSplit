# Native launcher and menu acceptance

Run from an unrelated directory with a fresh pinned developer Python:

```powershell
& <dev-python> -I -B <repo>/tests/run_tests.py --layer launcher --tool-root <shell-tools> --shell-path <PS51> --shell-path <PS7> --calibre-path <Calibre917> --report <external-new-report.json>
```

Both actual supported hosts and the exact reviewed QA Calibre 9.17.0 binary are
required by this regression layer; production accepts stable 9.17.0 or newer.
The 70 native
controls cover literal prompted PDF/EPUB/AZW3 paths, quoted paths, cancellation,
exact initial/fallback answers, reprompt, closed input, and BAT no/multi-file
behavior. The authored fixtures include genuine Calibre-generated AZW3. Every
published PDF is reopened and compared to its reference physical content, ranges
and manifest. Sources are read-only; neighbor/prior files and policies are checked.

BAT bytes are exact. Native successful BAT controls explicitly change only copied
PowerShell dependency/output/NoPause defaults; rejection controls replace copied
PowerShell with a launch sentinel. These declared copies isolate the process tests.
The strict validator rejects missing, contradictory or relabeled receipts. Failed
or timed-out workspaces are retained when descendant shutdown is unproved.

This is redirected-input native acceptance. The historical M3-T03 specification
included separate Explorer gestures. Active functional-repair authority excludes
all human testing, sign-off and report requests; no such work is pending.
Historical UI receipts retain their original source scope. Redirected controls
alone do not certify an extracted release package, clean OS or rendered fidelity.
Current exact-package evidence is in the functional-repair checkpoint records.
