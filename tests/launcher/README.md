# Native launcher and menu acceptance

Run from an unrelated directory with a fresh pinned developer Python:

```powershell
& <dev-python> -I -B <repo>/tests/run_tests.py --layer launcher --tool-root <shell-tools> --shell-path <PS51> --shell-path <PS7> --calibre-path <Calibre915> --report <external-new-report.json>
```

Both actual supported hosts and pinned real Calibre are required. The 70 native
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

This is redirected-input native acceptance, not Explorer interaction or human
testing. M3-T03 also requires actual Explorer drag/drop, double-click selection
and cancellation, and a two-file drop. Those separate UI procedures use exact
committed app bytes and authored books with explicit workstation dependency setup.
They do not certify an extracted release package, clean OS, PDF rendered fidelity
or CI. The historical M3-T01 human tests remain closed and retain their own source.
