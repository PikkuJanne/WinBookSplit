# V102-08 C1 required hosted CI execution

The [actual C1 run](https://github.com/PikkuJanne/WinBookSplit/actions/runs/38286167981) passed all four required jobs on source `420ae4f9b57acb4034d7336c922b32e32a08f90a`. Independent authenticated artifact downloads matched GitHub ZIP hashes and sizes; the maintained strict validator passed, and every report's 207-path source map matched exact C1 Git blobs. Common source digest: `98fadd281eeacf8b0de879c11caf986087d28d57fc1f7d01007a42141cd7c929`.

The application jobs ran actual regular x64 Python 3.13.16, 3.14.8 and 3.15.0 with policy 3 and pypdf 6.20.0. Each runtime contained only pypdf plus pip, ran 150 tests across the required nine modules with zero skips, and passed eight authored PDF/application cases through 15 supervised native probes. Exact selected executable/probe identity, package origin, unchanged before/after inventory and both wrapper runtime selections were checked by the application harness. The fixture driver remained 3.14.8.

The reference developer Python 3.14.8 suite reported `Ran 664 tests in 365.798s` with `OK`. Both actual PowerShell suites passed: 5.1.26100.33438 204/204, 7.6.6 204/204; syntax and scoped static analysis had no findings. All four committed package-prerequisite checks passed without building a package or using worktree inputs.

Actual CI OS was Microsoft Windows Server 2025 Datacenter, build 26100; actual application hosts were PowerShell 5.1.26100.33438 and 7.6.6. This is Server CI evidence. Native Windows 10/11 final ZIP checks and the full local suite have their own receipts. No fallback, missing-report success or absent 3.15 target was accepted.

The companion JSON records actual artifact/report hashes, commands, jobs, runtimes, unit counts and full source map. Private inspection receipt SHA-256: `0c286e0fdfb17d1457365d6eabfbd7c1ae3c5616abf743938747bd506db9ca4a`. V102-08 completion still requires its independent full-regression and native-kit gates.
