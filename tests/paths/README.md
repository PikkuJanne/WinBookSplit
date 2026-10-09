# Current Windows path and filename acceptance

Invoke the fresh supported developer Python from an unrelated directory:

```powershell
& $DevPython -I -B "$Repo\tests\run_tests.py" --layer paths `
  --shell-path $WindowsPowerShell --shell-path $PowerShell7 `
  --report "$ExternalRoot\paths.json"
```

Both actual supported hosts are required. `--layer full` includes this as its
tenth stage. Original/extraction harnesses, guards, fixtures and oracles remain
unchanged; this layer checks AC-036 through AC-039 against current production.

Three actual PS5.1/PS7/BAT successes use a mixed-case `.PdF` input containing
spaces, brackets, an apostrophe, ampersand, parentheses, percent/exclamation
variable-like text and Unicode. A different wildcard-matching generated PDF
has different content and size. Exact source binding, displayed metadata,
manifest filenames, reopened page identities/content and original hashes must
match the literal input. PS uses relative input arguments and explicit external
owned output bases. BAT uses its actual default Documents base, only its emitted
authenticated final/log children, and the existing held-object cleanup helper;
private Documents entries are never enumerated or processed.
The exact owned source is marked read-only for each actual success probe; the
native attribute is observed and restored before temporary cleanup.

The BAT stimulus is an authored fixed command file invoked with `cmd /d /v:off`.
Its last line directly invokes shipped BAT with an environment-supplied quoted
path. It does not use `CALL`, which would introduce a second expansion absent
from the intended stimulus. Controlled Python PATH, stdin and PS UTF-8 wrapper
encoding are recorded; these are launcher tests, not human Explorer acceptance.

Eight actual PS rejection probes cover `.pdf` directories, non-filesystem
provider items, corrupt real files and held-unreadable sources, on each host.
The file-only/provider/readability preflight must reject before engine execution;
the parser must reject corrupt bytes. Native failures must create no successful
final files. Both hosts parse the production PS entrypoint and new Paths helper
and record unchanged stored execution policies.

Nine independent filename expectations cover forbidden/control characters,
trailing dots/spaces, a reserved title, punctuation-only fallback, Unicode,
whitespace, case aliases and astral-character truncation. Each mode writes 120
one-page sections with `001` through `120` numbering, complete identities and
chronological lexical order. Title text never authorizes a path or deletion.

A destination-bound long-base preview executes unchanged with shortened Unicode
names. Stage/final paths are independently measured in UTF-16 units: files at
most 259, directories at most 247, components at most 255. Five failures cover
an impossible budget, impossible legacy unbound names, a changed bound base,
an actual held-exclusive Windows base and explicit ENOSPC injection after one
real slice. Every failure has an actionable category, zero published output and
owned cleanup. Held-lock controls are sharing denials; no ACL, elevation,
registry, global execution-policy or long-path settings are changed. The full
target case is explicitly simulated, not a filled physical disk claim.
Direct API writers install prepare/planner and source-specific PdfReader traps
during execution, allowing the required staged-output reopens. Actual launcher
records compare the child result with an independent destination-bound preview;
their no-replanning property is supported by the direct API/shared-plan controls,
not by a claim that the child process was patched.

Reports require exact case sets, actual requested hosts, full writer/manifest
and page-content evidence, measured path budgets, preserved source/input/neighbor
bytes, import safety, historical guards and removal of owned temporary data.
Missing or malformed promised evidence fails the central runner even if its
child exits zero. Arbitrary long paths, UNC, conversion, Explorer, PDF fidelity
and release-package acceptance remain separate gates.
