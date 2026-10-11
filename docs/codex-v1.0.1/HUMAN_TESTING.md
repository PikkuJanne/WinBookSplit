# Minimal-human acceptance budget

## Default: no manual cases
Codex runs policy/launcher/setup tests, the existing regression suite, runtime-matrix jobs and genuine Windows 10/11 final-package verification automatically. Use already available and authorized local/remote machines or VMs. An existing machine plus a new isolated application extraction is sufficient; a clean OS install is **not** required.

Inspect native Windows 10 access in V101-01. The reported machine's existence is not permission to enable remote access, install a hypervisor, change its configuration or sign it up as a public self-hosted runner. Use an existing authorized capability or prepare the kit below. Do not delay all independent work while a native machine is unavailable.

## At most one planned owner session
Only when Codex cannot run native Win10 acceptance itself, V101-08 prepares a self-contained verification kit. Validate the kit on the accessible Windows host first. After the exact v1.0.1 package is frozen, ask once:

> Run this prepared verification command on your Windows 10 machine and return the generated redacted report. No source editing or manual dependency matching is needed.

The request contains the exact local command, a short explanation of approved per-user setup/downloads, expected output location and no private-document requirement. Prefer approved local runtimes/wheelhouse; if prerequisites are absent, the kit may offer an explicit approved per-user extraction of an official runtime/converter, not an invisible global install. It must identify what it downloads, verify published artifacts and ask no repeated questions during checks. No elevation, PATH/global-package changes or firewall changes.

The kit runs the whole native smoke itself, including synthetic PDF/EPUB/AZW3 fixtures, and exports **one small redacted JSON report plus its needed sanitized attachments in one result bundle**. No screenshots, Excel forms, copying dozens of commands or manual chapter counting. A user-run automated test is recorded as NATIVE_AUTOMATED with execution mode USER_LAUNCHED, not an independent audit or mere owner attestation.

One actual Explorer drag onto the usual shortcut can be included in the same session only if launcher/shortcut integration changed and suitable UI automation is unavailable. No repeated viewer/render checks. If the BAT/shortcut behavior is unchanged, preserve that scope and test actual BAT argument forwarding automatically; document that live Explorer interaction was not repeated. `CreateProcess`/ShellExecute/shortcut resolution tests are not falsely described as a physical drag-and-drop event.

## Failure and retest policy
A substantive defect gets a targeted fix and rerun of affected tests. Changing the final ZIP after native acceptance invalidates the affected final-package evidence and requires the appropriate automated rerun. Do not hide the need for a justified additional owner launch; explain why and consolidate it. The objective is one planned session, not a waiver that permits shipping a known failure.

Missing native Windows 10 execution stays BLOCKED, not SKIPPED-AS-PASS. Complete all independent code/docs/CI work and leave one concrete ready-to-run request. User-requested minimal human testing does not authorize fabricating evidence or silently removing the core Windows 10 goal. Do not invent a new owner approval gate when all genuine required checks have passed.
