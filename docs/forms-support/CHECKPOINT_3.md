# Checkpoint 3 — Static form output and ebook repair

The owner continued checkpoints 3 through 5 in one thread and added the
reported EPUB failure to v1.0.5. The checkout started at
`87f3e805dbc97779c63b651d22abb076da2b15e4`. The scoped branch matched live
origin, and live GitHub had neither a v1.0.5 tag nor a release.

## Static page content

The shared reader prepares visible saved form values and widget appearances
before chapter writing. It resolves bounded parent-field inheritance, chooses
button and choice states, and regenerates supported variable text if the saved
appearance is absent or stale. Chapter projection paints the selected
appearance with its bounds, transformation and local resources. The fresh
catalog carries neither interactive fields nor source signature validity.
Hidden values and unrelated appearances stay outside chapter graphs.

The copied-source form matrix passed 26 native cases across inbox PowerShell
5.1, PowerShell 7 and BAT. Twenty-three runs produced 70 reopened chapters.
Three fixtures with missing required appearances returned their exact refusal.
Poppler and PDFium verified field regions, including selected and unselected
controls, multiselect and multiline fields, Unicode, offset and transformed
appearance bounds, and visible signature artwork. Report SHA-256
`29f797e37c48e6f71a59096964b58fd8dde19a9499634cc9f8560d77dff487c6`.
A separate actual PowerShell 5.1 split of a ReportLab-authored form had two
pixel-identical Poppler source/output pages, including accented text, checkbox,
radio and combo artwork.

## Calibre discovery and first use

The original report did not include its underlying attempts. An unchanged
published v1.0.4 launcher run on this Windows host reproduced
`converter_not_found` / 3. Its bounded version probe observed the installed
per-user Calibre 9.15.0 and rejected it against the existing 9.17.0 minimum.
The old install was preserved.

Automatic discovery now verifies the publisher and executable identity before
probing PATH, registered and known-folder candidates. The result retains the
observed path, version and rejection reason. An invalid explicit selection
remains authoritative. With consent, the application can prepare a separately
managed Portable copy and resume the original ebook.

The reviewed official Portable installer is 259,196,168 bytes. The downloaded
bytes matched both official hashes and had a valid Authenticode signature from
Kovid Goyal. The application reads its verified embedded data resource and
publishes the resulting complete tree separately from existing installations.
It verifies the tree again before reuse.

| Identity | Digest |
| --- | --- |
| Installer SHA-256 | `bb3b5423e1d64a735322a006edce15c6fb012980c2c5e55c1c45eaeba21a28dc` |
| Installer SHA-512 | `6ed05950375b098b58ea2e71cde927fe88623074f5f703246f5cd36d8c91c79625ba5a184134b3e3ff005843491af3d1ee2a1424d8a3e2e2bee552465d559dfe` |
| Extracted tree SHA-256 | `5de8584ed4aaa7516cbac1aae102ca6f40f7af9a16eb351b100de8fce418a3a4` |

[Official download](https://download.calibre-ebook.com/9.18.0/) ·
[Upstream GitHub release](https://github.com/kovidgoyal/calibre/releases/tag/v9.18.0) ·
[Upstream SHA-512](https://calibre-ebook.com/signatures/calibre-portable-installer-9.18.0.exe.sha512)

The initial managed conversion reproduced a separate Calibre renderer error
with a deep temporary path. Controlled invocations of the exact 9.18.0
executable failed with 166–171-character temporary roots and succeeded with
60-character roots using the same illustrated EPUB. The app now uses short,
owned conversion temporary state and isolates converter settings and caches.
The fresh source offline route prepared 9.18.0 beside the old installation,
then wrote two manual chapters covering all three generated pages and preserved
Unicode text, images, the installer, and the original 9.15.0 tree. Report
SHA-256 `6c65151da6ba1b0738e8b46d7c12ec843a7f464df160e0977b8d593122fdcc4c`.

Final frozen-package acceptance is recorded separately at checkpoint 4.
