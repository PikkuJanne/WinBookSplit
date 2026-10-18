# New security candidate: one Windows 10 test

The owner authorized the pypdf security repair on 18 October 2026. This candidate
uses pypdf 6.21.1. The accepted T07 report remains valid for the original R/ZIP;
it cannot accept these changed bytes. This is one necessary new-candidate test.
Original OWNER_KIT.md, OWNER_TRANSFER.md and their archives remain historical
contracts. Do not repeat the unchanged original-R test.

## Start

1. Copy **WinBookSplit-Windows10-Test-afd75b4.zip** to Downloads on the same
   Windows 10 computer with the existing Microsoft Store Python installation.
2. Right-click that ZIP, choose **Extract All**, and open the extracted folder.
3. Double-click **START-TEST.cmd** once. Keep its window open while it runs.
   It needs no typed commands or runtime selection. Follow an organization-policy
   refusal as a reported blocker; do not change policy or install another Python.
4. When the window says **REPORT READY**, send back only
   **SEND-THIS-REPORT.zip** from the displayed RESULTS folder, including when the
   test reports a failure. Keep other diagnostic files on that computer. If no
   report is produced, return the short error shown in the window.

The test verifies a fresh extraction of the contained distribution, automatically
discovers the existing runtime, performs normal application-local setup and
checks the BAT/PDF route. The Store scenario disables runtime acquisition and
does not accept an explicit Python path or hidden selection. Keep its extraction
and outputs. Existing application directories and shortcuts are not replaced.

## Reviewed identities

| Item | Bytes | SHA256 / identity |
| --- | ---: | --- |
| Product source F | - | `afd75b4f41ab7135e13fbb40818ade9e47daba03` |
| Product source tree | - | `08812489efdfb9921f87dca5ceba2c3ab5b37a67` |
| WinBookSplit-v1.0.3.zip | 935676 | `216c310e395a098e085f8c3a4db994a2c94072e58bc5982234fad1e5db7b9359` |
| release-manifest.json | 16594 | `78613d69ccc876a043947156476673e522ba4103fd82d8800e47b2a408397ab2` |
| SHA256SUMS.txt | 178 | `67a41d8fc33c57614866992bc24a52081bc1eac341dedfd96facb15bbf9abb73` |
| WinBookSplit-v103-native-kit-afd75b4.zip | 1564435 | `af4bf7107149f33c3a8d139f4ca5302d1427b7549dc982ee97a2572585ea49fb` |
| kit-bindings.json | - | `1d9677add6f4237edf850006f783161b33b59df9175547cc3625a94c5f971fe9` |
| Delivery source D | - | `f99c6960478f2d2d789935f1d54111d217130037` |
| WinBookSplit-Windows10-Test-afd75b4.zip | 3157973 | `c9672265f3a2b5a3653e9c36967f8f64a83bd8ac3799489ecdf412b92d0a18a3` |

The exact ZIP is already prepared in the agent computer's Downloads folder.
The owner need not rebuild it. Rebuilding this carrier uses the builder and four
delivery assets at D; old-carrier reconstruction uses their original delivery
commit. A later documentation checkpoint does not change F or these bytes.

## Review boundary

The actual new carrier was launched on Windows 11 from a path containing spaces,
apostrophe, ampersand and a Unicode character. Its Windows 10 preflight refused
the host, preserved exit 1, and delivered the unchanged failure ZIP. The offline
Store checker rejected that schema 1 failure with exit 2. This validates the
launcher/failure-report flow; it is not Windows 10 or Store acceptance. No
physical Explorer drag-and-drop or successful report GUI selection is claimed.

New-candidate native Store acceptance and separate STORE-024 genuine enforced
organization-policy denial remain NOT_RUN. The Store report intentionally keeps
the latter NOT_RUN; rerunning this test cannot clear it. No enforced target was
supplied, and desktop policy must not be changed to manufacture that evidence.
T08 publication remains BLOCKED. Scripts are unsigned.
