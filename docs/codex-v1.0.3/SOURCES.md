# Sources and attribution

Observed incident facts come from OWNER, with original line ranges in ROOT_CAUSE.
Code findings come from RUNTIME, SUMMARY and POLICY. Old acceptance limitations
come from OLD_NATIVE and KIT/KIT_VALIDATOR. All URLs and reference IDs are in
SOURCE_MAP.json; implementation recommendations are this handoff's design, not
claims that upstream already implements them.

Official Python documentation confirms two MSIX family identities and directs
programmatic use to documented WindowsApps entry points rather than the physical
Program Files executable. Its installed-runtime listing supports explicit machine
formats; online listing is a separate option. Hard-linked global entry points are
a documented default. These explain why the captured installation is legitimate;
they do not certify a WinBookSplit patch. Recheck current documentation and actual
registered package metadata in the native environment.

Official Microsoft Get-AppxPackage documentation explains the current-user query
and its filtering. No recommendation here authorizes elevation or all-user access.

Repository reads used the connected GitHub tool. Direct raw-file download into
the analysis container was unavailable (DNS resolution); no independently
byte-for-byte verified repository snapshot is bundled. Recorded local source
fingerprints are the owner's observations and must be compared by Codex. The
raw owner report was read locally to make the sanitized fixture and verify its
fingerprint. No Windows process was executed during bundle preparation.
