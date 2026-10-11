# V101-11 — preserve guards while correcting version fixtures

Task 11 remains IN_PROGRESS. On implementation **087deb7aff3c6de102336085088ce570cd28f5b2**,
the native ten-module subset passed 161/162 tests. Complete
[CI](https://github.com/PikkuJanne/WinBookSplit/actions/runs/38166990809) remained
**FAIL**: 644/645 Python tests and 185/192 Pester tests per shell passed. Both
application jobs passed. Parser/scoped-analysis findings were zero. The
[receipt](V101-11-fixture-correction.json) preserves all failure identities,
downloaded archive hashes, rosters and qualified focused correction results.

The failures were stale 1.0.0 expectations/positive manifests after the actual
version became 1.0.1. Four test files now use the intended shipped version or
the existing strict reader. Replacement-ref/provenance assertions, malformed
version tokens and support secret/allowlist refusals remain. No application
payload file changed. The actual corrected replacement case passed 1/1, and
both native PowerShell hosts passed all 52 affected-suite tests (29 Outcome,
four Version, 19 Support), zero skips/failures and zero parser/PSSA findings.
Owned SDK trees stopped safely and each recorded source map stayed unchanged.
These runs bind the precise uncommitted test bytes plus base 087; they do not
claim the earlier clean commit passed or certify the next commit.

The actual builder build/verify also passed a 33-member unpublished development
package from 087. Both native source Version queries returned 1.0.1 / exit 0. Main
integration, a new release source R/assets, matching corrected complete CI,
native final-ZIP Windows10/Windows11 and publication remain separate gates.
