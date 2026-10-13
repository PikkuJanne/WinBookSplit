# V102-12 — Anonymous public release verification

**PASS.** The unchanged reviewed verifier downloaded the public [WinBookSplit v1.0.2 release](https://github.com/PikkuJanne/WinBookSplit/releases/tag/v1.0.2) into a fresh owned directory with no authorization headers, cookies, netrc, inherited proxies or environment-token use. TLS certificate/hostname verification remained enabled. All four public metadata requests and three asset streams returned HTTP 200; each asset followed one ordinary HTTPS GitHub 302 redirect to `release-assets.githubusercontent.com`. Only host/status and safe date/request-ID headers are recorded; signed query URLs are omitted.

Actual verification ran 2026-10-13 17:29:23.162082–17:29:31.375364 UTC, exit 0, without transient failures. Public release ID `411934002` is non-draft and non-prerelease, published at `2026-10-13T17:28:52Z`. Its exact approved title, target R2 and UTF-8 notes bytes matched. Annotated `v1.0.2` object `c7ef8c4aa0421c553089319cf170ab9958fc2195` peels to R2 `483baf858569bc838ac8a809d329b2595cdacf57`; public commit metadata confirms tree `818de909908c95395691dc962395f270c1cd5735`. The exact tag-message hash matches, including its newline.

| Anonymous asset | Bytes | SHA256 |
| --- | ---: | --- |
| WinBookSplit-v1.0.2.zip | 909557 | `37d017e2a7bf704676b637f89239ec2cbaaa67e76f190f3dbbd17fec87e5092e` |
| release-manifest.json | 16593 | `a8f8109bc8a49dadf890a10989a92accf6c0f6a038f29aac4ea51474f0792c03` |
| SHA256SUMS.txt | 178 | `5aea43906c7cd55b841fad51e4d1e36b3c2425edd162f8b0cc388440f979a00e` |

Stream lengths, header lengths and hashes match the unchanged accepted assets and public API digests. The existing independent read-only inspector verified all 34 safe unique regular ZIP members, exact modes/bytes against committed R2 Git blobs, canonical manifest/builder records, dependency policy and lock, VERSION, new Setup BAT and checksum lines. No application or builder was run, no archive was extracted, and no repository, release or original accepted asset was modified.

Verifier source SHA256: `9924ab48f4fbadaacd6b0365f9f7c93dc1405b84d9e8c960a9c44651f1cc0564`. Actual receipt SHA256: `15d1ef44f473e6584979795fd4cd19ce43b7d1358bb90bf618fff870b4f0f77f`. Downloaded-package inspection SHA256: `e782da06733d84911002ef52b0b0a0682bc29c3de076068249e477578e087c79`, using authenticated inspector `7b253bc5632278839f1d4ac037d3f8e6ced26da56f0b5dcbfa9829117cfdefd3`. Invocation used the trusted developer Python with `-I -B Verify-AnonymousV102Assets.py --release-id 411934002 --tag-object c7ef8c4aa0421c553089319cf170ab9958fc2195`.

This record covers anonymous v1.0.2 metadata/assets and exact source correspondence. Old-release preservation and final clean live-main synchronization are recorded separately. The unsigned package status remains disclosed in the manifest and release notes.
