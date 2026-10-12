# V102-10 frozen R2 package review

The actual package was built once from committed source `483baf858569bc838ac8a809d329b2595cdacf57` (tree `818de909908c95395691dc962395f270c1cd5735`) using the four exact exported R2 builder members and observed regular x64 developer Python 3.14.8. The native build exited 0. A separate read-only inspector imported no builder or application code and independently verified all 34 ZIP member bytes against R2 Git blobs, the complete deterministic archive recipe, canonical manifest/dependency/sibling/builder projection and exact checksums.

The accepted ZIP is 909557 bytes, SHA256 `37d017e2a7bf704676b637f89239ec2cbaaa67e76f190f3dbbd17fec87e5092e`; the manifest is 16593 bytes, SHA256 `a8f8109bc8a49dadf890a10989a92accf6c0f6a038f29aac4ea51474f0792c03`; the checksums are 178 bytes, SHA256 `5aea43906c7cd55b841fad51e4d1e36b3c2425edd162f8b0cc388440f979a00e`. All three inspected assets were reread unchanged. No asset rebuild or mutation occurred during this review.

Clean local main equals a fresh live origin/main at the review time. This unsigned package contains source only, with no private handoff/dependencies. Final-Z native Windows 10/11, fixed tag and public anonymous-download verification remain separate gates. Prior source tests and Python 3.15 Server CI are not relabeled as final-Z client evidence.
