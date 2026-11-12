# Checkpoint 2 — Form classification and optional PDF behavior

The owner asked to keep the current v1.0.5 form repair and correct the additional
document-policy failure in the same worktree. The handoff ZIP supplies the
original form specification and authored fixtures. The later pasted request
expands the active page-copy policy. The PDF from the latest failure was unavailable.

The starting branch `codex/forms-v1.0.5` was clean at
`64426d069ccb49772652c92dd119d3c2ea1d1aa2` and matched live origin. The live
v1.0.5 tag and release were absent. The existing v1.0.4 tag and public assets
were preserved.

## Released baseline

The existing v1.0.4 package had ZIP SHA-256
`9d402f72092b20ad4d2a1985d8affef0ccfae525f520817e276185db4920a309`.
On Windows 11, its PowerShell 5.1 launcher rejected independent authored
JavaScript-only and attachment-only fixtures in Manual mode with exit 7,
`unsupported_document`, and the exact text
“Document scripts and embedded-file name trees are not supported.” Neither case
prepared a plan or published a chapter. A separate eight-case run covered
empty trees, both features together, and named-navigation combinations in
Manual, Auto1, Auto2 and Preview through PowerShell 5.1, PowerShell 7 and BAT.
All eight reached the same refusal. Its report SHA-256 is
`35853c18f3960cd622dc0dd8743af72f4ac2b24795e768ac481384cbf8d78de6`.

## Implementation

- The shared reader classifies AcroForm fields, widgets and XFA separately.
  Harmless metadata can proceed. Widget-bearing sources require preservation of
  their appearances. A nonempty XFA source requires verified complete rendering.
- Chapter writing projects the required page and static annotation graph into a
  fresh catalog, including the effective inherited page data and layer settings.
  Supported local links are rebuilt. Attachment appearances can be retained
  without the file relationship. Reopening checks output structure and serialized
  object reachability. See the [PDF policy audit](../PDF_POLICY.md).
- Plan and diagnostic form summaries contain only classifications and counts.
  Support exports include fixed omission categories and codes. An explicitly
  retained converted working PDF is still a byte-for-byte copy of the converter
  output. The interactive viewer route withholds that copy when it detects
  optional document behavior excluded from chapters.

## Frozen checkpoint package

The tested payload commit is `547875b9faa5f505da30ffec69b340b534e896d7`,
tree `20c6c2458037f6de63a9ff1f9a283000f86e66eb`. The existing builder produced
and verified the unpublished ZIP with SHA-256
`8b0c6e8fba2cdb6560c1391e7b2b30731e4ea9bd80c89c1b275dede8e4a8ef8d`.
The checkpoint still uses the inherited application VERSION, so the local
builder named this ZIP `WinBookSplit-v1.0.4.zip` in an isolated output directory.
All 34 freshly extracted files matched the package manifest.

The first archive construction preserved checkout CRLF conversions and the
builder rejected it with `builder_source_mismatch`. The successful build used
`git -c core.autocrlf=false archive` to read the committed builder bytes. The
repository configuration and package guards were preserved.

Commands used with absolute paths to the isolated developer interpreter,
separate application runtime, hosts, renderer and external outputs:

```powershell
& $DevPython -I -B $Builder build --repo $Repo --commit $SourceCommit --output $Assets
& $DevPython -I -B $Builder verify --repo $Repo --commit $SourceCommit --output $Assets
& $DevPython -I -B tests/document_policy/characterize_form_support.py --checkpoint-2 `
  --app-root $FreshExtraction --application-python-path $RuntimePython `
  --shell-path $PS51 --shell-path $PS7 --renderer-path $Poppler --report $PackageReport
& $DevPython -I -B tests/document_policy/characterize_document_policy.py `
  --shell-path $PS51 --shell-path $PS7 --report $PolicyReport
& $DevPython -I -B tests/run_tests.py --_application-units-child
```

## Automated results

Native runs used Windows 11 Pro build 26300, Windows PowerShell 5.1.26100.9711,
PowerShell 7.6.9, CPython 3.14.8, pypdf 6.21.4 and Poppler 26.07.0. The
runtime was selected explicitly. The BAT controls use the unchanged launcher
and the three documented copied PowerShell parameter defaults from checkpoint 1.

| Run | Observed result |
| --- | --- |
| Exact ZIP form and optional-feature matrix | 25/25 across PowerShell 5.1, PowerShell 7 and BAT. 18 split runs produced 54 reopened chapters, two Preview runs produced complete plans, and five required refusals used their declared reason and exit code. |
| Exact ZIP loopback URI and attachment canary | Both PowerShell hosts produced complete chapters. A separate health request reached the server. The document URL received zero requests and no attachment file appeared. |
| Exact ZIP success support export | Both PowerShell hosts exported fixed omission codes without warning text. |
| Exact ZIP real Calibre 9.17.0 EPUB | Both PowerShell hosts split three physical pages in order. The three chapter markers and decoded content matched the separately retained converter PDF. |
| Native document-policy matrix | 75/75 on the frozen source across PowerShell 5.1, PowerShell 7 and BAT. It verified 73 application controls and two redacted failure exports, with owned temporary cleanup complete. |
| Application unit layer | 178/178 in 204.316 seconds on the frozen source, including conversion, captured reader, password, outline, and resource cases. |
| Receipt and outline unit layers | 161 receipt tests and 47 Level 1/Level 2 tests passed. |

The package matrix exercised Manual, Auto1, Auto2, Preview and captured-reader
execution. It checked source and neighbor hashes, exact page ranges, decoded
content, rendered pixels and selected annotation regions, chapter-local links
and all serialized objects. Empty trees produced no script or attachment claim.
Indirect payload aliases and cyclic omitted name-tree branches passed. The
visible attachment fixture retained its supplied appearance. An attachment
without a usable supplied appearance uses an inert vector icon.

Selected package-run input SHA-256 values:

| Authored fixture | SHA-256 |
| --- | --- |
| JavaScript and attachment | `de4fde3e82be5235856bc219a3bbf7553764a9c4cf6c7efdf5e0186622a993f2` |
| Associated files and visible attachment | `abd26baf43fa72d770c5551b894b335161a16228c97834881cf778c07660efb3` |
| Shared indirect payload aliases | `d3c7ca1f3575680993261df6a6ca464e8ae6abf1db7d8d585651f181d8199866` |
| Empty-password metadata form, JavaScript and attachment | `16e674075fa173f96b348c6d6167063110f1e4ed77b3632df4a538eac5700a9c` |

Reports remain local and contain original authored inputs only. Report SHA-256:

- Exact ZIP native matrix: `a867e9e451be0e64e2620b4eb148d9aede85d2520a1fc0741decda709c8bea26`
- Loopback canary: `3f9a08d5b75be709fbbaab3647c992fb6c85b97fa50481c20f3eba02aa6a9e94`
- Calibre and canary combined: `16f102b89e0feb9f50488155f8b1a8db124f5c37da941e67ecbdefda5ccc71d1`
- Native document policy: `7f6b0438883ba842fed6afa88254abfe7e8a77a8df2c944be1666999f52b565c`
- Application units log: `9877f032f59952ccd08c87aec5c4b14a8bcf172c62d9ec26f30cbe5c84382c92`

An earlier checkout run overlapped implementation changes, including a transient
syntax error and a writer bug that left an unused object for a content array.
The frozen package is the acceptance receipt for those cases.

## Coverage limits

Successful filled-field and visible signature-widget normalization is
`NOT_RUN` at this checkpoint. The two authored filled-form combinations,
including the empty-password input, correctly returned
`form_appearance_required`. The XFA sentinel correctly returned
`xfa_rendering_required`. Real XFA rendering, cryptographic signature
verification, Windows 10, real AZW3 conversion, and ordinary Store first-run
consent/resume are `NOT_RUN` here. Publication remains at checkpoint 5.
