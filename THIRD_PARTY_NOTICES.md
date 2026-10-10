# Third-party notices

WinBookSplit's source, documentation and existing repository artwork are covered
by the project's unchanged [MIT license](LICENSE), copyright 2025 Janne Vuorela.
This does not relicense third-party components.

## Separately installed runtime components

The repository and intended small application distribution do not bundle the
Python interpreter, pypdf wheel, PowerShell installation or Calibre binaries.
Requirements and executable paths refer to components installed separately.
Their upstream license terms continue to apply to those installations.

| Component used by WinBookSplit | Upstream license and primary source |
| --- | --- |
| CPython 3.14.8 | [PSF License Version 2 and incorporated-software notices](https://docs.python.org/3.14/license.html) |
| pypdf 6.20.0 | [BSD 3-Clause license](https://github.com/py-pdf/pypdf/blob/6.20.0/LICENSE); [pinned package declaration](https://github.com/py-pdf/pypdf/blob/6.20.0/pyproject.toml) |
| PowerShell 7.6.6 | [MIT license, Microsoft Corporation](https://github.com/PowerShell/PowerShell/blob/v7.6.6/LICENSE.txt) |
| Calibre 9.15.0, for EPUB/AZW3 | [GNU GPL version 3](https://github.com/kovidgoyal/calibre/blob/v9.15.0/LICENSE) |

Windows PowerShell 5.1 is a Windows component governed by the applicable Windows
terms; the PowerShell 7 MIT notice is not a license for Windows or its inbox shell.
Calibre contains additional components with their own upstream notices. The links
above describe the named upstream projects, not a complete license inventory of
an installed Python/Windows/Calibre system.

Developer-only fixture packages, renderers, shell test modules and GitHub Actions
are not required for application use and are not bundled as application payload.
Their dependencies and verification are described in the source repository's
[tests/README.md](https://github.com/PikkuJanne/WinBookSplit/blob/main/tests/README.md)
and [tools/ci/README.md](https://github.com/PikkuJanne/WinBookSplit/blob/main/tools/ci/README.md).

No third-party runtime binary, wheel or vendored library is presently
redistributed in this repository. Before distributing
different bytes, audit that exact package and retain any required attribution,
license texts and source offers. A requirements file or this notice is not a
substitute for that package review. The project is not endorsed by these upstream
projects. Unsigned status and checksums are explained in [setup](docs/SETUP.md).
