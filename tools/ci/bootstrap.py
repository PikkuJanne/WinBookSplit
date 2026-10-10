"""Download hash-pinned developer modules into a new, isolated CI tool root."""
import argparse
from datetime import datetime, timezone
from hashlib import sha256
import io
import json
from pathlib import Path, PurePosixPath
import re
import stat
import sys
from urllib.request import Request, urlopen
import zipfile

ROOT = Path(__file__).resolve().parents[2]
# Reviewed official Gallery package bytes; never substitute inbox/global modules.
PACKAGES = (
    ("Pester", "6.2.0", "e6ac7418d4f12500269aaca58ae56cf1caafbbf1afa2cec334e289d1cf50a239"),
    ("PSScriptAnalyzer", "1.25.0", "14e634c828eb98efb9f40b2918ba90f139ed5eccdf663a2a747736d996995d60"),
)
MAX_PACKAGE = 64 * 1024 * 1024
# Reviewed PSScriptAnalyzer 1.25.0 expands to 299,357,415 bytes (53 members).
MAX_EXPANDED = 320 * 1024 * 1024
MAX_MEMBERS = 4096


def extract_package(data, target, expected_sha256):
    """Verify bytes first, then create only ordinary bounded relative members."""
    if not isinstance(data, bytes) or not 0 < len(data) <= MAX_PACKAGE:
        raise ValueError("Developer package size exceeds its bound")
    if sha256(data).hexdigest() != expected_sha256:
        raise ValueError("Developer package SHA256 differs from its reviewed pin")
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        members = archive.infolist()
        if not 0 < len(members) <= MAX_MEMBERS or sum(item.file_size for item in members) > MAX_EXPANDED:
            raise ValueError("Developer package expanded size/member count exceeds its bound")
        names = set()
        for item in members:
            original = item.orig_filename
            name = original[:-1] if original.endswith("/") else original
            path = PurePosixPath(name)
            parts = name.split("/")
            if (not name or item.filename != original or "\\" in name or path.is_absolute()
                    or any(not part or part in {".", ".."} or part.endswith((".", " "))
                           or re.fullmatch(r"(?i)(?:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\..*)?", part)
                           or any(ord(char) < 32 or ord(char) == 127 or char in ':<>"|?*' for char in part)
                           for part in parts)
                    or str(path) != name or name.casefold() in names
                    or item.flag_bits & 1 or item.file_size > MAX_PACKAGE
                    or stat.S_IFMT(item.external_attr >> 16) not in (0, stat.S_IFREG, stat.S_IFDIR)):
                raise ValueError("Developer package has an unsafe/duplicate member")
            names.add(name.casefold())
        target.mkdir(parents=True, exist_ok=False)
        for item in members:
            path = target.joinpath(*PurePosixPath(item.filename.rstrip("/")).parts)
            if item.is_dir():
                path.mkdir(parents=True, exist_ok=True)
            else:
                path.parent.mkdir(parents=True, exist_ok=True)
                with path.open("xb") as stream:
                    stream.write(archive.read(item))
    return {str(path.relative_to(target)).replace("\\", "/"): sha256(path.read_bytes()).hexdigest()
            for path in sorted(target.rglob("*")) if path.is_file()}


def bootstrap(tool_root, report, observations=None):
    if (not tool_root.is_absolute() or tool_root.exists() or tool_root.is_relative_to(ROOT)
            or not report.is_absolute() or report.exists() or report.is_relative_to(ROOT)
            or not report.parent.is_dir()):
        raise ValueError("Choose new absolute external tool/report paths")
    tool_root.mkdir()
    if observations is None:
        observations = []
    for name, version, expected in PACKAGES:
        url = f"https://www.powershellgallery.com/api/v2/package/{name}/{version}"
        observation = {"name": name, "version": version, "url": url,
                       "expected_sha256": expected, "state": "download_attempted"}
        observations.append(observation)
        request = Request(url, headers={"User-Agent": "WinBookSplit-CI-bootstrap/1"})
        with urlopen(request, timeout=60) as response:
            data = response.read(MAX_PACKAGE + 1)
        actual = sha256(data).hexdigest()
        observation.update(package_sha256=actual, package_bytes=len(data), state="downloaded")
        target = tool_root / name / version
        files = extract_package(data, target, expected)
        if name + ".psd1" not in files:
            raise ValueError("Pinned developer module manifest is absent")
        observation.update(files_sha256=files, state="verified_and_extracted")
    return {"schema_version": 1, "result": "CI_TOOL_BOOTSTRAP_PASSED", "success": True,
            "exit_code": 0, "observed_at": datetime.now(timezone.utc).isoformat(),
            "python": sys.version, "packages": observations,
            "scope": "Isolated developer modules only; no global install, repository trust or policy change."}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--tool-root", required=True, type=Path)
    parser.add_argument("--report", required=True, type=Path)
    args = parser.parse_args()
    report = args.report.resolve()
    observations = []
    try:
        if not sys.flags.isolated or not sys.dont_write_bytecode:
            raise ValueError("Invoke isolated developer Python with -I -B")
        if not args.tool_root.is_absolute() or not args.report.is_absolute():
            raise ValueError("Tool/report paths must be absolute")
        record = bootstrap(args.tool_root.resolve(), report, observations)
    except Exception as error:
        record = {"schema_version": 1, "result": "CI_TOOL_BOOTSTRAP_FAILED", "success": False,
                  "exit_code": 1, "error_type": type(error).__name__, "error": str(error), "packages": observations,
                  "scope": "Developer bootstrap failure; no test pass or cleanup claim."}
    if report.parent.is_dir() and not report.exists() and not report.is_relative_to(ROOT):
        with report.open("x", encoding="utf-8", newline="\n") as stream:
            json.dump(record, stream, ensure_ascii=True, indent=2)
            stream.write("\n")
    print(record["result"])
    return record["exit_code"]


if __name__ == "__main__":
    raise SystemExit(main())
