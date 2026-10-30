"""Create the checkpoint-4 native workflow kit from an already verified package.

No runtime acquisition, application setup or global installation belongs here.
The existing developer pypdf is used only to author a public six-page fixture.
"""
from __future__ import annotations

import argparse
from hashlib import sha256
import importlib.util
import json
from pathlib import Path
import shutil
import zipfile

ROOT = Path(__file__).resolve().parents[2]
HELPERS = (
    "tools/compatibility/Invoke-FunctionalAcceptance.ps1",
    "tools/compatibility/functional_oracle.py",
    "tools/compatibility/Invoke-NativeAcceptance.ps1",
    "tools/compatibility/Invoke-StoreExistingAcceptance.ps1",
    "tools/compatibility/Invoke-NativeDiscoveryCollector.ps1",
)


def digest(data):
    return {"sha256": sha256(data).hexdigest(), "bytes": len(data)}


def build(package_dir, destination):
    spec = importlib.util.spec_from_file_location("functional_fixture", ROOT / "tests/python/test_functional_pdf_repair.py")
    fixture = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(fixture)
    from pypdf import PdfReader
    from io import BytesIO
    package_dir, destination = Path(package_dir).resolve(), Path(destination).absolute()
    if destination.exists() or destination.is_relative_to(ROOT):
        raise ValueError("new external destination required")
    manifest_data = (package_dir / "release-manifest.json").read_bytes()
    manifest = json.loads(manifest_data)
    archive_data = (package_dir / manifest["archive"]["name"]).read_bytes()
    if digest(archive_data) != {"sha256": manifest["archive"]["sha256"], "bytes": manifest["archive"]["bytes"]}:
        raise ValueError("package hash mismatch")
    payload = {}
    with zipfile.ZipFile(BytesIO(archive_data)) as archive:
        for row in manifest["files"]:
            payload[row["path"]] = archive.read(row["path"])
            if digest(payload[row["path"]]) != {"sha256": row["sha256"], "bytes": row["bytes"]}:
                raise ValueError("package payload mismatch")
        if sorted(archive.namelist()) != sorted(payload):
            raise ValueError("package member mismatch")
    destination.mkdir()
    files = {}
    def put(name, data):
        target = destination / name
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(data)
        files[name] = digest(data)
    put("package.zip", archive_data)
    put("release-manifest.json", manifest_data)
    for name, data in payload.items():
        if name.startswith("engine/"):
            put("support/" + name, data)
    for name in HELPERS:
        put(name, (ROOT / name).read_bytes())
    data = fixture.build_fixture("empty_password_and_local_goto")
    put("combined.pdf", data)
    reader = PdfReader(BytesIO(data))
    if not reader.decrypt(""):
        raise ValueError("authored empty password fixture invalid")
    # The oracle is derived from authored literal content and ranges, never an
    # application plan or a copy of the application's output.
    page_hashes = [sha256(f"BT /F1 16 Tf 72 720 Td (Synthetic test page {number}) Tj ET\n".encode("ascii")).hexdigest()
                   for number in range(1, 7)]
    if [sha256(page.get_contents().get_data()).hexdigest() for page in reader.pages] != page_hashes:
        raise ValueError("fixture differs from independently authored page markers")
    expected = {"source": digest(data), "ranges": [[0, 2], [2, 4], [4, 6]], "pages": page_hashes}
    put("expected.json", (json.dumps(expected, indent=2) + "\n").encode())
    bindings = {"schema_version": 1, "source_commit": manifest["source_commit"],
                "source_tree": manifest["source_tree"], "files": files}
    raw = (json.dumps(bindings, sort_keys=True, indent=2) + "\n").encode()
    (destination / "bindings.json").write_bytes(raw)
    start = ('@echo off\r\nsetlocal DisableDelayedExpansion\r\n'
             '"%SystemRoot%\\System32\\WindowsPowerShell\\v1.0\\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File '
             '"%~dp0tools\\compatibility\\Invoke-FunctionalAcceptance.ps1" -KitRoot "%~dp0." -BindingsSha256 "'
             + sha256(raw).hexdigest() + '"\r\nset "FunctionalKitExitCode=%ERRORLEVEL%"\r\necho.\r\npause\r\nexit /b %FunctionalKitExitCode%\r\n')
    (destination / "START-HERE.bat").write_bytes(start.encode("ascii"))
    (destination / "READ-ME.txt").write_text(
        "1. Extract this complete kit into an ordinary local folder.\n"
        "2. Double-click START-HERE.bat. It runs automatic checks and may perform\n"
        "   the application's explicit consent-based LOCAL setup after verifying\n"
        "   an already installed supported Store Python without pypdf.\n"
        "   It never installs Python or changes PATH/global packages.\n"
        "3. Return the one SEND-THIS-REPORT.zip printed at the end. Failures and\n"
        "   missing prerequisites also produce that report. Keep the folder.\n"
        "Raw paths/logs stay local; the report redacts profile/kit/work roots.\n"
        "No screenshots or manual chapter inspection are needed. No publication.\n",
        encoding="utf-8")
    archive_path = destination.with_suffix(".zip")
    if archive_path.exists():
        raise ValueError("kit archive already exists")
    with zipfile.ZipFile(archive_path, "x", zipfile.ZIP_STORED) as archive:
        for item in sorted(destination.rglob("*")):
            if item.is_file():
                archive.write(item, str(item.relative_to(destination)).replace("\\", "/"))
    return {"kit": str(destination), "archive": str(archive_path), **digest(archive_path.read_bytes()),
            "bindings_sha256": sha256(raw).hexdigest(), "source_commit": manifest["source_commit"]}


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--package-dir", required=True, type=Path)
    parser.add_argument("--destination", required=True, type=Path)
    args = parser.parse_args()
    print(json.dumps(build(args.package_dir, args.destination)))
