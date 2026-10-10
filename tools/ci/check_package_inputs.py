#!/usr/bin/env python3
"""Check committed source prerequisites; this does not build or certify a package."""
from __future__ import annotations

import argparse
import ast
from hashlib import sha256
import json
import os
from pathlib import Path, PurePosixPath
import re
import stat
import subprocess
import sys


RUNTIME_PATHS = (
    "WinBookSplit.bat", "WinBookSplit.ps1", "Export-WinBookSplitDiagnostics.ps1", "requirements.txt",
    "engine/WinBookSplit.Runtime.ps1", "engine/WinBookSplit.Process.ps1",
    "engine/WinBookSplit.Diagnostics.ps1", "engine/WinBookSplit.Paths.ps1",
    "engine/WinBookSplit.Logging.ps1", "engine/WinBookSplit.Support.ps1",
    "engine/WinBookSplit.Outcomes.json", "engine/winbooksplit_engine.py",
    "engine/winbooksplit_conversion.py", "engine/winbooksplit_windows.py", "engine/winbooksplit_job.py",
)
# Current user/support source prerequisites, not the future M5 package payload.
SUPPORT_PATHS = (
    "README.md", "LICENSE", "WinBookSplit_icon.ico", "docs/PDF_POLICY.md",
    "docs/PDF_FIDELITY.md", "docs/EBOOK_SUPPORT.md", "docs/SUPPORT_DIAGNOSTICS.md",
    "docs/codex-v1.0.0/SUPPORT_AND_SETUP.md",
)
INPUT_PATHS = RUNTIME_PATHS + SUPPORT_PATHS
RUNTIME_REQUIREMENT = "pypdf==6.19.0 --hash=sha256:7e5d6e730e7dae87d560a2cee218b852f6498c8be61966f3cd02ead971e48d14"
MAX_FILE_BYTES = 8 * 1024 * 1024
GIT_TIMEOUT_SECONDS = 30
SHA1 = re.compile(r"[0-9a-f]{40}\Z")
JOIN_SIBLING = re.compile(r"Join-Path\s+\$PSScriptRoot\s+(['\"])([^'\"\r\n]+)\1", re.IGNORECASE)
WINDOWS_RESERVED = re.compile(r"(?:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\..*)?\Z", re.IGNORECASE)
REPOSITORY_ENV_KEYS = frozenset({"GIT_DIR", "GIT_WORK_TREE", "GIT_INDEX_FILE", "GIT_OBJECT_DIRECTORY",
    "GIT_ALTERNATE_OBJECT_DIRECTORIES", "GIT_COMMON_DIR", "GIT_NAMESPACE"})


class InputError(ValueError):
    def __init__(self, code: str, path: str | None = None):
        super().__init__(code)
        self.code, self.path = code, path


def safe_relative(path: str) -> bool:
    if not isinstance(path, str) or not path or "\\" in path:
        return False
    parts = path.split("/")
    return not PurePosixPath(path).is_absolute() and all(
        part not in {"", ".", ".."} and part[-1] not in {".", " "}
        and not WINDOWS_RESERVED.fullmatch(part)
        and not any(ord(char) < 32 or ord(char) == 127 or char in '<>:"|?*' for char in part)
        for part in parts)


def git(repo: Path, *args: str) -> bytes:
    environment = {key: value for key, value in os.environ.items() if key.upper() not in REPOSITORY_ENV_KEYS}
    try:
        process = subprocess.run(["git", "-C", str(repo), *args], stdin=subprocess.DEVNULL,
                                 stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                 timeout=GIT_TIMEOUT_SECONDS, check=False, env=environment)
    except (OSError, subprocess.TimeoutExpired) as error:
        raise InputError("git_unavailable") from error
    if process.returncode:
        # Git stderr can include local paths or credential-bearing remote URLs.
        raise InputError("git_query_failed")
    return process.stdout


def _text(data: bytes, path: str) -> str:
    try:
        value = data.decode("utf-8-sig", errors="strict")
    except UnicodeError as error:
        raise InputError("invalid_text_input", path) from error
    if "\x00" in value:
        raise InputError("invalid_text_input", path)
    return value


def _dependency(path: str, target: str) -> str:
    if not safe_relative(target):
        raise InputError("unsafe_dependency", path)
    dependency = (PurePosixPath(path).parent / target).as_posix()
    if dependency not in RUNTIME_PATHS:
        raise InputError("undeclared_runtime_dependency", path)
    return dependency


def _sibling_call(node: ast.AST) -> str | None:
    # Only the current explicit Path(__file__).resolve().with_name('helper') pattern.
    if not isinstance(node, ast.Call) or not isinstance(node.func, ast.Attribute) or node.func.attr != "with_name":
        return None
    resolved = node.func.value
    if len(node.args) != 1 or node.keywords or not isinstance(node.args[0], ast.Constant) or type(node.args[0].value) is not str \
            or not isinstance(resolved, ast.Call) or resolved.args or resolved.keywords \
            or not isinstance(resolved.func, ast.Attribute) or resolved.func.attr != "resolve":
        raise InputError("unresolved_runtime_dependency")
    constructor = resolved.func.value
    if not isinstance(constructor, ast.Call) or not isinstance(constructor.func, ast.Name) or constructor.func.id != "Path" \
            or len(constructor.args) != 1 or constructor.keywords \
            or not isinstance(constructor.args[0], ast.Name) or constructor.args[0].id != "__file__":
        raise InputError("unresolved_runtime_dependency")
    return node.args[0].value


def _python_dependencies(text: str, path: str) -> set[str]:
    try:
        tree = ast.parse(text, filename=path)
    except SyntaxError as error:
        raise InputError("invalid_python_input", path) from error
    dependencies = set()
    for node in ast.walk(tree):
        modules = [alias.name for alias in node.names] if isinstance(node, ast.Import) else \
            [node.module or ""] if isinstance(node, ast.ImportFrom) else []
        for name in modules:
            root = name.partition(".")[0]
            if isinstance(node, ast.ImportFrom) and node.level:
                raise InputError("undeclared_runtime_dependency", path)
            if root.startswith("winbooksplit_"):
                dependencies.add(_dependency(path, root + ".py"))
            elif root not in sys.stdlib_module_names and root != "pypdf":
                raise InputError("undeclared_runtime_dependency", path)
        if isinstance(node, ast.Call):
            sibling = _sibling_call(node)
            if sibling is not None:
                dependencies.add(_dependency(path, sibling))
            if isinstance(node.func, ast.Name) and node.func.id == "__import__" or \
                    isinstance(node.func, ast.Attribute) and node.func.attr == "import_module":
                raise InputError("unresolved_runtime_dependency", path)

    def scope(node: ast.AST, inherited: dict[str, str]) -> None:
        bindings = dict(inherited)
        local = []
        def visit(item: ast.AST) -> None:
            if isinstance(item, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
                return
            local.append(item)
            for child in ast.iter_child_nodes(item):
                visit(child)
        for child in ast.iter_child_nodes(node):
            visit(child)
        for item in local:
            if isinstance(item, ast.Assign):
                sibling = _sibling_call(item.value)
                for target in item.targets:
                    if isinstance(target, ast.Name):
                        if sibling is not None:
                            bindings[target.id] = sibling
                        else:
                            bindings.pop(target.id, None)
        for item in local:
            if isinstance(item, ast.Call) and isinstance(item.func, ast.Attribute) and item.func.attr == "spec_from_file_location":
                if len(item.args) != 2 or item.keywords:
                    raise InputError("unresolved_runtime_dependency", path)
                target = item.args[1]
                sibling = bindings.get(target.id) if isinstance(target, ast.Name) else _sibling_call(target)
                if sibling is None:
                    raise InputError("unresolved_runtime_dependency", path)
                dependencies.add(_dependency(path, sibling))
        for item in ast.iter_child_nodes(node):
            if isinstance(item, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
                scope(item, bindings)
    scope(tree, {})
    return dependencies


def validate_inputs(repo: Path, commit: str) -> dict:
    if not isinstance(commit, str) or SHA1.fullmatch(commit) is None:
        raise InputError("explicit_commit_required")
    repo = Path(repo).resolve(strict=True)
    actual_root = Path(os.fsdecode(git(repo, "rev-parse", "--show-toplevel")).strip()).resolve(strict=True)
    if actual_root != repo:
        raise InputError("repository_root_required")
    actual_commit = git(repo, "rev-parse", "--verify", commit + "^{commit}").decode("ascii").strip()
    if actual_commit != commit:
        raise InputError("commit_identity_mismatch")
    tree = git(repo, "rev-parse", commit + "^{tree}").decode("ascii").strip()
    if not SHA1.fullmatch(tree):
        raise InputError("invalid_tree_identity")
    if len(INPUT_PATHS) != len(set(path.casefold() for path in INPUT_PATHS)) or not all(map(safe_relative, INPUT_PATHS)):
        raise InputError("invalid_input_allowlist")
    entries = {}
    for record in git(repo, "ls-tree", "-r", "-z", commit, "--", *INPUT_PATHS).split(b"\0"):
        if not record:
            continue
        try:
            metadata, raw_path = record.split(b"\t", 1)
            mode, kind, oid = metadata.decode("ascii").split(" ")
            path = raw_path.decode("utf-8", errors="strict")
        except (ValueError, UnicodeError) as error:
            raise InputError("invalid_git_entry") from error
        if not safe_relative(path) or path not in INPUT_PATHS or path in entries:
            raise InputError("invalid_git_entry")
        if mode not in {"100644", "100755"} or kind != "blob" or not SHA1.fullmatch(oid):
            raise InputError("nonordinary_input", path)
        entries[path] = (mode, oid)
    missing = sorted(set(INPUT_PATHS) - entries.keys())
    if missing:
        raise InputError("missing_input", missing[0])
    blobs, files = {}, []
    for path in sorted(INPUT_PATHS):
        mode, oid = entries[path]
        size = int(git(repo, "cat-file", "-s", oid).decode("ascii").strip())
        if not 0 < size <= MAX_FILE_BYTES:
            raise InputError("invalid_input_size", path)
        data = git(repo, "cat-file", "blob", oid)
        if len(data) != size:
            raise InputError("blob_size_mismatch", path)
        blobs[path] = data
        files.append({"path": path, "scope": "runtime" if path in RUNTIME_PATHS else "support_source",
                      "git_mode": mode, "git_blob": oid, "bytes": size, "sha256": sha256(data).hexdigest()})
    requirements = [line.strip() for line in _text(blobs["requirements.txt"], "requirements.txt").splitlines()
                    if line.strip() and not line.lstrip().startswith("#")]
    if requirements != [RUNTIME_REQUIREMENT]:
        raise InputError("runtime_dependency_pin_changed", "requirements.txt")
    dependencies = {}
    for path in RUNTIME_PATHS:
        if path.endswith(".py"):
            dependencies[path] = sorted(_python_dependencies(_text(blobs[path], path), path))
        elif path.endswith(".ps1"):
            dependencies[path] = sorted({_dependency(path, match.group(2).replace("\\", "/"))
                                         for match in JOIN_SIBLING.finditer(_text(blobs[path], path))})
    bat = _text(blobs["WinBookSplit.bat"], "WinBookSplit.bat")
    references = re.findall(r"%~dp0([^\"\r\n]+)", bat, re.IGNORECASE)
    if not references or any(reference != "WinBookSplit.ps1" for reference in references):
        raise InputError("launcher_dependency_changed", "WinBookSplit.bat")
    dependencies["WinBookSplit.bat"] = ["WinBookSplit.ps1"]
    canonical = json.dumps(files, sort_keys=True, separators=(",", ":"), ensure_ascii=True).encode("ascii")
    return {"schema_version": 1, "check": "package_source_inputs", "success": True, "exit_code": 0,
            "source_commit": commit, "source_tree": tree, "input_sha256": sha256(canonical).hexdigest(),
            "runtime_count": len(RUNTIME_PATHS), "support_source_count": len(SUPPORT_PATHS),
            "files": files, "runtime_sibling_dependencies": dependencies,
            "scope": "Committed source prerequisites only; no package build, release readiness or private-data scan.",
            "working_tree_inputs_used": False, "package_built": False}


def report_destination(repo: Path, report: Path) -> Path:
    repo = Path(repo).resolve(strict=True)
    report = Path(report)
    if not report.is_absolute() or report.name in {"", ".", ".."}:
        raise InputError("external_report_required")
    lexical = Path(os.path.abspath(report))
    parent = report.parent.resolve(strict=True)
    resolved = parent / report.name
    if lexical.is_relative_to(repo) or resolved.is_relative_to(repo):
        raise InputError("external_report_required")
    info = report.parent.lstat()
    if not stat.S_ISDIR(info.st_mode) or stat.S_ISLNK(info.st_mode) or getattr(info, "st_file_attributes", 0) & 0x400:
        raise InputError("ordinary_report_parent_required")
    return resolved


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repo", type=Path, required=True)
    parser.add_argument("--commit", required=True)
    parser.add_argument("--report", type=Path, required=True)
    arguments = parser.parse_args(argv)
    try:
        destination = report_destination(arguments.repo, arguments.report)
        # Reserve a new external file before reading inputs; never overwrite evidence.
        with destination.open("x", encoding="utf-8", newline="\n") as stream:
            try:
                result = validate_inputs(arguments.repo, arguments.commit)
            except InputError as error:
                result = {"schema_version": 1, "check": "package_source_inputs", "success": False,
                          "exit_code": 1, "error_code": error.code,
                          "input_path": error.path if error.path in INPUT_PATHS else None}
            json.dump(result, stream, indent=2, ensure_ascii=True)
            stream.write("\n")
        print("PACKAGE_INPUTS_PASSED" if result["success"] else "PACKAGE_INPUTS_FAILED: " + result["error_code"])
        return result["exit_code"]
    except (InputError, OSError, UnicodeError, ValueError):
        print("PACKAGE_INPUTS_FAILED: report_or_repository_invalid", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
