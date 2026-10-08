"""Standard-library safety helpers for the Codex handoff, not WinBookSplit runtime."""
from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path, PurePosixPath
import re
import stat
import subprocess
from typing import Any, Sequence
from urllib.parse import urlsplit

EXPECTED_REPO = "PikkuJanne/WinBookSplit"
ALLOWED_PREFIXES = ("docs/codex-v1.0.0/", "tools/codex-handoff/")

class HandoffError(RuntimeError):
    """An operation failed safely; caller must not report success."""


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def load_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        raise HandoffError(f"Cannot read valid JSON: {path.name}") from exc


def relative_parts(value: str) -> tuple[str, ...]:
    if not isinstance(value, str) or not value or "\\" in value or ":" in value:
        raise HandoffError("Invalid relative path in handoff metadata")
    p = PurePosixPath(value)
    if p.is_absolute() or any(v in ("", ".", "..") for v in value.split("/")):
        raise HandoffError("Unsafe relative path in handoff metadata")
    if any(v.lower() == ".git" for v in p.parts):
        raise HandoffError("Git internals are not a handoff destination")
    return p.parts


def is_link_or_reparse(path: Path) -> bool:
    try:
        s = path.lstat()
    except FileNotFoundError:
        return False
    return stat.S_ISLNK(s.st_mode) or bool(
        getattr(s, "st_file_attributes", 0) & getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0x400)
    )


def safe_path(root: Path, relative: str) -> Path:
    """Reject traversal and existing symlink/junction ancestors, including leaf."""
    root = root.resolve(strict=True)
    parts = relative_parts(relative)
    current = root
    for i, part in enumerate(parts):
        current = current / part
        if is_link_or_reparse(current):
            raise HandoffError("Refusing a symlink/junction/reparse handoff path")
        if current.exists() and i < len(parts) - 1 and not current.is_dir():
            raise HandoffError("A handoff directory is occupied by a file")
    try:
        current.resolve(strict=False).relative_to(root)
    except ValueError as exc:
        raise HandoffError("Handoff path escaped its root") from exc
    return current


def git(repo: Path, args: Sequence[str], *, timeout: int = 30) -> subprocess.CompletedProcess[str]:
    env = os.environ.copy()
    env.update({"GIT_TERMINAL_PROMPT": "0", "GCM_INTERACTIVE": "Never"})
    # Avoid environment variables selecting an unrelated repository/index.
    for key in ("GIT_DIR", "GIT_WORK_TREE", "GIT_INDEX_FILE", "GIT_OBJECT_DIRECTORY", "GIT_ALTERNATE_OBJECT_DIRECTORIES"):
        env.pop(key, None)
    try:
        return subprocess.run(
            ["git", "--no-optional-locks", "-C", str(repo), *args],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            timeout=timeout, check=False, env=env,
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        # Do not echo stderr, URLs or credential material from Git configuration.
        raise HandoffError("Git unavailable or timed out; operation is UNKNOWN") from exc


def must_git(repo: Path, args: Sequence[str]) -> str:
    r = git(repo, args)
    if r.returncode:
        raise HandoffError("Git inspection failed; inspect configuration/authentication locally")
    return r.stdout.strip()


def normalize_remote(url: str) -> str:
    """Accept canonical GitHub SSH/HTTPS remotes; never expose embedded secrets."""
    u = url.strip()
    if u.startswith("git@github.com:"):
        path = u[len("git@github.com:"):]
    else:
        try:
            p = urlsplit(u)
            if p.hostname is None or p.hostname.lower() != "github.com" or p.query or p.fragment:
                raise ValueError()
            if p.scheme == "https":
                if p.username is not None or p.password is not None or p.port not in (None, 443):
                    raise ValueError()
            elif p.scheme == "ssh":
                if p.username != "git" or p.password is not None or p.port not in (None, 22):
                    raise ValueError()
            else:
                raise ValueError()
            path = p.path.lstrip("/")
        except ValueError as exc:
            raise HandoffError("Remote identity is unsupported or contains credentials; inspect locally") from exc
    path = path.rstrip("/")
    if path.endswith(".git"):
        path = path[:-4]
    if not re.fullmatch(r"[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+", path):
        raise HandoffError("Malformed GitHub repository identity")
    return path.lower()


def inspect_repository(repo: Path, *, require_clean: bool = False) -> dict[str, Any]:
    repo = repo.expanduser().resolve(strict=True)
    top = Path(must_git(repo, ["rev-parse", "--show-toplevel"])).resolve()
    if top != repo:
        raise HandoffError("--repo must name the Git working-tree root, not a subdirectory")
    if must_git(repo, ["rev-parse", "--is-bare-repository"]) != "false":
        raise HandoffError("A normal working tree is required")
    # A rewrite can cause a canonical GitHub ls-remote URL to query a different host.
    rewrites = git(repo, ["config", "--get-regexp", r"^url\..*\.(insteadof|pushinsteadof)$"])
    if rewrites.returncode == 0 and rewrites.stdout.strip():
        raise HandoffError("Git URL rewrite configuration needs manual review before identity verification")
    if rewrites.returncode not in (0, 1):
        raise HandoffError("Cannot inspect Git URL rewrite configuration")
    fetch_urls = must_git(repo, ["remote", "get-url", "--all", "origin"]).splitlines()
    push_urls = must_git(repo, ["remote", "get-url", "--push", "--all", "origin"]).splitlines()
    if len(fetch_urls) != 1 or len(push_urls) != 1:
        raise HandoffError("Exactly one origin fetch URL and one push URL are required")
    for url in fetch_urls + push_urls:
        if normalize_remote(url) != EXPECTED_REPO.lower():
            raise HandoffError("Origin does not identify the expected PikkuJanne/WinBookSplit repository")
    for state in ("MERGE_HEAD", "CHERRY_PICK_HEAD", "REVERT_HEAD", "rebase-merge", "rebase-apply", "BISECT_LOG"):
        raw = must_git(repo, ["rev-parse", "--git-path", state])
        target = Path(raw)
        if not target.is_absolute():
            target = repo / target
        if target.exists():
            raise HandoffError("An existing Git operation must be reconciled before this handoff")
    head = must_git(repo, ["rev-parse", "--verify", "HEAD"])
    if not re.fullmatch(r"[0-9a-f]{40}|[0-9a-f]{64}", head):
        raise HandoffError("Unexpected HEAD identity")
    branch_r = git(repo, ["symbolic-ref", "--quiet", "--short", "HEAD"])
    if branch_r.returncode != 0:
        raise HandoffError("Detached HEAD is not a synchronized work branch")
    branch = branch_r.stdout.strip()
    if not branch or git(repo, ["check-ref-format", "--branch", branch]).returncode:
        raise HandoffError("Invalid branch identity")
    dirty = bool(must_git(repo, ["status", "--porcelain=v1", "--untracked-files=all"]))
    if require_clean and dirty:
        raise HandoffError("Worktree is dirty/untracked; preserve and reconcile it before importing")
    return {"root": repo, "head": head, "branch": branch, "dirty": dirty,
            "fetch_url": fetch_urls[0], "push_url": push_urls[0]}


def verify_bundle(root: Path) -> dict[str, dict[str, Any]]:
    root = root.resolve(strict=True)
    manifest = load_json(safe_path(root, "BUNDLE_MANIFEST.json"))
    if manifest.get("schema_version") != 1 or not isinstance(manifest.get("files"), list):
        raise HandoffError("Unsupported bundle manifest")
    rows: dict[str, dict[str, Any]] = {}
    for row in manifest["files"]:
        name = row.get("path")
        if name in rows or name == "BUNDLE_MANIFEST.json":
            raise HandoffError("Duplicate or self-referential manifest path")
        p = safe_path(root, name)
        if not p.is_file():
            raise HandoffError("A manifested bundle file is missing")
        data = p.read_bytes()
        if len(data) != row.get("bytes") or sha256_bytes(data) != row.get("sha256"):
            raise HandoffError(f"Bundle integrity mismatch: {name}")
        rows[name] = row
    observed = set()
    for p in root.rglob("*"):
        if is_link_or_reparse(p):
            raise HandoffError("Unexpected linked/reparse item in bundle")
        if p.is_file():
            name = p.relative_to(root).as_posix()
            if name != "BUNDLE_MANIFEST.json":
                observed.add(name)
    if observed != set(rows):
        raise HandoffError("Unmanifested or missing files in the bundle")
    return rows
