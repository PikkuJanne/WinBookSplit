"""Invented local run records for pure validator units, never native evidence."""

from copy import deepcopy
from hashlib import sha256
import json


def make(outcome, log_hash, *, log_text=None, plan=None, source_path="C:/synthetic/source.pdf",
         source_observation=None, run_id="c" * 32, unfinalized=False):
    owner = {"run_id": run_id, "kind": "console"}
    members = [".WinBookSplit-console-owner.json", "console.log"]
    length = len(log_text.encode("utf-8")) if log_text is not None else 128
    identities = {name: {"device": 1, "inode": number, "size_bytes": length if name == "console.log" else 64,
                         "sha256": log_hash if name == "console.log" else "e" * 64}
                  for number, name in enumerate(members, 1)}
    result = {"state": "unfinalized" if unfinalized else "finalized", "members": members, "owner": owner,
              "directory_identity": {"device": 1, "inode": 4},
              "file_identities": identities, "log_sha256": log_hash, "log_size_bytes": length,
              "operation_outcome": deepcopy(outcome), "confirmed_plan": deepcopy(plan),
              "run_manifest": None, "run_manifest_raw": None, "run_manifest_sha256": None}
    if log_text is not None and not any(json.loads(line[len("[INTERACTION-REPLY] "):]).get("action") == "execute"
        for line in log_text.splitlines() if line.startswith("[INTERACTION-REPLY] ")):
        result["confirmed_plan"] = None
    if unfinalized:
        return result
    frame = outcome.get("engine_result") or {}
    plan = frame.get("plan") or result["confirmed_plan"]
    observed = source_observation or {"size_bytes": 10, "sha256": "a" * 64}
    source = (plan.get("original_ebook_identity", plan.get("source_identity")) if plan else
              {"path": source_path, "size_bytes": observed["size_bytes"], "last_write_utc": "2026-10-10T00:00:00Z",
               "sha256": None, "binding": "metadata_only"})
    parser = (frame.get("diagnostic") or {}).get("parser_warnings") or {}
    manifest = {"protocol": "winbooksplit.run", "version": 1, "run_id": run_id,
        "application_version": "1.0.0", "started_utc": "2026-10-10T00:00:00Z", "finished_utc": "2026-10-10T00:00:01Z",
        "diagnostics_finalized": True, "runtime_versions": {"powershell": "5.1.26100.9444", "python": "3.14.8", "pypdf": "6.20.0", "calibre": None},
        "settings": {"mode": outcome.get("mode"), "input_kind": "pdf", "preview": False, "non_interactive": False,
            "no_pause": True, "keep_converted_pdf": False, "conversion_timeout": 1800, "process_timeout": 3600,
            "normalized_inputs": plan.get("normalized_inputs") if plan else None},
        "source_identity": deepcopy(source), "plan": deepcopy(plan), "engine_result": deepcopy(outcome.get("engine_result")), "outcome": deepcopy(outcome),
        "warnings": {"planning": deepcopy(frame.get("warnings", [])), "parser": deepcopy(parser.get("records", [])),
            "parser_suppressed_count": parser.get("suppressed_count", 0), "parser_message_truncated_count": parser.get("message_truncated_count", 0)},
        "log": {"filename": "console.log", "max_bytes": 50331648, "body_limit_bytes": 33554432, "bytes_written": length, "limit_reached": False}}
    raw = json.dumps(manifest, ensure_ascii=True, separators=(",", ":"))
    hashed = sha256(raw.encode("utf-8")).hexdigest()
    result.update(members=sorted([*members, "WinBookSplit_Run.json"]), run_manifest=manifest, run_manifest_raw=raw, run_manifest_sha256=hashed)
    identities["WinBookSplit_Run.json"] = {"device": 1, "inode": 3, "size_bytes": len(raw.encode("utf-8")), "sha256": hashed}
    return result


def reseal(receipt):
    """Rehash authored mutations so semantic guards, not stale hashes, reject."""
    raw = json.dumps(receipt["run_manifest"], ensure_ascii=True, separators=(",", ":"))
    hashed = sha256(raw.encode("utf-8")).hexdigest()
    receipt.update(run_manifest_raw=raw, run_manifest_sha256=hashed)
    receipt["file_identities"]["WinBookSplit_Run.json"].update(size_bytes=len(raw.encode("utf-8")), sha256=hashed)
