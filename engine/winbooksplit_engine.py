"""WinBookSplit PDF logic with complete manual and parent-aware bookmark coverage."""

import sys
import os
import re
import logging
import json
import stat
import uuid
import importlib.util
import time
import unicodedata
from datetime import datetime, timezone
from pathlib import Path
from collections.abc import Mapping
from dataclasses import dataclass
from hashlib import sha256
from io import BytesIO
from types import MappingProxyType

# Both entrypoints consume this shipped contract; unknown outcomes fail closed.
with Path(__file__).resolve().with_name("WinBookSplit.Outcomes.json").open(encoding="utf-8") as _contract_stream:
    _outcome_contract = json.load(_contract_stream)
if _outcome_contract.get("schema_version") != 1 or not isinstance(_outcome_contract.get("codes"), dict):
    raise ValueError("Invalid shipped outcome contract.")
EXIT_CODES = MappingProxyType(_outcome_contract["codes"])


def result_exit_code(code):
    value = EXIT_CODES.get(code)
    if type(value) is not int or value not in {0, 2, 3, 4, 5, 6, 7, 130}:
        raise ValueError("Unknown or invalid outcome code: " + str(code))
    return value


try:
    from pypdf import PdfReader, PdfWriter
    from pypdf.generic import IndirectObject, NullObject
except ImportError as error:
    if __name__ != "__main__":
        raise
    # Direct engine invocation has the same dependency outcome as preflight.
    dependency_exit = result_exit_code("dependency_missing")
    print(json.dumps({"protocol": "winbooksplit.result", "version": 1,
        "mode": sys.argv[3] if len(sys.argv) > 3 else "", "status": "error",
        "code": "dependency_missing", "message": "Cannot import pypdf: " + str(error),
        "warnings": [], "fallback_modes": [], "exit_code": dependency_exit, "written_count": 0,
        "execution": None, "diagnostic": None}, ensure_ascii=True))
    raise SystemExit(dependency_exit) from error


def log(msg): print(msg)


class ManualPlanError(ValueError):
    def __init__(self, code, message):
        super().__init__(message)
        self.code = code


MAX_FILE_PATH_UNITS = 259
MAX_DIRECTORY_PATH_UNITS = 247
MAX_COMPONENT_UNITS = 255
MAX_TITLE_UNITS = 50


def _utf16_units(value):
    """Count Windows path units, including two units for astral characters."""
    return len(value.encode("utf-16-le")) // 2


def _truncate_units(value, limit):
    used, result = 0, []
    for character in value:
        width = 2 if ord(character) > 0xffff else 1
        if used + width > limit:
            break
        used += width
        result.append(character)
    return "".join(result)


def _safe_title(title, fallback, limit):
    # Keep Unicode text without normalization; controls and surrogate code
    # points cannot form a usable Windows basename. Titles are never paths.
    cleaned = "".join(character for character in title
                      if character not in '<>:"/\\|?*'
                      and unicodedata.category(character) not in {"Cc", "Cf", "Cs"})
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" .")
    cleaned = _truncate_units(cleaned, limit).rstrip(" .")
    if not any(unicodedata.category(character)[0] not in {"P", "Z"} for character in cleaned):
        if _utf16_units(fallback) > limit:
            raise PlanError("output_path_too_long", "Choose a shorter output base; a stable section name cannot fit.")
        cleaned = fallback
    return cleaned


def _assign_filenames(entries, filename_budget=MAX_COMPONENT_UNITS):
    width = max(2, len(str(len(entries))))
    for sequence, entry in enumerate(entries, start=1):
        prefix = f"{sequence:0{width}d} - "
        title_budget = min(MAX_TITLE_UNITS, filename_budget - _utf16_units(prefix + ".pdf"))
        safe_title = _safe_title(entry["title"], f"Section {sequence}", title_budget)
        entry.update(sequence=sequence, filename=prefix + safe_title + ".pdf")


def _reserved_basename(filename):
    device = filename.split(".", 1)[0].rstrip(" ").upper()
    return device in {"CON", "PRN", "AUX", "NUL", "CONIN$", "CONOUT$"} or re.fullmatch(r"(?:COM|LPT)[1-9¹²³]", device) is not None


def plan_manual_starts(manual_data, total_pages):
    """Validate every token before normalizing into complete half-open ranges."""
    if type(total_pages) is not int or total_pages < 1:
        raise ManualPlanError("invalid_document", "The PDF must contain at least one page.")
    if not isinstance(manual_data, str):
        raise ManualPlanError("invalid_start_pages", "Enter comma-separated start pages.")

    starts = []
    upper = str(total_pages)
    for position, raw in enumerate(manual_data.split(','), start=1):
        token = raw.strip()
        if re.fullmatch(r'[0-9]+', token) is None:
            raise ManualPlanError("invalid_start_pages",
                                  f"Token {position} must be a nonempty ASCII decimal page number.")
        # Compare bounded decimal text before int(): even very long inputs and
        # valid leading-zero tokens do not depend on Python's digit limit.
        canonical = token.lstrip('0') or '0'
        if canonical == '0' or len(canonical) > len(upper) or (
                len(canonical) == len(upper) and canonical > upper):
            raise ManualPlanError("invalid_start_pages",
                                  f"Token {position} must be between 1 and {total_pages}.")
        starts.append(int(canonical))

    notices = []
    if starts != sorted(starts):
        notices.append("Start pages sorted into physical page order.")
    if len(set(starts)) != len(starts):
        notices.append("Duplicate start pages removed.")
    normalized = sorted(set(starts))
    if normalized[0] != 1:
        normalized.insert(0, 1)
        notices.append("Physical page 1 added to preserve opening pages.")
    indices = [page - 1 for page in normalized]
    ranges = list(zip(indices, indices[1:] + [total_pages]))
    if len(ranges) == 1:
        notices.append("One section; no internal split.")
    return {"starts": normalized, "ranges": ranges, "notices": notices}


MAX_OUTLINE_DEPTH = 64
MAX_OUTLINE_NODES = 10000


class BookmarkPlanError(ValueError):
    def __init__(self, code, message, warnings=None):
        super().__init__(message)
        self.code = code
        self.warnings = warnings or []


def _outline_failure(code, message, depth=None):
    return BookmarkPlanError("invalid_outline", message,
                             [{"code": code, "source_order": None, "depth": depth, "message": message}])


def _pdf_object(value):
    return value.get_object() if isinstance(value, IndirectObject) else value


def _preflight_bookmark_trees(reader):
    """Bound only outline/name-tree retrieval before pypdf's recursive parsing."""
    catalog = getattr(reader, "root_object", None)
    if catalog is None:  # Callable normalization also accepts synthetic readers.
        return {}
    if not isinstance(catalog, dict):
        raise _outline_failure("malformed_outline", "The PDF catalog is not a dictionary.")
    named = {}
    roots = []
    if "/Outlines" in catalog:
        root = _pdf_object(catalog["/Outlines"])
        if not isinstance(root, NullObject):
            if not isinstance(root, dict):
                raise _outline_failure("malformed_outline", "The outline root is not a dictionary.")
            if "/First" in root:
                roots.append((root["/First"], "outline"))
    if "/Dests" in catalog:
        roots.append((catalog["/Dests"], "names"))
    elif "/Names" in catalog:
        names = _pdf_object(catalog["/Names"])
        if isinstance(names, dict) and "/Dests" in names:
            roots.append((names["/Dests"], "names"))
    for root, kind in roots:
        pending = [(root, 1)]
        visited = set()
        count = 0
        while pending:
            value, depth = pending.pop()
            node = _pdf_object(value)
            if node is None or isinstance(node, NullObject):
                continue
            if not isinstance(node, dict):
                raise _outline_failure("malformed_outline", f"A {kind} tree node is not a dictionary.", depth)
            if id(node) in visited:
                raise _outline_failure("outline_cycle", f"A {kind} tree repeats a node or contains a cycle.", depth)
            visited.add(id(node))
            count += 1
            if depth > MAX_OUTLINE_DEPTH or count > MAX_OUTLINE_NODES:
                raise _outline_failure("outline_limit", f"The {kind} tree exceeds the traversal limit.", depth)
            if kind == "outline":
                if "/Next" in node:
                    pending.append((node["/Next"], depth))
                if "/First" in node:
                    pending.append((node["/First"], depth + 1))
            elif "/Kids" in node:
                kids = _pdf_object(node["/Kids"])
                if not isinstance(kids, list) or len(kids) > MAX_OUTLINE_NODES:
                    raise _outline_failure("malformed_outline", "Invalid named-destination child list.", depth)
                pending.extend((kid, depth + 1) for kid in reversed(kids))
            else:
                if "/Names" in node:
                    names = _pdf_object(node["/Names"])
                    if not isinstance(names, list) or len(names) % 2:
                        raise _outline_failure("malformed_outline", "Invalid named-destination pair list.", depth)
                    if len(names) > MAX_OUTLINE_NODES * 2:
                        raise _outline_failure("outline_limit", "Too many named destinations.", depth)
                    pairs = zip(names[::2], names[1::2])
                else:
                    if len(node) > MAX_OUTLINE_NODES:
                        raise _outline_failure("outline_limit", "Too many named destinations.", depth)
                    pairs = node.items()
                for key, target in pairs:
                    key = _pdf_object(key)
                    if isinstance(key, (str, bytes)):
                        if key in named:
                            raise _outline_failure("malformed_outline", "Duplicate named-destination definition.", depth)
                        named[key] = target
                        if len(named) > MAX_OUTLINE_NODES:
                            raise _outline_failure("outline_limit", "Too many named destinations.", depth)
    return named


def _raw_destination_problem(node, named_targets):
    """Do not trust pypdf's synthesized page 0 for malformed destination fits."""
    raw = getattr(node, "node", None)
    if raw is None:
        raw = node if "/A" in node or "/Dest" in node else None
    if raw is None:
        return None  # Already-normalized callable/mock destination.
    if not isinstance(raw, dict):
        return "invalid_destination", "The original destination node is not a dictionary."
    if "/A" in raw:
        action = _pdf_object(raw["/A"])
        if not isinstance(action, dict):
            return "invalid_destination", "The bookmark action is not a dictionary."
        if _pdf_object(action.get("/S")) != "/GoTo":
            return "external_destination", "The bookmark action is external or unsupported."
        target = action.get("/D")
    else:
        target = raw.get("/Dest")
    target = _pdf_object(target)
    if isinstance(target, (str, bytes)):
        if target not in named_targets:
            return "invalid_destination", "The named destination cannot be resolved."
        target = _pdf_object(named_targets[target])
    if isinstance(target, dict):
        target = _pdf_object(target.get("/D"))
    fit_types = {"/XYZ", "/Fit", "/FitH", "/FitV", "/FitR", "/FitB", "/FitBH", "/FitBV"}
    if not isinstance(target, list) or len(target) < 2:
        return "invalid_destination", "The destination array or fit type is invalid."
    # Internal PDF destinations identify a page by indirect reference. pypdf
    # also accepts integers as object IDs, which can accidentally name a page.
    if not isinstance(target[0], IndirectObject):
        return "invalid_destination", "The internal destination must reference a PDF page object."
    fit_type = _pdf_object(target[1])
    if not isinstance(fit_type, str) or fit_type not in fit_types:
        return "invalid_destination", "The destination array or fit type is invalid."
    return None


def normalize_outline(reader, outline, total_pages, *, named_targets=None):
    """Retain source order, depth and lineage with bounded iterative traversal."""
    if type(total_pages) is not int or total_pages < 1:
        raise BookmarkPlanError("invalid_document", "The PDF must contain at least one page.")
    if not isinstance(outline, list):
        raise _outline_failure("malformed_outline", "The outline must be a list.")
    records, warnings = [], []
    visited_lists, visited_nodes = {id(outline)}, set()
    stack = [{"items": iter(outline), "depth": 1, "lineage": [], "parent_id": None, "last_id": None}]
    examined = 0
    while stack:
        frame = stack[-1]
        try:
            node = next(frame["items"])
        except StopIteration:
            stack.pop()
            continue
        examined += 1
        if examined > MAX_OUTLINE_NODES:
            raise _outline_failure("outline_limit", "The outline exceeds the traversal limit.", frame["depth"])
        if isinstance(node, list):
            if id(node) in visited_lists:
                raise _outline_failure("outline_cycle", "An outline list repeats or contains a cycle.", frame["depth"])
            if frame["last_id"] is None:
                raise _outline_failure("malformed_outline", "An outline child list has no preceding parent.", frame["depth"])
            depth = frame["depth"] + 1
            if depth > MAX_OUTLINE_DEPTH:
                raise _outline_failure("outline_limit", "The outline exceeds the depth limit.", depth)
            visited_lists.add(id(node))
            stack.append({"items": iter(node), "depth": depth,
                          "lineage": frame["lineage"] + [frame["last_id"]],
                          "parent_id": frame["last_id"], "last_id": None})
            continue
        order = len(records) + 1
        record = {"id": f"bookmark-{order}", "source_order": order, "depth": frame["depth"],
                  "parent_id": frame["parent_id"], "lineage": list(frame["lineage"]),
                  "title": "Untitled", "page": None}
        records.append(record)
        frame["last_id"] = record["id"]

        def warn(code, message):
            warnings.append({"code": code, "source_order": order, "depth": frame["depth"], "message": message})

        if not isinstance(node, dict):
            warn("invalid_destination", "The outline entry is not a destination dictionary.")
            continue
        if id(node) in visited_nodes:
            raise _outline_failure("outline_cycle", "A destination is reused with contradictory source lineage.", frame["depth"])
        visited_nodes.add(id(node))
        try:
            title = getattr(node, "title", node.get("/Title", "Untitled"))
            if isinstance(title, str) and title.strip():
                record["title"] = title
            problem = _raw_destination_problem(node, named_targets or {})
            if problem:
                warn(*problem)
                continue
            page = reader.get_destination_page_number(node)
            if not isinstance(page, int) or isinstance(page, bool) or not 0 <= page < total_pages:
                warn("invalid_destination", f"The destination must be an integer page in 0..{total_pages - 1}.")
                continue
            record["page"] = int(page)
        except Exception as error:
            warn("destination_error", f"Could not resolve the destination: {type(error).__name__}: {error}")
    return {"bookmarks": records, "warnings": warnings}


def normalize_bookmarks(reader, total_pages):
    try:
        named = _preflight_bookmark_trees(reader)
        outline = reader.outline
        return normalize_outline(reader, outline, total_pages, named_targets=named)
    except BookmarkPlanError:
        raise
    except Exception as error:
        raise _outline_failure("outline_read_error", f"Could not read the outline: {type(error).__name__}: {error}") from error


def plan_level1(reader, total_pages):
    normalized = normalize_bookmarks(reader, total_pages)
    records, warnings = normalized["bookmarks"], normalized["warnings"]
    if not records:
        raise BookmarkPlanError("no_bookmarks", "The PDF has no bookmarks.", warnings)
    parents = []
    first_at_page = {}
    for record in records:
        if record["depth"] != 1 or record["page"] is None:
            continue
        if record["page"] in first_at_page:
            record["alias_of"] = first_at_page[record["page"]]["id"]
            warnings.append({"code": "duplicate_destination", "source_order": record["source_order"], "depth": 1,
                             "message": f"Ignored duplicate destination; first bookmark is {record['alias_of']}."})
        else:
            first_at_page[record["page"]] = record
            parents.append(record)
    if not parents:
        raise BookmarkPlanError("no_usable_bookmarks", "The PDF has no usable Level 1 bookmarks.", warnings)
    ordered = sorted(parents, key=lambda record: record["page"])
    if parents != ordered:
        warnings.append({"code": "outline_reordered", "source_order": None, "depth": 1,
                         "message": "Bookmarks sorted into physical page order."})
    entries = []
    if ordered[0]["page"] > 0:
        entries.append({"title": "Front matter", "start": 0, "end": ordered[0]["page"],
                        "parent_id": None, "reason": "front_matter", "warnings": []})
    for index, record in enumerate(ordered):
        end = ordered[index + 1]["page"] if index + 1 < len(ordered) else total_pages
        entries.append({"title": record["title"], "start": record["page"], "end": end,
                        "parent_id": None, "bookmark_id": record["id"], "reason": "bookmark", "warnings": []})
    _assign_filenames(entries)
    return {"mode": "1", "total_pages": total_pages, "entries": entries,
            "ranges": [(entry["start"], entry["end"]) for entry in entries],
            "bookmarks": records, "warnings": warnings}


def plan_level2(reader, total_pages):
    """Partition each retained Level 1 parent using only its own direct children."""
    parent_plan = plan_level1(reader, total_pages)
    records, warnings = parent_plan["bookmarks"], parent_plan["warnings"]
    parents = [entry for entry in parent_plan["entries"] if entry["reason"] == "bookmark"]
    selected_ids = {entry["bookmark_id"] for entry in parents}
    children_by_parent = {}
    parents_with_descendants = set()
    for record in records:
        if record["lineage"]:
            parents_with_descendants.add(record["lineage"][0])
        if record["depth"] == 2:
            children_by_parent.setdefault(record["parent_id"], []).append(record)

    for record in records:
        if record["depth"] != 1 or record["id"] in selected_ids or record["id"] not in parents_with_descendants:
            continue
        reason = "duplicate_parent_subtree" if "alias_of" in record else "invalid_parent_subtree"
        record["subtree_ignored_reason"] = reason
        warnings.append({"code": reason, "source_order": record["source_order"], "depth": 1,
                         "parent_id": record["id"],
                         "message": "Ignored this parent's subtree because its start is an alias or unusable."})

    entries = [dict(entry) for entry in parent_plan["entries"] if entry["reason"] == "front_matter"]
    usable_children = 0
    for parent in parents:
        parent_id = parent["bookmark_id"]
        first_at_page, children = {}, []
        for child in children_by_parent.get(parent_id, []):
            if child["page"] is None:
                continue  # Normalization already records the resolution diagnostic.
            if not parent["start"] <= child["page"] < parent["end"]:
                warnings.append({"code": "child_outside_parent", "source_order": child["source_order"],
                                 "depth": 2, "parent_id": parent_id,
                                 "message": f"Ignored child outside parent pages {parent['start'] + 1}..{parent['end']}."})
                continue
            if child["page"] in first_at_page:
                child["alias_of"] = first_at_page[child["page"]]["id"]
                warnings.append({"code": "duplicate_destination", "source_order": child["source_order"],
                                 "depth": 2, "parent_id": parent_id,
                                 "message": f"Ignored duplicate child destination; first bookmark is {child['alias_of']}."})
            else:
                first_at_page[child["page"]] = child
                children.append(child)
        ordered = sorted(children, key=lambda child: child["page"])
        if children != ordered:
            warnings.append({"code": "outline_reordered", "source_order": None, "depth": 2,
                             "parent_id": parent_id, "message": "Child bookmarks sorted into physical page order."})
        usable_children += len(ordered)
        if not ordered:
            entries.append({"title": parent["title"], "start": parent["start"], "end": parent["end"],
                            "parent_id": parent_id, "bookmark_id": parent_id,
                            "reason": "parent_fallback", "warnings": []})
            continue
        if ordered[0]["page"] > parent["start"]:
            entries.append({"title": parent["title"] + " - Opening pages", "start": parent["start"],
                            "end": ordered[0]["page"], "parent_id": parent_id,
                            "reason": "parent_opening", "warnings": []})
        for index, child in enumerate(ordered):
            end = ordered[index + 1]["page"] if index + 1 < len(ordered) else parent["end"]
            entries.append({"title": child["title"], "start": child["page"], "end": end,
                            "parent_id": parent_id, "bookmark_id": child["id"],
                            "reason": "bookmark", "warnings": []})
    if not usable_children:
        raise BookmarkPlanError("no_bookmarks_at_level", "The PDF has no usable direct Level 2 bookmarks.", warnings)
    _assign_filenames(entries)
    return {"mode": "2", "total_pages": total_pages, "entries": entries,
            "ranges": [(entry["start"], entry["end"]) for entry in entries],
            "bookmarks": records, "warnings": warnings}


def log_bookmark_warnings(warnings):
    for warning in warnings:
        log(f"[WARNING] {warning['code']} (outline {warning['source_order']}, depth {warning['depth']}): {warning['message']}")


class PlanError(ValueError):
    def __init__(self, code, message):
        super().__init__(message)
        self.code = code


def _freeze(value):
    """Detach planning metadata and make nested containers read-only."""
    if isinstance(value, Mapping):
        if not all(isinstance(key, str) for key in value):
            raise PlanError("invalid_plan", "Plan metadata keys must be strings.")
        return MappingProxyType({key: _freeze(item) for key, item in value.items()})
    if isinstance(value, (list, tuple)):
        return tuple(_freeze(item) for item in value)
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    raise PlanError("invalid_plan", "Plan metadata contains an unsupported mutable value.")


def validate_plan(plan):
    """Return a frozen, complete ordered partition; never trust length sums alone."""
    if not isinstance(plan, Mapping):
        raise PlanError("invalid_plan", "A split plan must be a mapping.")
    pages = plan.get("total_pages")
    if type(pages) is not int or pages < 1:
        raise PlanError("invalid_document", "The PDF must contain at least one page.")
    if not isinstance(plan.get("mode"), str) or plan["mode"] not in {"manual", "1", "2"}:
        raise PlanError("invalid_mode", "Choose manual, Level 1 or Level 2 splitting.")
    entries = plan.get("entries")
    if not isinstance(entries, (list, tuple)) or not entries:
        raise PlanError("invalid_plan", "A split plan must contain at least one section.")
    previous_end, filenames, ranges = 0, set(), []
    for sequence, entry in enumerate(entries, start=1):
        if not isinstance(entry, Mapping) or not {"sequence", "title", "start", "end", "parent_id",
                                                  "reason", "filename", "warnings"}.issubset(entry):
            raise PlanError("invalid_plan", "Every plan entry must contain the complete section metadata.")
        start, end = entry.get("start"), entry.get("end")
        if type(start) is not int or type(end) is not int or not 0 <= start < end <= pages:
            raise PlanError("invalid_plan", "Section ranges must be nonempty integer pages inside the PDF.")
        if start != previous_end:
            raise PlanError("invalid_plan", "Sections must cover every page in order without gaps or overlaps.")
        if type(entry.get("sequence")) is not int or entry["sequence"] != sequence:
            raise PlanError("invalid_plan", "Section sequences must be consecutive in physical page order.")
        filename = entry.get("filename")
        if not isinstance(filename, str) or not filename.endswith(".pdf") \
                or re.search(r'[<>:"/\\|?*\x00-\x1f]', filename) or filename in {".pdf", "..pdf"} \
                or filename != filename.strip() or filename.casefold() in filenames \
                or any(unicodedata.category(character) in {"Cc", "Cf", "Cs"} for character in filename) \
                or _utf16_units(filename) > MAX_COMPONENT_UNITS or _reserved_basename(filename):
            raise PlanError("invalid_plan", "Section filenames must be unique safe PDF basenames.")
        if not isinstance(entry.get("title"), str) or not isinstance(entry.get("reason"), str) \
                or entry.get("parent_id") is not None and not isinstance(entry["parent_id"], str) \
                or not isinstance(entry.get("warnings", ()), (list, tuple)):
            raise PlanError("invalid_plan", "Section metadata is invalid.")
        filenames.add(filename.casefold())
        ranges.append((start, end))
        previous_end = end
    if previous_end != pages:
        raise PlanError("invalid_plan", "The final section must reach the final physical page.")
    if "ranges" in plan:
        declared = plan["ranges"]
        if not isinstance(declared, (list, tuple)) or len(declared) != len(ranges) \
                or any(not isinstance(item, (list, tuple)) or len(item) != 2
                       or any(type(bound) is not int for bound in item) for item in declared) \
                or tuple(tuple(item) for item in declared) != tuple(ranges):
            raise PlanError("invalid_plan", "Declared ranges disagree with the section entries.")
    data = dict(plan)
    data.update(ranges=ranges, coverage={"complete": True, "covered_pages": pages,
                                       "section_count": len(entries)})
    return _freeze(data)


@dataclass(frozen=True)
class PreparedSplit:
    """A validated immutable plan bound to its private captured PDF reader."""
    plan: Mapping
    _reader: object
    _binding: tuple


def prepare_split(input_path, mode, manual_data=None, *, output_base=None, _pdf_bytes=None, _conversion_metadata=None):
    """Probe and plan without chapter writes; pypdf captures path inputs in memory."""
    if not isinstance(mode, str) or mode not in {"manual", "1", "2"}:
        raise PlanError("invalid_mode", "Choose manual, Level 1 or Level 2 splitting.")
    if _pdf_bytes is not None and not isinstance(_pdf_bytes, bytes):
        raise PlanError("invalid_prepared_split", "Converted PDF input must be a captured byte snapshot.")
    reader = PdfReader(input_path if _pdf_bytes is None else BytesIO(_pdf_bytes))
    pages = len(reader.pages)
    if mode == "manual":
        raw = plan_manual_starts(manual_data, pages)
        entries = [{"sequence": index, "title": f"Section (Page {start + 1}-{end})",
                    "start": start, "end": end, "parent_id": None, "reason": "manual",
                    "filename": "", "warnings": []}
                   for index, (start, end) in enumerate(raw["ranges"], start=1)]
        _assign_filenames(entries)
        data = {"mode": mode, "total_pages": pages, "entries": entries, "ranges": raw["ranges"],
                "normalized_inputs": {"starts": raw["starts"]}, "notices": raw["notices"], "warnings": []}
    else:
        raw = plan_level1(reader, pages) if mode == "1" else plan_level2(reader, pages)
        data = {**raw, "normalized_inputs": {"bookmarks": raw["bookmarks"]}, "notices": []}
    snapshot = reader.stream.getvalue()
    digest = sha256(snapshot).hexdigest()
    source = {"path": os.path.abspath(os.fspath(input_path)),
              "resolved_path": os.path.realpath(input_path), "sha256": digest,
              "size_bytes": len(snapshot), "binding": "reader_snapshot"}
    naming_source = source["path"]
    if _conversion_metadata is not None:
        original = _conversion_metadata["original_ebook_identity"]
        conversion = _conversion_metadata["conversion"]
        generated = conversion["generated_pdf_identity"]
        if generated["sha256"] != digest or generated["size_bytes"] != len(snapshot) or generated["page_count"] != pages:
            raise PlanError("conversion_output_invalid", "Converted PDF identity disagrees with the captured reader.")
        data.update(original_ebook_identity=original, conversion=conversion,
                    keep_converted_pdf=_conversion_metadata["keep_converted_pdf"])
        naming_source = original["path"]
    plan = validate_plan({**data, "source_identity": source})
    if output_base is not None:
        naming = _output_naming(output_base, naming_source)
        entries = [dict(entry) for entry in plan["entries"]]
        _assign_filenames(entries, naming["filename_budget"])
        _check_output_budget(output_base, naming["run_stem"], [entry["filename"] for entry in entries])
        plan = validate_plan({**plan, "entries": entries, "output_naming": naming})
    return PreparedSplit(plan, reader, (id(plan), id(reader), digest))



def prepare_ebook(input_path, mode, manual_data=None, *, output_base, calibre_path,
                  keep_converted_pdf=False, conversion_timeout=1800):
    """Convert in a flat owned workspace; return only a captured immutable PDF job."""
    if not isinstance(mode, str) or mode not in {"manual", "1", "2"}:
        raise PlanError("invalid_mode", "Choose manual, Level 1 or Level 2 splitting.")
    if not isinstance(keep_converted_pdf, bool):
        raise PlanError("invalid_prepared_split", "Intermediate retention must be an explicit Boolean.")
    path = Path(__file__).resolve().with_name("winbooksplit_conversion.py")
    spec = importlib.util.spec_from_file_location("_winbooksplit_conversion", path)
    module = importlib.util.module_from_spec(spec)
    # Dataclasses resolve their defining module while loading; publish this
    # exact sibling under the private name, never an import from CWD.
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    converted = module.convert_ebook(input_path, calibre_path, output_base,
                              new_run=OutputRun, record_failure=_failed_run,
                              timeout_seconds=conversion_timeout)
    _log_conversion(converted.conversion)
    metadata = {"original_ebook_identity": converted.original_source_identity,
                "conversion": converted.conversion, "keep_converted_pdf": keep_converted_pdf}
    return prepare_split(converted.generated_pdf_identity["path"], mode, manual_data,
                         output_base=output_base, _pdf_bytes=converted.pdf_bytes,
                         _conversion_metadata=metadata)


def _check_prepared(prepared):
    if not isinstance(prepared, PreparedSplit) or not isinstance(prepared.plan, MappingProxyType) \
            or not isinstance(prepared._binding, tuple) or len(prepared._binding) != 3 \
            or prepared._binding[:2] != (id(prepared.plan), id(prepared._reader)):
        raise PlanError("invalid_prepared_split", "Execute the unchanged job returned by prepare_split.")
    if sha256(prepared._reader.stream.getvalue()).hexdigest() != prepared._binding[2] \
            or len(prepared._reader.pages) != prepared.plan["total_pages"]:
        raise PlanError("source_changed", "The captured PDF reader changed; prepare a new plan.")


def preview_plan(prepared):
    """Expose the same read-only plan without writing or reopening the input."""
    _check_prepared(prepared)
    return prepared.plan


OWNER_FILENAME = ".WinBookSplit-owner.json"
MANIFEST_FILENAME = "WinBookSplit_Manifest.json"


def _check_output_budget(base, run_stem, filenames=()):
    """Check every created path before reservation, without long-path opt-in."""
    if not isinstance(run_stem, str) or not run_stem or run_stem != run_stem.strip(" .") \
            or any(character in '<>:"/\\|?*' or unicodedata.category(character) in {"Cc", "Cf", "Cs"}
                   for character in run_stem):
        raise PlanError("invalid_plan", "The output run stem must be a safe basename.")
    base = os.path.abspath(os.fspath(base))
    directories = [".WinBookSplit-stage-" + "0" * 32, ".WinBookSplit-failed-" + "0" * 32,
                   run_stem + "_20000101-000000_" + "0" * 32]
    members = [(directories[0], (OWNER_FILENAME, MANIFEST_FILENAME, *filenames)),
               (directories[1], (OWNER_FILENAME, "failure.json")),
               (directories[2], (OWNER_FILENAME, MANIFEST_FILENAME, *filenames))]
    for directory, names in members:
        path = os.path.join(base, directory)
        if _utf16_units(directory) > MAX_COMPONENT_UNITS or _utf16_units(path) > MAX_DIRECTORY_PATH_UNITS:
            raise PlanError("output_path_too_long", "Choose a shorter output base; the run directory exceeds the Windows path budget.")
        for name in names:
            if _utf16_units(name) > MAX_COMPONENT_UNITS or _utf16_units(os.path.join(path, name)) > MAX_FILE_PATH_UNITS:
                raise PlanError("output_path_too_long", "Choose a shorter output base; output or diagnostic files exceed the Windows path budget.")


def _output_naming(base, source_path, *, required_filename_units=0):
    """Freeze a stem/title budget for stage, final and diagnostic locations."""
    base = os.path.abspath(os.fspath(base))
    # Completion evidence must fit even when a short chapter title could fit.
    required = max(_utf16_units(MANIFEST_FILENAME), required_filename_units)
    stem_budget = min(64, MAX_COMPONENT_UNITS - 49,
                      MAX_DIRECTORY_PATH_UNITS - _utf16_units(base) - 1 - 49,
                      MAX_FILE_PATH_UNITS - _utf16_units(base) - 2 - required - 49)
    if stem_budget < 1:
        raise PlanError("output_path_too_long", "Choose a shorter output base; there is no room for a complete run and manifest.")
    stem = _safe_title(Path(source_path).stem, "Book", stem_budget)
    _check_output_budget(base, stem)
    stage = os.path.join(base, ".WinBookSplit-stage-" + "0" * 32)
    final = os.path.join(base, stem + "_20000101-000000_" + "0" * 32)
    budget = min(MAX_COMPONENT_UNITS, MAX_FILE_PATH_UNITS - max(_utf16_units(stage), _utf16_units(final)) - 1)
    return {"resolved_base": os.path.realpath(base), "run_stem": stem, "filename_budget": budget}


class OutputError(PlanError):
    def __init__(self, code, message, diagnostic=None):
        super().__init__(code, message)
        self.diagnostic = diagnostic


def _timestamp():
    return datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S")


def _new_run_id():
    return uuid.uuid4().hex


def _windows_output():
    # Load only the shipped sibling, never a similarly named CWD module.
    path = Path(__file__).resolve().with_name("winbooksplit_windows.py")
    spec = importlib.util.spec_from_file_location("_winbooksplit_windows", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _identity(details):
    return details.st_dev, details.st_ino


def _ordinary(path, directory=False):
    details = os.lstat(path)
    if details.st_file_attributes & stat.FILE_ATTRIBUTE_REPARSE_POINT or not (
            stat.S_ISDIR(details.st_mode) if directory else stat.S_ISREG(details.st_mode)):
        raise OutputError("output_ownership_failed", "Output ownership refuses a reparse point or unexpected file type.")
    return details


class OutputRun:
    """One flat owned stage; no recursive cleanup or manifest-directed deletion."""
    def __init__(self, base, source_path, *, run_stem=None):
        self.base = os.path.abspath(os.fspath(base))
        if run_stem is None:
            run_stem = _output_naming(self.base, source_path)["run_stem"]
        _check_output_budget(self.base, run_stem)
        timestamp = _timestamp()
        if not isinstance(timestamp, str) or re.fullmatch(r"[0-9]{8}-[0-9]{6}", timestamp) is None:
            raise OutputError("output_ownership_failed", "Invalid output timestamp.")
        self.stage = None
        self.files, self.file_guards, self.guards = {}, {}, []
        self.stage_guard = None
        self.run_id = None
        self.marker_bytes = None
        self.windows = _windows_output()
        try:
            # Hold each ancestor against rename/reparse mutation while operating.
            chain = list(reversed(Path(self.base).parents)) + [Path(self.base)]
            for ancestor in chain:
                guard = self.windows.DirectoryGuard(str(ancestor))
                self.guards.append(guard)
                _ordinary(ancestor, directory=True)
            self.base_guard = self.guards[-1]
            for _ in range(100):
                self.run_id = _new_run_id()
                if re.fullmatch(r"[0-9a-f]{32}", self.run_id) is None:
                    raise OutputError("output_ownership_failed", "Invalid run identity.")
                stage = os.path.join(self.base, ".WinBookSplit-stage-" + self.run_id)
                try:
                    os.mkdir(stage)
                except FileExistsError:
                    continue
                self.stage = stage
                self.stage_identity = _identity(_ordinary(stage, directory=True))
                self.stage_guard = self.windows.DirectoryGuard(stage, delete_access=True)
                if _identity(_ordinary(stage, directory=True)) != self.stage_identity:
                    raise OutputError("output_ownership_failed", "The reserved stage was replaced.")
                # Keep the pinned stage/nonempty base invariant while allowing
                # other already prepared runs to publish into the same base.
                self.base_guard.allow_publication(self.stage_guard)
                break
            else:
                raise OutputError("output_exists", "Cannot reserve a unique output run.")
            self.final_prefix = run_stem + "_" + timestamp + "_"
            self.marker_bytes = serialize_result({"schema_version": 1, "kind": "run", "run_id": self.run_id}).encode("utf-8")
            self.write_owned(OWNER_FILENAME, self.marker_bytes)
        except (Exception, KeyboardInterrupt) as error:
            failure = OutputError("output_cancelled", "Output preparation was cancelled.") if isinstance(error, KeyboardInterrupt) else \
                error if isinstance(error, OutputError) else OutputError(
                    "output_base_invalid", "Cannot reserve a safe output base: " + str(error))
            diagnostic = _failed_run(self, failure, initializing=True) if self.stage is not None else None
            try:
                self.close()
            except Exception as close_error:
                if diagnostic is not None:
                    diagnostic["close_error"] = str(close_error)[:2048]
                failure.args = (str(failure) + "; " + str(close_error)[:2048],)
            raise OutputError(failure.code, str(failure), diagnostic) from error

    def record_created(self, path, stream):
        path = os.path.abspath(os.fspath(path))
        if os.path.dirname(path) != self.stage or path in self.files:
            raise OutputError("output_ownership_failed", "The created file is outside this run's stage.")
        self.files[path] = _identity(os.fstat(stream.fileno()))

    def seal_file(self, path):
        path = os.path.abspath(os.fspath(path))
        if path not in self.files:
            raise OutputError("output_ownership_failed", "The writer did not register its exclusive file creation.")
        if path not in self.file_guards:
            guard = self.windows.FileGuard(path)
            try:
                if _identity(_ordinary(path)) != self.files[path]:
                    raise OutputError("output_ownership_failed", "An owned file was replaced.")
            except Exception:
                guard.close()
                raise
            self.file_guards[path] = guard

    def write_owned(self, name, data):
        # Compatible metadata handles can mutate an empty directory's reparse
        # attributes. Reject an observed change before creating the owner file;
        # sharing flags are not an atomic check/create or account sandbox.
        try:
            for guard in self.guards:
                guard.assert_unchanged()
            self.stage_guard.assert_unchanged()
            if _identity(_ordinary(self.stage, directory=True)) != self.stage_identity:
                raise OutputError("output_ownership_failed", "The stage changed before owned file creation.")
        except OSError as error:
            raise OutputError("output_ownership_failed", "Output directory changed before file creation: " + str(error)) from error
        path = os.path.join(self.stage, name)
        with open(path, "xb") as stream:
            self.record_created(path, stream)
            stream.write(data)
            stream.flush()
            os.fsync(stream.fileno())
        self.seal_file(path)

    def assert_owned(self, check_marker=True):
        for guard in self.guards:
            guard.assert_unchanged()
        self.stage_guard.assert_unchanged()
        if _identity(_ordinary(self.stage, directory=True)) != self.stage_identity:
            raise OutputError("output_ownership_failed", "The owned stage changed.")
        if {os.path.join(self.stage, name) for name in os.listdir(self.stage)} != set(self.files):
            raise OutputError("output_ownership_failed", "The stage contains unexpected members; it is preserved.")
        for path in self.files:
            self.seal_file(path)
            self.file_guards[path].assert_unchanged()
            if _identity(_ordinary(path)) != self.files[path]:
                raise OutputError("output_ownership_failed", "An owned file changed identity.")
        if check_marker and Path(self.stage, OWNER_FILENAME).read_bytes() != self.marker_bytes:
            raise OutputError("output_ownership_failed", "The run ownership marker changed.")

    def cleanup(self, check_marker=True):
        # Only constructor failure can use the created identities before the
        # ownership marker has finished writing. The ledger still must be exact.
        self.assert_owned(check_marker=check_marker)
        # Acquire DELETE handles and recheck identity after acquisition; never
        # unlink a path between a check and a potentially foreign replacement.
        for guard in self.file_guards.values():
            guard.close()
        self.file_guards.clear()
        delete_guards = []
        try:
            for path, identity in self.files.items():
                guard = self.windows.FileGuard(path, delete_access=True)
                delete_guards.append(guard)
                if _identity(_ordinary(path)) != identity:
                    raise OutputError("output_ownership_failed", "Cleanup refuses a replaced file.")
            if {os.path.join(self.stage, name) for name in os.listdir(self.stage)} != set(self.files):
                raise OutputError("output_ownership_failed", "Cleanup refuses unexpected members.")
            for guard in delete_guards:
                guard.delete()
        finally:
            for guard in delete_guards:
                guard.close()
        self.stage_guard.delete()
        self.stage_guard.close()
        self.stage_guard = None
        return not os.path.lexists(self.stage)

    def remove_owned(self, name):
        self.assert_owned()
        path = os.path.join(self.stage, name)
        self.file_guards.pop(path).close()
        guard = self.windows.FileGuard(path, delete_access=True)
        try:
            if _identity(_ordinary(path)) != self.files[path]:
                raise OutputError("output_ownership_failed", "Refusing to replace changed completion evidence.")
            guard.delete()
        finally:
            guard.close()
        del self.files[path]

    def check_publication(self):
        # Windows cannot rename a directory while protective child handles are
        # open. Retain the directory/ancestor guards, release only child locks,
        # then check the exact closed-file bytes and identities once more.
        # This is not a sandbox against arbitrary same-account file mutation.
        for guard in self.file_guards.values():
            guard.close()
        self.file_guards.clear()
        for guard in self.guards:
            guard.assert_unchanged()
        self.stage_guard.assert_unchanged()
        if _identity(_ordinary(self.stage, directory=True)) != self.stage_identity or \
                {os.path.join(self.stage, name) for name in os.listdir(self.stage)} != set(self.files):
            raise OutputError("output_ownership_failed", "Publication refuses changed stage members.")
        for path, identity in self.files.items():
            if _identity(_ordinary(path)) != identity:
                raise OutputError("output_ownership_failed", "Publication refuses a replaced file.")
        if Path(self.stage, OWNER_FILENAME).read_bytes() != self.marker_bytes:
            raise OutputError("output_ownership_failed", "Publication refuses changed ownership.")
        manifest = json.loads(Path(self.stage, MANIFEST_FILENAME).read_text(encoding="utf-8"))
        if manifest != self.publication_manifest:
            raise OutputError("output_ownership_failed", "Publication refuses changed completion evidence.")
        verified_pdfs = list(manifest["outputs"])
        if manifest.get("retained_intermediate") is not None:
            verified_pdfs.append(manifest["retained_intermediate"])
        for entry in verified_pdfs:
            data = Path(self.stage, entry["filename"]).read_bytes()
            if len(data) != entry["size_bytes"] or sha256(data).hexdigest() != entry["sha256"]:
                raise OutputError("output_validation_failed", "Publication refuses changed PDF bytes.")

    def close(self):
        guards = list(self.file_guards.values())
        self.file_guards.clear()
        if self.stage_guard is not None:
            guards.append(self.stage_guard)
            self.stage_guard = None
        guards.extend(reversed(self.guards))
        self.guards.clear()
        errors = []
        for guard in guards:
            try:
                guard.close()
            except Exception as error:
                errors.append(str(error)[:512])
        if errors:
            raise OutputError("output_handle_close_failed", "Cannot close output handles: " + "; ".join(errors)[:2048])


def _validate_output(path, entry):
    try:
        reader = PdfReader(path)
        if len(reader.pages) != entry["end"] - entry["start"] or len(reader.pages) < 1:
            raise ValueError("The staged PDF page count disagrees with its planned range.")
        # Force page objects to be resolved before publication.
        for page in reader.pages:
            page.get_object()
        data = Path(path).read_bytes()
        return {"page_count": len(reader.pages), "sha256": sha256(data).hexdigest(), "size_bytes": len(data)}
    except Exception as error:
        raise OutputError("output_validation_failed", "Cannot validate staged PDF: " + str(error)) from error


def _write_manifest(run, manifest):
    run.write_owned(MANIFEST_FILENAME, (serialize_result(manifest) + "\n").encode("utf-8"))


def _publish_run(run, final_name):
    run.base_guard.allow_publication(run.stage_guard)
    run.check_publication()
    run.stage_guard.rename_to(run.base_guard, final_name)


def _failed_run(run, error, initializing=False):
    diagnostic = {"run_id": run.run_id, "cleanup_complete": False, "retained_staging": run.stage,
                  "record_path": None, "cleanup_error": None}
    try:
        run.base_guard.restrict_writes(run.stage_guard)
        diagnostic["cleanup_complete"] = run.cleanup(check_marker=not initializing)
        if diagnostic["cleanup_complete"]:
            diagnostic["retained_staging"] = None
    except Exception as cleanup_error:
        diagnostic["cleanup_error"] = str(cleanup_error)[:2048]
    # A distinct diagnostic folder can never be confused with completed chapters.
    directory = os.path.join(run.base, ".WinBookSplit-failed-" + run.run_id)
    try:
        for guard in run.guards:
            guard.assert_unchanged()
        os.mkdir(directory)
        identity = _identity(_ordinary(directory, directory=True))
        guard = run.windows.DirectoryGuard(directory, delete_access=True)
        try:
            if _identity(_ordinary(directory, directory=True)) != identity:
                raise OutputError("output_ownership_failed", "The diagnostic reservation was replaced.")
            marker = {"schema_version": 1, "kind": "failed", "run_id": run.run_id}
            record = {"schema_version": 1, "status": "failed", "code": getattr(error, "code", "output_write_failed"),
                      "message": str(error)[:2048], **diagnostic,
                      "record_path": os.path.join(directory, "failure.json")}
            if getattr(error, "conversion", None):
                # Converter streams are already bounded to 64 KiB per tail.
                record["conversion"] = error.conversion
            for name, value in ((OWNER_FILENAME, marker), ("failure.json", record)):
                for ancestor_guard in run.guards:
                    ancestor_guard.assert_unchanged()
                guard.assert_unchanged()
                with open(os.path.join(directory, name), "xb") as stream:
                    stream.write((serialize_result(value) + "\n").encode("utf-8"))
            diagnostic["record_path"] = os.path.join(directory, "failure.json")
        finally:
            guard.close()
    except Exception as record_error:
        diagnostic["cleanup_error"] = ((diagnostic["cleanup_error"] or "") +
                                       "; failure record: " + str(record_error))[:2048]
    return diagnostic


def execute_split(prepared, output_dir):
    """Stage the unchanged plan; publish only a complete validated unique run."""
    _check_prepared(prepared)
    plan = prepared.plan
    naming = plan.get("output_naming")
    if naming is None:
        naming = _output_naming(output_dir, plan["source_identity"]["path"],
                                required_filename_units=max(_utf16_units(entry["filename"]) for entry in plan["entries"]))
    elif os.path.normcase(os.path.realpath(os.path.abspath(os.fspath(output_dir)))) != os.path.normcase(naming["resolved_base"]):
        raise PlanError("output_destination_changed", "Execute at the output base used for this preview, or prepare a new plan.")
    _check_output_budget(output_dir, naming["run_stem"], [entry["filename"] for entry in plan["entries"]])
    run = OutputRun(output_dir, plan["source_identity"]["path"], run_stem=naming["run_stem"])
    published, close_error = False, None
    try:
        retained_intermediate = None
        if plan.get("keep_converted_pdf", False):
            data = prepared._reader.stream.getvalue()
            name = "WinBookSplit_Converted.pdf"
            run.write_owned(name, data)
            path = os.path.join(run.stage, name)
            retained_intermediate = {"filename": name, **_validate_output(path, {"start": 0, "end": plan["total_pages"]})}
            if retained_intermediate["sha256"] != plan["source_identity"]["sha256"] or retained_intermediate["size_bytes"] != len(data):
                raise OutputError("output_validation_failed", "Retained converted PDF disagrees with the captured reader.")
        outputs = []
        for entry in plan["entries"]:
            run.assert_owned()
            path = os.path.join(run.stage, entry["filename"])
            write_slice(prepared._reader, entry["start"], entry["end"], path, on_created=run.record_created)
            run.seal_file(path)
            outputs.append({**entry, **_validate_output(path, entry)})
        if not outputs or len(outputs) != len(plan["entries"]):
            raise OutputError("invalid_execution", "The writer produced no complete split.")
        # A final-name collision changes only the random suffix and completion
        # evidence; it never replaces the existing directory.
        for attempt in range(100):
            suffix = run.run_id if attempt == 0 else _new_run_id()
            final_name = run.final_prefix + suffix
            final_directory = os.path.join(run.base, final_name)
            manifest = {"schema_version": 1, "status": "complete", "run_id": run.run_id,
                        "final_directory": final_directory, "mode": plan["mode"],
                        "total_pages": plan["total_pages"], "source_identity": plan["source_identity"],
                        "coverage": plan["coverage"], "written_count": len(outputs), "outputs": outputs}
            if "conversion" in plan:
                manifest.update(original_ebook_identity=plan["original_ebook_identity"],
                                conversion=plan["conversion"], retained_intermediate=retained_intermediate)
            try:
                if attempt:
                    run.remove_owned(MANIFEST_FILENAME)
                _write_manifest(run, manifest)
            except OutputError:
                raise
            except Exception as error:
                raise OutputError("output_manifest_failed", "Cannot write completion manifest: " + str(error)) from error
            run.assert_owned()
            result = _freeze({**manifest, "manifest_filename": MANIFEST_FILENAME, "manifest": manifest})
            run.publication_manifest = json.loads(serialize_result(manifest))
            try:
                # A concurrently starting/cleaning run may briefly hold a
                # stricter base handle. Retry that native sharing conflict only;
                # ownership/content validation still runs on each attempt.
                for retry in range(100):
                    try:
                        _publish_run(run, final_name)
                        break
                    except OSError as sharing_error:
                        if getattr(sharing_error, "winerror", None) != 32 or retry == 99:
                            raise
                        time.sleep(0.01)
            except FileExistsError:
                continue
            except OutputError:
                raise
            except Exception as error:
                raise OutputError("output_publish_failed", "Cannot publish the completed run: " + str(error)) from error
            # Publication is the commit point: no subsequent filesystem writes.
            published = True
            break
        else:
            raise OutputError("output_exists", "Cannot publish to a unique final directory.")
    except (Exception, KeyboardInterrupt) as error:
        if isinstance(error, KeyboardInterrupt):
            failure = OutputError("output_cancelled", "Output writing was cancelled.")
        else:
            failure = error if isinstance(error, PlanError) else OutputError("output_write_failed", str(error))
        diagnostic = _failed_run(run, failure)
        if isinstance(error, KeyboardInterrupt):
            raise OutputError(failure.code, str(failure), diagnostic) from error
        if isinstance(error, PlanError):
            raise OutputError(error.code, str(error), diagnostic) from error
        raise OutputError("output_write_failed", str(error), diagnostic) from error
    finally:
        pending_error = sys.exception()
        try:
            run.close()
        except (Exception, KeyboardInterrupt) as error:
            if not published:
                if pending_error is None:
                    raise
                if isinstance(getattr(pending_error, "diagnostic", None), dict):
                    pending_error.diagnostic["close_error"] = str(error)[:2048]
                pending_error.args = (str(pending_error) + "; " + str(error)[:2048],)
            close_error = (str(error) or "Output handle finalization was interrupted.")[:2048]
    if close_error is not None:
        # Keep the completed result and disclose the post-commit close problem;
        # never relabel already published chapters as a failed transaction.
        return _freeze({**result, "post_publication_warnings": [
            {"code": "output_handle_close_failed", "message": close_error}]})
    return result


def _split_result(mode, status, code, message, warnings=(), execution=None, diagnostic=None):
    """One explicit outcome; choices are data and never execute a fallback."""
    fallback = ()
    if status == "no_plan":
        fallback = ("1", "manual") if code == "no_bookmarks_at_level" and mode == "2" else ("manual",)
    exit_code = result_exit_code(code)
    primary_code = diagnostic.get("primary_code") if isinstance(diagnostic, Mapping) else None
    if code == "conversion_cleanup_failed" and primary_code in {"conversion_cancelled", "conversion_timeout"}:
        exit_code = 130
    if exit_code == 130:
        status = "timeout" if code == "conversion_timeout" or primary_code == "conversion_timeout" else "cancelled"
    return _freeze({"protocol": "winbooksplit.result", "version": 1, "mode": mode,
                    "status": status, "code": code, "message": message, "warnings": warnings,
                    "fallback_modes": fallback, "exit_code": exit_code,
                    "written_count": execution["written_count"] if execution is not None else 0,
                    "execution": execution, "diagnostic": diagnostic})


def _planning_failure(mode, error):
    warnings = getattr(error, "warnings", ())
    if isinstance(error, FileNotFoundError):
        status, code = "invalid_input", "input_invalid"
    elif isinstance(error, BookmarkPlanError) and error.code in {
            "no_bookmarks", "no_usable_bookmarks", "no_bookmarks_at_level"}:
        status, code = "no_plan", error.code
    elif isinstance(error, (ManualPlanError, BookmarkPlanError, PlanError)):
        code = error.code
        status = "invalid_input" if code in {"invalid_document", "invalid_outline",
                                             "invalid_start_pages", "invalid_mode"} else "error"
    else:
        status, code = "read_error", "unreadable_document"
    return _split_result(mode, status, code, str(error), warnings)


def _log_conversion(conversion):
    """Keep bounded converter diagnostics visible without introducing result frames."""
    if not conversion:
        return
    log(f"[CONVERSION] Calibre: {conversion.get('converter_path', '')}; profile: tablet; exit: {conversion.get('exit_code')}")
    for stream in ("stdout", "stderr"):
        tail = conversion.get(stream + "_tail", "")
        if conversion.get(stream + "_truncated", False):
            log(f"[CONVERTER {stream}] Earlier output omitted; showing final 65536 bytes.")
        for line in tail.splitlines():
            log(f"[CONVERTER {stream}] {line}")


def run_split(input_path, output_dir, mode, manual_data=None, *, calibre_path=None, keep_converted_pdf=False, conversion_timeout=1800):
    """Return a frozen diagnostic for planning and writing, without implicit retries."""
    try:
        if Path(input_path).suffix.lower() in {".epub", ".azw3"}:
            prepared = prepare_ebook(input_path, mode, manual_data, output_base=output_dir,
                                     calibre_path=calibre_path, keep_converted_pdf=keep_converted_pdf,
                                     conversion_timeout=conversion_timeout)
        else:
            if keep_converted_pdf:
                return _split_result(mode, "invalid_input", "invalid_arguments",
                                     "KeepConvertedPdf applies only to EPUB or AZW3 conversion.")
            prepared = prepare_split(input_path, mode, manual_data, output_base=output_dir)
        plan = preview_plan(prepared)
    except KeyboardInterrupt:
        return _split_result(mode, "cancelled", "processing_cancelled", "PDF preparation was cancelled.")
    except Exception as error:
        if getattr(error, "code", "").startswith("conversion_") or getattr(error, "code", "") == "converter_not_found":
            conversion = getattr(error, "conversion", None)
            _log_conversion(conversion)
            diagnostic = getattr(error, "diagnostic", None)
            if conversion:
                diagnostic = {**(diagnostic or {}), "conversion": conversion}
            return _split_result(mode, "error", error.code, str(error), diagnostic=diagnostic)
        return _planning_failure(mode, error)
    try:
        log_bookmark_warnings(plan["warnings"])
        if mode == "manual":
            for notice in plan["notices"]:
                log(f"[NOTICE] {notice}")
            log(f"[*] Manual Split Points (Page #): {list(plan['normalized_inputs']['starts'])}")
        else:
            log(f"[*] Found {len(plan['entries'])} sections based on Level {mode} bookmarks.")
        execution = execute_split(prepared, output_dir)
        if execution["written_count"] < 1 or not execution["outputs"]:
            raise PlanError("invalid_execution", "The writer produced no sections.")
        if execution.get("post_publication_warnings"):
            return _split_result(mode, "incomplete", "output_handle_close_failed",
                "Chapter PDFs were published, but output handles could not be finalized; the completed folder is retained.",
                (*plan["warnings"], *execution["post_publication_warnings"]), execution)
        return _split_result(mode, "success", "split_complete", "PDF split completed.",
                             (*plan["warnings"], *execution.get("post_publication_warnings", ())), execution)
    except KeyboardInterrupt:
        return _split_result(mode, "cancelled", "processing_cancelled", "PDF processing was cancelled.", plan["warnings"])
    except PlanError as error:
        return _split_result(mode, "write_error" if error.code == "output_write_failed" else "error",
                             error.code, str(error), plan["warnings"], diagnostic=getattr(error, "diagnostic", None))
    except Exception as error:
        return _split_result(mode, "write_error", "output_write_failed", str(error), plan["warnings"])


def serialize_result(result):
    """JSON cannot encode frozen mappings directly; detach only for serialization."""
    def thaw(value):
        if isinstance(value, Mapping):
            return {key: thaw(item) for key, item in value.items()}
        if isinstance(value, (list, tuple)):
            return [thaw(item) for item in value]
        return value
    return json.dumps(thaw(result), ensure_ascii=True, separators=(",", ":"))


def split_pdf(input_path, output_dir, mode, manual_data=None, **conversion_options):
    result = run_split(input_path, output_dir, mode, manual_data, **conversion_options)
    if result["status"] != "success":
        log_bookmark_warnings(result["warnings"])
        log(f"[ERROR] {result['code']}: {result['message']}")
    log(serialize_result(result))
    if result["exit_code"]:
        raise SystemExit(result["exit_code"])
    return result["execution"]

def write_slice(reader, start, end, out_path, on_created=None):
    log(f"    [Writing] {os.path.basename(out_path)}")
    writer = PdfWriter()
    for p in range(start, end):
        writer.add_page(reader.pages[p])
    with open(out_path, "xb") as f:
        if on_created is not None:
            on_created(out_path, f)
        writer.write(f)


def main():
    # Preserve CLI configuration while keeping callable imports free of it.
    logging.getLogger("pypdf").setLevel(logging.ERROR)
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace", line_buffering=True)
    sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")
    mode = sys.argv[3] if len(sys.argv) > 3 else ""
    try:
        if len(sys.argv) < 4:
            raise ValueError("Provide input, output and a split mode.")
        # Keep the existing fourth positional manual token literal. New options
        # follow it, so invalid manual text cannot turn into a conversion switch.
        manual_data = sys.argv[4] if len(sys.argv) > 4 else None
        extras = sys.argv[5:]
        options, index = {}, 0
        while index < len(extras):
            option = extras[index]
            if option == "--keep-converted-pdf" and "keep_converted_pdf" not in options:
                options["keep_converted_pdf"] = True
                index += 1
            elif option in {"--calibre-path", "--conversion-timeout"} and index + 1 < len(extras):
                key = "calibre_path" if option == "--calibre-path" else "conversion_timeout"
                if key in options:
                    raise ValueError("Repeated conversion option.")
                value = extras[index + 1]
                if key == "conversion_timeout":
                    if not re.fullmatch(r"[0-9]+", value) or not 1 <= int(value) <= 86400:
                        raise ValueError("Conversion timeout must be 1 through 86400 seconds.")
                    value = int(value)
                options[key] = value
                index += 2
            else:
                raise ValueError("Unknown or incomplete conversion option.")
    except ValueError as error:
        result = _split_result(mode, "invalid_input", "invalid_arguments", str(error))
        log(serialize_result(result))
        return result["exit_code"]
    split_pdf(sys.argv[1], sys.argv[2], mode, manual_data, **options)


if __name__ == "__main__":
    raise SystemExit(main())
