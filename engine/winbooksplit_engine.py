"""WinBookSplit PDF logic with complete manual and parent-aware bookmark coverage."""

import sys
import os
import re
import logging
import json
from collections.abc import Mapping
from dataclasses import dataclass
from hashlib import sha256
from types import MappingProxyType

from pypdf import PdfReader, PdfWriter
from pypdf.generic import IndirectObject, NullObject


def log(msg): print(msg)


class ManualPlanError(ValueError):
    def __init__(self, code, message):
        super().__init__(message)
        self.code = code


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
    for sequence, entry in enumerate(entries, start=1):
        safe_title = re.sub(r'[<>:""/\\|?*]', '', entry["title"]).strip()[:50]
        entry.update(sequence=sequence, filename=f"{sequence:02d} - {safe_title}.pdf")
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
    for sequence, entry in enumerate(entries, start=1):
        safe_title = re.sub(r'[<>:""/\\|?*]', '', entry["title"]).strip()[:50]
        entry.update(sequence=sequence, filename=f"{sequence:02d} - {safe_title}.pdf")
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
                or filename != filename.strip() or filename.casefold() in filenames:
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


def prepare_split(input_path, mode, manual_data=None):
    """Probe and plan without chapter writes; pypdf captures path inputs in memory."""
    if not isinstance(mode, str) or mode not in {"manual", "1", "2"}:
        raise PlanError("invalid_mode", "Choose manual, Level 1 or Level 2 splitting.")
    reader = PdfReader(input_path)
    pages = len(reader.pages)
    if mode == "manual":
        raw = plan_manual_starts(manual_data, pages)
        entries = [{"sequence": index, "title": f"Section (Page {start + 1}-{end})",
                    "start": start, "end": end, "parent_id": None, "reason": "manual",
                    "filename": f"{index:02d} - Section (Page {start + 1}-{end}).pdf", "warnings": []}
                   for index, (start, end) in enumerate(raw["ranges"], start=1)]
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
    plan = validate_plan({**data, "source_identity": source})
    return PreparedSplit(plan, reader, (id(plan), id(reader), digest))


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


def execute_split(prepared, output_dir):
    """Execute exactly the prepared entries against their captured source."""
    _check_prepared(prepared)
    plan = prepared.plan
    paths = [os.path.join(output_dir, entry["filename"]) for entry in plan["entries"]]
    source_paths = {os.path.normcase(plan["source_identity"][key]) for key in ("path", "resolved_path")}
    for path in paths:
        if os.path.normcase(os.path.realpath(path)) in source_paths \
                or os.path.normcase(os.path.abspath(path)) in source_paths:
            raise PlanError("source_output_alias", "A planned output would replace the source PDF.")
        if os.path.lexists(path):
            raise PlanError("output_exists", "A planned output already exists; use a new output directory.")
    outputs = []
    for entry, path in zip(plan["entries"], paths):
        write_slice(prepared._reader, entry["start"], entry["end"], path)
        outputs.append({**entry, "page_count": entry["end"] - entry["start"]})
    return _freeze({"mode": plan["mode"], "total_pages": plan["total_pages"],
                    "source_identity": plan["source_identity"], "coverage": plan["coverage"],
                    "written_count": len(outputs), "outputs": outputs})


def _split_result(mode, status, code, message, warnings=(), execution=None):
    """One explicit outcome; choices are data and never execute a fallback."""
    fallback = ()
    if status == "no_plan":
        fallback = ("1", "manual") if code == "no_bookmarks_at_level" and mode == "2" else ("manual",)
    return _freeze({"protocol": "winbooksplit.result", "version": 1, "mode": mode,
                    "status": status, "code": code, "message": message, "warnings": warnings,
                    "fallback_modes": fallback, "exit_code": 0 if status == "success" else
                    55 if status == "no_plan" else 1,
                    "written_count": execution["written_count"] if execution is not None else 0,
                    "execution": execution})


def _planning_failure(mode, error):
    warnings = getattr(error, "warnings", ())
    if isinstance(error, BookmarkPlanError) and error.code in {
            "no_bookmarks", "no_usable_bookmarks", "no_bookmarks_at_level"}:
        status, code = "no_plan", error.code
    elif isinstance(error, (ManualPlanError, BookmarkPlanError, PlanError)):
        code = error.code
        status = "invalid_input" if code in {"invalid_document", "invalid_outline",
                                             "invalid_start_pages", "invalid_mode"} else "error"
    else:
        status, code = "read_error", "unreadable_document"
    return _split_result(mode, status, code, str(error), warnings)


def run_split(input_path, output_dir, mode, manual_data=None):
    """Return a frozen diagnostic for planning and writing, without implicit retries."""
    try:
        prepared = prepare_split(input_path, mode, manual_data)
        plan = preview_plan(prepared)
    except Exception as error:
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
        return _split_result(mode, "success", "split_complete", "PDF split completed.",
                             plan["warnings"], execution)
    except PlanError as error:
        return _split_result(mode, "error", error.code, str(error), plan["warnings"])
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


def split_pdf(input_path, output_dir, mode, manual_data=None):
    result = run_split(input_path, output_dir, mode, manual_data)
    if result["status"] != "success":
        log_bookmark_warnings(result["warnings"])
        log(f"[ERROR] {result['code']}: {result['message']}")
    log(serialize_result(result))
    if result["exit_code"]:
        raise SystemExit(result["exit_code"])
    return result["execution"]

def write_slice(reader, start, end, out_path):
    log(f"    [Writing] {os.path.basename(out_path)}")
    writer = PdfWriter()
    for p in range(start, end):
        writer.add_page(reader.pages[p])
    with open(out_path, "xb") as f:
        writer.write(f)


def main():
    # Preserve CLI configuration while keeping callable imports free of it.
    logging.getLogger("pypdf").setLevel(logging.ERROR)
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace", line_buffering=True)
    sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")
    if len(sys.argv) not in {4, 5}:
        mode = sys.argv[3] if len(sys.argv) > 3 else ""
        log(serialize_result(_split_result(mode, "invalid_input", "invalid_arguments",
                                          "Provide input, output, mode and at most one manual start-pages argument.")))
        return 1
    mode = sys.argv[3]
    manual_data = sys.argv[4] if len(sys.argv) > 4 else None
    split_pdf(sys.argv[1], sys.argv[2], mode, manual_data)


if __name__ == "__main__":
    raise SystemExit(main())
