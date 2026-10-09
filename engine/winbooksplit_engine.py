"""WinBookSplit PDF logic with complete manual and Level 1 page coverage."""

import sys
import os
import re
import logging

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


def log_bookmark_warnings(warnings):
    for warning in warnings:
        log(f"[WARNING] {warning['code']} (outline {warning['source_order']}, depth {warning['depth']}): {warning['message']}")


def split_pdf(input_path, output_dir, mode, manual_data=None):
    try:
        reader = PdfReader(input_path)
    except Exception as e:
        log(f"[CRITICAL] Could not read PDF: {e}")
        sys.exit(1)

    total_pages = len(reader.pages)

    if mode == '1':
        try:
            plan = plan_level1(reader, total_pages)
        except BookmarkPlanError as error:
            log_bookmark_warnings(error.warnings)
            log(f"[ERROR] {error.code}: {error}")
            if error.code in {'no_bookmarks', 'no_usable_bookmarks'}:
                log("[NO_BOOKMARKS_FOUND]")
                sys.exit(55)
            sys.exit(1)
        log_bookmark_warnings(plan['warnings'])
        log(f"[*] Found {len(plan['entries'])} sections based on Level 1 bookmarks.")
        for entry in plan['entries']:
            write_slice(reader, entry['start'], entry['end'], os.path.join(output_dir, entry['filename']))

    elif mode == '2':
        target_level = 2
        if not reader.outline:
            log("[NO_BOOKMARKS_FOUND]")
            sys.exit(55)

        bookmarks_found = []
        def extract(nodes, depth=1):
            for node in nodes:
                if isinstance(node, list):
                    extract(node, depth + 1)
                else:
                    if depth == target_level:
                        try:
                            pg = reader.get_destination_page_number(node)
                            if pg is not None and pg != -1:
                                title = node.title if node.title else "Untitled"
                                bookmarks_found.append({'page': pg, 'title': title})
                        except: pass

        extract(reader.outline)

        if not bookmarks_found:
            log("[NO_BOOKMARKS_FOUND]")
            sys.exit(55)

        bookmarks_found.sort(key=lambda x: x['page'])
        unique_map = []
        seen = set()
        for b in bookmarks_found:
            if b['page'] not in seen:
                unique_map.append(b)
                seen.add(b['page'])

        log(f"[*] Found {len(unique_map)} chapters based on bookmarks.")

        for i, mark in enumerate(unique_map):
            start = mark['page']
            end = unique_map[i+1]['page'] if i+1 < len(unique_map) else total_pages
            if start >= end: continue
            safe_title = re.sub(r'[<>:""/\\|?*]', '', mark['title']).strip()[:50]
            fname = f"{i+1:02d} - {safe_title}.pdf"
            write_slice(reader, start, end, os.path.join(output_dir, fname))

    elif mode == 'manual':
        try:
            plan = plan_manual_starts(manual_data, total_pages)
        except ManualPlanError as error:
            log(f"[ERROR] {error.code}: {error}")
            sys.exit(1)
        for notice in plan['notices']:
            log(f"[NOTICE] {notice}")
        log(f"[*] Manual Split Points (Page #): {plan['starts']}")
        for i, (start_idx, end_idx) in enumerate(plan['ranges']):
            fname = f"{i+1:02d} - Section (Page {start_idx+1}-{end_idx}).pdf"
            write_slice(reader, start_idx, end_idx, os.path.join(output_dir, fname))

def write_slice(reader, start, end, out_path):
    log(f"    [Writing] {os.path.basename(out_path)}")
    writer = PdfWriter()
    for p in range(start, end):
        writer.add_page(reader.pages[p])
    with open(out_path, "wb") as f:
        writer.write(f)


def main():
    # Preserve CLI configuration while keeping callable imports free of it.
    logging.getLogger("pypdf").setLevel(logging.ERROR)
    sys.stdout.reconfigure(line_buffering=True)
    mode = sys.argv[3]
    manual_data = sys.argv[4] if len(sys.argv) > 4 else None
    split_pdf(sys.argv[1], sys.argv[2], mode, manual_data)


if __name__ == "__main__":
    raise SystemExit(main())
