"""WinBookSplit PDF logic with strict physical-page manual starts."""

import sys
import os
import re
import logging

from pypdf import PdfReader, PdfWriter


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


def split_pdf(input_path, output_dir, mode, manual_data=None):
    try:
        reader = PdfReader(input_path)
    except Exception as e:
        log(f"[CRITICAL] Could not read PDF: {e}")
        sys.exit(1)

    total_pages = len(reader.pages)

    if mode in ['1', '2']:
        target_level = int(mode)
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
