"""Existing WinBookSplit PDF logic, mechanically extracted before behavior fixes."""

import sys
import os
import re
import logging

from pypdf import PdfReader, PdfWriter


def log(msg): print(msg)

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
            raw_nums = [int(x.strip()) for x in manual_data.split(',') if x.strip().isdigit()]
        except:
            log("[ERROR] Invalid number format.")
            sys.exit(1)
        raw_nums.sort()
        split_indices = []
        if 1 not in raw_nums: split_indices.append(0)
        for p in raw_nums:
            idx = p - 1
            if idx > 0 and idx < total_pages: split_indices.append(idx)
        split_indices = sorted(list(set(split_indices)))

        log(f"[*] Manual Split Points (Page #): {[x+1 for x in split_indices]}")
        for i, start_idx in enumerate(split_indices):
            end_idx = split_indices[i+1] if i + 1 < len(split_indices) else total_pages
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
