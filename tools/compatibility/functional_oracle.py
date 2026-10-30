"""Independent physical-page oracle for the public combined native fixture."""
from __future__ import annotations
import argparse
from hashlib import sha256
import json
from pathlib import Path
import sys
from pypdf import PdfReader


def validate(expected, source, outcome):
    if outcome["code"] != "split_complete" or outcome["exit_code"] != 0 or outcome["written_count"] != 3:
        raise ValueError("split_result_invalid")
    data = Path(source).read_bytes()
    if {"sha256": sha256(data).hexdigest(), "bytes": len(data)} != expected["source"]:
        raise ValueError("source_changed")
    result = outcome["engine_result"]
    if result["plan"]["ranges"] != expected["ranges"] or result["plan"]["total_pages"] != 6:
        raise ValueError("independent_ranges_differ")
    warnings = [row for row in result["warnings"] if row["code"] == "opening_destination_ignored"]
    if len(warnings) != 1:
        raise ValueError("opening_warning_count_invalid")
    outputs = result["execution"]["outputs"]
    if len(outputs) != 3:
        raise ValueError("output_count_invalid")
    root = Path(outcome["final_directory"]).resolve(strict=True)
    pages, identities = [], []
    for row, (start, end) in zip(outputs, expected["ranges"], strict=True):
        filename = row["filename"]
        if Path(filename).name != filename:
            raise ValueError("output_filename_invalid")
        path = root / filename
        reader = PdfReader(path)
        if reader.is_encrypted or "/OpenAction" in reader.root_object or len(reader.pages) != end - start:
            raise ValueError("chapter_structure_invalid")
        pages.extend(sha256(page.get_contents().get_data()).hexdigest() for page in reader.pages)
        identities.append({"filename": filename, "sha256": sha256(path.read_bytes()).hexdigest()})
    if pages != expected["pages"]:
        raise ValueError("physical_page_content_or_order_differ")
    return {"status": "PASS", "pages": 6, "chapters": 3, "ranges": expected["ranges"],
            "source_unchanged": True, "output_identities": identities}


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--expected", required=True, type=Path)
    parser.add_argument("--source", required=True, type=Path)
    parser.add_argument("--outcome", required=True, type=Path)
    args = parser.parse_args()
    try:
        print(json.dumps(validate(json.loads(args.expected.read_bytes()), args.source,
                                  json.loads(args.outcome.read_bytes()))))
    except (ValueError, KeyError, OSError) as error:
        print(json.dumps({"status": "FAIL", "reason": str(error)}))
        sys.exit(1)
