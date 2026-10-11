"""Developer-only original final-ZIP fixture preparation using existing generators.

The owner's kit reads the frozen files; it installs no ReportLab/Pillow. Genuine
AZW3 is produced by the exact reviewed Calibre executable under the existing
conversion/job supervisor, with provenance kept separate from shareable assets.
"""
from __future__ import annotations

import argparse
from hashlib import sha256
import importlib.metadata
import importlib.util
import json
import os
from pathlib import Path
import stat
import sys


def need(condition, message):
    if not condition:
        raise ValueError(message)


def ordinary(path, directory=False):
    path = Path(path)
    need(path.is_absolute() and path.resolve(strict=True) == path, "Ordinary absolute path required")
    for ancestor in [*reversed(path.parents), path]:
        details = ancestor.lstat()
        need(not ancestor.is_symlink() and not getattr(details, "st_file_attributes", 0) & 1024, "Reparse path rejected")
        need(stat.S_ISDIR(details.st_mode) if ancestor != path or directory else stat.S_ISREG(details.st_mode), "Ordinary path kind required")
    return path


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, ordinary(path))
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


def digest(path):
    return sha256(ordinary(path).read_bytes()).hexdigest()


def write_json(path, value):
    ordinary(path.parent, directory=True)
    with path.open("x", encoding="utf-8", newline="\n") as stream:
        json.dump(value, stream, ensure_ascii=True, allow_nan=False, indent=2)
        stream.write("\n")


def generate(repository, output, calibre, report):
    need(os.name == "nt" and sys.flags.isolated and sys.dont_write_bytecode, "Actual Windows developer Python -I -B required")
    repository = ordinary(repository, directory=True)
    output = ordinary(output, directory=True)
    need(not any(output.iterdir()) and not output.is_relative_to(repository), "Fresh external fixture directory required")
    need(not report.exists() and not report.is_relative_to(output) and not report.is_relative_to(repository), "New external raw provenance report required")
    paths = [repository / name for name in ("tests/fixtures/generate_pdf_fixtures.py", "tests/conversion/generate_ebook_fixtures.py",
             "engine/winbooksplit_conversion.py", "engine/winbooksplit_job.py", "LICENSE")]
    source_before = {str(path): digest(path) for path in paths}
    pdf = load("wbs_native_pdf_fixtures", paths[0])
    ebook = load("wbs_native_ebook_fixtures", paths[1])
    conversion = load("wbs_native_fixture_conversion", paths[2])
    calibre = ordinary(calibre)
    need(digest(calibre) == ebook.EXPECTED_CALIBRE_SHA256, "Reviewed Calibre bytes required")
    process_root = report.parent / (report.stem + "-process")
    process_root.mkdir()
    environment = dict(os.environ)
    for variable, name in (("CALIBRE_CONFIG_DIRECTORY", "cfg"), ("CALIBRE_CACHE_DIRECTORY", "cache"), ("TEMP", "tmp"), ("TMP", "tmp")):
        target = process_root / name
        target.mkdir(exist_ok=True)
        environment[variable] = str(target)
    operations = []
    def run(arguments):
        observed = []
        def factory(argv, *, cwd, env):
            job = conversion._default_process_factory(argv, cwd=cwd, env=env)
            observed.append(job)
            return job
        record = conversion._run_converter([str(value) for value in arguments], cwd=str(output), env=environment, timeout_seconds=240, process_factory=factory)
        need(len(observed) == 1 and observed[0].assigned_before_resume and observed[0]._tree_stopped,
             "Owned converter assignment and stopped descendants required")
        need(not record["stdout_truncated"] and not record["stderr_truncated"], "Complete bounded converter streams required")
        result = {**record, "job_assigned_before_resume": observed[0].assigned_before_resume,
                  "descendants_stopped": observed[0]._tree_stopped}
        operations.append(result)
        return result
    version = run([calibre, "--version"])
    need("calibre " + ebook.EXPECTED_CALIBRE_VERSION in version["stdout_tail"] and version["exit_code"] == 0, "Actual Calibre version required")
    files = {}
    for name in ("simple10", "nested12"):
        definition = pdf.FIXTURE_DEFINITIONS[name]
        target = output / definition["filename"]
        with target.open("xb") as stream:
            stream.write(pdf._pdf_bytes(definition))
        need(pdf.page_ids(target) == list(range(1, definition["pages"] + 1)), "Original PDF markers differ")
        files[name] = target
    flat = output / "flat-6.pdf"
    with flat.open("xb") as stream:
        stream.write(pdf._pdf_bytes({"title": "Original safe final-ZIP no-bookmark fixture", "pages": 6, "outline": []}))
    need(pdf.page_ids(flat) == list(range(1, 7)), "Original flat PDF markers differ")
    files["flat6"] = flat
    epub = output / "original-three-chapters.epub"
    ebook.authored_epub(epub)
    azw3 = output / "original-three-chapters.azw3"
    azw3_generation = run([calibre, epub, azw3, "--output-profile", "tablet"])
    need(azw3_generation["exit_code"] == 0 and azw3.is_file() and azw3.stat().st_size > 0, "Genuine Calibre AZW3 generation required")
    files.update(epub=epub, azw3=azw3)
    source_after = {str(path): digest(path) for path in paths}
    need(source_before == source_after, "Generator/process/license source changed")
    catalog = {"protocol": "winbooksplit.native-fixtures", "schema_version": 1, "license": "MIT", "authored_original": True,
               "remote_resources": False, "chapter_markers": ebook.CHAPTER_MARKERS, "chapter_titles": ebook.CHAPTER_TITLES,
               "calibre": {"version": ebook.EXPECTED_CALIBRE_VERSION, "sha256": digest(calibre)},
               "azw3_generation": {key: azw3_generation[key] for key in ("exit_code", "job_assigned_before_resume", "descendants_stopped")},
               "generator_sha256": {str(path.relative_to(repository)).replace("\\", "/"): source_before[str(path)] for path in paths},
               "files": {key: {"name": path.name, "sha256": digest(path), "size_bytes": path.stat().st_size} for key, path in files.items()}}
    write_json(output / "native-fixtures.json", catalog)
    write_json(report, {"protocol": "winbooksplit.native-fixture-preparation", "schema_version": 1, "success": True,
                       "python": sys.version, "python_path": sys.executable, "pypdf": importlib.metadata.version("pypdf"),
                       "reportlab": importlib.metadata.version("reportlab"), "calibre_path": str(calibre), "operations": operations,
                       "process_environment_overrides": {key: environment[key] for key in ("CALIBRE_CONFIG_DIRECTORY", "CALIBRE_CACHE_DIRECTORY", "TEMP", "TMP")},
                       "source_before": source_before, "source_after": source_after, "source_unchanged": True,
                       "catalog_sha256": digest(output / "native-fixtures.json")})
    return catalog


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repository-root", type=Path, required=True)
    parser.add_argument("--output-root", type=Path, required=True)
    parser.add_argument("--calibre-path", type=Path, required=True)
    parser.add_argument("--report", type=Path, required=True)
    args = parser.parse_args(argv)
    try:
        generate(args.repository_root, args.output_root, args.calibre_path, args.report)
        return 0
    except Exception as error:
        # Only this explicitly requested raw report may be created on failure.
        # Partial fixtures/process work remain; no uncertain cleanup is attempted.
        if not args.report.exists() and args.report.is_absolute():
            try:
                write_json(args.report, {"protocol": "winbooksplit.native-fixture-preparation", "schema_version": 1, "success": False,
                           "error": str(error), "code": getattr(error, "code", "preparation_failed"),
                           "tree_stopped": getattr(error, "tree_stopped", None), "partial_process": dict(getattr(error, "conversion", {}) or {})})
            except (OSError, ValueError, TypeError):
                pass
        print("Fixture preparation failed; preserve the authored output and raw provenance: " + str(error), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
