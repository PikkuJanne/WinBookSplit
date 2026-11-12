"""Run real Calibre discovery, preparation, resumption and chapter checks on Windows."""

from __future__ import annotations

import argparse
from hashlib import sha256
import importlib.util
import json
import os
from pathlib import Path
import platform
import re
import runpy
import socket
import stat
import subprocess
import sys
import tempfile

from pypdf import PdfReader, PdfWriter, __version__ as pypdf_version
from pypdf.generic import NameObject, NumberObject


ROOT = Path(__file__).resolve().parents[2]
PREPARE = "Prepare the verified Calibre Portable converter for WinBookSplit now? [Y/N]"
CASES = ("existing-path", "explicit-path", "explicit-rejection", "preview-no-download", "noninteractive-no-download",
         "setup-check-no-download", "declined", "download-unavailable", "invalid-offline-artifact",
         "managed-install", "managed-reuse", "older-preserved-offline", "existing-azw3",
         "conversion-failure", "pdf-only")


def load(name, relative):
    spec = importlib.util.spec_from_file_location(name, ROOT / relative)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


support = load("wbs_managed_support", "tests/support/characterize_support.py")
launcher = load("wbs_managed_launcher", "tests/launcher/characterize_launcher.py")
fixtures = runpy.run_path(str(ROOT / "tests/conversion/generate_ebook_fixtures.py"))
oracle = runpy.run_path(str(ROOT / "tests/ebooks/managed_fixture.py"))
cli, require = support.cli, support.require


def tree_digest(root):
    """Fingerprint only the supplied application/installation tree, without following reparse points."""
    items = {}
    if root.is_dir():
        for base, directories, names in os.walk(root, followlinks=False):
            for name in directories + names:
                path = Path(base) / name
                details = path.lstat()
                relative = path.relative_to(root).as_posix()
                if details.st_file_attributes & stat.FILE_ATTRIBUTE_REPARSE_POINT:
                    items[relative] = ["reparse", str(os.readlink(path))]
                    if name in directories:
                        directories.remove(name)
                elif stat.S_ISREG(details.st_mode):
                    items[relative] = [details.st_size, cli.digest(path)]
                elif stat.S_ISDIR(details.st_mode):
                    items[relative] = ["directory"]
    encoded = json.dumps(items, sort_keys=True, separators=(",", ":")).encode()
    return {"members": len(items), "sha256": sha256(encoded).hexdigest()}


def application_receipt(app, manifest_path):
    if manifest_path is None:
        return {"files": {name: cli.digest(app / name) for name in support.validator.APPLICATION}}
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    files = {}
    for entry in manifest["files"]:
        path = app / entry["path"]
        require(path.resolve(strict=True).is_relative_to(app) and path.is_file(), "Release manifest path escaped extraction")
        value = cli.digest(path)
        require(value == entry["sha256"], "Extracted application does not match its release manifest")
        files[entry["path"]] = value
    require(set(support.validator.APPLICATION) <= files.keys(), "Release manifest lacks a required application file")
    return {"files": files, "manifest_sha256": cli.digest(manifest_path), "source_commit": manifest["source_commit"],
            "source_tree": manifest["source_tree"]}


def environment_for(work, shell, scope, candidate=None, proxy=None):
    env = support.controlled_environment(work, shell)
    env["LOCALAPPDATA"] = str(scope)
    if candidate is not None:
        env["PATH"] = str(candidate.parent) + os.pathsep + env["PATH"]
    if proxy is not None:
        env.update(HTTPS_PROXY=proxy, HTTP_PROXY=proxy, ALL_PROXY=proxy, NO_PROXY="")
    return env


def run_application(record, args, work, host, source, scope, *, base=None, candidate=None, explicit=None,
                    proxy=None, answers=None, preview=False, expected=("split_complete", 0), manual=False):
    base = base or work / ("output-" + record["id"])
    base.mkdir(exist_ok=True)
    prior = base / "prior-output.pdf"
    if not prior.exists():
        with prior.open("xb") as stream:
            stream.write(b"Original earlier output must remain unchanged.\n")
    neighbor = work / "neighbor.txt"
    original_members = {path.name for path in base.iterdir()}
    script = args.app_root / "WinBookSplit.ps1"
    command = [host, "-NoProfile", "-ExecutionPolicy", "RemoteSigned", "-File", str(script),
               "-InputFile", str(source), "-OutputDirectory", str(base),
               "-PythonPath", str(args.application_python_path), "-NoPause"]
    command += ["-Mode", "Manual", "-StartPages", "1,3"] if manual else ["-Mode", "Auto", "-BookmarkLevel", "1"]
    if answers is None and not preview:
        command.append("-NonInteractive")
    if preview:
        command.append("-Preview")
    elif source.suffix.lower() != ".pdf":
        command.append("-KeepConvertedPdf")
    if explicit is not None:
        command += ["-CalibrePath", str(explicit)]
    env = environment_for(work, host, scope, candidate, proxy)
    record.update(output_base=str(base), command=command, localappdata=str(scope),
                  path_candidate=str(candidate) if candidate else None, test_proxy=proxy)
    with support.paths.read_only_source(source):
        before = {item.name: cli.identity(item) for item in (source, prior, neighbor)}
        native = (cli.run_unavailable_stdin(command, work, env, timeout=360) if answers is None else
                  launcher.run_prompted(command, work, env, answers, timeout=480))
        record.update(process=native, protected_before=before,
                      protected_after={item.name: cli.identity(item) for item in (source, prior, neighbor)})
        require(record["protected_before"] == record["protected_after"], "The source or neighboring file changed")
    result = cli.outcome(native["stdout"])
    record["outcome"] = result
    record["temporary_converter_state"] = sorted(path.name for path in work.glob(".WinBookSplit-Calibre-*"))
    require(not record["temporary_converter_state"], "A stopped converter left temporary Calibre state")
    require((result["code"], native["exit_code"]) == expected and result["exit_code"] == native["exit_code"],
            f"Expected {expected}, observed {(result['code'], native['exit_code'])}")
    if expected[1] != 0:
        require(result["written_count"] == 0 and result["final_directory"] is None,
                "Failed or declined processing reported successful output")
        additions = {path.name for path in base.iterdir()} - original_members
        require(all(name.startswith((".WinBookSplit-console-", ".WinBookSplit-failed-")) for name in additions),
                "Failed or declined processing published or left staged chapter output")
        if preview:
            require(not additions, "Preview wrote an output or diagnostic directory")
        return None

    execution = result["engine_result"]["execution"]
    final = Path(execution["final_directory"])
    require(final.parent == base and final.name not in original_members, "Publication reused or escaped its output directory")
    chapters = [(row["start"], row["end"], final / row["filename"]) for row in execution["outputs"]]
    require(len(chapters) == result["written_count"] and all(cli.digest(path) == row["sha256"]
            for (_, _, path), row in zip(chapters, execution["outputs"])), "Reopened chapter digest or output count changed")
    if manual:
        require([(start, end) for start, end, _ in chapters] == [(0, 2), (2, execution["total_pages"])],
                "Manual split did not use the selected generated physical start pages")
    if source.suffix.lower() == ".pdf":
        actual = [int(page["/WBSFixturePage"]) for _, _, path in chapters for page in PdfReader(path).pages]
        require(actual == [1, 2, 3], "PDF-only processing lost physical pages")
        require("conversion" not in execution and "Calibre:" not in native["stdout"], "PDF-only processing entered converter discovery")
        return final
    retained = execution["retained_intermediate"]
    require(retained["filename"] == "WinBookSplit_Converted.pdf", "Application lost the explicitly retained generated PDF")
    record["chapter_oracle"] = oracle["verify_chapters"](final / retained["filename"], chapters, auto1=not manual)
    require(record["chapter_oracle"]["generated_sha256"] == retained["sha256"]
            and execution["original_ebook_identity"]["sha256"] == before[source.name]["sha256"],
            "Conversion output or source hash lost its binding")
    record["selected_converter"] = execution["conversion"]["converter_path"]
    return final


def storage(scope):
    return scope / "WinBookSplit" / "Calibre"


def run_setup(record, args, work, host, scope, proxy, *, candidate=None, options=()):
    command = [host, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "RemoteSigned",
               "-File", str(args.app_root / "Setup-WinBookSplit.ps1"), "-Calibre", "-NonInteractive",
               "-PythonPath", str(args.application_python_path), *options]
    process = cli.run_unavailable_stdin(command, work, environment_for(work, host, scope, candidate, proxy), timeout=360)
    record["offline_setup"] = process
    prefix = "[CALIBRE-SETUP] "
    frames = [json.loads(line.removeprefix(prefix)) for line in process["stdout"].splitlines() if line.startswith(prefix)]
    require(len(frames) == 1 and frames[0]["exit_code"] == process["exit_code"], "Calibre setup did not return its exact outcome")
    return frames[0]


def verify_managed_selection(record, scope):
    managed = storage(scope)
    selected = Path(record["selected_converter"])
    conversion = record["outcome"]["engine_result"]["execution"]["conversion"]
    require(selected.is_relative_to(managed) and selected.name.casefold() == "ebook-convert.exe"
            and conversion.get("managed_calibre") is True and (managed / "current.json").is_file(),
            "Conversion bypassed the persisted managed Calibre")
    installation = selected.parent.parent
    fingerprint = tree_digest(installation)
    require(fingerprint["members"] > 1, "Managed converter has no supporting installation files")
    state = managed / "State" / installation.name
    require(all((state / name).is_dir() for name in ("Config", "Cache")),
            "Managed conversion did not use its own configuration and cache directories")
    record["managed"] = {"installation": str(installation), "tree": fingerprint,
                         "state": str(state), "current_sha256": cli.digest(managed / "current.json")}
    return record["managed"]


def characterize(args, work, report):
    book = work / "Å illustrated chapters [1] & literal.epub"
    fixtures["authored_epub"](book, rich=True)
    bad_book = work / "Original malformed.epub"
    bad_book.write_bytes(b"Original bytes which are not an EPUB ZIP.\n")
    pdf = work / "PDF only Å [1].pdf"
    writer = PdfWriter()
    for index in range(1, 4):
        page = writer.add_blank_page(100, 120)
        page[NameObject("/WBSFixturePage")] = NumberObject(index)
    with pdf.open("xb") as stream:
        writer.write(stream)
    (work / "neighbor.txt").write_bytes(b"Keep this authored neighboring file.\n")
    probe = cli.process_tests.host_probe(work)
    hosts = {host["id"]: host for host in (support.paths.host_observation(shell, work, probe) for shell in args.shell_path)}
    require(set(hosts) == {"PS51", "PS7"}, "Supply the two distinct supported PowerShell hosts")
    ps51, ps7 = (hosts[key]["shell_executable"] for key in ("PS51", "PS7"))
    report.update(shells=hosts, source_sha256=cli.digest(book), application=application_receipt(args.app_root, args.package_manifest))
    chosen = args.case or CASES
    selected = {}
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as closed:
        closed.bind(("127.0.0.1", 0))
        proxy = f"http://127.0.0.1:{closed.getsockname()[1]}"
        for case in chosen:
            scope = work / ("local-" + ("managed-install" if case == "managed-reuse" else case))
            scope.mkdir(exist_ok=True)
            row = {"id": case, "status": "FAIL"}
            report["cases"].append(row)
            try:
                if case in ("managed-install", "managed-reuse") and not args.online:
                    row.update(status="NOT_RUN", reason="Live official download route not selected")
                    continue
                if case == "older-preserved-offline" and (args.older_calibre_path is None or args.calibre_installer is None):
                    row.update(status="NOT_RUN", reason="Original older converter or verified offline artifact unavailable")
                    continue
                if case == "existing-path":
                    run_application(row, args, work, ps51, book, scope, candidate=args.existing_calibre_path, proxy=proxy)
                    require(Path(row["selected_converter"]).samefile(args.existing_calibre_path), "PATH discovery chose a different converter")
                elif case == "explicit-path":
                    run_application(row, args, work, ps7, book, scope, explicit=args.existing_calibre_path, proxy=proxy)
                    require(Path(row["selected_converter"]).samefile(args.existing_calibre_path), "Explicit converter was substituted")
                elif case == "explicit-rejection":
                    run_application(row, args, work, ps7, book, scope, candidate=args.existing_calibre_path,
                                    explicit=work / "missing-converter.exe", proxy=proxy, expected=("converter_invalid", 3))
                elif case == "existing-azw3":
                    azw3 = work / "Original Å illustrated chapters.azw3"
                    argv = [str(args.existing_calibre_path), str(book), str(azw3), "--output-profile", "tablet"]
                    row["azw3_generation"] = cli.run_unavailable_stdin(
                        argv, work, environment_for(work, ps7, scope, proxy=proxy), timeout=180)
                    require(row["azw3_generation"]["exit_code"] == 0 and azw3.is_file()
                            and azw3.read_bytes()[60:68] == b"BOOKMOBI", "Actual Calibre did not produce a genuine AZW3")
                    run_application(row, args, work, ps7, azw3, scope, candidate=args.existing_calibre_path, proxy=proxy)
                    require(Path(row["selected_converter"]).samefile(args.existing_calibre_path),
                            "AZW3 conversion did not use the suitable existing Calibre")
                elif case in ("preview-no-download", "noninteractive-no-download"):
                    run_application(row, args, work, ps51, book, scope, proxy=proxy, preview=case.startswith("preview"),
                                    expected=("converter_not_found", 3))
                    require(PREPARE not in row["process"]["stdout"], "Non-installing request prompted for preparation")
                elif case == "setup-check-no-download":
                    setup = run_setup(row, args, work, ps51, scope, proxy, options=("-Check",))
                    require(setup["code"] == "converter_not_found" and setup["exit_code"] == 3 and setup["changed"] is False
                            and PREPARE not in row["offline_setup"]["stdout"], "Read-only converter setup changed state or prompted")
                elif case == "declined":
                    run_application(row, args, work, ps7, book, scope, proxy=proxy, answers="N\n", expected=("cancelled", 130))
                    require(row["process"]["stdout"].count(PREPARE) == 1, "Declined preparation did not show exactly one prompt")
                elif case == "download-unavailable":
                    run_application(row, args, work, ps51, book, scope, proxy=proxy, answers="Y\n",
                                    expected=("converter_download_failed", 3))
                    require(row["process"]["stdout"].count(PREPARE) == 1, "Download failure changed the one-prompt consent contract")
                elif case == "invalid-offline-artifact":
                    invalid = work / "Invalid authored Calibre installer.exe"
                    invalid.write_bytes(b"Original test bytes, deliberately different from the reviewed artifact.\n")
                    invalid_before = cli.identity(invalid)
                    setup = run_setup(row, args, work, ps51, scope, proxy, options=("-CalibreInstaller", str(invalid)))
                    require(setup["code"] == "converter_integrity_failed" and setup["exit_code"] == 3
                            and setup["changed"] is False and cli.identity(invalid) == invalid_before,
                            "Unverified offline bytes were accepted or changed")
                elif case == "managed-install":
                    require(not storage(scope).exists(), "First-use conversion scope unexpectedly contained managed Calibre files")
                    first = run_application(row, args, work, ps51, book, scope, answers="Y\nY\nN\n")
                    require(row["process"]["stdout"].count(PREPARE) == 1, "First-use preparation did not resume after one prompt")
                    selected[case] = (verify_managed_selection(row, scope), first, tree_digest(first))
                elif case == "managed-reuse":
                    if "managed-install" not in selected:
                        row.update(status="NOT_RUN", reason="Managed installation did not finish earlier in this run")
                        continue
                    managed, first, prior_tree = selected["managed-install"]
                    run_application(row, args, work, ps7, book, scope, base=first.parent, proxy=proxy, answers="Y\nN\n")
                    require(verify_managed_selection(row, scope) == managed and tree_digest(first) == prior_tree,
                            "Repeat processing changed the managed installation or first chapter output")
                    require(PREPARE not in row["process"]["stdout"], "A working managed converter requested another preparation")
                elif case == "older-preserved-offline":
                    older = args.older_calibre_path
                    old_before, installer_before = tree_digest(older.parent), cli.identity(args.calibre_installer)
                    old_probe = cli.run_unavailable_stdin([str(older), "--version"], work,
                        environment_for(work, ps51, scope, older, proxy), timeout=30)
                    old_version = re.search(r"\bcalibre (\d+\.\d+\.\d+)\b", old_probe["stdout"])
                    row["older_probe"] = old_probe
                    require(old_probe["exit_code"] == 0 and old_version
                            and tuple(map(int, old_version[1].split("."))) < (9, 17, 0),
                            "The supplied older converter did not establish a below-minimum real version")
                    refusal = {"id": case + "-before"}
                    row["before_setup"] = refusal
                    run_application(refusal, args, work, ps51, book, scope, candidate=older, proxy=proxy,
                                    expected=("converter_not_found", 3))
                    attempts = refusal["process"]["stdout"] + refusal["process"]["stderr"]
                    require(str(older) in attempts and old_version[1] in attempts,
                            "The actual rejected older candidate was absent from diagnostics")
                    setup = run_setup(row, args, work, ps51, scope, proxy, candidate=older,
                                      options=("-CalibreInstaller", str(args.calibre_installer)))
                    require(setup["exit_code"] == 0 and setup["code"] == "ready"
                            and setup["changed"] is True and setup["source"] == "managed",
                            "Verified offline preparation did not publish a real managed converter")
                    run_application(row, args, work, ps7, book, scope, candidate=older, proxy=proxy, manual=True)
                    verify_managed_selection(row, scope)
                    row["older_installation"] = {"before": old_before, "after": tree_digest(older.parent)}
                    require(row["older_installation"]["before"] == row["older_installation"]["after"]
                            and cli.identity(args.calibre_installer) == installer_before,
                            "Managed preparation changed the older installation or installer")
                elif case == "conversion-failure":
                    run_application(row, args, work, ps7, bad_book, scope, explicit=args.existing_calibre_path, proxy=proxy,
                                    expected=("conversion_failed", 4))
                elif case == "pdf-only":
                    run_application(row, args, work, ps51, pdf, scope, proxy=proxy, manual=True)
                if case not in ("managed-install", "managed-reuse", "older-preserved-offline"):
                    require(not (storage(scope) / "current.json").exists(), "A non-installing route published managed configuration")
                row["status"] = "PASS"
            except (RuntimeError, OSError, KeyError, TypeError, ValueError, StopIteration) as error:
                row["failure"] = str(error)
            finally:
                print(case + ": " + row["status"] + (": " + row["failure"] if "failure" in row else ""), flush=True)
    require(report["application"] == application_receipt(args.app_root, args.package_manifest),
            "Application payload changed during managed Calibre acceptance")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--app-root", type=Path, required=True)
    parser.add_argument("--application-python-path", type=Path, required=True)
    parser.add_argument("--shell-path", type=Path, action="append", required=True)
    parser.add_argument("--existing-calibre-path", type=Path, required=True)
    parser.add_argument("--older-calibre-path", type=Path)
    parser.add_argument("--calibre-installer", type=Path)
    parser.add_argument("--package-manifest", type=Path)
    parser.add_argument("--online", action="store_true", help="Exercise product-managed official portable download")
    parser.add_argument("--case", action="append", choices=CASES)
    parser.add_argument("--report", type=Path, required=True)
    args = parser.parse_args()
    require(os.name == "nt" and sys.flags.isolated and sys.dont_write_bytecode,
            "Use native Windows and isolated developer Python with -I -B")
    require(len(args.shell_path) == 2, "Supply Windows PowerShell 5.1 and PowerShell 7 executable paths")
    args.app_root = args.app_root.resolve(strict=True)
    for path in (args.application_python_path, args.existing_calibre_path, *args.shell_path,
                 args.older_calibre_path, args.calibre_installer, args.package_manifest):
        require(path is None or path.is_absolute() and path.is_file(), "Supply absolute existing test paths")
    report_path = support.runner.new_external_path(args.report)
    work = Path(tempfile.mkdtemp(prefix="WBC-")).resolve()
    report = {"cases": [], "workspace": str(work), "app_root": str(args.app_root), "windows": platform.win32_ver(),
              "developer_python": sys.version, "pypdf": pypdf_version,
              "checkout_head": subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=ROOT, text=True).strip()}
    try:
        characterize(args, work, report)
    except Exception as error:
        report["error"] = str(error)
    with report_path.open("x", encoding="utf-8", newline="\n") as stream:
        json.dump(report, stream, ensure_ascii=True, indent=2)
        stream.write("\n")
    if "error" in report:
        print(report["error"], file=sys.stderr)
    return int("error" in report or any(row["status"] == "FAIL" for row in report["cases"]))


if __name__ == "__main__":
    raise SystemExit(main())
