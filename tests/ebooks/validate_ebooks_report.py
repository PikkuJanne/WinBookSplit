"""Read-only, fail-closed M4-T04 authored ebook receipt validation."""
from hashlib import sha256
from datetime import datetime, timedelta
import importlib.util
import json
import math
from pathlib import Path
import re
import sys
from urllib.parse import urlsplit

ROOT = Path(__file__).resolve().parents[2]
HEX = re.compile(r"[0-9a-f]{64}\Z")
TITLES = ["Chapter One", "Chapter Two", "Chapter Three"]
MARKERS = ["WBS-PAGE-001", "WBS-PAGE-002", "WBS-PAGE-003"]
OUTLINE = [{"title": title, "page": number, "depth": 0} for number, title in enumerate(TITLES)]
RANGES = {"epub": [[0, 1], [1, 2], [2, 3]], "azw3": [[0, 1], [1, 2], [2, 4]]}
GEOMETRY = {"media_box": [0.0, 0.0, 612.0, 792.0], "crop_box": [0.0, 0.0, 612.0, 792.0], "rotation": 0}
MEMBERS = {"mimetype", "META-INF/container.xml", "OEBPS/content.opf", "OEBPS/nav.xhtml", "OEBPS/toc.ncx",
           "OEBPS/style.css", *(f"OEBPS/chapter{n}.xhtml" for n in range(1, 4))}
APPLICATION = ("VERSION", "WinBookSplit.ps1", "WinBookSplit.bat", "Export-WinBookSplitDiagnostics.ps1", "requirements.txt",
    "engine/WinBookSplit.Runtime.ps1", "engine/WinBookSplit.Process.ps1", "engine/WinBookSplit.Diagnostics.ps1",
    "engine/WinBookSplit.Paths.ps1", "engine/WinBookSplit.Logging.ps1", "engine/WinBookSplit.Support.ps1",
    "engine/WinBookSplit.Outcomes.json", "engine/winbooksplit_engine.py", "engine/winbooksplit_conversion.py",
    "engine/winbooksplit_windows.py", "engine/winbooksplit_job.py")
CALIBRE_SHA256 = "f46a01c9b8cd392e1190ce2e50a2a5a48b326c8c3eab670bbcd8b298bede501d"


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


console = load("wbs_ebook_console_guard", ROOT / "tests/manual/current_launchers.py")
fidelity = load("wbs_ebook_fidelity_guard", ROOT / "tests/fidelity/validate_fidelity_report.py")


def need(condition, message):
    if not condition:
        raise ValueError(message)


def valid_hash(value):
    return isinstance(value, str) and HEX.fullmatch(value) is not None


def integer(value, minimum=0):
    return type(value) is int and value >= minimum


def records(text, prefix):
    need(isinstance(text, str), "Missing actual UTF8 stream")
    return [fidelity.strict_json(line[len(prefix):]) for line in text.splitlines() if line.startswith(prefix)]


def utc(value):
    need(isinstance(value, str), "Actual UTC timestamp missing")
    stamp = datetime.fromisoformat(value)
    need(stamp.tzinfo is not None and stamp.utcoffset() == timedelta(0), "Actual UTC timestamp invalid")
    return stamp


def validate_process(process, *, expected_exit=0, command=None):
    need(isinstance(process, dict) and process.get("actual_process") is True and process.get("stdin") == "DEVNULL"
         and integer(process.get("exit_code")) and process["exit_code"] == expected_exit
         and process.get("timed_out") is False and integer(process.get("pid"), 1), "Native process/status evidence invalid")
    need(type(process.get("elapsed_seconds")) in (int, float) and math.isfinite(process["elapsed_seconds"])
         and 0 <= process["elapsed_seconds"] <= process.get("timeout_seconds", -1)
         and integer(process.get("timeout_seconds"), 1), "Actual bounded native timing missing")
    started, finished = utc(process.get("started_at")), utc(process.get("finished_at"))
    need(started <= finished and (finished-started).total_seconds() <= process["timeout_seconds"] + 5,
         "Native UTC ordering/deadline contradicts process observation")
    need(isinstance(process.get("command"), list) and process["command"] and all(isinstance(value, str) for value in process["command"])
         and isinstance(process.get("cwd"), str) and Path(process["cwd"]).is_absolute(), "Literal native argv/cwd missing")
    if command is not None:
        need(process["command"] == command, "Actual literal native command differs")
    for name in ("stdout", "stderr"):
        text = process.get(name)
        need(isinstance(text, str) and sha256(text.encode("utf-8")).hexdigest() == process.get(name + "_sha256"), "Raw native stream hash differs")


def validate_format_checks(checks, provenance, resource_closure):
    need(isinstance(checks, dict) and set(checks) == {"epub", "azw3"}, "Both independently identified ebook formats required")
    files = provenance.get("files", [])
    need(isinstance(files, list) and len(files) == 2 and {item.get("format") for item in files} == {"epub", "azw3"}, "Fixture provenance formats missing")
    for item in files:
        observed = checks[item["format"]]
        need(isinstance(observed, dict) and observed.get("format") == item["format"]
             and all(item.get(key) == observed.get(key) for key in ("sha256", "size_bytes"))
             and valid_hash(item.get("sha256")) and integer(item.get("size_bytes"), 1), "Actual format bytes/provenance binding differs")
    epub, azw3 = checks["epub"], checks["azw3"]
    need(epub == resource_closure and epub.get("epub_version") == "3.0" and epub.get("mimetype_first_uncompressed") is True
         and epub.get("spine") == ["c1", "c2", "c3"] and epub.get("navigation") == [{"title": title, "href": f"chapter{n}.xhtml"} for n, title in enumerate(TITLES, 1)]
         and epub.get("scripts_or_active_content") is False and epub.get("external_references") == []
         and epub.get("resource_scope") == "closed-authored-family", "Original EPUB is not the inspected local original EPUB3/TOC fixture")
    members = epub.get("members")
    need(isinstance(members, dict) and set(members) == MEMBERS and all(integer(item.get("size_bytes"), 1) and item["size_bytes"] <= 65536
         and valid_hash(item.get("sha256")) for item in members.values()), "Bounded original EPUB members missing")
    refs = epub.get("local_references")
    need(isinstance(refs, list) and len(refs) >= 15 and all(set(item) == {"from", "reference", "target"}
         and item["from"] in MEMBERS and item["target"] in MEMBERS and isinstance(item["reference"], str)
         and not urlsplit(item["reference"]).scheme and not urlsplit(item["reference"]).netloc for item in refs), "EPUB resource closure is not bound to inspected archive members")
    need(azw3.get("palmdb_identity") == "BOOKMOBI" and azw3.get("mobi_version") == 8 and type(azw3.get("mobi_version")) is int
         and azw3.get("encryption_type") == 0 and type(azw3.get("encryption_type")) is int and azw3.get("not_zip") is True
         and azw3.get("scope") == "bounded-generated-standalone-KF8-fixture" and integer(azw3.get("record_count"), 2)
         and azw3["record_count"] <= 4096 and integer(azw3.get("mobi_header_length"), 112) and azw3["mobi_header_length"] <= 500,
         "AZW3 is not the bounded unencrypted generated MOBI8/KF8 fixture")
    offsets = azw3.get("record_offsets")
    need(isinstance(offsets, list) and len(offsets) == azw3["record_count"] and all(integer(value) for value in offsets)
         and offsets[0] >= 78 + 8 * len(offsets) and offsets[-1] < azw3["size_bytes"]
         and all(a < b for a, b in zip(offsets, offsets[1:])) and offsets[1] - offsets[0] == azw3.get("record_zero_size_bytes")
         and azw3["record_zero_size_bytes"] >= 16 + azw3["mobi_header_length"] and valid_hash(azw3.get("record_zero_sha256")), "AZW3 record table/MOBI header bounds differ")


def validate_pdf_observation(observation, renderer_path, *, expected_count=None, expected_outline=None):
    need(isinstance(observation, dict) and Path(observation.get("path", "")).is_absolute() and valid_hash(observation.get("sha256"))
         and integer(observation.get("size_bytes"), 1) and integer(observation.get("page_count"), 1)
         and observation["page_count"] <= 16, "Actual positive PDF bytes/physical count missing")
    if expected_count is not None:
        need(observation["page_count"] == expected_count, "Actual PDF physical count differs from fixture/plan")
    if expected_outline is not None:
        need(observation.get("outline") == expected_outline, "Actual generated/chapter outline destinations differ")
    need(isinstance(observation.get("outline"), list) and all(set(item) == {"title", "page", "depth"}
         and isinstance(item["title"], str) and integer(item["page"]) and item["page"] < observation["page_count"]
         and integer(item["depth"]) for item in observation["outline"]), "Actual bounded PDF outline malformed")
    root = Path(observation.get("render_root", ""))
    need(root.is_absolute() and not root.is_relative_to(ROOT), "Persistent PDF render root escaped external scope")
    process = observation.get("render_process")
    validate_process(process, command=[renderer_path, "-r", "72", "-cropbox", "-png", observation["path"], str(root / "page")])
    need(process["cwd"] == str(root), "Renderer cwd differs from exact retained root")
    pages = observation.get("pages")
    need(isinstance(pages, list) and len(pages) == observation["page_count"], "Every physical PDF page must be independently reopened/rendered")
    for number, page in enumerate(pages, 1):
        need(integer(page.get("page"), 1) and page["page"] == number and isinstance(page.get("text"), str)
             and page["text"] and len(page["text"].encode("utf-8")) <= 65536
             and sha256(page["text"].encode("utf-8")).hexdigest() == page.get("text_sha256")
             and valid_hash(page.get("content_sha256")), "Reopened physical text/content/order differs")
        geometry = page.get("geometry")
        need(isinstance(geometry, dict) and all(isinstance(geometry.get(key), list) and len(geometry[key]) == 4
             and all(type(value) in (int, float) and math.isfinite(value) for value in geometry[key]) for key in ("media_box", "crop_box"))
             and geometry == GEOMETRY and type(geometry.get("rotation")) is int,
             "Observed tested letter geometry/rotation differs")
        render = page.get("render")
        fidelity.validate_png(render)
        match = re.fullmatch(r"page-([0-9]+)\.png", render["filename"])
        need(match is not None and int(match[1]) == render["page"] == number and Path(render["path"]) == root / render["filename"]
             and (render["width"], render["height"]) == (612, 792), "Actual retained page PNG path/order/geometry differs")
    return pages


def compare_physical_pages(source, output, start, end):
    need(integer(start) and integer(end, start + 1) and end <= len(source) and len(output) == end - start, "Physical range is not a complete nonempty source slice")
    for original, actual in zip(source[start:end], output):
        need(all(original[key] == actual[key] for key in ("text", "text_sha256", "content_sha256", "geometry")), "Reopened chapter text/content/geometry differs from captured reference")
        need(all(original["render"][key] == actual["render"][key] for key in ("width", "height", "png_sha256", "pixels_sha256")), "Same-renderer physical source/chapter pixels differ")


def validate_references(references, checks, calibre_path, renderer_path):
    need(isinstance(references, dict) and set(references) == {"epub", "azw3"}, "Both independent format references required")
    for fmt, profiles in references.items():
        need(isinstance(profiles, dict) and set(profiles) == {"tablet", "default"}, "Exact current-profile and diagnostic-default references required")
        count = 3 if fmt == "epub" else 4
        for profile, ref in profiles.items():
            options = ["--output-profile", "tablet"] if profile == "tablet" else []
            need(ref.get("input_format") == fmt and ref.get("input_sha256") == checks[fmt]["sha256"]
                 and ref.get("profile") == profile and ref.get("options") == options, "Independent reference format/input/profile binding differs")
            observation = ref.get("observation")
            need(isinstance(observation, dict), "Independent actual reference PDF missing")
            process = ref.get("process")
            command = [calibre_path, ref["input_path"], observation["path"], *options]
            validate_process(process, command=command)
            need(process.get("argv") == command, "Reference command compatibility alias differs")
            pages = validate_pdf_observation(observation, renderer_path, expected_count=count, expected_outline=OUTLINE)
            markers = re.findall(r"(?<![A-Z0-9-])WBS-PAGE-[0-9]{3}(?![0-9])", "\n".join(page["text"] for page in pages))
            need(markers == MARKERS, "Independent real conversion lost/duplicated authored chapter markers")


def validate_identity(item):
    need(isinstance(item, dict) and valid_hash(item.get("sha256"))
         and all(integer(item.get(key)) for key in ("size_bytes", "device", "inode", "attributes"))
         and item["inode"] > 0 and not item["attributes"] & 1024,
         "Ordinary immutable file identity missing")


def validate_tree(tree, *, path=None, members=None):
    need(isinstance(tree, dict) and Path(tree.get("path", "")).is_absolute()
         and integer(tree.get("device")) and integer(tree.get("inode"), 1)
         and isinstance(tree.get("members"), dict) and tree["members"], "Authenticated owned directory identity missing")
    if path is not None:
        need(Path(tree["path"]) == Path(path), "Authenticated directory path differs")
    if members is not None:
        need(set(tree["members"]) == set(members), "Authenticated exact owned member set differs")
    for name, item in tree["members"].items():
        need(isinstance(name, str) and name == Path(name).name and name not in {".", ".."}, "Owned directory member escaped")
        validate_identity(item)


def bind_raw(tree, name, raw):
    need(isinstance(raw, str) and name in tree["members"]
         and sha256(raw.encode("utf-8")).hexdigest() == tree["members"][name]["sha256"]
         and len(raw.encode("utf-8")) == tree["members"][name]["size_bytes"], "Authenticated owned raw record differs")
    return fidelity.strict_json(raw)


def validate_config(config, *, expected_root=None):
    variables = {"CALIBRE_CONFIG_DIRECTORY": "config", "CALIBRE_CACHE_DIRECTORY": "cache", "CALIBRE_TEMP_DIR": "temp"}
    need(isinstance(config, dict) and set(config.get("environment", {})) == set(variables)
         and config.get("new_empty_process_directories") is True and config.get("inherited_environment_restored") is True
         and isinstance(config.get("original_environment"), dict) and set(config["original_environment"]) == set(variables)
         and config.get("restored_environment") == config["original_environment"], "Process-local configuration/environment restoration missing")
    parents = set()
    for variable, name in variables.items():
        root = Path(config["environment"][variable])
        need(root.is_absolute() and root.name == name, "Converter configuration path escaped owned selection")
        if expected_root is not None:
            need(root == Path(expected_root) / name, "Converter configuration escaped the actual owned operation root")
        parents.add(root.parent)
        before, after = config.get("before", {}).get(variable), config.get("after", {}).get(variable)
        need(isinstance(before, dict) and isinstance(after, dict) and before.get("path") == after.get("path") == str(root)
             and integer(before.get("device")) and integer(before.get("inode"), 1)
             and all(before[key] == after.get(key) for key in ("device", "inode"))
             and before.get("files") == {} and before.get("directories") == {}, "New empty configuration identity changed")
        need(isinstance(after.get("files"), dict) and isinstance(after.get("directories"), dict)
             and len(after["files"]) + len(after["directories"]) <= 4096, "Bounded configuration snapshot missing")
        for relative, item in after["files"].items():
            need(not Path(relative).is_absolute() and ".." not in Path(relative).parts, "Configuration member escaped owned directory")
            validate_identity(item)
            need(item["size_bytes"] <= 33554432, "Configuration member exceeded fixture bounds")
        for relative, item in after["directories"].items():
            need(not Path(relative).is_absolute() and ".." not in Path(relative).parts
                 and isinstance(item, dict) and integer(item.get("device")) and integer(item.get("inode"), 1), "Configuration directory identity escaped")
    need(len(parents) == 1 and next(iter(parents)).name.endswith("settings"), "Process configuration directories do not share their owned root")


def validate_protected(row, cleanup_complete):
    initial, before = row.get("source_observations_initial"), row.get("source_observations_before")
    need(isinstance(initial, dict) and isinstance(before, dict) and set(initial) == set(before) == {"source", "neighbor", "prior"}
         and before == row.get("source_observations_after") and row.get("source_read_only_observed") is True,
         "Actual source/neighbor/prior identity changed")
    for key in initial:
        validate_identity(initial[key])
        validate_identity(before[key])
        expected = {**initial[key], "attributes": initial[key]["attributes"] | 1} if key == "source" else initial[key]
        need(before[key] == expected, "Read-only input or neighboring/prior identity differs")
        need(initial[key]["size_bytes"] > 0, "Authored input/neighbor/prior file is empty")
    if cleanup_complete:
        need(row.get("source_observations_restored") == row.get("source_observations_after_cleanup") == initial
             and row.get("source_attributes_restored") is True and row.get("owned_outputs_removed") is True
             and row.get("passed") is True and row.get("output_members_after_cleanup") == ["prior-output.pdf"],
             "Owned cleanup/source-attribute restoration proof missing")
    prior = row.get("prior_publications_before")
    need(isinstance(prior, dict) and prior == row.get("prior_publications_after"), "Earlier publication or console changed")
    for name, tree in prior.items():
        need(Path(tree.get("path", "")).name == name and Path(tree["path"]).parent == Path(row["output_base"]), "Prior run escaped shared base")
        validate_tree(tree)


def validate_case(row, cleanup_complete=True, *, calibre_path=None, renderer_path=None):
    need(isinstance(row, dict) and row.get("kind") in {"retained", "malformed", "canary"}
         and row.get("input_kind") in {"epub", "azw3"} and row.get("host_id") in {"PS51", "PS7"}, "Actual ebook case identity missing")
    kind, fmt = row["kind"], row["input_kind"]
    expected_id = row["host_id"] + "-" + fmt + "-" + ("auto1-keep" if kind == "retained" else kind)
    need(row.get("id") == expected_id and (kind != "canary" or fmt == "epub"), "Ebook matrix case identity differs")
    workspace, base, source = (Path(row.get(key, "")) for key in ("workspace", "output_base", "input_path"))
    need(workspace.is_absolute() and base.is_absolute() and source.parent == workspace / "b"
         and source.name == "Book Å 日本 [1] & %! $(literal)." + fmt and Path(row.get("cwd", "")) == workspace / "c",
         "Literal source/owned application cwd binding differs")
    params = row.get("parameters")
    need(isinstance(params, list) and len(params) == 14, "Exact NI ebook parameter list missing")
    calibre_path = str(calibre_path or params[7])
    expected_params = ["-InputFile", str(source), "-OutputDirectory", str(base), "-PythonPath", params[5],
                       "-CalibrePath", calibre_path, "-Mode", "Auto", "-BookmarkLevel", "1", "-KeepConvertedPdf", "-NonInteractive"]
    need(params == expected_params and Path(params[5]).is_absolute(), "Literal fixed tablet/Auto1/Keep/NonInteractive settings differ")
    command = [row["shell_executable"], "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "RemoteSigned", "-File",
               str(workspace / "a/WinBookSplit.ps1"), *expected_params]
    status, code, native = ("error", "conversion_failed", 4) if kind == "malformed" else ("success", "split_complete", 0)
    validate_process(row, expected_exit=native, command=command)
    validate_protected(row, cleanup_complete)
    need(row.get("application_unchanged") is True and isinstance(row.get("application_sha256"), dict)
         and set(row["application_sha256"]) == set(APPLICATION) and all(valid_hash(value) for value in row["application_sha256"].values()), "Copied application identity changed/missing")
    validate_config(row.get("calibre_environment"), expected_root=workspace / "settings")
    final, engine = row.get("outcome"), row.get("engine_record")
    for value, protocol in ((final, "winbooksplit.outcome"), (engine, "winbooksplit.result")):
        need(isinstance(value, dict) and value.get("protocol") == protocol and type(value.get("version")) is int and value["version"] == 1
             and value.get("status") == status and value.get("code") == code and type(value.get("exit_code")) is int and value["exit_code"] == native
             and value.get("mode") == "1" and type(value.get("written_count")) is int and value["written_count"] == (0 if native else 3),
             "Native terminal status/code/mode/count contradicts actual case")
    need(records(row["stdout"], "[OUTCOME] ") == [final] and final.get("engine_result") == engine and row.get("engine_records") == [engine], "Sole native stdout/engine outcome binding differs")
    text = row.get("console_log")
    need(isinstance(text, str) and sha256(text.encode("utf-8")).hexdigest() == row.get("console_log_sha256")
         and [fidelity.strict_json(line) for line in text.splitlines() if line.startswith("{")] == [engine]
         and records(text, "[PROCESS] ") == row.get("process_summaries"), "Actual raw engine/transport log binding differs")
    log = Path(row.get("console_path", ""))
    need(log.name == "console.log" and log.parent.parent == base
         and log.parent.name == ".WinBookSplit-console-" + row.get("console_evidence", {}).get("owner", {}).get("run_id", "")
         and [line[5:].strip() for line in row["stdout"].splitlines() if line.startswith("Log: ")] == [str(log)], "Native console location differs")
    console.validate_console_evidence(row.get("console_evidence"), final, row["console_log_sha256"], source_path=str(source),
                                     source_observation=row["source_observations_before"]["source"], log_text=text)
    settings = row["console_evidence"]["run_manifest"]["settings"]
    need(settings["non_interactive"] is True and settings["preview"] is False and settings["keep_converted_pdf"] is True
         and settings["input_kind"] == fmt, "Local manifest NI/retention/input-kind settings differ")
    transports = row["process_summaries"]
    need(isinstance(transports, list) and len(transports) == 1 and records(text, "[PROCESS] ") == transports
         and all(transports[0].get(field) is True for field in ("JobAssigned", "ParentStopped", "DescendantsStopped", "StreamsComplete"))
         and type(transports[0].get("ExitCode")) is int and transports[0]["ExitCode"] == native
         and type(transports[0].get("ResultRecordCount")) is int and transports[0]["ResultRecordCount"] == 1,
         "Owned process tree/stream/native exit proof missing")
    for field in ("InteractionCount", "ReplyCount", "QueuedReplyCount"):
        need(field not in transports[0] or type(transports[0][field]) is int and transports[0][field] == 0, "NI process recorded prompt activity")
    need(not any(marker in row["stdout"] or marker in text for marker in ("[WBS-INTERACTION] ", "[INTERACTION] ", "[REPLY] ", "[INTERACTION-REPLY] ", "Write these chapter PDFs", "Enter selection", "Open the completed folder", "[OPEN] ")), "NI ebook case prompted/opened")
    if native:
        validate_failure(row, calibre_path)
        need("Done." not in row["stdout"], "Failure printed a success message")
    else:
        renderer_path = str(renderer_path or row["retained_pdf"]["render_process"]["command"][0])
        validate_success(row, renderer_path, calibre_path)
        need("Done." in row["stdout"] and "Output: " + final["final_directory"] in row["stdout"], "Truthful successful publication display missing")
    if kind == "canary" and cleanup_complete:
        validate_canary_asset(row.get("canary_asset"), row["canary_urls"])
        need(row["source_observations_initial"]["source"]["sha256"] == row["canary_asset"]["sha256"], "Canary input byte identity differs")
        validate_block_diagnostics(row)
    elif kind != "canary":
        need(row.get("canary_urls") == [], "Ordinary fixture used remote resource URLs")


def validate_success(row, renderer_path, calibre_path=None):
    engine, final, fmt = row["engine_record"], row["outcome"], row["input_kind"]
    plan, execution = engine.get("plan"), engine.get("execution")
    count = 3 if fmt == "epub" else 4
    need(isinstance(plan, dict) and plan.get("mode") == "1" and type(plan.get("total_pages")) is int and plan["total_pages"] == count
         and isinstance(plan.get("ranges"), list) and all(isinstance(bounds, list) and len(bounds) == 2
             and all(type(value) is int for value in bounds) for bounds in plan["ranges"])
         and plan["ranges"] == RANGES[fmt] and plan.get("coverage") == {"complete": True, "covered_pages": count, "section_count": 3}
         and plan["coverage"].get("complete") is True
         and all(type(plan["coverage"][key]) is int for key in ("covered_pages", "section_count"))
         and plan.get("output_naming", {}).get("resolved_base") == row["output_base"] and plan.get("keep_converted_pdf") is True,
         "Complete fixed Auto1 captured plan/coverage/output base differs")
    original = plan.get("original_ebook_identity")
    need(isinstance(original, dict) and original.get("path") == row["input_path"] and original.get("sha256") == row["source_observations_before"]["source"]["sha256"]
         and original.get("size_bytes") == row["source_observations_before"]["source"]["size_bytes"], "Plan original ebook identity differs from immutable actual input")
    retained = row.get("retained_pdf")
    pages = validate_pdf_observation(retained, renderer_path, expected_count=count, expected_outline=OUTLINE)
    need(isinstance(execution, dict) and execution.get("mode") == "1" and execution.get("status") == "complete"
         and execution.get("written_count") == 3 and type(execution["written_count"]) is int
         and execution.get("total_pages") == count and type(execution["total_pages"]) is int
         and execution.get("coverage") == plan["coverage"] and execution.get("source_identity") == plan.get("source_identity")
         and execution.get("original_ebook_identity") == original and execution.get("conversion") == plan.get("conversion")
         and execution.get("final_directory") == final.get("final_directory") and Path(execution["final_directory"]).parent == Path(row["output_base"])
         and re.fullmatch(r"[0-9a-f]{32}", execution.get("run_id", "")), "Execution changed captured source/plan/destination")
    conversion = execution["conversion"]
    need(conversion.get("output_profile") == "tablet" and conversion.get("original_source_identity") == original,
         "Fixed production tablet profile/original conversion identity differs")
    need(conversion.get("workspace_cleanup") == {"cleanup_complete": True, "retained_staging": None}
         and row.get("converter_process") == conversion, "Owned conversion workspace/process evidence differs")
    generated = conversion.get("generated_pdf_identity", {})
    need(generated.get("sha256") == retained["sha256"] and generated.get("size_bytes") == retained["size_bytes"]
         and generated.get("page_count") == count and generated.get("page_content_sha256") == [page["content_sha256"] for page in pages]
         and plan.get("source_identity", {}).get("sha256") == retained["sha256"] and plan["source_identity"].get("size_bytes") == retained["size_bytes"],
         "Actual full PDF does not match captured conversion/reader bytes and physical content")
    generated_path = Path(generated.get("path", ""))
    need(generated.get("binding") == "reader_snapshot" and original.get("binding") == "ebook_snapshot"
         and generated_path.name == "WinBookSplit_Converted.pdf" and generated_path.parent.parent == Path(row["output_base"])
         and re.fullmatch(r"\.WinBookSplit-stage-[0-9a-f]{32}", generated_path.parent.name)
         and plan["source_identity"].get("path") == str(generated_path), "Captured generated PDF escaped conversion run/source identity")
    validate_converter(conversion, str(calibre_path or row["parameters"][7]), row["input_path"], str(generated_path), success=True)
    kept = execution.get("retained_intermediate")
    need(isinstance(kept, dict) and kept.get("filename") == "WinBookSplit_Converted.pdf"
         and all(kept.get(key) == retained.get(key) for key in ("sha256", "size_bytes", "page_count"))
         and Path(retained["path"]) == Path(execution["final_directory"]) / kept["filename"], "Complete retained working PDF byte/path binding differs")
    entries, outputs, chapters = plan.get("entries"), execution.get("outputs"), row.get("chapters")
    need(isinstance(entries, list) and isinstance(outputs, list) and isinstance(chapters, list) and len(entries) == len(outputs) == len(chapters) == 3,
         "Exact nonempty chapter plan/output/reopen records missing")
    for number, (planned, output, chapter, bounds, title) in enumerate(zip(entries, outputs, chapters, RANGES[fmt], TITLES), 1):
        need(all(type(planned.get(key)) is int and type(output.get(key)) is int for key in ("start", "end", "sequence"))
             and [planned["start"], planned["end"]] == bounds and planned.get("sequence") == number and planned.get("title") == title
             and planned.get("filename") == f"{number:02d} - {title}.pdf" and all(output.get(key) == value for key, value in planned.items())
             and output.get("page_count") == bounds[1] - bounds[0] and type(output["page_count"]) is int
             and valid_hash(output.get("sha256")) and integer(output.get("size_bytes"), 1) and chapter.get("entry") == output,
             "Execution changed chapter plan/range/filename/count/hash")
        observed = chapter.get("observation")
        selected = validate_pdf_observation(observed, renderer_path, expected_count=bounds[1]-bounds[0], expected_outline=[{"title": title, "page": 0, "depth": 0}])
        need(Path(observed["path"]) == Path(execution["final_directory"]) / output["filename"]
             and all(observed[key] == output[key] for key in ("sha256", "size_bytes"))
             and observed.get("metadata", {}).get("/Title") == title and observed["metadata"].get("/Author") == "WinBookSplit test authors",
             "Reopened chapter byte identity/title/author differs")
        compare_physical_pages(pages, selected, *bounds)
    markers = re.findall(r"(?<![A-Z0-9-])WBS-PAGE-[0-9]{3}(?![0-9])", "\n".join(page["text"] for page in pages))
    need(markers == MARKERS, "Real complete conversion lost/duplicated original chapter markers")
    need(engine.get("warnings") == plan.get("warnings"), "Terminal planning warning snapshot changed")
    if fmt == "azw3":
        warnings = plan.get("warnings")
        need(isinstance(warnings, list) and len(warnings) == 1 and warnings[0].get("code") == "cross_chapter_link_dropped"
             and warnings[0].get("message") == "Count: 2. Internal links to pages outside this chapter are omitted."
             and "[WARNING] cross_chapter_link_dropped:" in row["stdout"] and "[WARNING] cross_chapter_link_dropped:" in row["console_log"],
             "Converted AZW3 inline TOC cross-chapter exclusions were not visibly reported")
    else:
        need(plan.get("warnings") == [], "Ordinary authored EPUB produced unexpected planning exclusions")
    publication = row.get("publication")
    need(isinstance(publication, dict), "Authenticated complete publication missing")
    tree = publication.get("tree")
    names = {".WinBookSplit-owner.json", "WinBookSplit_Manifest.json", kept["filename"], *(entry["filename"] for entry in outputs)}
    validate_tree(tree, path=execution["final_directory"], members=names)
    need(bind_raw(tree, ".WinBookSplit-owner.json", publication.get("owner_raw")) == {"schema_version": 1, "kind": "run", "run_id": execution["run_id"]}, "Published owner binding differs")
    manifest = bind_raw(tree, "WinBookSplit_Manifest.json", publication.get("manifest_raw"))
    need(manifest == execution.get("manifest") and all(manifest.get(key) == execution.get(key) for key in ("schema_version", "status", "run_id", "final_directory", "mode", "total_pages", "source_identity", "coverage", "written_count", "outputs", "original_ebook_identity", "conversion", "retained_intermediate")), "Complete publication manifest differs from execution")
    for entry in [*outputs, kept]:
        need(all(tree["members"][entry["filename"]][key] == entry[key] for key in ("sha256", "size_bytes")), "Authenticated publication member bytes differ")
    checked = row.get("checked_publication", {})
    need(checked.get("manifest_validated") is True and checked.get("retained_sha256_verified") is True
         and checked.get("retained_page_content_verified") is True and checked.get("retained_absent_verified") is False
         and checked.get("page_content_sha256") == generated["page_content_sha256"] and checked.get("chapter_markers") == MARKERS,
         "Reopened retained-publication coverage proof differs")
    need(row.get("failure_record") is None, "Successful ebook invented a failure directory")


def validate_failure(row, calibre_path):
    engine, final = row["engine_record"], row["outcome"]
    need(engine.get("execution") is None and engine.get("plan") is None and final.get("final_directory") is None
         and row.get("publication") is None and row.get("retained_pdf") is None and row.get("chapters") == [], "Malformed ebook reached chapter extraction/publication")
    diagnostic = engine.get("diagnostic", {})
    process = row.get("converter_process")
    need(isinstance(process, dict) and process == diagnostic.get("conversion") and type(process.get("exit_code")) is int and process["exit_code"] != 0
         and process.get("converter_path") == calibre_path and process.get("output_profile") == "tablet"
         and process.get("original_source_identity", {}).get("sha256") == row["source_observations_before"]["source"]["sha256"]
         and process["original_source_identity"].get("path") == row["input_path"], "Real malformed input did not bind actual Calibre nonzero diagnostics")
    argv = process.get("argv")
    need(isinstance(argv, list) and len(argv) == 5 and argv[:2] == [calibre_path, row["input_path"]]
         and Path(argv[2]).name == "WinBookSplit_Converted.pdf" and argv[-2:] == ["--output-profile", "tablet"], "Malformed converter literal tablet argv differs")
    generated_path = Path(argv[2])
    need(generated_path.parent.parent == Path(row["output_base"])
         and generated_path.parent.name == ".WinBookSplit-stage-" + diagnostic.get("run_id", "")
         and re.fullmatch(r"[0-9a-f]{32}", diagnostic.get("run_id", "")), "Malformed converter output escaped its actual owned stage/base")
    validate_converter(process, calibre_path, row["input_path"], argv[2], success=False)
    for name in ("stdout", "stderr"):
        need(isinstance(process.get(name + "_tail"), str) and integer(process.get(name + "_total_bytes"))
             and type(process.get(name + "_truncated")) is bool, "Actual bounded converter error stream missing")
    need(process["stdout_total_bytes"] + process["stderr_total_bytes"] > 0 and diagnostic.get("cleanup_complete") is True
         and diagnostic.get("retained_staging") is None and diagnostic.get("cleanup_error") is None,
         "Malformed converter output or owned workspace cleanup evidence missing")
    record = row.get("failure_record")
    need(isinstance(record, dict), "Owned actual conversion failure record missing")
    path = Path(record.get("path", ""))
    need(str(path) == diagnostic.get("record_path") and path.name == "failure.json" and path.parent.parent == Path(row["output_base"])
         and path.parent.name == ".WinBookSplit-failed-" + diagnostic.get("run_id", ""), "Failure record escaped actual output base/run identity")
    validate_tree(record.get("tree"), path=path.parent, members={".WinBookSplit-owner.json", "failure.json"})
    need(bind_raw(record["tree"], ".WinBookSplit-owner.json", record.get("owner_raw")) == {"schema_version": 1, "kind": "failed", "run_id": diagnostic["run_id"]}, "Conversion failure ownership marker differs")
    saved = bind_raw(record["tree"], "failure.json", record.get("raw"))
    need(saved.get("run_id") == diagnostic["run_id"] and saved.get("code") == "conversion_failed" and saved.get("conversion") == process,
         "Owned failure raw bytes differ from actual converter diagnostics")


def validate_converter(process, calibre, source, generated, *, success):
    need(isinstance(process, dict) and process.get("converter_path") == calibre and process.get("output_profile") == "tablet"
         and type(process.get("exit_code")) is int and ((process["exit_code"] == 0) if success else (process["exit_code"] != 0))
         and process.get("argv") == [calibre, source, generated, "--output-profile", "tablet"], "Actual Calibre executable/status/literal tablet argv differs")
    need(type(process.get("elapsed_seconds")) in (int, float) and math.isfinite(process["elapsed_seconds"]) and process["elapsed_seconds"] >= 0
         and integer(process.get("timeout_seconds"), 1) and process["elapsed_seconds"] <= process["timeout_seconds"] + 5,
         "Actual converter elapsed/deadline evidence missing")
    for name in ("stdout", "stderr"):
        tail, count, truncated = (process.get(name + suffix) for suffix in ("_tail", "_total_bytes", "_truncated"))
        need(isinstance(tail, str) and len(tail.encode("utf-8")) <= 196608 and integer(count) and type(truncated) is bool
             and truncated == (count > 65536) and min(count, 65536) <= len(tail.encode("utf-8")) <= 3 * min(count, 65536),
             "Actual bounded converter dual-stream evidence differs")


def validate_canary_asset(asset, urls):
    need(isinstance(asset, dict) and asset.get("authored_original_derivative") is True and asset.get("urls") == urls
         and valid_hash(asset.get("sha256")) and valid_hash(asset.get("source_epub_sha256")) and integer(asset.get("size_bytes"), 1), "Authored canary byte provenance missing")
    closure = asset.get("resource_closure", {})
    need(closure.get("sha256") == asset["sha256"] and closure.get("size_bytes") == asset["size_bytes"]
         and closure.get("resource_scope") == "exact-authored-loopback-canary" and closure.get("scripts_or_active_content") is False
         and closure.get("external_references") == sorted([{"from": "OEBPS/chapter1.xhtml", "reference": url} for url in urls], key=lambda value: (value["from"], value["reference"])),
         "Canary archive does not bind exactly two declared loopback elements")


def validate_block_diagnostics(row):
    urls, conversion, observed = row.get("canary_urls"), row.get("converter_process"), row.get("loopback_diagnostics")
    need(isinstance(urls, list) and len(urls) == 2 and isinstance(conversion, dict) and isinstance(observed, dict)
         and set(observed) == set(urls), "Actual loopback diagnostic endpoint bindings missing")
    for url in urls:
        actual = []
        for stream in ("stdout_tail", "stderr_tail"):
            need(isinstance(conversion.get(stream), str), "Actual converter diagnostic tail missing")
            for line in conversion[stream].splitlines():
                if url in line and "block" in line.lower():
                    actual.append({"stream": stream, "line": line})
        need(observed[url] == actual, "Loopback diagnostic line was substituted or omitted")
    need(observed[urls[0]] and any("Blocking URL request " + urls[0] in item["line"] for item in observed[urls[0]]),
         "Actual image endpoint blocking diagnostic missing; no broader network inference allowed")


def validate_canary(canary, asset):
    need(isinstance(canary, dict) and canary.get("listener_stopped") is True and canary.get("system_network_settings_changed") is False
         and canary.get("conversion_requests") == [], "Loopback listener requests/stop/settings evidence contradicts claim")
    binding = canary.get("binding", {})
    need(binding.get("host") == "127.0.0.1" and integer(binding.get("port"), 1) and binding["port"] < 65536
         and re.fullmatch(r"[0-9a-f]{32}", binding.get("nonce", "")), "Owned loopback namespace missing")
    base = f'http://127.0.0.1:{binding["port"]}/{binding["nonce"]}/'
    urls = [base + "image.png", base + "style.css"]
    need(canary.get("urls") == urls, "Loopback exact endpoints changed")
    validate_canary_asset(asset, urls)
    events, positives = canary.get("events"), canary.get("positive_controls")
    need(isinstance(events, list) and len(events) == 4 and isinstance(positives, list) and len(positives) == 4,
         "Same-endpoint before/after positive controls missing")
    # Bind positive responses to the fixed authored bytes, not merely each other.
    import base64
    image = base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aM3sAAAAASUVORK5CYII=")
    css = b"/* Original loopback canary; no imported resource. */\nbody { color: black; }\n"
    bodies = {"image.png": image, "style.css": css}
    for index, (event, phase, url) in enumerate(zip(events, ["before", "before", "after", "after"], urls * 2), 1):
        need(event.get("sequence") == index and type(event["sequence"]) is int and event.get("phase") == "control-" + phase
             and event.get("method") == "GET" and event.get("path") == urlsplit(url).path and event.get("body_length_declared") == 0
             and type(event["body_length_declared"]) is int, "Observed loopback event differs from authored positive endpoint")
        positive = positives[index-1]
        need(positive.get("phase") == phase and positive.get("url") == url and positive.get("method") == "GET"
             and positive.get("status") == 200 and type(positive["status"]) is int
             and positive.get("body_sha256") == sha256(bodies[Path(urlsplit(url).path).name]).hexdigest()
             and positive.get("size_bytes") == len(bodies[Path(urlsplit(url).path).name]) and type(positive["size_bytes"]) is int,
             "Positive server liveness/control response missing")
        started, finished, observed = utc(positive.get("started_at")), utc(positive.get("finished_at")), utc(event.get("observed_at"))
        need(started <= observed <= finished and (finished-started).total_seconds() <= 5,
             "Same-endpoint positive event was not observed during its bounded live request")
        if index > 1:
            need(utc(positives[index-2]["finished_at"]) <= started, "Positive loopback controls reused/reversed chronology")


def validate_canary_intervals(canary, rows):
    need(isinstance(rows, list) and len(rows) == 2, "Both actual canary intervals required")
    ordered = sorted(rows, key=lambda row: utc(row.get("started_at")))
    need(utc(ordered[0]["finished_at"]) <= utc(ordered[1]["started_at"]), "Canary native intervals contradict sequential host controls")
    positives = canary["positive_controls"]
    need(all(utc(item["finished_at"]) <= utc(ordered[0]["started_at"]) for item in positives if item["phase"] == "before")
         and all(utc(item["started_at"]) >= utc(ordered[-1]["finished_at"]) for item in positives if item["phase"] == "after"),
         "Loopback liveness controls did not bracket both actual native conversion intervals")


def validate_prior_binding(first, second):
    need(first["output_base"] == second["output_base"] and first["prior_publications_before"] == {},
         "Sequential format controls did not share one fresh base")
    publication, receipt = first["publication"]["tree"], first["console_evidence"]
    console_tree = {"path": str(Path(first["console_path"]).parent), **receipt["directory_identity"],
                    "members": receipt["file_identities"]}
    expected = {Path(publication["path"]).name: publication, Path(console_tree["path"]).name: console_tree}
    actual = second["prior_publications_before"]
    need(set(actual) == set(expected), "Second-format prior proof omitted the first output or finalized console")
    for name, tree in expected.items():
        observed = actual[name]
        need(all(observed[key] == tree[key] for key in ("path", "device", "inode")) and set(observed["members"]) == set(tree["members"])
             and all(all(observed["members"][member][key] == item[key] for key in ("sha256", "size_bytes", "device", "inode")) for member, item in tree["members"].items()),
             "Earlier real publication/console identity binding differs")


def validate_ebooks_report(report, requested_shell_paths, requested_calibre_path, requested_renderer_path, *, cleanup_complete=True, expected_source_sha256=None):
    need(isinstance(report, dict) and report.get("schema_version") == 1 and type(report["schema_version"]) is int
         and report.get("task_id") == "M4-T04" and report.get("result") == "EBOOK_REGRESSION_PASSED"
         and report.get("success") is True and type(report.get("exit_code")) is int and report["exit_code"] == 0
         and report.get("acceptance_ids") == ["AC-076", "AC-077", "AC-078"], "Ebook stage did not pass its exact scoped acceptance")
    source = report.get("tested_path_sha256")
    need(isinstance(source, dict) and source and all(valid_hash(value) for value in source.values()) and report.get("source_unchanged") is True,
         "Source-bound native ebook observations missing")
    if expected_source_sha256 is not None:
        need(source == expected_source_sha256, "Ebook report source differs from requested tested map")
    need(report.get("input_neighbor_prior_unchanged") is True and report.get("machine_settings_unchanged") is True
         and report.get("human_or_GUI_tested") is False and report.get("render_artifacts_retained") is True,
         "Scoped source/settings/artifact/agent-only evidence flags missing")
    if cleanup_complete:
        need(report.get("owned_temp_removed") is True, "Owned ebook temporary cleanup missing")
    hosts = report.get("host_cases")
    need(isinstance(hosts, list) and len(hosts) == 2 and {host.get("id") for host in hosts} == {"PS51", "PS7"}
         and {host.get("shell_executable") for host in hosts} == set(requested_shell_paths), "Both requested actual supported hosts required")
    for host in hosts:
        need(host.get("passed") is True and host.get("exit_code") == 0 and type(host["exit_code"]) is int
             and host.get("syntax_error_count") == 0 and type(host["syntax_error_count"]) is int
             and host.get("host_version", "").startswith("5.1." if host["id"] == "PS51" else "7.")
             and host.get("stored_policies") == host.get("policies_after"), "Actual host syntax/version/policy observation differs")
    checks, provenance = report.get("format_checks"), report.get("fixture_provenance", {})
    validate_format_checks(checks, provenance, report.get("resource_closure"))
    need(provenance.get("schema_version") == 1 and provenance.get("kind") == "original-offline-ebook-fixtures"
         and provenance.get("license") == "MIT" and provenance.get("authored_original") is True and provenance.get("remote_resources") is False
         and provenance.get("chapter_markers") == MARKERS and provenance.get("chapter_titles") == TITLES,
         "Original authored provenance changed/missing")
    calibre, renderer = report.get("calibre", {}), report.get("renderer", {})
    need(calibre == provenance.get("calibre") and calibre.get("path") == requested_calibre_path and calibre.get("version") == "9.15.0"
         and calibre.get("sha256") == CALIBRE_SHA256 and renderer.get("path") == requested_renderer_path and renderer.get("version") == "26.07.0"
         and valid_hash(renderer.get("sha256")), "Actual pinned converter/renderer identity differs")
    validate_process(calibre.get("version_observation"), command=[requested_calibre_path, "--version"])
    need("calibre 9.15.0" in calibre["version_observation"]["stdout"], "Pinned native Calibre version output missing")
    validate_process(renderer.get("probe"), command=[requested_renderer_path, "-v"])
    need("pdftoppm version 26.07.0" in renderer["probe"]["stderr"], "Pinned actual renderer version output missing")
    files = {item["format"]: item for item in provenance["files"]}
    generation = provenance.get("azw3_generation")
    need(isinstance(generation, dict), "Independent genuine AZW3 generation observation missing")
    directory = Path(generation.get("cwd", ""))
    validate_process(generation, command=[requested_calibre_path, str(directory / files["epub"]["name"]),
                                        str(directory / files["azw3"]["name"]), "--output-profile", "tablet"])
    need(generation.get("argv") == generation["command"], "Genuine AZW3 generation command alias differs")
    validate_references(report.get("references"), checks, requested_calibre_path, requested_renderer_path)
    before = report.get("original_inputs_before")
    need(isinstance(before, dict) and set(before) == {"epub", "azw3"} and before == report.get("original_inputs_after"), "Original ebooks changed during native matrix")
    for fmt in before:
        validate_identity(before[fmt])
        need(all(before[fmt][key] == checks[fmt][key] for key in ("sha256", "size_bytes")), "Original ebook format identity changed")
    tools = report.get("tools_before")
    need(isinstance(tools, dict) and set(tools) == {"calibre", "renderer"} and tools == report.get("tools_after"), "Native executable identity changed")
    for name in tools:
        validate_identity(tools[name])
        need(tools[name]["sha256"] == report[name]["sha256"], "Pinned executable bytes differ")
    cases = report.get("cases")
    ids = {host + "-" + fmt + "-" + suffix for host in ("PS51", "PS7") for fmt in ("epub", "azw3") for suffix in ("auto1-keep", "malformed")} | {host + "-epub-canary" for host in ("PS51", "PS7")}
    need(isinstance(cases, list) and len(cases) == 10 and {row.get("id") for row in cases} == ids
         and all(type(report.get(key)) is int and report[key] == count for key, count in (("case_count", 10), ("application_cases", 4), ("failure_cases", 4), ("canary_cases", 2))),
         "Exact ten real application controls/counts missing")
    by_id = {row["id"]: row for row in cases}
    for row in cases:
        validate_case(row, cleanup_complete, calibre_path=requested_calibre_path, renderer_path=requested_renderer_path)
        need(row["application_sha256"] == {name: source[name] for name in APPLICATION}, "Native copied application differs from tested source map")
        host = next(host for host in hosts if host["id"] == row["host_id"])
        need(row["shell_executable"] == host["shell_executable"] and row["host_version"] == host["host_version"]
             and row["parameters"][5] == report.get("environment", {}).get("python_executable"), "Native host/interpreter binding differs")
        if row["kind"] == "retained":
            need(all(row["source_observations_initial"]["source"][key] == checks[row["input_kind"]][key] for key in ("sha256", "size_bytes")), "Actual copied original format differs")
            reference = report["references"][row["input_kind"]]["tablet"]["observation"]
            compare_physical_pages(reference["pages"], row["retained_pdf"]["pages"], 0, len(reference["pages"]))
    for host in ("PS51", "PS7"):
        first, second = by_id[host + "-epub-auto1-keep"], by_id[host + "-azw3-auto1-keep"]
        validate_prior_binding(first, second)
    behavior = report.get("behavior", {})
    validate_config(behavior.get("reference_config"), expected_root=directory.parent / "reference-settings")
    help_records = behavior.get("local_help")
    need(isinstance(help_records, dict) and set(help_records) == {"epub", "azw3"}, "Actual selected-format help observations missing")
    for fmt, item in help_records.items():
        command = [requested_calibre_path, item["input_path"], str(Path(item["process"]["cwd"]) / ("help-only-" + fmt + ".pdf")), "--help"]
        validate_process(item["process"], command=command)
        need(item.get("output_absent") is True and item.get("input_before") == item.get("input_after") == before[fmt]
             and item.get("selected_profile") == "tablet" and item.get("profile_option_observed") == "--output-profile"
             and "--output-profile" in item["process"]["stdout"] and "tablet" in item["process"]["stdout"], "Local format help/profile/source proof differs")
    validate_canary(behavior.get("canary"), behavior.get("canary_asset"))
    validate_canary_intervals(behavior["canary"], [by_id[host + "-epub-canary"] for host in ("PS51", "PS7")])
    for host in ("PS51", "PS7"):
        row = by_id[host + "-epub-canary"]
        need(row["canary_asset"] == behavior["canary_asset"] and row["canary_urls"] == behavior["canary"]["urls"]
             and row["canary_asset"]["source_epub_sha256"] == checks["epub"]["sha256"], "Actual canary/observer/original archive identity differs")
    root = Path(report.get("render_directory", ""))
    ledger = report.get("retained_render_files")
    need(root.is_absolute() and not root.is_relative_to(ROOT) and isinstance(ledger, dict) and ledger, "Persistent external render ledger missing")
    observed = [ref["observation"] for profiles in report["references"].values() for ref in profiles.values()]
    observed += [obs for row in cases if row["kind"] != "malformed" for obs in [row["retained_pdf"], *(chapter["observation"] for chapter in row["chapters"])]]
    expected_files = {}
    for observation in observed:
        need(Path(observation["render_root"]).is_relative_to(root), "Actual PDF render escaped retained ledger root")
        for page in observation["pages"]:
            png = page["render"]
            need(png["path"] not in expected_files, "Actual page PNG reused across controls")
            expected_files[png["path"]] = {"sha256": png["png_sha256"], "size_bytes": png["bytes"]}
    for profiles in report["references"].values():
        for ref in profiles.values():
            observation = ref["observation"]
            need(Path(observation["path"]).is_relative_to(root), "Independent reference PDF escaped persistent ledger")
            expected_files[observation["path"]] = {key: observation[key] for key in ("sha256", "size_bytes")}
    need(set(ledger) == set(expected_files), "Retained render ledger omitted or added unexplained files")
    for path, values in expected_files.items():
        validate_identity(ledger[path])
        need(all(ledger[path][key] == value for key, value in values.items()), "Retained actual artifact byte hash differs")
