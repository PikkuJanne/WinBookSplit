"""Invented ebook receipt controls; no native Calibre/Windows/PDF claims."""
from copy import deepcopy
from hashlib import sha256
import importlib.util
import json
from pathlib import Path
import sys
import unittest

ROOT = Path(__file__).resolve().parents[2]


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


guard = load("wbs_ebook_receipt_units", ROOT / "tests/ebooks/validate_ebooks_report.py")
fixture = load("wbs_ebook_console_units", ROOT / "tests/python/console_receipt_fixture.py")
behavior = load("wbs_ebook_behavior_units", ROOT / "tests/ebooks/behavior.py")
BASE = Path("C:/synthetic-ebook-unit")
RENDERER = str(BASE / "pdftoppm.exe")
CALIBRE = str(BASE / "ebook-convert.exe")


def identity(digest="1" * 64, size=200, inode=10, attributes=32):
    return {"sha256": digest, "size_bytes": size, "device": 1, "inode": inode, "attributes": attributes}


def process(command, cwd, code=0):
    return {"actual_process": True, "stdin": "DEVNULL", "command": command, "cwd": str(cwd), "exit_code": code,
            "timed_out": False, "pid": 123, "elapsed_seconds": 0.1, "timeout_seconds": 180,
            "started_at": "2026-10-10T00:00:15Z", "finished_at": "2026-10-10T00:00:16Z",
            "stdout": "", "stderr": "", "stdout_sha256": sha256(b"").hexdigest(), "stderr_sha256": sha256(b"").hexdigest()}


def config(directory):
    values = {variable: str(directory / name) for variable, name in zip(behavior.VARIABLES, ("config", "cache", "temp"))}
    snapshots = {key: {"path": value, "device": 1, "inode": number, "files": {}, "directories": {}}
                 for number, (key, value) in enumerate(values.items(), 1)}
    return {"environment": values, "before": deepcopy(snapshots), "after": deepcopy(snapshots),
            "new_empty_process_directories": True, "inherited_environment_restored": True,
            "original_environment": {key: None for key in values}, "restored_environment": {key: None for key in values}}


def pdf(path, root, numbers, outline, title=None):
    pages = []
    for local, original in enumerate(numbers, 1):
        text = f"WBS-PAGE-{original:03d}\n" if original <= 3 else "Authored inline table of contents\n"
        pages.append({"page": local, "text": text, "text_sha256": sha256(text.encode()).hexdigest(),
            "content_sha256": str(original) * 64, "geometry": deepcopy(guard.GEOMETRY),
            "render": {"page": local, "filename": f"page-{local}.png", "path": str(root / f"page-{local}.png"),
                       "bytes": 100, "width": 612, "height": 792, "png_sha256": str(original) * 64, "pixels_sha256": str(original + 1) * 64}})
    metadata = {"/Author": "WinBookSplit test authors"}
    if title is not None:
        metadata["/Title"] = title
    return {"path": str(path), "sha256": "8" * 64, "size_bytes": 200, "page_count": len(pages), "pages": pages,
            "outline": deepcopy(outline), "metadata": metadata, "render_root": str(root),
            "render_process": process([RENDERER, "-r", "72", "-cropbox", "-png", str(path), str(root / "page")], root)}


def seal_tree(path, raws, extra=None):
    members = {name: identity(sha256(raw.encode()).hexdigest(), len(raw.encode()), index)
               for index, (name, raw) in enumerate(raws.items(), 1)}
    members.update(extra or {})
    return {"path": str(path), "device": 1, "inode": 44, "members": members}


def reseal(row):
    engine, outcome = row["engine_record"], row["outcome"]
    outcome["engine_result"] = engine
    row["engine_records"] = [engine]
    if engine.get("execution"):
        execution = engine["execution"]
        execution["manifest"] = {key: deepcopy(value) for key, value in execution.items() if key not in {"manifest", "manifest_filename"}}
        raw = json.dumps(execution["manifest"])
        owner = json.dumps({"schema_version": 1, "kind": "run", "run_id": execution["run_id"]})
        extras = {item["filename"]: identity(item["sha256"], item["size_bytes"], index)
                  for index, item in enumerate([*execution["outputs"], execution["retained_intermediate"]], 10)}
        row["publication"] = {"manifest_raw": raw, "owner_raw": owner,
                              "tree": seal_tree(Path(execution["final_directory"]), {".WinBookSplit-owner.json": owner, "WinBookSplit_Manifest.json": raw}, extras)}
    elif row.get("failure_record"):
        failed = row["failure_record"]
        saved = json.loads(failed["raw"])
        saved.update(engine["diagnostic"])
        failed["raw"] = json.dumps(saved)
        failed["tree"] = seal_tree(Path(failed["path"]).parent, {".WinBookSplit-owner.json":failed["owner_raw"], "failure.json":failed["raw"]})
    text = "[PROCESS] " + json.dumps(row["process_summaries"][0]) + "\n" + json.dumps(engine) + "\n"
    if row["input_kind"] == "azw3" and row["kind"] != "malformed":
        text += "[WARNING] cross_chapter_link_dropped: Count: 2. Internal links to pages outside this chapter are omitted.\n"
    text += "[OPERATION-OUTCOME] " + json.dumps(outcome) + "\n"
    row["console_log"], row["console_log_sha256"] = text, sha256(text.encode()).hexdigest()
    receipt = fixture.make(outcome, row["console_log_sha256"], log_text=text, source_path=row["input_path"],
                           source_observation=row["source_observations_before"]["source"])
    receipt["run_manifest"]["settings"].update(input_kind=row["input_kind"], non_interactive=True, keep_converted_pdf=True)
    fixture.reseal(receipt)
    row["console_evidence"] = receipt
    log = Path(row["output_base"]) / (".WinBookSplit-console-" + receipt["owner"]["run_id"]) / "console.log"
    row["console_path"] = str(log)
    human = "" if row["kind"] == "malformed" else "Done.\nOutput: " + outcome["final_directory"] + "\n"
    if row["input_kind"] == "azw3" and row["kind"] != "malformed":
        human += "[WARNING] cross_chapter_link_dropped: Count: 2. Internal links to pages outside this chapter are omitted.\n"
    row["stdout"] = human + "Log: " + str(log) + "\n[OUTCOME] " + json.dumps(outcome) + "\n"
    row["stdout_sha256"] = sha256(row["stdout"].encode()).hexdigest()


def synthetic_row(kind="retained", fmt="epub"):
    suffix = "auto1-keep" if kind == "retained" else kind
    identifier = "PS51-" + fmt + "-" + suffix
    workspace, base = BASE / identifier, BASE / "o"
    source = workspace / ("b/Book Å 日本 [1] & %! $(literal)." + fmt)
    params = ["-InputFile", str(source), "-OutputDirectory", str(base), "-PythonPath", str(BASE / "python.exe"),
              "-CalibrePath", CALIBRE, "-Mode", "Auto", "-BookmarkLevel", "1", "-KeepConvertedPdf", "-NonInteractive"]
    shell = str(BASE / "powershell.exe")
    command = [shell, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "RemoteSigned", "-File", str(workspace / "a/WinBookSplit.ps1"), *params]
    native = 4 if kind == "malformed" else 0
    original = {"path": str(source), "resolved_path": str(source), "sha256": "1" * 64, "size_bytes": 200, "binding": "ebook_snapshot"}
    initial = {name: identity(inode=number) for number, name in enumerate(("source", "neighbor", "prior"), 1)}
    before = deepcopy(initial)
    before["source"]["attributes"] |= 1
    row = {**process(command, workspace / "c", native), "id": identifier, "kind": kind, "host_id": "PS51", "host_version": "5.1.26100.9444",
        "workspace": str(workspace), "output_base": str(base), "input_path": str(source), "input_kind": fmt, "shell_executable": shell,
        "parameters": params, "application_sha256": {name: "a" * 64 for name in guard.APPLICATION}, "application_unchanged": True,
        "source_observations_initial": initial, "source_observations_before": before, "source_observations_after": deepcopy(before),
        "source_observations_restored": deepcopy(initial), "source_observations_after_cleanup": deepcopy(initial),
        "source_attributes_restored": True, "source_read_only_observed": True, "owned_outputs_removed": True, "passed": True,
        "output_members_after_cleanup": ["prior-output.pdf"], "prior_publications_before": {}, "prior_publications_after": {},
        "calibre_environment": config(workspace / "settings"), "canary_urls": [], "retained_pdf": None, "chapters": [], "publication": None,
        "failure_record": None, "process_summaries": [{"JobAssigned": True, "ParentStopped": True, "DescendantsStopped": True,
            "StreamsComplete": True, "ExitCode": native, "ResultRecordCount": 1}]}
    conversion = {"converter_path": CALIBRE, "output_profile": "tablet", "original_source_identity": original,
        "argv": [CALIBRE, str(source), str(base / (".WinBookSplit-stage-"+"b"*32) / "WinBookSplit_Converted.pdf"), "--output-profile", "tablet"],
        "elapsed_seconds": 0.1, "timeout_seconds": 1800,
        "exit_code": 1 if native else 0, "stdout_tail": "Authored converter output", "stderr_tail": "Authored format error" if native else "",
        "stdout_total_bytes": 25, "stderr_total_bytes": 21 if native else 0, "stdout_truncated": False, "stderr_truncated": False}
    for name in ("stdout", "stderr"):
        conversion[name+"_total_bytes"] = len(conversion[name+"_tail"].encode("utf-8"))
    engine = {"protocol": "winbooksplit.result", "version": 1, "mode": "1", "status": "error" if native else "success",
        "code": "conversion_failed" if native else "split_complete", "exit_code": native, "written_count": 0 if native else 3,
        "warnings": [], "execution": None, "diagnostic": None}
    row["engine_record"] = engine
    row["outcome"] = {"protocol": "winbooksplit.outcome", "version": 1, "mode": "1", "status": engine["status"], "code": engine["code"],
                      "exit_code": native, "written_count": engine["written_count"], "final_directory": None}
    if native:
        run_id = "f" * 32
        conversion["argv"][2] = str(base / (".WinBookSplit-stage-"+run_id) / "WinBookSplit_Converted.pdf")
        path = base / (".WinBookSplit-failed-" + run_id) / "failure.json"
        diagnostic = {"run_id": run_id, "record_path": str(path), "cleanup_complete": True, "retained_staging": None,
                      "cleanup_error": None, "conversion": conversion}
        engine["diagnostic"] = diagnostic
        raw = json.dumps({"schema_version": 1, "status": "failed", "code": "conversion_failed", **diagnostic})
        owner = json.dumps({"schema_version": 1, "kind": "failed", "run_id": run_id})
        row["failure_record"] = {"path": str(path), "raw": raw, "owner_raw": owner,
                                "tree": seal_tree(path.parent, {".WinBookSplit-owner.json": owner, "failure.json": raw})}
    else:
        count = 3 if fmt == "epub" else 4
        final = base / "published"
        retained = pdf(final / "WinBookSplit_Converted.pdf", BASE / identifier / "renders/retained", range(1, count+1), guard.OUTLINE)
        row["retained_pdf"] = retained
        generated = {"path": conversion["argv"][2], "binding":"reader_snapshot", "sha256": retained["sha256"], "size_bytes": retained["size_bytes"], "page_count": count,
                     "page_content_sha256": [page["content_sha256"] for page in retained["pages"]]}
        conversion.update(generated_pdf_identity=generated, workspace_cleanup={"cleanup_complete": True, "retained_staging": None})
        warnings = [] if fmt == "epub" else [{"code": "cross_chapter_link_dropped", "source_order": None, "depth": None,
                       "message": "Count: 2. Internal links to pages outside this chapter are omitted."}]
        entries = [{"start": start, "end": end, "sequence": number, "title": title, "filename": f"{number:02d} - {title}.pdf"}
                   for number, ((start, end), title) in enumerate(zip(guard.RANGES[fmt], guard.TITLES), 1)]
        plan = {"mode": "1", "total_pages": count, "ranges": deepcopy(guard.RANGES[fmt]), "entries": entries,
                "coverage": {"complete": True, "covered_pages": count, "section_count": 3}, "original_ebook_identity": original,
                "source_identity": {"path": generated["path"], "sha256": retained["sha256"], "size_bytes": retained["size_bytes"], "binding": "reader_snapshot"},
                "conversion": conversion, "output_naming": {"resolved_base": str(base)}, "keep_converted_pdf": True,
                "warnings": warnings, "normalized_inputs": {"bookmarks": []}}
        outputs = [{**entry, "page_count": entry["end"]-entry["start"], "sha256": "8" * 64, "size_bytes": 200} for entry in entries]
        kept = {"filename": "WinBookSplit_Converted.pdf", **{key: retained[key] for key in ("sha256", "size_bytes", "page_count")}}
        execution = {"schema_version": 1, "status": "complete", "mode": "1", "total_pages": count, "run_id": "a" * 32,
                     "written_count": 3, "final_directory": str(final), "coverage": deepcopy(plan["coverage"]),
                     "source_identity": deepcopy(plan["source_identity"]), "original_ebook_identity": original, "conversion": conversion,
                     "outputs": outputs, "retained_intermediate": kept}
        engine.update(plan=plan, execution=execution, warnings=warnings)
        row["outcome"]["final_directory"] = str(final)
        row["chapters"] = [{"entry": entry, "observation": pdf(final / entry["filename"], BASE / identifier / ("renders/chapter-"+str(number)),
                                range(entry["start"]+1,entry["end"]+1), [{"title": entry["title"], "page": 0, "depth": 0}], entry["title"])}
                           for number, entry in enumerate(outputs, 1)]
        row["checked_publication"] = {"manifest_validated": True, "retained_sha256_verified": True, "retained_page_content_verified": True,
            "retained_absent_verified": False, "page_content_sha256": generated["page_content_sha256"], "chapter_markers": guard.MARKERS}
    row["converter_process"] = conversion
    reseal(row)
    return row


def synthetic_canary():
    nonce, port = "a" * 32, 1234
    urls = [f"http://127.0.0.1:{port}/{nonce}/" + name for name in ("image.png", "style.css")]
    closure = {"sha256": "1"*64, "size_bytes": 200, "resource_scope": "exact-authored-loopback-canary", "scripts_or_active_content": False,
               "external_references": sorted([{"from": "OEBPS/chapter1.xhtml", "reference": url} for url in urls], key=lambda item: item["reference"])}
    asset = {"authored_original_derivative": True, "urls": urls, "sha256": "1"*64, "source_epub_sha256": "2"*64, "size_bytes": 200, "resource_closure": closure}
    events, positives = [], []
    for number, (phase, url, body) in enumerate(zip(["before"]*2+["after"]*2, urls*2, [behavior.PNG, behavior.CSS]*2), 1):
        second = (10 if phase == "before" else 30) + ((number-1)%2)*2
        stamp = lambda offset:f"2026-10-10T00:00:{second+offset:02d}Z"
        events.append({"sequence": number, "phase": "control-"+phase, "method": "GET", "path": "/"+nonce+"/"+url.rsplit("/",1)[1],
                       "body_length_declared": 0, "observed_at":stamp(0)})
        positives.append({"phase": phase, "url": url, "method": "GET", "status": 200, "body_sha256": sha256(body).hexdigest(),
                          "size_bytes": len(body), "started_at":stamp(0), "finished_at":stamp(1)})
    return {"binding": {"host": "127.0.0.1", "port": port, "nonce": nonce}, "urls": urls, "events": events,
            "positive_controls": positives, "conversion_requests": [], "listener_stopped": True, "system_network_settings_changed": False}, asset


class EbookReceiptTests(unittest.TestCase):
    def check_mutations(self, factory, validator, mutations, resealer=None):
        validator(factory())
        for number, mutate in enumerate(mutations):
            row = factory()
            mutate(row)
            if resealer:
                resealer(row)
            with self.subTest(mutation=number), self.assertRaises((ValueError, KeyError, TypeError)):
                validator(row)

    def test_positive_invented_retained_and_malformed_rows(self):
        for kind in ("retained", "malformed"):
            for fmt in ("epub", "azw3"):
                guard.validate_case(synthetic_row(kind, fmt))

    def test_resealed_terminal_status_code_count_mode_and_native_contradictions(self):
        self.check_mutations(synthetic_row, guard.validate_case, [
            lambda r: r["outcome"].update(status="error"), lambda r: r["outcome"].update(code="preview_complete"),
            lambda r: r["outcome"].update(written_count=1), lambda r: r["engine_record"].update(mode="manual"),
            lambda r: r.update(exit_code=4), lambda r: r.update(exit_code=False)], reseal)

    def test_resealed_plan_source_coverage_range_and_destination_contradictions(self):
        self.check_mutations(synthetic_row, guard.validate_case, [
            lambda r: r["engine_record"]["plan"]["original_ebook_identity"].update(sha256="9"*64),
            lambda r: r["engine_record"]["plan"]["coverage"].update(complete=1),
            lambda r: r["engine_record"]["plan"]["ranges"][0].__setitem__(0,False),
            lambda r: r["engine_record"]["plan"]["entries"][0].update(start=1),
            lambda r: r["engine_record"]["execution"].update(final_directory=str(BASE/"foreign/published")),
            lambda r: r["engine_record"]["plan"]["output_naming"].update(resolved_base=str(BASE/"foreign"))], reseal)

    def test_reopened_text_pixels_metadata_and_outline_cannot_be_substituted(self):
        self.check_mutations(synthetic_row, guard.validate_case, [
            lambda r: r["chapters"][0]["observation"]["pages"][0]["render"].update(pixels_sha256="9"*64),
            lambda r: r["chapters"][0]["observation"]["pages"][0].update(text="foreign", text_sha256=sha256(b"foreign").hexdigest()),
            lambda r: r["chapters"][0]["observation"]["metadata"].update({"/Author":"foreign"}),
            lambda r: r["chapters"][0]["observation"]["outline"][0].update(page=1),
            lambda r: [page["geometry"]["media_box"].__setitem__(0,False) for observation in
                       [r["retained_pdf"], *(chapter["observation"] for chapter in r["chapters"])] for page in observation["pages"]],
            lambda r: r["retained_pdf"]["pages"][0]["geometry"].update(rotation=False)])

    def test_retained_pdf_hash_member_and_cleanup_contradictions(self):
        self.check_mutations(synthetic_row, guard.validate_case, [
            lambda r: r["retained_pdf"].update(sha256="9"*64),
            lambda r: r["publication"]["tree"]["members"].update({"foreign.pdf":identity()}),
            lambda r: r.update(owned_outputs_removed=False),
            lambda r: r["source_observations_after_cleanup"]["neighbor"].update(inode=999)])

    def test_real_converter_failure_frame_must_keep_no_plan_no_outputs(self):
        self.check_mutations(lambda:synthetic_row("malformed"), guard.validate_case, [
            lambda r: r["engine_record"].update(plan={}), lambda r: r["engine_record"].update(execution={}),
            lambda r: r["converter_process"].update(exit_code=0), lambda r: r["converter_process"]["argv"].__setitem__(4,"default"),
            lambda r: r["converter_process"]["argv"].__setitem__(2,str(BASE/"foreign/WinBookSplit_Converted.pdf")),
            lambda r: r["engine_record"]["diagnostic"].update(cleanup_complete=False)], reseal)

    def test_source_readonly_prior_identity_and_transport_stop_proofs(self):
        self.check_mutations(synthetic_row, guard.validate_case, [
            lambda r: r["source_observations_after"]["source"].update(sha256="9"*64),
            lambda r: r["source_observations_before"]["source"].update(attributes=32),
            lambda r: r["prior_publications_after"].update({"foreign": {}}),
            lambda r: r["process_summaries"][0].update(DescendantsStopped=False),
            lambda r: r.update(calibre_environment=config(BASE/"foreign/settings")),
            lambda r: r["process_summaries"][0].update(ReplyCount=1)], reseal)

    def test_config_keeps_original_environment_and_owned_directory_identity(self):
        self.check_mutations(lambda:config(BASE/"settings"), guard.validate_config, [
            lambda c: c["restored_environment"].update(CALIBRE_CONFIG_DIRECTORY="foreign"),
            lambda c: c["after"]["CALIBRE_CACHE_DIRECTORY"].update(inode=999),
            lambda c: c["before"]["CALIBRE_TEMP_DIR"]["files"].update({"prior":identity()}),
            lambda c: c["after"]["CALIBRE_TEMP_DIR"]["files"].update({"../foreign":identity()})])

    def test_canary_exact_four_events_reject_delayed_idle_fetch_and_dead_positive(self):
        validator = lambda pair:guard.validate_canary(*pair)
        self.check_mutations(synthetic_canary, validator, [
            lambda p: p[0]["events"].append({"phase":"idle","method":"GET"}),
            lambda p: p[0]["events"][2].update(phase="idle"),
            lambda p: p[0]["positive_controls"][0].update(body_sha256="9"*64),
            lambda p: p[0].update(listener_stopped=False),
            lambda p: p[0]["conversion_requests"].append({"method":"GET"})])

    def test_image_block_diagnostic_requires_exact_native_tail_and_url(self):
        canary, _ = synthetic_canary()
        url = canary["urls"][0]
        line = "Blocking URL request " + url + " as it is not for a resource in the book"
        row = {"canary_urls":canary["urls"], "converter_process":{"stdout_tail":line,"stderr_tail":""},
               "loopback_diagnostics":{url:[{"stream":"stdout_tail","line":line}],canary["urls"][1]:[]}}
        guard.validate_block_diagnostics(row)
        for change in (lambda r:r["converter_process"].update(stdout_tail=""),
                       lambda r:r["loopback_diagnostics"][url][0].update(line="invented block"),
                       lambda r:r["loopback_diagnostics"].__setitem__(canary["urls"][1],[{"stream":"stdout_tail","line":line}])):
            mutated = deepcopy(row)
            change(mutated)
            with self.assertRaises(ValueError):
                guard.validate_block_diagnostics(mutated)

    def test_raw_stream_hash_and_literal_native_argv_cannot_be_substituted(self):
        self.check_mutations(synthetic_row, guard.validate_case, [
            lambda r:r.update(stderr="foreign"), lambda r:r["command"].__setitem__(6,str(BASE/"foreign.ps1")),
            lambda r:r["parameters"].__setitem__(11,"2"), lambda r:r.update(actual_process=False)])

    def test_resealed_converter_options_status_and_bounded_streams(self):
        self.check_mutations(synthetic_row, guard.validate_case, [
            lambda r:r["converter_process"]["argv"].__setitem__(4,"default"),
            lambda r:r["converter_process"].update(converter_path=str(BASE/"foreign.exe")),
            lambda r:r["converter_process"].update(exit_code=1),
            lambda r:r["converter_process"].update(stdout_total_bytes=True),
            lambda r:r["converter_process"].update(stdout_truncated=True),
            lambda r:r["converter_process"].update(stdout_tail=""),
            lambda r:r["converter_process"].update(stdout_total_bytes=1, stdout_tail="x"*100),
            lambda r:r["converter_process"].update(timeout_seconds=0)], reseal)

    def test_missing_reversed_native_times_fail(self):
        self.check_mutations(synthetic_row, guard.validate_case, [
            lambda r:r.pop("started_at"), lambda r:r.update(finished_at="2026-10-09T00:00:00Z"),
            lambda r:r.update(started_at="2026-10-10T00:00:15"), lambda r:r.update(finished_at="2026-10-10T00:05:15Z")])

    def test_liveness_controls_must_bracket_actual_native_canary_intervals(self):
        canary, asset = synthetic_canary()
        rows = [{"started_at":"2026-10-10T00:00:15Z","finished_at":"2026-10-10T00:00:16Z"},
                {"started_at":"2026-10-10T00:00:20Z","finished_at":"2026-10-10T00:00:21Z"}]
        guard.validate_canary(canary,asset)
        guard.validate_canary_intervals(canary,rows)
        for change in (lambda c:c["positive_controls"][0].pop("started_at"),
                       lambda c:c["positive_controls"][0].update(finished_at="2026-10-09T00:00:00Z"),
                       lambda c:c["events"][0].update(observed_at="2026-10-10T00:00:09Z"),
                       lambda c:c["positive_controls"][2].update(started_at="2026-10-10T00:00:11Z",finished_at="2026-10-10T00:00:12Z")):
            changed=deepcopy(canary)
            change(changed)
            with self.assertRaises(ValueError):
                guard.validate_canary(changed,asset)
                guard.validate_canary_intervals(changed,rows)
        changed=deepcopy(rows)
        changed[1]["finished_at"]="2026-10-10T00:00:31Z"
        with self.assertRaises(ValueError):
            guard.validate_canary_intervals(canary,changed)

    def test_second_format_prior_snapshot_binds_actual_first_publication_and_console(self):
        first, second = synthetic_row(), synthetic_row(fmt="azw3")
        publication, receipt = first["publication"]["tree"], first["console_evidence"]
        prior_console = {"path":str(Path(first["console_path"]).parent), **receipt["directory_identity"],
                         "members":deepcopy(receipt["file_identities"])}
        second["prior_publications_before"] = {Path(publication["path"]).name:deepcopy(publication),
                                               Path(prior_console["path"]).name:prior_console}
        guard.validate_prior_binding(first,second)
        for change in (lambda r:r["prior_publications_before"].pop(Path(prior_console["path"]).name),
                       lambda r:r["prior_publications_before"][Path(publication["path"]).name]["members"]["01 - Chapter One.pdf"].update(inode=999),
                       lambda r:r["prior_publications_before"][Path(prior_console["path"]).name]["members"]["WinBookSplit_Run.json"].update(sha256="9"*64)):
            changed=deepcopy(second)
            change(changed)
            changed["prior_publications_after"]=deepcopy(changed["prior_publications_before"])
            with self.assertRaises(ValueError):
                guard.validate_prior_binding(first,changed)


if __name__ == "__main__":
    unittest.main(verbosity=2)
