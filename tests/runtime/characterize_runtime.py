"""Actual Windows dependency selection, isolation and portable conversion checks."""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
from hashlib import sha256
import importlib.util
import json
import os
from pathlib import Path
import re
import shutil
import sys
import tempfile

from pypdf import PdfReader, PdfWriter

ROOT = Path(__file__).resolve().parents[2]
PYTHON_FAILURE_KINDS = {"missing", "non-python", "wrong-version", "missing-pypdf", "wrong-pypdf"}
PYTHON_SELECTION_KINDS = {"explicit-literal", "app-venv", "multiple-path", "py-launcher-listing"}
CONVERTER_FAILURE_KINDS = {"missing", "wrong-version", "book-cwd-decoys"}
CONVERTER_SUCCESS_KINDS = {"explicit-portable", "path-portable"}
ACCEPTANCE_IDS = ["AC-043", "AC-044", "AC-045"]
APPLICATION_FILES = ("VERSION", "WinBookSplit.ps1", "WinBookSplit.bat", "requirements.txt",
                     "engine/WinBookSplit.Paths.ps1", "engine/WinBookSplit.Diagnostics.ps1",
                     "engine/WinBookSplit.Runtime.ps1", "engine/winbooksplit_engine.py",
                     "engine/WinBookSplit.Process.ps1", "engine/WinBookSplit.Outcomes.json",
                     "engine/winbooksplit_windows.py", "engine/winbooksplit_conversion.py",
                     "engine/winbooksplit_job.py", "engine/WinBookSplit.Logging.ps1",
                     "engine/WinBookSplit.Support.ps1", "Export-WinBookSplitDiagnostics.ps1")


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


manual = load("wbs_runtime_manual", ROOT / "tests/manual/characterize_manual.py")
launchers = load("wbs_runtime_launchers", ROOT / "tests/manual/current_launchers.py")
paths = load("wbs_runtime_paths", ROOT / "tests/paths/characterize_paths.py")
conversion = load("wbs_runtime_conversion", ROOT / "tests/conversion/characterize_conversion.py")
runner = load("wbs_runtime_runner", ROOT / "tests/run_tests.py")
history, require = manual.history, manual.require


def digest(path):
    return sha256(path.read_bytes()).hexdigest()


def prefixed_json(text, prefix):
    return [json.loads(line[len(prefix):]) for line in text.splitlines() if line.startswith(prefix)]


def source_identity(path):
    details = path.lstat()
    return {"sha256": digest(path), "size_bytes": details.st_size, "device": details.st_dev,
            "inode": details.st_ino, "attributes": details.st_file_attributes}


def copy_application(target):
    target.mkdir()
    for name in APPLICATION_FILES:
        destination = target / name
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(ROOT / name, destination)
    return {name: digest(target / name) for name in APPLICATION_FILES}


def clone_venv(source, target, *, wrong_pypdf=False):
    """Relocate only an owned copy; never edit the pinned developer environment."""
    shutil.copytree(source, target)
    if wrong_pypdf:
        version_file = target / "Lib/site-packages/pypdf/_version.py"
        text = version_file.read_text(encoding="utf-8")
        changed, count = re.subn(r'(?m)^__version__\s*=\s*["\'][^"\']+["\']', '__version__ = "9.99.0"', text)
        require(count == 1, "Copied pypdf version fixture has changed its declaration")
        version_file.write_text(changed, encoding="utf-8")
    return target / "Scripts/python.exe"


def native_stub(work, ps51):
    source, executable = work / "Stub.cs", work / "Stub.exe"
    source.write_text('''using System;
using System.IO;
public class Stub {
    static void Mark(string name) { string path=Environment.GetEnvironmentVariable(name); if (!String.IsNullOrEmpty(path)) File.WriteAllText(path,"Owned native control reached"); }
    public static int Main(string[] args) {
        if (args.Length==1 && args[0]=="-0p") {
            Mark("WBS_LISTING_MARKER");
            Console.WriteLine(" -V:3.14 * " + Environment.GetEnvironmentVariable("WBS_LISTED_PYTHON")); return 0;
        }
        if (args.Length==1 && args[0]=="--version") {
            Mark("WBS_VERSION_MARKER"); Mark("WBS_DECOY_MARKER");
            Console.WriteLine("ebook-convert.exe (calibre " + (Environment.GetEnvironmentVariable("WBS_STUB_VERSION") ?? "9.15.0") + ")"); return 0;
        }
        Mark("WBS_NATIVE_MARKER"); Mark("WBS_DECOY_MARKER");
        if (args.Length>1 && args[0]!="-I") Mark("WBS_CONVERSION_MARKER");
        Console.WriteLine("OWNED-NON-PYTHON-NATIVE-CONTROL"); return 7;
    }
}''', encoding="utf-8")
    environment = {**history.clean_environment(work), "PSModulePath": str(ps51.parent / "Modules"),
                   "WBS_STUB_SOURCE": str(source), "WBS_STUB_EXE": str(executable)}
    command = [str(ps51), "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "RemoteSigned", "-Command",
               "Add-Type -Path $env:WBS_STUB_SOURCE -OutputAssembly $env:WBS_STUB_EXE -OutputType ConsoleApplication"]
    result = history.run(command, work, environment=environment)
    require(result["exit_code"] == 0 and executable.is_file(), "Owned native dependency control did not compile: " + repr(result))
    return executable, {"command": command, "cwd": str(work), **result,
                        "source_sha256": digest(source), "executable_sha256": digest(executable)}


def host_probe(work):
    probe = work / "Host.ps1"
    probe.write_text("[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)\n"
        "$total=0;$checked=@('WinBookSplit.ps1','engine/WinBookSplit.Paths.ps1','engine/WinBookSplit.Diagnostics.ps1','engine/WinBookSplit.Runtime.ps1','engine/WinBookSplit.Process.ps1')\n"
        "foreach($name in $checked){$tokens=$null;$errors=$null;[Management.Automation.Language.Parser]::ParseFile((Join-Path $env:WBS_PATHS_ROOT $name),[ref]$tokens,[ref]$errors)|Out-Null;$total+=$errors.Count}\n"
        "$policies=@(Get-ExecutionPolicy -List|Where-Object {$_.Scope -ne 'Process'}|ForEach-Object {[pscustomobject]@{scope=$_.Scope.ToString();policy=$_.ExecutionPolicy.ToString()}})\n"
        "[pscustomobject]@{host_major=$PSVersionTable.PSVersion.Major;host_version=$PSVersionTable.PSVersion.ToString();syntax_checked=$checked;syntax_error_count=$total;stored_policies=$policies}|ConvertTo-Json -Depth 6 -Compress\n", encoding="utf-8")
    return probe


def shadow_module(directory):
    target = directory / "pypdf.py"
    target.write_text('import os\nfrom pathlib import Path\nPath(os.environ["WBS_MODULE_MARKER"]).write_text("Owned module control imported")\nraise RuntimeError("OWNED-PYPDF-MODULE-DECOY")\n', encoding="utf-8")
    return target


def positive_decoy_controls(work, stub, core):
    native_marker = work / "positive-native.txt"
    native = history.run([str(stub), "--version"], work,
        environment={**history.clean_environment(work), "WBS_DECOY_MARKER": str(native_marker)})
    module_dir = work / "positive-module"
    module_dir.mkdir()
    module = shadow_module(module_dir)
    module_marker = work / "positive-module.txt"
    imported = history.run([str(core), "-B", "-c", "import pypdf"], work,
        environment={**history.clean_environment(work), "PYTHONPATH": str(module_dir), "WBS_MODULE_MARKER": str(module_marker)})
    require(native["exit_code"] == 0 and native_marker.is_file() and imported["exit_code"] != 0 and module_marker.is_file(),
            "Owned execution/import marker controls did not work")
    return {"native": {**native, "command": [str(stub), "--version"], "cwd": str(work),
                       "marker_path": str(native_marker), "marker_sha256": digest(native_marker), "control_reached": True},
            "module": {**imported, "command": [str(core), "-B", "-c", "import pypdf"], "cwd": str(work),
                       "controlled_PYTHONPATH": str(module_dir), "module_path": str(module), "module_sha256": digest(module),
                       "marker_path": str(module_marker), "marker_sha256": digest(module_marker), "control_reached": True}}


class Cases:
    def __init__(self, work, shells, python, calibre):
        self.work, self.shells, self.python, self.calibre = work, shells, python, calibre
        self.venv = python.parent.parent
        self.core = Path(next(line.split(" = ", 1)[1] for line in (self.venv / "pyvenv.cfg").read_text().splitlines() if line.startswith("home = "))) / "python.exe"
        self.wrong_python = Path(os.environ.get("WBS_TEST_UNSUPPORTED_PYTHON", str(Path(os.environ["LOCALAPPDATA"]) / "Python/pythoncore-3.14-64/python.exe")))
        require(self.wrong_python.is_file(), "Recorded real unsupported Python 3.14.7 control is required")
        self.wrong_observation = history.run([str(self.wrong_python), "-I", "-B", "-c", "import sys;print(sys.executable);print(sys.version)"], work)
        require(self.wrong_observation["exit_code"] == 0 and "3.14.7" in self.wrong_observation["stdout"], "Actual unsupported Python control changed")
        self.system = Path(os.environ["SystemRoot"])
        self.ps51 = self.system / "System32/WindowsPowerShell/v1.0/powershell.exe"
        self.stub, self.compilation = native_stub(work, self.ps51)
        self.controls = positive_decoy_controls(work, self.stub, self.core)
        venv_dir = work / "Python [1] O'Neil & (日本) %!"
        self.literal_python = clone_venv(self.venv, venv_dir)
        self.wrong_pypdf_python = clone_venv(self.venv, work / "wrong-pypdf", wrong_pypdf=True)
        self.fixture = work / "original.epub"
        conversion.fixtures.authored_epub(self.fixture)
        self.reference = conversion.reference_pdf(self.fixture, work / "reference.pdf", calibre, work)
        self.generator = history.load_generator()
        self.index = 0

    def one(self, identifier, shell, kind, *, category, batch=False):
        self.index += 1
        case = self.work / ("c" + str(self.index))
        case.mkdir()
        app, book, cwd, base = case / "a", case / "d", case / "c", case / "o"
        application_hashes = copy_application(app)
        for directory in (book, cwd, base):
            directory.mkdir()
        ebook = category in {"python_failure", "converter_failure", "converter_success"}
        source = book / ("Book [1] & %!.epub" if ebook else "Book [1] & %!.pdf")
        if ebook:
            source.write_bytes(self.fixture.read_bytes())
        else:
            source.write_bytes(self.generator._pdf_bytes({"pages": 3, "title": "Original dependency acceptance", "outline": []}))
        neighbor = source.with_suffix(".pdf") if ebook else book / "neighbor.txt"
        if ebook:
            with neighbor.open("xb") as stream:
                writer = PdfWriter()
                writer.add_blank_page(width=103, height=107)
                writer.write(stream)
        else:
            neighbor.write_bytes(b"Original owned source neighbor\n")
        originals = {str(path): source_identity(path) for path in (source, neighbor)}
        original_content = self.reference["page_content_sha256"] if ebook else conversion.page_content(PdfReader(source))
        version_marker, conversion_marker = case / "version-called.txt", case / "conversion-called.txt"
        native_marker, module_marker, decoy_marker = case / "native-called.txt", case / "module-called.txt", case / "decoy-called.txt"
        listing_marker, launcher_control = case / "listing-called.txt", None
        environment = history.clean_environment(case)
        parent_local_app_data = os.environ.get("LOCALAPPDATA")
        isolated_local_app_data = None
        if category == "converter_failure" and kind in {"missing", "book-cwd-decoys"}:
            # Only these absence controls hide an actual per-user installation.
            require(all(not path.exists() for path in (
                Path("C:/Program Files/Calibre2/ebook-convert.exe"),
                Path("C:/Program Files (x86)/Calibre2/ebook-convert.exe"))),
                "Missing-converter fixture requires both fixed system Calibre paths absent; do not alter an installed converter")
            isolated_local_app_data = case / "empty-local-app-data"
            isolated_local_app_data.mkdir()
            environment["LOCALAPPDATA"] = str(isolated_local_app_data)
        environment.update(PSModulePath=str(shell.parent / "Modules"), WBS_RUNTIME_APP=str(app / "WinBookSplit.ps1"),
            WBS_RUNTIME_BAT=str(app / "WinBookSplit.bat"), WBS_RUNTIME_INPUT=str(source), WBS_RUNTIME_BASE=str(base),
            WBS_VERSION_MARKER=str(version_marker), WBS_CONVERSION_MARKER=str(conversion_marker),
            WBS_NATIVE_MARKER=str(native_marker), WBS_MODULE_MARKER=str(module_marker))
        system_path = [str(self.ps51.parent), str(self.system / "System32"), str(self.system)]
        python_path, converter_path = None, None
        candidate, expected_source, expected_converter_source = str(self.python), "PATH", None
        path = [str(self.python.parent), *system_path]
        decoys = []
        if category == "python_failure":
            candidates = {"missing": case / "missing-python.exe", "non-python": self.stub,
                "wrong-version": self.wrong_python, "missing-pypdf": self.core, "wrong-pypdf": self.wrong_pypdf_python}
            python_path = candidates[kind]
            candidate = str(python_path)
            converter_path = self.stub
            expected_source = "explicit"
        elif kind == "explicit-literal":
            python_path, candidate, expected_source = self.literal_python, str(self.literal_python), "explicit"
            path = system_path
        elif kind == "app-venv":
            candidate = str(clone_venv(self.venv, app / ".venv"))
            expected_source = "application_venv"
            path = system_path
        elif kind == "multiple-path":
            path.insert(0, str(self.wrong_python.parent))
        elif kind == "py-launcher-listing":
            launcher_directory = case / "l"
            launcher_directory.mkdir()
            launcher_control = launcher_directory / "py.exe"
            shutil.copyfile(self.stub, launcher_control)
            environment.update(WBS_LISTED_PYTHON=str(self.python), WBS_LISTING_MARKER=str(listing_marker))
            path = [str(launcher_directory), *system_path]
            expected_source = "py_launcher"
        if category == "shadow":
            env_dir = case / "e"
            env_dir.mkdir()
            for directory in (book, cwd, env_dir):
                decoys.append(shadow_module(directory))
            for directory in (book, cwd):
                for name in ("python.exe", "py.exe"):
                    target = directory / name
                    shutil.copyfile(self.stub, target)
                    decoys.append(target)
            environment.update(PYTHONPATH=str(env_dir), WBS_DECOY_MARKER=str(decoy_marker))
            path = [str(book), str(cwd), str(self.python.parent), *system_path]
        if category == "pdf_without_calibre":
            converter_path = self.stub  # A real version-spy control must remain uncalled for PDFs.
        if category == "converter_failure":
            if kind == "missing":
                path = [str(self.python.parent), *system_path]
            elif kind == "wrong-version":
                converter_path = self.stub
                environment["WBS_STUB_VERSION"] = "9.14.0"
            else:
                for directory in (book, cwd):
                    target = directory / "ebook-convert.exe"
                    shutil.copyfile(self.stub, target)
                    decoys.append(target)
                environment["WBS_DECOY_MARKER"] = str(decoy_marker)
                path = [str(book), str(cwd), str(self.python.parent), *system_path]
        if category == "converter_success":
            if kind == "explicit-portable":
                converter_path, expected_converter_source = self.calibre, "explicit"
            else:
                path.insert(0, str(self.calibre.parent))
                expected_converter_source = "PATH"
        environment["PATH"] = os.pathsep.join(path)
        if python_path is not None:
            environment["WBS_RUNTIME_PYTHON"] = str(python_path)
        if converter_path is not None:
            environment["WBS_RUNTIME_CALIBRE"] = str(converter_path)
        wrapper = case / ("Invoke.cmd" if batch else "Invoke.ps1")
        if batch:
            # No CALL/re-expansion of percent or exclamation characters.
            wrapper.write_text('@echo off\nsetlocal DisableDelayedExpansion\n"%WBS_RUNTIME_BAT%" "%WBS_RUNTIME_INPUT%"\n', encoding="ascii", newline="\r\n")
            command = [str(self.system / "System32/cmd.exe"), "/d", "/v:off", "/c", str(wrapper)]
            observed = history.run([str(self.ps51), "-NoProfile", "-Command", "[Environment]::GetFolderPath('MyDocuments')"], case, environment=environment)
            require(observed["exit_code"] == 0 and observed["stdout"].strip(), "Actual BAT Documents base observation failed")
            base = Path(observed["stdout"].strip()).resolve(strict=True)
        else:
            wrapper.write_text('[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)\n'
                '$options=@{InputFile=$env:WBS_RUNTIME_INPUT;OutputDirectory=$env:WBS_RUNTIME_BASE}\n'
                'if($env:WBS_RUNTIME_PYTHON){$options.PythonPath=$env:WBS_RUNTIME_PYTHON}\n'
                'if($env:WBS_RUNTIME_CALIBRE){$options.CalibrePath=$env:WBS_RUNTIME_CALIBRE}\n'
                '& $env:WBS_RUNTIME_APP @options\nexit $LASTEXITCODE\n', encoding="utf-8")
            command = [str(shell), "-NoProfile", "-ExecutionPolicy", "RemoteSigned", "-File", str(wrapper)]
        decoy_hashes = {str(path): digest(path) for path in decoys}
        output_before = [] if not batch else None
        expected_failure = category in {"python_failure", "converter_failure"}
        process, log, console_owner, execution, members, record = None, None, None, None, None, None
        cleanup_safe = True
        try:
            with paths.read_only_source(source):
                readonly_observed = bool(source.lstat().st_file_attributes & 1)
                try:
                    answers = "" if expected_failure else ("M\nN\n2,3\nY\nN\n\n" if ebook else "M\n2,3\nY\nN\n\n")
                    process = history.run_entrypoint(command, cwd, environment=environment, stdin=answers, timeout=120)
                except history.EntryPointFailure as error:
                    cleanup_safe = error.cleanup_safe
                    raise
                except BaseException as error:
                    cleanup_safe = False
                    raise history.EntryPointFailure("Entrypoint interrupted; process tree state is uncertain: " + repr(error),
                                                    cleanup_safe=False, pid=None) from error
            require(process["exit_code"] == (3 if expected_failure else 0), "Actual runtime launcher native status differs: " + identifier + ": " + repr(process))
            if expected_failure:
                require(not any(base.iterdir()), "Dependency failure created output/console/stage members")
                errors = prefixed_json(process["stdout"], "[DEPENDENCY-ERROR] ")
                require(len(errors) == 1 and errors[0]["Code"] in {"runtime_invalid", "runtime_not_found", "converter_invalid", "converter_not_found"}, "Dependency failure omitted its actual rejection metadata")
                require("Dependency preflight failed" in process["stdout"] and "Done." not in process["stdout"]
                        and "[Writing]" not in process["stdout"] and "[CONVERSION]" not in process["stdout"]
                        and "[ENGINE] " not in process["stdout"] and "[DEPENDENCY] " not in process["stdout"], "Dependency failure announced processing/success")
                require(not conversion_marker.exists(), "Dependency rejection started conversion")
                if category == "python_failure":
                    require(not version_marker.exists() and "3.14.8" in process["stdout"] and "6.20.0" in process["stdout"], "Python failure attempted converter or omitted exact setup targets")
                else:
                    require("Calibre 9.15.0" in process["stdout"] and "-CalibrePath" in process["stdout"], "Converter failure omitted actionable exact setup guidance")
                record = {"dependency_error": errors[0], "candidate_path": candidate if category == "python_failure" else str(converter_path) if converter_path else None,
                    "preflight_before_output": True, "output_base_before": output_before, "output_base_after": [],
                    "conversion_started": False, "converter_version_marker_absent": not version_marker.exists(),
                    "engine_invocation_record_absent": True, "dependency_success_record_absent": True,
                    "no_success_summary": True, "written_count": 0, "outputs": [], "setup_guidance_verified": True}
            else:
                log = launchers.exact_line(process["stdout"], "Log: ")
                require(log.is_absolute() and log.name == "console.log" and log.parent.parent == base
                        and re.fullmatch(r"\.WinBookSplit-console-[0-9a-f]{32}", log.parent.name), "Selected-runtime console escaped its owned base")
                console_owner = json.loads((log.parent / ".WinBookSplit-console-owner.json").read_text(encoding="utf-8"))
                require(console_owner == {"run_id": log.parent.name.removeprefix(".WinBookSplit-console-"), "kind": "console"}, "Selected-runtime console owner differs")
                text = log.read_text(encoding="utf-8")
                dependency, invocation = prefixed_json(text, "[DEPENDENCY] "), prefixed_json(text, "[ENGINE] ")
                frames = [json.loads(line) for line in text.splitlines() if line.startswith("{")]
                require(len(dependency) == len(invocation) == len(frames) == 1 and frames[0]["status"] == "success", "Missing actual dependency/execution records")
                selected = dependency[0]["Runtime"]
                require(os.path.normcase(selected["Path"]) == os.path.normcase(candidate) and selected["Version"] == "3.14.8"
                        and selected["Source"] == expected_source and selected["PypdfVersion"] == "6.20.0", "Selected actual interpreter differs from case")
                require(invocation[0]["path"] == selected["Path"] and invocation[0]["arguments"][:4] == ["-I", "-B", "-X", "utf8"]
                        and invocation[0]["arguments"][4:9] == [str(app / "engine/winbooksplit_engine.py"), str(source), str(base), "manual", ""], "Engine did not use exact selected interpreter and literal isolated UTF-8 argv")
                require(selected["Details"]["isolated"] is True and selected["Details"]["dont_write_bytecode"] is True
                        and selected["Details"]["gil_disabled"] is False and selected["Details"]["bits"] == 64,
                        "Runtime probe did not verify required isolation/regular architecture")
                execution = frames[0]["execution"]
                require(launchers.exact_line(process["stdout"], "Output: ") == Path(execution["final_directory"]), "Actual successful output path differs")
                if ebook:
                    checked, members = conversion.check_publication(base, execution, self.reference["page_content_sha256"])
                    require(dependency[0]["Converter"]["Path"] == str(self.calibre)
                            and dependency[0]["Converter"]["Version"] == "9.15.0"
                            and dependency[0]["Converter"]["Source"] == expected_converter_source,
                            "Actual chosen portable converter differs")
                else:
                    actual = manual.published_outputs(base, self.generator, execution)
                    require([page for item in actual for page in item["page_ids"]] == [1, 2, 3], "Selected-runtime PDF lost physical page IDs")
                    members = {".WinBookSplit-owner.json", "WinBookSplit_Manifest.json", *[item["filename"] for item in actual]}
                    checked = {"outputs": actual, "page_content_sha256": [value for item in execution["outputs"] for value in
                        conversion.page_content(PdfReader(Path(execution["final_directory"]) / item["filename"]))], "manifest_validated": True}
                    require(checked["page_content_sha256"] == original_content, "Selected-runtime PDF content differs from original source pages")
                    require(dependency[0]["Converter"] is None and not version_marker.exists(), "PDF-only execution unnecessarily probed Calibre")
                record = {"dependency_selection": dependency[0], "engine_invocation": invocation[0], "expected_python_path": candidate,
                    "expected_python_source": expected_source, "expected_converter_source": expected_converter_source,
                    "writer_result": execution, "engine_record": frames[0], "written_count": execution["written_count"],
                    "selected_interpreter_used_for_engine": True, "all_physical_pages_preserved": True, **checked}
                record["independent_source_page_content_sha256"] = original_content
                interaction = launchers.interactions.capture(text)
                summaries = prefixed_json(text, "[PROCESS] ")
                launchers.interactions.validate(interaction, summaries, ["input_ready", "plan_ready"],
                    ["starts", "execute"], starts="2,3", execution=execution)
                record.update(interaction=interaction, process_summaries=summaries)
                record["console_evidence"] = launchers.authenticate_console_manifest(log,
                    source_path=source, source_observation=originals[str(source)])
                if launcher_control is not None:
                    listings = [item for item in selected["Attempts"] if item["Source"] == "py_launcher_listing" and item["Accepted"] is True]
                    require(len(listings) == 1 and listings[0]["Path"] == str(launcher_control)
                            and str(self.python) in listings[0]["Probe"]["Stdout"] and listing_marker.is_file(),
                            "Authored native read-only launcher listing control was not used")
                    record["launcher_listing_control"] = {"scope": "authored-native-py-listing-control; actual selected interpreter",
                        "path": str(launcher_control), "sha256": digest(launcher_control), "arguments": ["-0p"],
                        "listed_python_path": str(self.python), "marker_path": str(listing_marker),
                        "marker_sha256": digest(listing_marker), "control_reached": True}
            require(originals == {str(path): source_identity(path) for path in (source, neighbor)}, "Preflight/execution changed source or neighbor")
            require(decoy_hashes == {path: digest(Path(path)) for path in decoy_hashes}
                    and not decoy_marker.exists() and not module_marker.exists(), "Untrusted dependency decoy executed/imported or changed")
            require(application_hashes == {name: digest(app / name) for name in APPLICATION_FILES}, "Copied application source changed while running")
            if isolated_local_app_data is not None:
                require(isolated_local_app_data.is_dir() and not isolated_local_app_data.is_symlink()
                        and not isolated_local_app_data.is_junction() and not any(isolated_local_app_data.iterdir())
                        and os.environ.get("LOCALAPPDATA") == parent_local_app_data,
                        "Missing-converter child isolation or parent environment changed")
                record["converter_absence_environment"] = {
                    "child_local_app_data_isolated": True, "owned_directory_ordinary_and_empty": True,
                    "parent_local_app_data_unchanged": True, "fixed_system_calibre_paths_absent": True}
            record.update(id=identifier, category=category, kind=kind, passed=True, actual_process=True,
                shell_executable=str(shell), command=command, cwd=str(cwd), output_base=str(base),
                source_path=str(source), source_format=source.suffix[1:], source_read_only_attribute_observed=readonly_observed,
                source_attributes_restored=True, input_and_neighbor_identity=originals,
                input_and_neighbor_identity_after={str(path): source_identity(path) for path in (source, neighbor)},
                input_unchanged=True, neighbor_unchanged=True, decoy_sha256=decoy_hashes,
                decoy_sha256_after={path: digest(Path(path)) for path in decoy_hashes},
                decoy_marker_observation={"native_path": str(decoy_marker), "native_absent": not decoy_marker.exists(),
                    "module_path": str(module_marker), "module_absent": not module_marker.exists(),
                    "converter_version_path": str(version_marker), "converter_version_absent": not version_marker.exists(),
                    "conversion_path": str(conversion_marker), "conversion_absent": not conversion_marker.exists()},
                decoys_unchanged=True, decoy_execution_marker_absent=True, module_import_marker_absent=True,
                controlled_PATH=environment["PATH"], controlled_PYTHONPATH=environment.get("PYTHONPATH"),
                application_path_sha256=application_hashes, copied_application_unchanged=True, stdin_utf8=answers, **process)
        finally:
            if cleanup_safe:
                if execution is not None and members is not None:
                    launchers.remove_known_directory(Path(execution["final_directory"]), base, members, ".WinBookSplit-owner.json",
                        {"schema_version": 1, "kind": "run", "run_id": execution["run_id"]})
                if log is not None and console_owner is not None:
                    launchers.remove_console_directory(log, base, source_path=source, source_observation=originals[str(source)])
        if not batch:
            require(not any(base.iterdir()), "Known runtime outputs or unknown base members remain")
        record["owned_outputs_removed"] = True
        return record


def characterize(work, shells, calibre):
    manual.trusted_original_sources()
    source_before = runner.source_manifest()
    probe = host_probe(work)
    hosts = [paths.host_observation(shell, work, probe) for shell in shells]
    require({host["id"] for host in hosts} == {"PS51", "PS7"}, "Both actual supported hosts required")
    cases = Cases(work, shells, Path(sys.executable), calibre)
    groups = {name: [] for name in ("python_failure_cases", "python_selection_cases", "shadow_cases",
        "pdf_without_calibre_cases", "converter_failure_cases", "converter_success_cases")}
    for host in hosts:
        shell, host_id = Path(host["shell_executable"]), host["id"]
        for kind in sorted(PYTHON_FAILURE_KINDS):
            groups["python_failure_cases"].append(cases.one(host_id + "-python-" + kind, shell, kind, category="python_failure"))
        for kind in sorted(PYTHON_SELECTION_KINDS):
            groups["python_selection_cases"].append(cases.one(host_id + "-" + kind, shell, kind, category="python_selection"))
        groups["shadow_cases"].append(cases.one(host_id + "-book-cwd-env-decoys", shell, "book-cwd-env-decoys", category="shadow"))
        groups["pdf_without_calibre_cases"].append(cases.one(host_id + "-pdf-no-calibre", shell, "pdf-no-calibre", category="pdf_without_calibre"))
        for kind in sorted(CONVERTER_FAILURE_KINDS):
            groups["converter_failure_cases"].append(cases.one(host_id + "-converter-" + kind, shell, kind, category="converter_failure"))
        for kind in sorted(CONVERTER_SUCCESS_KINDS):
            groups["converter_success_cases"].append(cases.one(host_id + "-" + kind, shell, kind, category="converter_success"))
    groups["converter_success_cases"].append(cases.one("BAT-path-portable", cases.ps51, "path-portable", category="converter_success", batch=True))
    for host in hosts:
        after = paths.host_observation(Path(host["shell_executable"]), work, probe)
        require(after["stored_policies"] == host["stored_policies"], "Runtime preflight changed stored execution policies")
        host["policies_after"] = after["stored_policies"]
    require(source_before == runner.source_manifest(), "Source changed during runtime acceptance")
    return {"schema_version": 1, "task_id": "M2-T04", "result": "RUNTIME_REGRESSION_PASSED", "success": True, "exit_code": 0,
        "acceptance_ids": ACCEPTANCE_IDS, "observed_at": datetime.now(timezone.utc).isoformat(), "host_cases": hosts,
        **groups, "case_count": sum(len(value) for value in groups.values()), "native_control_compilation": cases.compilation,
        "positive_decoy_controls": cases.controls, "unsupported_real_python_observation": cases.wrong_observation,
        "independent_real_conversion_reference": cases.reference,
        "calibre": {"path": str(calibre), "version": "9.15.0", "sha256": digest(calibre)},
        "source_unchanged": True, "baseline_guards_preserved": True, "input_and_neighbor_unchanged": True,
        "machine_settings_unchanged": True, "immutable_original_commit": history.ORIGINAL_COMMIT,
        "tested_path_sha256": source_before, "environment": {"python": sys.version, "python_executable": sys.executable,
            "pypdf": history.pypdf.__version__},
        "not_run": ["human Explorer/interactive discovery", "clean OS/VM", "unsupported runtimes", "network/package installation", "release package"],
        "limits": ["Actual whole PS/BAT processes use controlled environment/stdin and original owned fixtures.",
            "Successful BAT uses observed Documents and only exact authenticated emitted final/log children; no Documents listing.",
            "Application venv and literal-path controls relocate owned copies of the pinned venv; the developer environment is unchanged.",
            "Wrong-pypdf controls alter only an owned copied package's declared version to 9.99.0; the pinned developer package is unchanged.",
            "Positive py_launcher cases use an authored native -0p listing control; they do not claim real manager registration or installation.",
            "ENGINE log rows record launch intent; successful actual engine frames and independently reopened outputs establish execution."]}


def main():
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
    sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--report", type=Path, required=True)
    parser.add_argument("--shell-path", type=Path, action="append", required=True)
    parser.add_argument("--calibre-path", type=Path, required=True)
    args = parser.parse_args()
    try:
        require(sys.flags.isolated and sys.dont_write_bytecode, "Use explicit developer Python -I -B")
        report_path = runner.new_external_path(args.report)
        require(len(args.shell_path) == 2 and all(path.is_absolute() and path.is_file() for path in args.shell_path), "Both actual absolute shell hosts required")
        require(args.calibre_path.is_absolute() and args.calibre_path.is_file()
                and digest(args.calibre_path) == conversion.fixtures.EXPECTED_CALIBRE_SHA256, "Actual pinned trusted Calibre required")
        temporary = tempfile.TemporaryDirectory(prefix="R-")
        directory = temporary.name
        try:
            report = characterize(Path(directory).resolve(), [path.resolve() for path in args.shell_path], args.calibre_path.resolve())
        except history.EntryPointFailure as error:
            if not error.cleanup_safe:
                temporary._finalizer.detach()
                print("Unproved process stop; owned acceptance workspace retained: " + directory, file=sys.stderr)
            else:
                temporary.cleanup()
            raise
        except BaseException:
            temporary.cleanup()
            raise
        else:
            temporary.cleanup()
        report["owned_temp_removed"] = not Path(directory).exists()
        validator = load("wbs_runtime_independent_validator", ROOT / "tests/runtime/validate_runtime_report.py")
        validator.validate_runtime_report(report, [str(path.resolve()) for path in args.shell_path], str(args.calibre_path.resolve()))
        with report_path.open("x", encoding="utf-8", newline="\n") as stream:
            json.dump(report, stream, indent=2, ensure_ascii=True)
            stream.write("\n")
        print("Runtime acceptance passed: 33 actual dependency/isolation/portable-converter cases under both hosts and BAT")
        return 0
    except (OSError, RuntimeError, ValueError) as error:
        print("Runtime acceptance failed: " + str(error), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
