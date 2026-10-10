"""Actual Git-blob prerequisite checks; no archive or release acceptance claim."""
import importlib.util
import io
import json
import os
from pathlib import Path
import subprocess
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[2]
SPEC = importlib.util.spec_from_file_location("wbs_ci_package_inputs", ROOT / "tools/ci/check_package_inputs.py")
checker = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(checker)


class PackageInputTests(unittest.TestCase):
    def setUp(self):
        self.owned = tempfile.TemporaryDirectory(prefix="WinBookSplit-package-inputs-")
        self.addCleanup(self.owned.cleanup)
        self.work = Path(self.owned.name)
        self.repo = self.work / "repo"
        self.repo.mkdir()
        self.git("init", "--quiet")
        for path in checker.INPUT_PATHS:
            target = self.repo / path
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes((ROOT / path).read_bytes())
        self.commit = self.commit_all()

    def git(self, *args, data=None):
        result = subprocess.run(["git", "-c", "core.autocrlf=false", "-C", str(self.repo), *args], input=data,
                                stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=False, timeout=30)
        self.assertEqual(result.returncode, 0, "Owned fixture Git operation failed")
        return result.stdout.decode("utf-8").strip()

    def commit_all(self):
        self.git("add", "--all")
        self.git("-c", "user.name=Authored input test", "-c", "user.email=authored@example.invalid",
                 "commit", "--quiet", "-m", "Authored prerequisite fixture")
        return self.git("rev-parse", "HEAD")

    def rejected(self, code, commit=None):
        with self.assertRaises(checker.InputError) as caught:
            checker.validate_inputs(self.repo, commit or self.commit)
        self.assertEqual(caught.exception.code, code)

    def test_current_committed_inputs_and_helper_closure_are_bound(self):
        result = checker.validate_inputs(self.repo, self.commit)
        self.assertTrue(result["success"])
        self.assertEqual((result["runtime_count"], result["support_source_count"]), (15, 8))
        self.assertEqual(result["source_commit"], self.commit)
        self.assertEqual(result["source_tree"], self.git("rev-parse", "HEAD^{tree}"))
        self.assertFalse(result["working_tree_inputs_used"])
        self.assertFalse(result["package_built"])
        self.assertEqual({row["path"] for row in result["files"]}, set(checker.INPUT_PATHS))
        dependencies = result["runtime_sibling_dependencies"]
        self.assertIn("engine/winbooksplit_conversion.py", dependencies["engine/winbooksplit_engine.py"])
        self.assertIn("engine/winbooksplit_job.py", dependencies["engine/winbooksplit_conversion.py"])
        self.assertIn("engine/WinBookSplit.Support.ps1", dependencies["Export-WinBookSplitDiagnostics.ps1"])

    def test_dirty_and_untracked_bytes_are_never_adopted_or_reported(self):
        before = checker.validate_inputs(self.repo, self.commit)
        sentinel = "AUTHORED_PRIVATE_SENTINEL_NOT_A_PACKAGE_INPUT"
        (self.repo / "engine/winbooksplit_engine.py").write_text(sentinel, encoding="utf-8")
        (self.repo / "private-document.txt").write_text(sentinel, encoding="utf-8")
        after = checker.validate_inputs(self.repo, self.commit)
        self.assertEqual(before, after)
        self.assertNotIn(sentinel, json.dumps(after))
        self.assertNotIn(str(self.repo), json.dumps(after))

    def test_inherited_git_selectors_cannot_redirect_the_explicit_repository(self):
        baseline = checker.validate_inputs(self.repo, self.commit)
        foreign = self.work / "foreign"
        foreign.mkdir()
        created = subprocess.run(["git", "-c", "core.autocrlf=false", "-C", str(foreign), "init", "--quiet"],
                                 stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=False, timeout=30)
        self.assertEqual(created.returncode, 0)
        selectors = {"GIT_DIR": str(foreign / ".git"), "GIT_WORK_TREE": str(foreign),
                     "GIT_COMMON_DIR": str(foreign / ".git"), "GIT_INDEX_FILE": str(foreign / "foreign-index"),
                     "GIT_OBJECT_DIRECTORY": str(foreign / ".git/objects"),
                     "GIT_ALTERNATE_OBJECT_DIRECTORIES": str(foreign / ".git/objects"), "GIT_NAMESPACE": "authored"}
        with patch.dict(os.environ, selectors):
            redirected = subprocess.run(["git", "-C", str(self.repo), "rev-parse", "--show-toplevel"],
                                        stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=False, timeout=30)
            self.assertEqual(redirected.returncode, 0)
            self.assertEqual(Path(redirected.stdout.decode("utf-8").strip()).resolve(), foreign.resolve())
            self.assertEqual(checker.validate_inputs(self.repo, self.commit), baseline)

    def test_missing_committed_helper_is_not_rescued_by_working_file(self):
        path = "engine/winbooksplit_job.py"
        self.git("rm", "--cached", "--", path)
        self.git("-c", "user.name=Authored input test", "-c", "user.email=authored@example.invalid",
                 "commit", "--quiet", "-m", "Authored missing helper")
        self.assertTrue((self.repo / path).is_file())
        self.rejected("missing_input", self.git("rev-parse", "HEAD"))

    def test_symlink_and_gitlink_modes_cannot_be_runtime_inputs(self):
        path = "engine/winbooksplit_job.py"
        for mode in ("120000", "160000"):
            with self.subTest(mode=mode):
                oid = self.commit if mode == "160000" else self.git("hash-object", "-w", "--stdin", data=b"foreign-helper.py")
                self.git("update-index", "--add", "--cacheinfo", f"{mode},{oid},{path}")
                self.git("-c", "user.name=Authored input test", "-c", "user.email=authored@example.invalid",
                         "commit", "--quiet", "-m", "Authored nonordinary helper")
                self.rejected("nonordinary_input", self.git("rev-parse", "HEAD"))

    def test_unsafe_windows_and_traversal_names_are_rejected(self):
        for path in ("../engine/x.py", "/absolute", "C:/file", "engine\\x.py", "engine//x.py",
                     "engine/CON.py", "engine/a.", "engine/a ", "engine/a:b", "engine/a\n.py", "./README.md"):
            with self.subTest(path=repr(path)):
                self.assertFalse(checker.safe_relative(path))
        self.assertTrue(checker.safe_relative("engine/winbooksplit_job.py"))
        for target in ("../private.py", "C:/private.py", "private.py"):
            with self.subTest(target=target):
                with self.assertRaises(checker.InputError):
                    checker._dependency("engine/winbooksplit_engine.py", target)

    def test_new_untracked_sibling_and_external_imports_fail_closed(self):
        original = (self.repo / "engine/winbooksplit_engine.py").read_text(encoding="utf-8")
        additions = (
            "\nPath(__file__).resolve().with_name('private_helper.py')\n",
            "\nimport reportlab\n", "\nfrom private_helper import secret\n",
            "\nimportlib.import_module('private_helper')\n",
            "\nimportlib.util.spec_from_file_location('private', 'private_helper.py')\n",
        )
        (self.repo / "engine/private_helper.py").write_text("# Authored untracked decoy", encoding="utf-8")
        for addition in additions:
            with self.subTest(addition=addition):
                (self.repo / "engine/winbooksplit_engine.py").write_text(original + addition, encoding="utf-8")
                # Leave the decoy untracked while committing only the changed engine.
                self.git("add", "--", "engine/winbooksplit_engine.py")
                self.git("-c", "user.name=Authored input test", "-c", "user.email=authored@example.invalid",
                         "commit", "--quiet", "-m", "Authored undeclared import")
                with self.assertRaises(checker.InputError) as caught:
                    checker.validate_inputs(self.repo, self.git("rev-parse", "HEAD"))
                self.assertIn(caught.exception.code, {"undeclared_runtime_dependency", "unresolved_runtime_dependency"})

    def test_powershell_sibling_and_batch_launcher_changes_are_rejected(self):
        original = (self.repo / "WinBookSplit.ps1").read_text(encoding="utf-8")
        (self.repo / "WinBookSplit.ps1").write_text(original + "\n. (Join-Path $PSScriptRoot 'private-helper.ps1')\n", encoding="utf-8")
        self.rejected("undeclared_runtime_dependency", self.commit_all())
        (self.repo / "WinBookSplit.ps1").write_text(original, encoding="utf-8")
        (self.repo / "WinBookSplit.bat").write_text('@echo off\nPowerShell -File "%~dp0private-helper.ps1"\n', encoding="utf-8")
        self.rejected("launcher_dependency_changed", self.commit_all())

    def test_runtime_requirement_rejects_unpinned_dev_extra_url_and_include(self):
        for line in ("pypdf==6.19.0", checker.RUNTIME_REQUIREMENT + "\nreportlab==5.0.1",
                     "-r requirements-dev.txt", "pypdf[crypto]==6.19.0", "https://example.invalid/private.whl"):
            with self.subTest(line=line):
                (self.repo / "requirements.txt").write_text(line + "\n", encoding="utf-8")
                self.rejected("runtime_dependency_pin_changed", self.commit_all())

    def test_empty_or_oversized_input_is_rejected_before_blob_read(self):
        path = self.repo / "README.md"
        for size in (0, checker.MAX_FILE_BYTES + 1):
            with self.subTest(size=size):
                path.write_bytes(b"x" * size)
                self.rejected("invalid_input_size", self.commit_all())

    def test_explicit_commit_and_new_external_report_are_required(self):
        for commit in ("HEAD", "main", self.commit[:8], "-bad", "f" * 40):
            with self.subTest(commit=commit):
                with self.assertRaises(checker.InputError):
                    checker.validate_inputs(self.repo, commit)
        with self.assertRaises(checker.InputError):
            checker.validate_inputs(self.repo / "engine", self.commit)
        external = self.work / "inputs.json"
        arguments = ["--repo", str(self.repo), "--commit", self.commit, "--report", str(external)]
        with redirect_stdout(io.StringIO()), redirect_stderr(io.StringIO()):
            self.assertEqual(checker.main(arguments), 0)
            original = external.read_bytes()
            self.assertEqual(checker.main(arguments), 1)
            self.assertEqual(external.read_bytes(), original)
            inside = self.repo / "inputs.json"
            self.assertEqual(checker.main([*arguments[:-1], str(inside)]), 1)
            self.assertFalse(inside.exists())

    def test_failed_check_emits_nonzero_safe_report_without_git_stderr(self):
        report = self.work / "failed.json"
        with redirect_stdout(io.StringIO()) as output, redirect_stderr(io.StringIO()) as errors:
            status = checker.main(["--repo", str(self.repo), "--commit", "f" * 40, "--report", str(report)])
        result = json.loads(report.read_text(encoding="utf-8"))
        self.assertEqual(status, 1)
        self.assertFalse(result["success"])
        self.assertEqual(result["error_code"], "git_query_failed")
        self.assertNotIn(str(self.repo), report.read_text(encoding="utf-8") + output.getvalue() + errors.getvalue())


if __name__ == "__main__":
    unittest.main()
