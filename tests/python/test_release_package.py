"""Owned real-Git release-package tests; no Windows smoke/publication claim."""
from __future__ import annotations

from copy import copy
from hashlib import sha256
import importlib.util
import io
import json
import os
from pathlib import Path
import shutil
import stat
import struct
import subprocess
import sys
import tempfile
import unittest
import warnings
import zipfile
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[2]
BUILDER = ROOT / "tools/release/build_package.py"
SPEC = importlib.util.spec_from_file_location("wbs_release_package", BUILDER)
builder = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(builder)
INPUT_SPEC = importlib.util.spec_from_file_location("wbs_release_package_inputs", ROOT / "tools/ci/check_package_inputs.py")
checker = importlib.util.module_from_spec(INPUT_SPEC)
INPUT_SPEC.loader.exec_module(checker)

PAYLOAD_PATHS = tuple(sorted(set(checker.INPUT_PATHS) - {"docs/codex-v1.0.0/SUPPORT_AND_SETUP.md"}))
BUILDER_PATHS = ("tools/release/build_package.py", "tools/release/payload.json", "tools/ci/check_package_inputs.py")
ASSET_NAMES = ("WinBookSplit-v1.0.0.zip", "release-manifest.json", "SHA256SUMS.txt")


def fixture_git(repo, *args, data=None):
    environment = {key: value for key, value in os.environ.items() if key.upper() not in checker.REPOSITORY_ENV_KEYS}
    environment["GIT_NO_REPLACE_OBJECTS"] = "1"
    process = subprocess.run(["git", "-c", "core.autocrlf=false", "-C", str(repo), *args], input=data,
                             stdout=subprocess.PIPE, stderr=subprocess.PIPE, env=environment,
                             timeout=30, check=False)
    if process.returncode:
        raise AssertionError("Owned fixture Git operation failed")
    return process.stdout


def fixture_commit(repo):
    fixture_git(repo, "add", "--all")
    fixture_git(repo, "-c", "user.name=Authored package test", "-c", "user.email=authored@example.invalid",
                "commit", "--quiet", "-m", "Authored release-package fixture")
    return fixture_git(repo, "rev-parse", "HEAD").decode("ascii").strip()


class ReleasePackageTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.shared = tempfile.TemporaryDirectory(prefix="WinBookSplit-release-base-")
        cls.addClassCleanup(cls.shared.cleanup)
        cls.base = Path(cls.shared.name)
        cls.base_repo = cls.base / "repo"
        cls.base_repo.mkdir()
        fixture_git(cls.base_repo, "init", "--quiet")
        for path in sorted(set(checker.INPUT_PATHS) | set(BUILDER_PATHS)):
            destination = cls.base_repo / path
            destination.parent.mkdir(parents=True, exist_ok=True)
            destination.write_bytes((ROOT / path).read_bytes())
        cls.base_commit = fixture_commit(cls.base_repo)
        # Reuse this actual Git-validated immutable source for asset-only mutations.
        # Public build/verify/source/CLI cases below still query the real repository.
        cls.base_source, cls.base_payload, _ = builder._source(cls.base_repo, cls.base_commit)
        cls.base_output = cls.base / "package"
        cls.base_manifest = builder.build_package(cls.base_repo, cls.base_commit, cls.base_output)

    def setUp(self):
        self.owned = tempfile.TemporaryDirectory(prefix="WinBookSplit-release-test-")
        self.addCleanup(self.owned.cleanup)
        self.work = Path(self.owned.name)
        self.repo = self.work / "repo"
        shutil.copytree(self.base_repo, self.repo)
        self.commit = self.base_commit
        self.output = self.work / "package"
        shutil.copytree(self.base_output, self.output)
        self.archive = self.output / ASSET_NAMES[0]

    def git(self, *args, data=None):
        return fixture_git(self.repo, *args, data=data).decode("utf-8").strip()

    def commit_all(self):
        return fixture_commit(self.repo)

    def package_bytes(self, output=None):
        directory = output or self.output
        return {path.name: path.read_bytes() for path in directory.iterdir() if path.is_file()}

    def rejected(self, operation=None, code=None):
        with self.assertRaises((builder.PackageError, builder.checker.InputError)) as caught:
            (operation or (lambda: builder._validate(self.base_source, self.base_payload, self.output)))()
        self.assertTrue(caught.exception.code)
        if code is not None:
            self.assertEqual(caught.exception.code, code)
        return caught.exception.code

    def write_manifest(self, value):
        self.output.joinpath("release-manifest.json").write_bytes(
            (json.dumps(value, sort_keys=True, indent=2, ensure_ascii=True) + "\n").encode("ascii"))

    def refresh_sums(self):
        manifest_bytes = self.output.joinpath("release-manifest.json").read_bytes()
        archive_bytes = self.archive.read_bytes()
        self.output.joinpath("SHA256SUMS.txt").write_bytes(
            (sha256(archive_bytes).hexdigest() + "  " + ASSET_NAMES[0] + "\n" +
             sha256(manifest_bytes).hexdigest() + "  release-manifest.json\n").encode("ascii"))

    def refresh_public_hashes(self):
        value = json.loads(self.output.joinpath("release-manifest.json").read_bytes())
        data = self.archive.read_bytes()
        value["archive"] = {"name": ASSET_NAMES[0], "bytes": len(data), "sha256": sha256(data).hexdigest()}
        self.write_manifest(value)
        self.refresh_sums()

    def restore_package(self):
        for name, data in self.package_bytes(self.base_output).items():
            self.output.joinpath(name).write_bytes(data)

    def rewrite_zip(self, mutate):
        with zipfile.ZipFile(self.archive) as source:
            entries = [(copy(info), source.read(info)) for info in source.infolist()]
        stream = io.BytesIO()
        with warnings.catch_warnings(), zipfile.ZipFile(stream, "w") as target:
            warnings.simplefilter("ignore", UserWarning)
            for info, data in mutate(entries):
                target.writestr(info, data)
        self.archive.write_bytes(stream.getvalue())

    def cli(self, command, commit=None, output=None, script=None):
        return subprocess.run([sys.executable, "-I", "-B", str(script or BUILDER), command,
                               "--repo", str(self.repo), "--commit", commit or self.commit,
                               "--output", str(output or self.output)],
                              stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                              timeout=90, check=False, cwd=self.work)

    def test_payload_members_and_independent_git_blob_hashes_are_exact(self):
        self.assertEqual(len(PAYLOAD_PATHS), 28)
        with zipfile.ZipFile(self.archive) as archive:
            self.assertEqual(archive.namelist(), list(PAYLOAD_PATHS))
            for info in archive.infolist():
                with self.subTest(path=info.filename):
                    committed = fixture_git(self.repo, "cat-file", "blob", self.commit + ":" + info.filename)
                    self.assertEqual(archive.read(info), committed)
                    self.assertEqual(info.date_time, (1980, 1, 1, 0, 0, 0))
                    self.assertEqual(info.compress_type, zipfile.ZIP_STORED)
                    self.assertEqual(info.create_system, 3)
                    self.assertEqual(stat.S_IFMT(info.external_attr >> 16), stat.S_IFREG)
                    self.assertEqual(stat.S_IMODE(info.external_attr >> 16), 0o644)
                    self.assertFalse(info.flag_bits & 1)
                    self.assertFalse(info.extra)
                    self.assertFalse(info.comment)
            self.assertFalse(archive.comment)
        self.assertEqual(set(self.package_bytes()), set(ASSET_NAMES))

    def test_read_only_validation_and_repeated_builds_are_reproducible(self):
        before = self.package_bytes()
        source_status = self.git("status", "--porcelain=v1", "--untracked-files=all")
        self.assertEqual(builder.validate_package(self.repo, self.commit, self.output), self.base_manifest)
        self.assertEqual(self.package_bytes(), before)
        rebuilt = self.work / "rebuilt"
        self.assertEqual(builder.build_package(self.repo, self.commit, rebuilt), self.base_manifest)
        self.assertEqual(self.package_bytes(rebuilt), before)
        self.assertEqual(self.git("status", "--porcelain=v1", "--untracked-files=all"), source_status)

    def test_manifest_and_sums_match_independent_actual_hashes_without_self_cycles(self):
        manifest = json.loads(self.output.joinpath("release-manifest.json").read_bytes())
        self.assertEqual(manifest, self.base_manifest)
        self.assertEqual(manifest["source_commit"], self.commit)
        self.assertEqual(manifest["source_tree"], self.git("rev-parse", self.commit + "^{tree}"))
        self.assertEqual(manifest["version"], "1.0.0")
        self.assertFalse(manifest["signed"])
        self.assertEqual([row["path"] for row in manifest["files"]], list(PAYLOAD_PATHS))
        for row in manifest["files"] + manifest["builder"]["files"]:
            with self.subTest(path=row["path"]):
                data = fixture_git(self.repo, "cat-file", "blob", self.commit + ":" + row["path"])
                self.assertEqual(row["sha256"], sha256(data).hexdigest())
                self.assertEqual(row["bytes"], len(data))
                self.assertEqual(row["git_blob"], self.git("rev-parse", self.commit + ":" + row["path"]))
        self.assertEqual({row["path"] for row in manifest["builder"]["files"]}, set(BUILDER_PATHS))
        self.assertNotIn("sha256", {key: value for key, value in manifest.items() if key != "archive"})
        self.assertEqual(manifest["archive"]["sha256"], sha256(self.archive.read_bytes()).hexdigest())
        self.assertEqual(manifest["archive"]["bytes"], self.archive.stat().st_size)
        sums = self.output.joinpath("SHA256SUMS.txt").read_text(encoding="ascii").splitlines()
        self.assertEqual(len(sums), 2)
        for line in sums:
            digest, name = line.split("  ", 1)
            self.assertIn(name, ASSET_NAMES[:2])
            self.assertEqual(digest, sha256(self.output.joinpath(name).read_bytes()).hexdigest())
        for dependency in manifest["dependencies"].values():
            self.assertFalse(dependency["bundled"])

    def test_dirty_untracked_ignored_and_committed_nonpayload_files_are_excluded(self):
        sentinel = b"AUTHORED_PACKAGE_EXCLUSION_SENTINEL"
        for relative in ("private-book.pdf", ".venv/private_helper.py", "tests/private_fixture.txt",
                         "cache/converter-output.pdf", "poster.png"):
            path = self.repo / relative
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_bytes(sentinel)
        self.repo.joinpath(".gitignore").write_text(".venv/\ncache/\n", encoding="utf-8")
        self.repo.joinpath("engine/winbooksplit_engine.py").write_bytes(sentinel)
        before = self.git("status", "--porcelain=v1", "--untracked-files=all")
        new_output = self.work / "dirty-source-package"
        self.assertEqual(builder.build_package(self.repo, self.commit, new_output), self.base_manifest)
        self.assertEqual(self.package_bytes(new_output), self.package_bytes())
        self.assertEqual(self.git("status", "--porcelain=v1", "--untracked-files=all"), before)
        # Restore only the known owned fixture engine; commit authored excluded files.
        self.repo.joinpath("engine/winbooksplit_engine.py").write_bytes((self.base_repo / "engine/winbooksplit_engine.py").read_bytes())
        new_commit = self.commit_all()
        committed_output = self.work / "committed-decoy-package"
        manifest = builder.build_package(self.repo, new_commit, committed_output)
        self.assertEqual(committed_output.joinpath(ASSET_NAMES[0]).read_bytes(), self.archive.read_bytes())
        self.assertNotIn(sentinel.decode("ascii"), json.dumps(manifest))
        self.assertNotIn(str(self.repo), json.dumps(manifest))
        self.assertEqual(manifest["source_commit"], new_commit)

    def test_explicit_commit_required_and_validation_is_bound_to_that_commit(self):
        for commit in ("HEAD", "main", self.commit[:8], "-bad", "f" * 40):
            with self.subTest(commit=commit):
                self.rejected(lambda: builder.validate_package(self.repo, commit, self.output))
        self.repo.joinpath("authored-extra.txt").write_text("Extra committed source", encoding="utf-8")
        other = self.commit_all()
        self.assertNotEqual(other, self.commit)
        self.rejected(lambda: builder.validate_package(self.repo, other, self.output), "manifest_mismatch")
        self.assertEqual(builder.validate_package(self.repo, self.commit, self.output), self.base_manifest)

    def test_missing_committed_engine_cannot_be_rescued_and_creates_no_output(self):
        path = "engine/winbooksplit_engine.py"
        self.git("rm", "--cached", "--", path)
        self.git("-c", "user.name=Authored package test", "-c", "user.email=authored@example.invalid",
                 "commit", "--quiet", "-m", "Authored missing engine")
        missing = self.git("rev-parse", "HEAD")
        self.assertTrue(self.repo.joinpath(path).is_file())
        output = self.work / "missing-input-output"
        self.rejected(lambda: builder.build_package(self.repo, missing, output), "missing_input")
        self.assertFalse(output.exists())

    def test_release_version_mismatch_creates_no_output(self):
        self.repo.joinpath("VERSION").write_bytes(b"2.0.0\n")
        bad_commit = self.commit_all()
        output = self.work / "wrong-version-output"
        self.rejected(lambda: builder.build_package(self.repo, bad_commit, output), "release_version_mismatch")
        self.assertFalse(output.exists())

    def test_uncommitted_or_changed_builder_provenance_is_rejected(self):
        for relative in BUILDER_PATHS:
            with self.subTest(builder_path=relative):
                path = self.repo / relative
                original = path.read_bytes()
                path.write_bytes(original + b"\n")
                altered_commit = self.commit_all()
                output = self.work / ("changed-builder-" + str(BUILDER_PATHS.index(relative)))
                self.rejected(lambda: builder.build_package(self.repo, altered_commit, output), "builder_source_mismatch")
                self.assertFalse(output.exists())
                path.write_bytes(original)
                self.commit_all()

    def test_dirty_local_builder_files_fail_before_creating_output(self):
        for index, relative in enumerate(BUILDER_PATHS):
            with self.subTest(local_builder_path=relative):
                path = self.repo / relative
                original = path.read_bytes()
                path.write_bytes(original + b"\n")
                output = self.work / ("dirty-local-tool-" + str(index))
                result = self.cli("build", output=output, script=self.repo / "tools/release/build_package.py")
                self.assertEqual(result.returncode, 1)
                self.assertIn(b"builder_source_mismatch", result.stderr)
                self.assertFalse(output.exists())
                path.write_bytes(original)

    def test_committed_local_allowlist_cannot_omit_add_duplicate_or_escape_payload(self):
        path = self.repo / "tools/release/payload.json"
        original = json.loads(path.read_bytes())
        changes = (
            lambda value: value["paths"].remove("engine/winbooksplit_engine.py"),
            lambda value: value["paths"].append("tests/private.pdf"),
            lambda value: value["paths"].append("../private.pdf"),
            lambda value: value["paths"].append("VERSION"),
            lambda value: value["paths"].append("version"),
            lambda value: value.update(version="2.0.0"),
            lambda value: value.update(schema_version=True),
        )
        for index, mutation in enumerate(changes):
            with self.subTest(allowlist_mutation=index):
                value = json.loads(json.dumps(original))
                mutation(value)
                path.write_text(json.dumps(value, sort_keys=True, indent=2) + "\n", encoding="utf-8")
                commit = self.commit_all()
                output = self.work / ("invalid-allowlist-" + str(index))
                # Run the committed fixture builder itself, so its local allowlist
                # genuinely matches the commit before the independent membership gate.
                result = self.cli("build", commit=commit, output=output,
                                  script=self.repo / "tools/release/build_package.py")
                self.assertEqual(result.returncode, 1)
                self.assertIn(b"invalid_payload_allowlist", result.stderr)
                self.assertFalse(output.exists())

    def test_existing_output_and_neighbor_are_never_overwritten(self):
        before = self.package_bytes()
        neighbor = self.work / "neighbor.txt"
        neighbor.write_bytes(b"Authored immutable neighbor")
        self.rejected(lambda: builder.build_package(self.repo, self.commit, self.output), "output_already_exists")
        self.assertEqual(self.package_bytes(), before)
        self.assertEqual(neighbor.read_bytes(), b"Authored immutable neighbor")

    def test_relative_and_unignored_checkout_destinations_are_rejected(self):
        for destination in (Path("relative-output"), self.repo / "package"):
            with self.subTest(destination=destination):
                self.rejected(lambda: builder.build_package(self.repo, self.commit, destination))
                self.assertFalse(destination.exists())

    def test_post_write_validation_failure_retains_incomplete_marker(self):
        output = self.work / "incomplete-output"
        neighbor = self.work / "neighbor.txt"
        neighbor.write_bytes(b"Authored neighbor")
        with patch.object(builder, "_validate", side_effect=builder.PackageError("authored_validation_failure")):
            self.rejected(lambda: builder.build_package(self.repo, self.commit, output), "authored_validation_failure")
        self.assertEqual({path.name for path in output.iterdir()}, set(ASSET_NAMES) | {".package-incomplete"})
        self.assertTrue(output.joinpath(".package-incomplete").read_bytes())
        self.rejected(lambda: builder.validate_package(self.repo, self.commit, output), "asset_membership_mismatch")
        self.assertEqual(neighbor.read_bytes(), b"Authored neighbor")

    def test_manifest_mutations_fail_even_after_recalculating_manifest_checksum(self):
        mutations = (
            lambda value: value.update(source_commit="f" * 40),
            lambda value: value.update(version="2.0.0"),
            lambda value: value.update(source_tree="e" * 40),
            lambda value: value["dependencies"]["pypdf"].update(bundled=True),
            lambda value: value["builder"]["files"][0].update(sha256="0" * 64),
            lambda value: value["builder"]["archive_recipe"].update(mode="100755"),
        )
        for index, mutation in enumerate(mutations):
            with self.subTest(mutation=index):
                self.restore_package()
                value = json.loads(self.output.joinpath("release-manifest.json").read_bytes())
                mutation(value)
                self.write_manifest(value)
                self.refresh_sums()
                self.rejected(code="manifest_mismatch")

    def test_malformed_noncanonical_manifest_and_checksum_cycle_are_rejected(self):
        for malformed in (b"{}", b"{", b"[]", b"null"):
            with self.subTest(manifest=malformed):
                self.restore_package()
                self.output.joinpath("release-manifest.json").write_bytes(malformed)
                self.refresh_sums()
                self.rejected()
        self.restore_package()
        value = json.loads(self.output.joinpath("release-manifest.json").read_bytes())
        self.output.joinpath("release-manifest.json").write_bytes(json.dumps(value).encode("ascii"))
        self.refresh_sums()
        self.rejected(code="manifest_mismatch")
        self.restore_package()
        sums = self.output.joinpath("SHA256SUMS.txt")
        sums.write_bytes(sums.read_bytes() + sha256(sums.read_bytes()).hexdigest().encode("ascii") + b"  SHA256SUMS.txt\n")
        self.rejected(code="checksum_mismatch")

    def test_wrong_or_missing_checksums_and_asset_membership_fail(self):
        sums = self.output.joinpath("SHA256SUMS.txt")
        original = sums.read_bytes()
        for data in (b"0" * 64 + original[64:], original.splitlines(keepends=True)[0], original[::-1]):
            with self.subTest(checksums=data[:70]):
                sums.write_bytes(data)
                self.rejected(code="checksum_mismatch")
        self.restore_package()
        self.output.joinpath("authored-private.txt").write_text("Synthetic excluded member", encoding="utf-8")
        self.rejected(code="asset_membership_mismatch")
        self.output.joinpath("authored-private.txt").unlink()
        sums.unlink()
        self.rejected(code="asset_membership_mismatch")

    def test_payload_and_version_tampering_fail_after_refreshing_all_untrusted_hashes(self):
        for relative in ("README.md", "VERSION", "engine/winbooksplit_engine.py"):
            with self.subTest(path=relative):
                self.restore_package()
                def mutate(entries):
                    return [(info, (b"2.0.0\n" if relative == "VERSION" else data + b"AUTHORED_PAYLOAD_TAMPER"))
                            if info.filename == relative else (info, data) for info, data in entries]
                self.rewrite_zip(mutate)
                self.refresh_public_hashes()
                # Forge per-file size/hash and the aggregate in the external manifest too.
                value = json.loads(self.output.joinpath("release-manifest.json").read_bytes())
                with zipfile.ZipFile(self.archive) as archive:
                    for row in value["files"]:
                        if row["path"] == relative:
                            data = archive.read(relative)
                            row.update(bytes=len(data), sha256=sha256(data).hexdigest())
                canonical_files = (json.dumps(value["files"], sort_keys=True, indent=2, ensure_ascii=True) + "\n").encode("ascii")
                value["payload_sha256"] = sha256(canonical_files).hexdigest()
                self.write_manifest(value)
                self.refresh_sums()
                self.rejected(code="manifest_mismatch")

    def test_member_bytes_cannot_be_replaced_while_claiming_original_git_hashes(self):
        def mutate(entries):
            return [(info, b"X" * len(data)) if info.filename == "VERSION" else (info, data) for info, data in entries]
        self.rewrite_zip(mutate)
        self.refresh_public_hashes()
        self.rejected(code="zip_payload_mismatch")

    def test_unsafe_nonallowed_duplicate_and_casefold_zip_members_fail(self):
        unsafe = ("../private.pdf", "/absolute.pdf", "C:/private.pdf", "engine\\private.py", "engine//private.py",
                  "engine/CON.py", "engine/private.py", "engine/", "version", "VERSION")
        for name in unsafe:
            with self.subTest(member=name):
                self.restore_package()
                def mutate(entries):
                    changed = copy(entries[0][0])
                    changed.filename = changed.orig_filename = name
                    # Adding preserves the original payload and exposes duplicates/casefold aliases.
                    return entries + [(changed, b"Authored unsafe or duplicate member")]
                self.rewrite_zip(mutate)
                self.refresh_public_hashes()
                self.rejected(code="zip_membership_mismatch")

    def test_missing_engine_and_wrong_member_order_fail_with_refreshed_hashes(self):
        for change in (lambda entries: [entry for entry in entries if entry[0].filename != "engine/winbooksplit_engine.py"],
                       lambda entries: list(reversed(entries))):
            with self.subTest(change=change):
                self.restore_package()
                self.rewrite_zip(change)
                self.refresh_public_hashes()
                self.rejected(code="zip_membership_mismatch")

    def test_symlink_directory_compression_and_metadata_members_fail(self):
        def symlink(info):
            info.external_attr = (stat.S_IFLNK | 0o777) << 16
        def directory(info):
            info.external_attr = ((stat.S_IFDIR | 0o755) << 16) | 0x10
        def compressed(info):
            info.compress_type = zipfile.ZIP_DEFLATED
        def timestamp(info):
            info.date_time = (2020, 1, 2, 3, 4, 6)
        def mode(info):
            info.external_attr = (stat.S_IFREG | 0o755) << 16
        def comment(info):
            info.comment = b"Authored member comment"
        def extra(info):
            info.extra = b"\xfe\xca\x00\x00"
        for mutation in (symlink, directory, compressed, timestamp, mode, comment, extra):
            with self.subTest(metadata=mutation.__name__):
                self.restore_package()
                def mutate(entries):
                    mutation(entries[0][0])
                    return entries
                self.rewrite_zip(mutate)
                self.refresh_public_hashes()
                self.rejected(code="invalid_zip_member")

    def test_encrypted_flags_are_rejected_before_any_member_read(self):
        data = bytearray(self.archive.read_bytes())
        # Authored flag mutation is applied to both local and central header.
        local = data.index(b"PK\x03\x04")
        central = data.index(b"PK\x01\x02")
        struct.pack_into("<H", data, local + 6, struct.unpack_from("<H", data, local + 6)[0] | 1)
        struct.pack_into("<H", data, central + 8, struct.unpack_from("<H", data, central + 8)[0] | 1)
        self.archive.write_bytes(data)
        self.refresh_public_hashes()
        self.rejected(code="invalid_zip_member")

    def test_archive_suffix_corruption_and_comment_cannot_hide_behind_refreshed_hashes(self):
        original = self.archive.read_bytes()
        for data in (original + b"Authored hidden suffix", b"Not an archive"):
            with self.subTest(archive=data[-30:]):
                self.archive.write_bytes(data)
                self.refresh_public_hashes()
                self.rejected()
        self.restore_package()
        with zipfile.ZipFile(self.archive, "a") as archive:
            archive.comment = b"Authored archive comment"
        self.refresh_public_hashes()
        self.rejected(code="zip_membership_mismatch")

    def test_small_zip64_metadata_override_fails_before_zipfile_construction(self):
        original = self.archive.read_bytes()
        end = len(original) - 22
        fields = struct.unpack("<4s4H2LH", original[end:])
        self.assertEqual(fields[0], b"PK\x05\x06")
        self.assertEqual(fields[4], len(PAYLOAD_PATHS))
        central_size, central_offset = fields[5:7]
        central = original[central_offset:central_offset + central_size]
        # This harmless 56-member fixture demonstrates that ZIP64 can override a
        # normal EOCD that still advertises the expected 28 members and CD size.
        expanded = central * 2
        record_offset = central_offset + len(expanded)
        record = struct.pack("<4sQ2H2L4Q", b"PK\x06\x06", 44, 45, 45, 0, 0,
                             len(PAYLOAD_PATHS) * 2, len(PAYLOAD_PATHS) * 2,
                             len(expanded), central_offset)
        locator = struct.pack("<4sLQL", b"PK\x06\x07", 0, record_offset, 1)
        variant = original[:central_offset] + expanded + record + locator + original[end:]
        with zipfile.ZipFile(io.BytesIO(variant)) as independently_parsed:
            self.assertEqual(len(independently_parsed.infolist()), len(PAYLOAD_PATHS) * 2)
        self.archive.write_bytes(variant)
        self.refresh_public_hashes()
        with patch.object(builder.zipfile, "ZipFile", side_effect=AssertionError("Unexpected archive parsing")) as parsing:
            self.rejected(code="zip_recipe_mismatch")
        parsing.assert_not_called()

    def test_cli_build_verify_and_failure_keep_native_exit_and_safe_output(self):
        fresh = self.work / "cli-package"
        result = self.cli("build", output=fresh)
        self.assertEqual(result.returncode, 0, result.stderr.decode("utf-8", errors="replace"))
        self.assertIn(b"PACKAGE_PASSED", result.stdout)
        result = self.cli("verify", output=fresh)
        self.assertEqual(result.returncode, 0, result.stderr.decode("utf-8", errors="replace"))
        before = self.package_bytes(fresh)
        result = self.cli("build", output=fresh)
        self.assertEqual(result.returncode, 1)
        self.assertIn(b"output_already_exists", result.stderr)
        self.assertEqual(self.package_bytes(fresh), before)
        result = self.cli("build", commit="HEAD", output=self.work / "cli-invalid")
        self.assertEqual(result.returncode, 1)
        self.assertFalse(self.work.joinpath("cli-invalid").exists())
        self.archive.write_bytes(b"Authored invalid ZIP")
        result = self.cli("verify")
        self.assertEqual(result.returncode, 1)
        self.assertNotIn(str(self.repo).encode("utf-8"), result.stdout + result.stderr)


if __name__ == "__main__":
    unittest.main()
