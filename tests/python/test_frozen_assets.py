"""Git-free frozen-byte checks on actual authored Git-built package fixtures.

Repinned malformed fixtures below are explicit synthetic validator controls;
they never authorize adopting external identities from an untrusted manifest.
"""
from copy import copy, deepcopy
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
SPEC = importlib.util.spec_from_file_location("wbs_frozen_fixture", ROOT / "tests/python/test_release_package.py")
fixture = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(fixture)
builder = fixture.builder
frozen = builder.assets


class FrozenAssetsTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.shared = tempfile.TemporaryDirectory(prefix="WinBookSplit-frozen-base-")
        cls.addClassCleanup(cls.shared.cleanup)
        cls.base = Path(cls.shared.name)
        cls.repo = cls.base / "repo"
        cls.repo.mkdir()
        fixture.fixture_git(cls.repo, "init", "--quiet")
        for relative in sorted(set(fixture.checker.INPUT_PATHS) | set(fixture.BUILDER_PATHS)):
            target = cls.repo / relative
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes((ROOT / relative).read_bytes())
        cls.commit = fixture.fixture_commit(cls.repo)
        cls.package = cls.base / "package"
        cls.manifest = builder.build_package(cls.repo, cls.commit, cls.package)
        cls.base_payload = builder._source(cls.repo, cls.commit)[1]
        cls.expectations = {
            "version": cls.manifest["version"], "source_commit": cls.commit,
            "source_tree": cls.manifest["source_tree"], "payload_sha256": cls.manifest["payload_sha256"],
            "policy_sha256": sha256(cls.base_payload["engine/WinBookSplit.Compatibility.json"]).hexdigest(),
            "assets": {name: {"bytes": (cls.package / name).stat().st_size,
                              "sha256": sha256((cls.package / name).read_bytes()).hexdigest()}
                       for name in fixture.ASSET_NAMES}}

    def setUp(self):
        self.owned = tempfile.TemporaryDirectory(prefix="WinBookSplit-frozen-case-")
        self.addCleanup(self.owned.cleanup)
        self.work = Path(self.owned.name)
        self.assets = self.work / "assets"
        shutil.copytree(self.package, self.assets)
        self.expected = deepcopy(self.expectations)
        self.zip_path = self.assets / fixture.ASSET_NAMES[0]

    def verify(self):
        return frozen.verify_frozen_assets(self.assets, self.expected)

    def repin(self, manifest=None):
        """Synthetic control only: explicitly change externally trusted pins."""
        archive = self.zip_path.read_bytes()
        if manifest is None:
            manifest = json.loads((self.assets / "release-manifest.json").read_bytes())
        manifest["archive"] = {"name": fixture.ASSET_NAMES[0], "bytes": len(archive),
                               "sha256": sha256(archive).hexdigest()}
        manifest_bytes = frozen.canonical(manifest)
        (self.assets / "release-manifest.json").write_bytes(manifest_bytes)
        (self.assets / "SHA256SUMS.txt").write_bytes(
            (sha256(archive).hexdigest() + "  " + fixture.ASSET_NAMES[0] + "\n" +
             sha256(manifest_bytes).hexdigest() + "  release-manifest.json\n").encode("ascii"))
        for name in fixture.ASSET_NAMES:
            data = (self.assets / name).read_bytes()
            self.expected["assets"][name] = {"bytes": len(data), "sha256": sha256(data).hexdigest()}

    def rewrite_zip(self, mutate):
        with zipfile.ZipFile(self.zip_path) as archive:
            entries = [(copy(member), archive.read(member)) for member in archive.infolist()]
        stream = io.BytesIO()
        with warnings.catch_warnings(), zipfile.ZipFile(stream, "w") as archive:
            warnings.simplefilter("ignore", UserWarning)
            for member, data in mutate(entries):
                archive.writestr(member, data)
        self.zip_path.write_bytes(stream.getvalue())

    def test_actual_git_built_bytes_verify_without_git_or_package_imports(self):
        with patch.object(fixture.checker, "git", side_effect=AssertionError("Unexpected Git invocation")):
            manifest, payload = self.verify()
        self.assertEqual(manifest, self.manifest)
        self.assertEqual(payload, self.base_payload)
        self.assertEqual(builder.validate_package(self.repo, self.commit, self.package), manifest)
        (self.work / "expectations.json").write_bytes(frozen.canonical(self.expected))
        code = """import importlib.util,json,pathlib,sys
p=pathlib.Path(sys.argv[1]); s=importlib.util.spec_from_file_location('frozen',p)
m=importlib.util.module_from_spec(s);s.loader.exec_module(m)
manifest,payload=m.verify_frozen_assets(pathlib.Path(sys.argv[2]),json.loads(pathlib.Path(sys.argv[3]).read_bytes()))
print(json.dumps({'source_commit':manifest['source_commit'],'count':len(payload)}))
"""
        environment = dict(os.environ)
        environment["PATH"] = ""
        observed = subprocess.run([sys.executable, "-I", "-B", "-c", code,
            str(ROOT / "tools/release/frozen_assets.py"), str(self.assets), str(self.work / "expectations.json")],
            cwd=self.work, env=environment, stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=20, check=False)
        self.assertEqual(observed.returncode, 0, observed.stderr)
        self.assertEqual(json.loads(observed.stdout), {"source_commit": self.commit, "count": 34})

    def test_dependency_contract_keeps_policy_ranges_separate_from_build_runtime(self):
        policy_bytes = self.base_payload["engine/WinBookSplit.Compatibility.json"]
        policy = json.loads(policy_bytes)
        contract = self.manifest["dependencies"]
        self.assertNotIn("version", contract["python"])
        self.assertEqual(contract["python"]["branches"], policy["python"]["branches"])
        self.assertEqual(contract["python"]["excluded_branches"], policy["excluded_python_branches"])
        self.assertEqual(contract["python"]["policy_sha256"], sha256(policy_bytes).hexdigest())
        self.assertEqual(contract["pypdf"]["wheel_sha256"], policy["pypdf"]["wheel_sha256"])
        self.assertEqual(contract["calibre"]["version"], policy["calibre"]["version"])
        self.assertEqual(contract["windows"]["supported"], ["Windows 10 22H2 x64", "Windows 11 x64"])
        self.assertEqual(self.manifest["builder"]["python"], sys.version.split()[0])
        self.assertEqual(contract, frozen.dependency_metadata(policy_bytes, self.base_payload["requirements.txt"]))
        self.verify()

    def test_repinned_manifest_cannot_change_floors_or_restore_single_patch_gate(self):
        mutations = (
            lambda value: value["dependencies"]["python"].update(version="3.14.8"),
            lambda value: value["dependencies"]["python"]["branches"][0].update(
                security_min_patch=value["dependencies"]["python"]["branches"][0]["security_min_patch"] + 1),
            lambda value: value["dependencies"]["python"].update(policy_schema_version=True),
            lambda value: value["dependencies"]["python"]["branches"][0].update(compatibility_min_patch=False),
            lambda value: value["dependencies"]["python"]["branches"][0].update(security_min_patch=7.0),
            lambda value: value["dependencies"]["python"]["branches"][0].update(blocked_patches=[9]),
            lambda value: value["dependencies"]["python"].update(policy_sha256="0" * 64),
            lambda value: value["dependencies"]["pypdf"].update(wheel_sha256="0" * 64),
            lambda value: value["dependencies"]["windows"].update(supported=["Windows 11 x64"]),
            lambda value: value["dependencies"]["calibre"].update(version="9.14.0"))
        for mutation in mutations:
            value = deepcopy(self.manifest)
            mutation(value)
            self.repin(value)  # Deliberately trusted malformed SYNTHETIC metadata.
            with self.assertRaisesRegex(frozen.PackageError, "manifest_mismatch"):
                self.verify()
        self.repin(deepcopy(self.manifest))
        self.verify()

    def test_metadata_projection_refuses_boolean_floor_and_split_wheel_pin(self):
        original = json.loads(self.base_payload["engine/WinBookSplit.Compatibility.json"])
        mutations = (
            lambda value: value["python"]["branches"][0].update(security_min_patch=True),
            lambda value: value["python"].update(bits=True),
            lambda value: value["python"].update(stable_only=False),
            lambda value: value["python"].update(gil_disabled=True),
            lambda value: next(branch for branch in value["python"]["branches"] if branch["minor"] == 14).update(recommended_patch=0),
            lambda value: value["pypdf"].update(wheel_sha256="0" * 64))
        for mutation in mutations:
            value = deepcopy(original)
            mutation(value)
            with self.assertRaisesRegex(frozen.PackageError, "invalid_dependency_policy"):
                frozen.dependency_metadata(json.dumps(value).encode("utf-8"), self.base_payload["requirements.txt"])

    def test_external_pins_are_checked_before_json_or_zip_parsing(self):
        for name in fixture.ASSET_NAMES:
            with self.subTest(asset=name):
                original = (self.assets / name).read_bytes()
                (self.assets / name).write_bytes(b"X" + original[1:])
                with patch.object(frozen, "_strict_json", side_effect=AssertionError("Untrusted manifest parsed")), \
                        patch.object(frozen.zipfile, "ZipFile", side_effect=AssertionError("Untrusted ZIP parsed")), \
                        self.assertRaisesRegex(frozen.PackageError, "asset_hash_mismatch"):
                    self.verify()
                (self.assets / name).write_bytes(original)

    def test_bindings_require_exact_typed_source_asset_and_policy_identities(self):
        for field, value in (("source_commit", "HEAD"), ("source_tree", "A" * 40),
                             ("payload_sha256", "z" * 64), ("policy_sha256", None),
                             ("version", "../1.0.0"), ("assets", {})):
            with self.subTest(field=field), self.assertRaisesRegex(frozen.PackageError, "invalid_asset_binding"):
                frozen.verify_frozen_assets(self.assets, {**self.expected, field: value})
        wrong_size = deepcopy(self.expected)
        wrong_size["assets"][fixture.ASSET_NAMES[0]]["bytes"] = True
        with self.assertRaisesRegex(frozen.PackageError, "invalid_asset_binding"):
            frozen.verify_frozen_assets(self.assets, wrong_size)
        wrong_policy = {**self.expected, "policy_sha256": "0" * 64}
        with self.assertRaisesRegex(frozen.PackageError, "manifest_mismatch"):
            frozen.verify_frozen_assets(self.assets, wrong_policy)

    def test_noncanonical_duplicate_and_nonfinite_manifest_fail_closed(self):
        original = (self.assets / "release-manifest.json").read_bytes()
        variants = [json.dumps(self.manifest).encode("ascii"), b'{"schema_version":1,"schema_version":1}',
                    b'{"schema_version":NaN}', b"[]", b"{"]
        for malformed in variants:
            with self.subTest(data=malformed[:40]):
                (self.assets / "release-manifest.json").write_bytes(malformed)
                self.expected["assets"]["release-manifest.json"] = {"bytes": len(malformed), "sha256": sha256(malformed).hexdigest()}
                with self.assertRaises(frozen.PackageError):
                    self.verify()
        (self.assets / "release-manifest.json").write_bytes(original)

    def test_forged_rows_traversal_casefold_duplicates_and_size_are_rejected(self):
        variants = [lambda m: m.update(source_commit="f" * 40),
                    lambda m: m["files"][0].update(path="../private.pdf"),
                    lambda m: m["files"][0].update(path="engine/CON.py"),
                    lambda m: m["files"][0].update(bytes=True),
                    lambda m: m["files"][0].update(bytes=frozen.MAX_FILE_BYTES + 1),
                    lambda m: m["files"].append({**m["files"][-1], "path": "version"}),
                    lambda m: m["files"][0].update(git_blob="0" * 40)]
        for mutate in variants:
            with self.subTest(mutation=mutate):
                manifest = deepcopy(self.manifest)
                mutate(manifest)
                manifest["payload_sha256"] = sha256(frozen.canonical(manifest["files"])).hexdigest()
                self.expected = deepcopy(self.expectations)
                self.expected["payload_sha256"] = manifest["payload_sha256"]
                self.repin(manifest)
                with self.assertRaises(frozen.PackageError):
                    self.verify()

    def test_all_headers_are_inspected_before_any_member_read(self):
        def compressed_last(entries):
            entries[-1][0].compress_type = zipfile.ZIP_DEFLATED
            return entries
        self.rewrite_zip(compressed_last)
        self.repin()
        with patch.object(frozen.zipfile.ZipFile, "read", side_effect=AssertionError("Member read before all headers checked")) as reading:
            with self.assertRaisesRegex(frozen.PackageError, "invalid_zip_member"):
                self.verify()
        reading.assert_not_called()

    def test_file_directory_prefix_and_ancestor_case_aliases_fail_before_zip_parsing(self):
        for path in ("engine", "ENGINE/fresh.py"):
            with self.subTest(path=path):
                manifest = deepcopy(self.manifest)
                manifest["files"].append({**manifest["files"][0], "path": path})
                manifest["files"].sort(key=lambda row: row["path"])
                manifest["payload_sha256"] = sha256(frozen.canonical(manifest["files"])).hexdigest()
                self.expected = deepcopy(self.expectations)
                self.expected["payload_sha256"] = manifest["payload_sha256"]
                self.repin(manifest)
                with patch.object(frozen.zipfile, "ZipFile", side_effect=AssertionError("Ambiguous members parsed")) as parsing:
                    with self.assertRaisesRegex(frozen.PackageError, "manifest_mismatch"):
                        self.verify()
                parsing.assert_not_called()

    def test_duplicate_and_unsafe_zip_entries_cannot_be_repinned_into_membership(self):
        for name in ("../private.pdf", "version", "VERSION"):
            with self.subTest(name=name):
                shutil.copyfile(self.package / fixture.ASSET_NAMES[0], self.zip_path)
                def extra(entries):
                    info = copy(entries[0][0])
                    info.filename = info.orig_filename = name
                    return entries + [(info, b"Authored rejected member")]
                self.rewrite_zip(extra)
                self.repin(deepcopy(self.manifest))
                with self.assertRaisesRegex(frozen.PackageError, "zip_membership_mismatch"):
                    self.verify()

    def test_zip64_override_is_rejected_before_zipfile_construction(self):
        original = self.zip_path.read_bytes()
        end = len(original) - 22
        fields = struct.unpack("<4s4H2LH", original[end:])
        central_size, central_offset = fields[5:7]
        central = original[central_offset:central_offset + central_size]
        expanded = central * 2
        offset = central_offset + len(expanded)
        record = struct.pack("<4sQ2H2L4Q", b"PK\x06\x06", 44, 45, 45, 0, 0,
                             len(self.base_payload) * 2, len(self.base_payload) * 2, len(expanded), central_offset)
        locator = struct.pack("<4sLQL", b"PK\x06\x07", 0, offset, 1)
        self.zip_path.write_bytes(original[:central_offset] + expanded + record + locator + original[end:])
        self.repin()
        with patch.object(frozen.zipfile, "ZipFile", side_effect=AssertionError("ZIP64 parsing allocated members")) as parsing:
            with self.assertRaisesRegex(frozen.PackageError, "zip_recipe_mismatch"):
                self.verify()
        parsing.assert_not_called()

    def test_raw_recipe_and_payload_contradictions_fail_after_external_repinning(self):
        self.zip_path.write_bytes(self.zip_path.read_bytes() + b"Authored trailing bytes")
        self.repin()
        with self.assertRaisesRegex(frozen.PackageError, "zip_recipe_mismatch"):
            self.verify()
        shutil.copyfile(self.package / fixture.ASSET_NAMES[0], self.zip_path)
        self.rewrite_zip(lambda entries: [(info, b"X" * len(data)) if info.filename == "VERSION" else (info, data)
                                          for info, data in entries])
        self.repin(deepcopy(self.manifest))
        with self.assertRaisesRegex(frozen.PackageError, "zip_payload_mismatch"):
            self.verify()

    def test_extra_assets_incomplete_marker_and_checksum_cycle_refuse_payload(self):
        marker = self.assets / ".package-incomplete"
        marker.write_bytes(b"Owned failed build")
        with self.assertRaisesRegex(frozen.PackageError, "asset_membership_mismatch"):
            self.verify()
        marker.unlink()
        sums = self.assets / "SHA256SUMS.txt"
        sums.write_bytes(sums.read_bytes() + b"0" * 64 + b"  SHA256SUMS.txt\n")
        self.expected["assets"][sums.name] = {"bytes": sums.stat().st_size, "sha256": sha256(sums.read_bytes()).hexdigest()}
        with self.assertRaisesRegex(frozen.PackageError, "checksum_mismatch"):
            self.verify()

    def test_same_existing_windows_relative_path_rules_are_preserved(self):
        for path in ("", "../escape", "/absolute", "C:/drive", "a\\b", "a//b", "a/./b", "a/../b",
                     "CON", "engine/NUL.txt", "a:stream", "trailing.", "trailing ", "a\x00b", "a\x7fb",
                     "VERSION", "engine/winbooksplit_engine.py"):
            with self.subTest(path=path):
                self.assertEqual(frozen.safe_relative(path), fixture.checker.safe_relative(path))


if __name__ == "__main__":
    unittest.main(verbosity=2)
