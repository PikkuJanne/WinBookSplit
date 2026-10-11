"""Offline, dependency-free validation of the shipped runtime compatibility policy.

Policy acceptance describes a supported version range, not a claim that every
accepted maintenance release was individually tested. Dependency import origins
and process transport remain the responsibility of the existing entry points.
"""

import json
import importlib
from importlib import metadata
import platform
import re
import struct
import sys
import sysconfig
from datetime import date
from hashlib import sha256
from pathlib import Path
from urllib.parse import urlsplit


POLICY_MAX_BYTES = 32768
_MAX_COMPONENT = 1000000
_VERSION = re.compile(r"(0|[1-9][0-9]{0,6})\.(0|[1-9][0-9]{0,6})\.(0|[1-9][0-9]{0,6})\Z")
_UNSTABLE_VERSION = re.compile(
    r"(?:0|[1-9][0-9]{0,6})\.(?:0|[1-9][0-9]{0,6})\.(?:0|[1-9][0-9]{0,6})"
    r"(?:(?:a|b|rc)[0-9]+|\.dev[0-9]+|\+)\Z")
_SOURCE_ID = re.compile(r"[A-Za-z][A-Za-z0-9_-]{0,63}\Z")
_WINDOWS_ABSOLUTE = re.compile(r"(?:[A-Za-z]:[\\/]|\\\\[^\\]+\\[^\\]+\\)")


class PolicyError(ValueError):
    """The shipped compatibility policy is missing or malformed."""


def _integer(value):
    return type(value) is int and 0 <= value <= _MAX_COMPONENT


def _require(condition, message):
    if not condition:
        raise PolicyError(message)


def _object(value, keys, label):
    _require(type(value) is dict and set(value) == set(keys),
             "Invalid " + label + " object or fields.")


def _text(value, maximum=1024):
    return (type(value) is str and 0 < len(value) <= maximum
            and bool(value.strip()) and "\x00" not in value)


def _version(value):
    if type(value) is str:
        match = _VERSION.fullmatch(value)
        if match is None:
            return None
        result = tuple(int(part) for part in match.groups())
    elif type(value) in (list, tuple) and len(value) == 3:
        result = tuple(value)
    else:
        return None
    return result if all(_integer(part) for part in result) else None


def _unique_object(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise PolicyError("Duplicate compatibility policy key: " + key)
        result[key] = value
    return result


def _invalid_constant(value):
    raise PolicyError("Nonfinite compatibility policy number: " + value)


def _validate_policy(policy):
    """Validate values again even when a caller did not use load_policy()."""
    _object(policy, ("schema_version", "policy_version", "reviewed_on", "python",
                     "pypdf", "calibre", "excluded_python_branches", "sources"), "policy")
    _require(type(policy["schema_version"]) is int and policy["schema_version"] == 1,
             "Unsupported compatibility policy schema.")
    _require(_integer(policy["policy_version"]) and policy["policy_version"] > 0,
             "Invalid compatibility policy version.")
    reviewed = policy["reviewed_on"]
    _require(type(reviewed) is str and re.fullmatch(r"[0-9]{4}-[0-9]{2}-[0-9]{2}", reviewed),
             "Invalid compatibility policy review date.")
    try:
        date.fromisoformat(reviewed)
    except ValueError as error:
        raise PolicyError("Invalid compatibility policy review date.") from error

    sources = policy["sources"]
    _require(type(sources) is list and 1 <= len(sources) <= 64, "Invalid policy sources.")
    source_ids = set()
    for source in sources:
        _object(source, ("id", "url", "relevance"), "source")
        identifier = source["id"]
        _require(type(identifier) is str and _SOURCE_ID.fullmatch(identifier)
                 and identifier not in source_ids, "Invalid or duplicate policy source ID.")
        _require(_text(source["url"], 2048) and _text(source["relevance"]), "Invalid policy source text.")
        try:
            url = urlsplit(source["url"])
            port = url.port  # Access validates numeric syntax and the TCP port range.
            valid_url = (url.scheme == "https" and bool(url.hostname)
                         and (port is None or 0 <= port <= 65535)
                         and url.username is None and url.password is None
                         and not any(character.isspace() for character in source["url"]))
        except ValueError:
            valid_url = False
        _require(valid_url, "Policy sources require absolute HTTPS URLs without credentials.")
        source_ids.add(identifier)

    python = policy["python"]
    _object(python, ("implementation", "platform", "machine", "bits", "stable_only",
                     "gil_disabled", "branches"), "Python policy")
    _require(python["implementation"] == "cpython" and python["platform"] == "win32"
             and python["machine"] == "AMD64" and type(python["bits"]) is int
             and python["bits"] == 64 and python["stable_only"] is True
             and python["gil_disabled"] is False, "Invalid supported Python identity.")
    branches = python["branches"]
    _require(type(branches) is list and 1 <= len(branches) <= 8, "Invalid Python branch list.")
    branch_ids = set()
    for branch in branches:
        _object(branch, ("major", "minor", "compatibility_min_patch", "security_min_patch",
                         "recommended_patch", "blocked_patches", "security_sources"), "Python branch")
        _require(all(_integer(branch[field]) for field in (
            "major", "minor", "compatibility_min_patch", "security_min_patch", "recommended_patch")),
            "Python branch versions require bounded nonnegative integers.")
        branch_id = (branch["major"], branch["minor"])
        _require(branch_id[0] == 3 and branch_id not in branch_ids, "Invalid or duplicate Python branch.")
        blocked = branch["blocked_patches"]
        _require(type(blocked) is list and len(blocked) <= 32 and all(_integer(item) for item in blocked)
                 and len(set(blocked)) == len(blocked), "Invalid or duplicate blocked Python patches.")
        _require(branch["recommended_patch"] >= max(branch["compatibility_min_patch"], branch["security_min_patch"])
                 and branch["recommended_patch"] not in blocked, "Recommended patch contradicts branch floors or exclusions.")
        references = branch["security_sources"]
        _require(type(references) is list and 1 <= len(references) <= 64
                 and all(type(item) is str and item in source_ids for item in references)
                 and len(set(references)) == len(references), "Missing, unknown or duplicate Python security source reference.")
        branch_ids.add(branch_id)

    exclusions = policy["excluded_python_branches"]
    _require(type(exclusions) is list and len(exclusions) <= 16, "Invalid excluded Python branches.")
    excluded_ids = set()
    for branch in exclusions:
        _object(branch, ("major", "minor", "reason"), "excluded Python branch")
        _require(_integer(branch["major"]) and _integer(branch["minor"]) and _text(branch["reason"]),
                 "Invalid excluded Python branch identity or reason.")
        branch_id = (branch["major"], branch["minor"])
        _require(branch_id not in excluded_ids and branch_id not in branch_ids,
                 "Duplicate or supported-and-excluded Python branch.")
        excluded_ids.add(branch_id)

    _object(policy["pypdf"], ("version", "wheel_sha256"), "pypdf policy")
    _require(_version(policy["pypdf"]["version"]) is not None
             and type(policy["pypdf"]["version"]) is str, "Invalid locked pypdf version.")
    _require(type(policy["pypdf"]["wheel_sha256"]) is str
             and re.fullmatch(r"[0-9a-f]{64}", policy["pypdf"]["wheel_sha256"]), "Invalid locked pypdf wheel SHA256.")
    _object(policy["calibre"], ("version",), "Calibre policy")
    _require(type(policy["calibre"]["version"]) is str
             and _version(policy["calibre"]["version"]) is not None, "Invalid locked Calibre version.")
    return policy


def _read_policy(path=None):
    """Return the policy and exactly the bytes validated, for digest binding."""
    try:
        target = Path(path) if path is not None else Path(__file__).resolve().with_name("WinBookSplit.Compatibility.json")
        with target.open("rb") as stream:
            data = stream.read(POLICY_MAX_BYTES + 1)
        _require(0 < len(data) <= POLICY_MAX_BYTES, "Compatibility policy is empty or exceeds 32 KiB.")
        policy = json.loads(data.decode("utf-8"), object_pairs_hook=_unique_object,
                            parse_constant=_invalid_constant)
        return _validate_policy(policy), data
    except PolicyError:
        raise
    except (OSError, UnicodeError, ValueError, TypeError, RecursionError) as error:
        raise PolicyError("Missing or malformed shipped compatibility policy.") from error


def load_policy(path=None):
    """Read at most 32 KiB of strict UTF-8 JSON; damaged policy fails closed."""
    return _read_policy(path)[0]


def _result(ok, reason, message):
    return {"ok": ok, "reason": reason, "message": message}


def validate_interpreter(record, policy):
    """Validate a runtime identity independently of a candidate's claimed ok flag.

    The caller validates its transport/protocol and package origin separately.
    Setup may therefore use this check before pypdf has been installed.
    """
    try:
        _validate_policy(policy)
    except (PolicyError, TypeError, RecursionError) as error:
        return _result(False, "malformed_policy", str(error))
    fields = ("version", "version_info", "releaselevel", "serial", "implementation", "platform",
              "machine", "bits", "gil_disabled", "isolated", "dont_write_bytecode", "executable")
    if type(record) is not dict or any(field not in record for field in fields):
        return _result(False, "malformed_runtime", "Runtime identity is missing required fields.")
    info = record["version_info"]
    if (type(info) not in (list, tuple) or len(info) != 3 or not all(_integer(item) for item in info)
            or not _integer(record["serial"]) or not _integer(record["bits"])
            or any(type(record[field]) is not bool for field in ("gil_disabled", "isolated", "dont_write_bytecode"))
            or any(not _text(record[field], 32768 if field == "executable" else 128)
                   for field in ("implementation", "platform", "machine", "releaselevel", "executable"))
            or not _WINDOWS_ABSOLUTE.match(record["executable"])):
        return _result(False, "malformed_runtime", "Runtime identity has malformed scalar fields.")
    if record["releaselevel"] not in ("final", "alpha", "beta", "candidate"):
        return _result(False, "malformed_runtime", "Runtime release level is malformed.")
    version = _version(record["version"]) if type(record["version"]) is str else None
    if version is None:
        if type(record["version"]) is str and _UNSTABLE_VERSION.fullmatch(record["version"]):
            return _result(False, "unstable_python", "A stable final Python release is required.")
        return _result(False, "malformed_runtime", "Runtime version must contain three bounded integers.")
    if version != tuple(info):
        return _result(False, "malformed_runtime", "Runtime version and version_info disagree.")
    if record["releaselevel"] != "final" or record["serial"] != 0:
        return _result(False, "unstable_python", "A stable final Python release is required.")
    if not record["isolated"] or not record["dont_write_bytecode"]:
        return _result(False, "malformed_runtime", "Isolated mode and no-bytecode mode are required.")
    python = policy["python"]
    if record["implementation"] != python["implementation"] or record["gil_disabled"]:
        return _result(False, "unsupported_flavor", "Regular GIL-enabled CPython is required.")
    if (record["platform"] != python["platform"] or record["machine"] != python["machine"]
            or record["bits"] != python["bits"]):
        return _result(False, "unsupported_architecture", "Regular Windows x64 Python is required.")
    branch = next((item for item in python["branches"]
                   if (item["major"], item["minor"]) == version[:2]), None)
    observed = ".".join(str(item) for item in version)
    if branch is None:
        exclusion = next((item for item in policy["excluded_python_branches"]
                          if (item["major"], item["minor"]) == version[:2]), None)
        detail = exclusion["reason"] if exclusion else "This Python minor branch is outside the reviewed compatibility policy."
        return _result(False, "unsupported_python", "Python " + observed + ": " + detail)
    branch_name = str(branch["major"]) + "." + str(branch["minor"])
    if version[2] < branch["security_min_patch"] or version[2] in branch["blocked_patches"]:
        floor = branch_name + "." + str(branch["security_min_patch"])
        return _result(False, "python_security_floor", "Python " + observed + " is excluded by the reviewed security policy; "
                       "the branch security floor is " + floor + ". See " + ", ".join(branch["security_sources"]) + ".")
    if version[2] < branch["compatibility_min_patch"]:
        floor = branch_name + "." + str(branch["compatibility_min_patch"])
        return _result(False, "python_compatibility_floor", "Python " + observed + " is below compatibility floor " + floor + ".")
    return _result(True, "supported", "Python " + observed + " is accepted by compatibility policy "
                   + str(policy["policy_version"]) + ".")


def _runtime_identity():
    return {
        "protocol": "winbooksplit.runtime", "schema_version": 1, "ok": False,
        "executable": sys.executable, "version": platform.python_version(),
        "version_info": list(sys.version_info[:3]), "releaselevel": sys.version_info.releaselevel,
        "serial": sys.version_info.serial, "implementation": sys.implementation.name,
        "platform": sys.platform, "machine": platform.machine(), "bits": struct.calcsize("P") * 8,
        "gil_disabled": bool(sysconfig.get_config_var("Py_GIL_DISABLED")),
        "isolated": bool(sys.flags.isolated), "dont_write_bytecode": bool(sys.flags.dont_write_bytecode),
        "policy_version": 0, "policy_sha256": "", "reason": "malformed_runtime", "message": "Runtime probe has not completed.",
    }


class _PackageError(ValueError):
    def __init__(self, reason, message):
        super().__init__(message)
        self.reason = reason


def _locked_pypdf(policy):
    """Import only after interpreter acceptance; retain package and metadata guards."""
    try:
        package = importlib.import_module("pypdf")
        distribution = metadata.distribution("pypdf")
    except (ImportError, metadata.PackageNotFoundError) as error:
        raise _PackageError("missing_pypdf", "The locked pypdf package is missing or cannot be imported: " + str(error)) from error
    locked = policy["pypdf"]["version"]
    actual = getattr(package, "__version__", None)
    if type(actual) is not str or actual != locked or distribution.version != locked:
        raise _PackageError("pypdf_version_mismatch", "Expected locked pypdf " + locked + " in both import and distribution metadata.")
    try:
        package_file = getattr(package, "__file__", None)
        if type(package_file) is not str or not package_file:
            raise ValueError("Missing package import origin.")
        origin = Path(package_file).resolve(strict=True)
        if not origin.is_file():
            raise ValueError("Package import origin is not a file.")
        roots = {Path(sysconfig.get_paths()[key]).resolve() for key in ("purelib", "platlib")}
        if not any(origin.is_relative_to(root) for root in roots):
            raise ValueError("Package imported outside this interpreter's package roots.")
        distribution_root = Path(distribution.locate_file("")).resolve(strict=True)
        if distribution_root not in roots or Path(distribution.locate_file("pypdf/__init__.py")).resolve(strict=True) != origin:
            raise ValueError("Distribution metadata and imported package roots disagree.")
        if re.sub(r"[-_.]+", "-", distribution.metadata.get("Name", "")).lower() != "pypdf":
            raise ValueError("Distribution metadata names a different project.")
        files = distribution.files
        if files is None:
            raise ValueError("Distribution metadata has no file inventory.")
        metadata_files = [item for item in files if Path(str(item)).name == "METADATA"
                          and Path(str(item)).parent.name.startswith("pypdf-")
                          and Path(str(item)).parent.suffix == ".dist-info"]
        if len(metadata_files) != 1:
            raise ValueError("Expected one pypdf distribution METADATA origin.")
        metadata_path = Path(distribution.locate_file(metadata_files[0])).resolve(strict=True)
        if not metadata_path.is_file() or metadata_path.parent.parent != distribution_root:
            raise ValueError("Distribution METADATA origin is outside this interpreter's package root.")
        package_root = origin.parent
        for name, module in tuple(sys.modules.items()):
            if name == "pypdf" or name.startswith("pypdf."):
                module_file = getattr(module, "__file__", None)
                if type(module_file) is not str or not module_file:
                    raise ValueError("Missing pypdf module import origin.")
                module_origin = Path(module_file).resolve(strict=True)
                if not module_origin.is_file() or not module_origin.is_relative_to(package_root):
                    raise ValueError("A pypdf module imported outside the selected package.")
    except (OSError, ValueError, TypeError, KeyError) as error:
        raise _PackageError("pypdf_origin_invalid", "Unexpected pypdf package or distribution origin: " + str(error)) from error
    return {"pypdf_version": actual, "pypdf_path": str(origin)}


def probe_interpreter(require_pypdf=True, policy_path=None):
    """Produce one offline identity record; interpreter-only probing imports no pypdf."""
    record = _runtime_identity()
    try:
        policy, policy_bytes = _read_policy(policy_path)
        record["policy_version"] = policy["policy_version"]
        record["policy_sha256"] = sha256(policy_bytes).hexdigest()
    except PolicyError as error:
        record.update(ok=False, reason="malformed_policy", message=str(error))
        return record
    decision = validate_interpreter(record, policy)
    record.update(decision)
    if not decision["ok"]:
        return record
    if type(require_pypdf) is not bool:
        record.update(ok=False, reason="malformed_runtime", message="The package probe option must be a boolean.")
        return record
    if require_pypdf:
        try:
            record.update(_locked_pypdf(policy))
        except _PackageError as error:
            record.update(ok=False, reason=error.reason, message=str(error))
            return record
        except Exception as error:
            record.update(ok=False, reason="pypdf_origin_invalid", message="The locked pypdf probe failed: " + str(error))
            return record
        record["message"] += " Locked pypdf " + record["pypdf_version"] + " imported from this interpreter's package roots."
    else:
        record["message"] += " Interpreter-only probe; pypdf was not imported."
    return record


def main(argv=None):
    """No setup, discovery or installation side effects; emit exactly one JSON record."""
    import argparse
    parser = argparse.ArgumentParser(description="Read-only WinBookSplit interpreter preflight")
    parser.add_argument("--probe", action="store_true", required=True)
    parser.add_argument("--interpreter-only", action="store_true")
    parser.add_argument("--policy", type=str)
    arguments = parser.parse_args(argv)
    record = probe_interpreter(require_pypdf=not arguments.interpreter_only, policy_path=arguments.policy)
    print(json.dumps(record, ensure_ascii=True))
    return 0 if record["ok"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
