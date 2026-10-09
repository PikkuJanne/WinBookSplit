"""Owned-copy adapter injecting faults into the real shipped PDF engine.

The application uses this file only in a copied acceptance fixture. The saved
engine still reads, plans and writes the authored PDF; no positive result is
invented. All introduced children are finite and report their own actual PIDs.
"""

import importlib.util
import json
import os
from pathlib import Path
import subprocess
import sys
import time


def receipt(value):
    target = Path(os.environ["WBS_OUTCOME_RECEIPT"])
    with target.open("x", encoding="utf-8") as stream:
        json.dump(value, stream)


def main():
    if len(sys.argv) == 2 and sys.argv[1] == "owned-descendant":
        Path(os.environ["WBS_OUTCOME_CHILD_RECEIPT"]).write_text(
            json.dumps({"pid": os.getpid(), "python": sys.executable}), encoding="utf-8")
        time.sleep(20)
        return 0
    original = Path(__file__).with_name("saved_engine.py")
    spec = importlib.util.spec_from_file_location("wbs_owned_outcome_engine", original)
    engine = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = engine
    spec.loader.exec_module(engine)
    kind = os.environ["WBS_OUTCOME_KIND"]
    written = engine.write_slice

    def record(stage):
        stage = Path(stage)
        receipt({"kind": kind, "pid": os.getpid(), "python": sys.executable,
            "stage": str(stage), "stage_identity": [stage.stat().st_dev, stage.stat().st_ino],
            "members": {path.name: {"identity": [path.stat().st_dev, path.stat().st_ino],
                "sha256": engine.sha256(path.read_bytes()).hexdigest()}
                for path in stage.iterdir()}})

    def slice_fault(reader, start, end, out_path, on_created=None):
        written(reader, start, end, out_path, on_created=on_created)
        stage = Path(out_path).parent
        if kind == "cleanup-refusal":
            (stage / "authored-unrelated.txt").write_bytes(b"Keep this unregistered acceptance member\n")
        record(stage)
        if kind in {"engine-cancel", "engine-timeout", "engine-ctrlc"}:
            subprocess.Popen([str(Path(sys.base_prefix) / "python.exe"), "-I", "-B", "-X", "utf8",
                __file__, "owned-descendant"], stdout=sys.stdout, stderr=sys.stderr)
            deadline = time.monotonic() + 2
            while not Path(os.environ["WBS_OUTCOME_CHILD_RECEIPT"]).exists() and time.monotonic() < deadline:
                time.sleep(.005)
            time.sleep(20)
        raise OSError("Authored outcome mid-write failure")

    def manifest_fault(run, manifest):
        record(run.stage)
        raise OSError("Authored outcome completion-manifest failure")

    def publish_fault(run, final_name):
        record(run.stage)
        raise OSError("Authored outcome no-replace publish-rename failure")

    if kind in {"output-write", "cleanup-refusal", "engine-cancel", "engine-timeout", "engine-ctrlc"}:
        engine.write_slice = slice_fault
    elif kind == "manifest-finalize":
        engine._write_manifest = manifest_fault
    elif kind == "publish-rename":
        engine._publish_run = publish_fault
    else:
        raise ValueError("Unknown owned outcome fault")
    return engine.main()


if __name__ == "__main__":
    raise SystemExit(main())
