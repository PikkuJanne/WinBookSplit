"""Controlled whole-entrypoint engine substitute; never a passing split claim."""

import json
import os
from pathlib import Path
import sys

# Loaded only from the owned application copy, with isolated -I -B.
import importlib.util
spec = importlib.util.spec_from_file_location("owned_native_control", os.environ["WBS_PROCESS_CHILD"])
control = importlib.util.module_from_spec(spec)
spec.loader.exec_module(control)
receipt = {"argv": sys.argv[1:], "python": sys.executable, "encoding": sys.stdout.encoding,
           "utf8_env": os.environ.get("PYTHONUTF8"), "io_env": os.environ.get("PYTHONIOENCODING")}
Path(os.environ["WBS_PROCESS_RECEIPT"]).write_text(json.dumps(receipt, ensure_ascii=False), encoding="utf-8")
kind = os.environ["WBS_PROCESS_KIND"]
mode = sys.argv[3]
if kind == "flood":
    control.flood(mode)
elif kind == "fast-tail":
    control.emit(1, b"\n\nstdout blank\n\n" + control.encoded_frame(control.UNICODE, mode))
    control.emit(2, b"\n\nFINAL_STDERR_NO_NEWLINE_" + control.UNICODE.encode("utf-8"))
else:
    raise ValueError("Unknown controlled engine case")
raise SystemExit(2)
