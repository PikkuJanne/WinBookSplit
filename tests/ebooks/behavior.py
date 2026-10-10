"""Owned process configuration and bounded loopback observations for M4-T04.

The observer is not a firewall or a network sandbox. Its evidence concerns only
the two declared resources on its own live loopback listener.
"""
from __future__ import annotations

import base64
from contextlib import contextmanager
from datetime import datetime, timezone
from hashlib import sha256
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import os
from pathlib import Path
import stat
import subprocess
import threading
import time
from urllib.request import ProxyHandler, build_opener
import uuid

VARIABLES = ("CALIBRE_CONFIG_DIRECTORY", "CALIBRE_CACHE_DIRECTORY", "CALIBRE_TEMP_DIR")
PNG = base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aM3sAAAAASUVORK5CYII=")
CSS = b"/* Original loopback canary; no imported resource. */\nbody { color: black; }\n"
NATIVE_OPERATIONS = []


def need(condition, message):
    if not condition:
        raise ValueError(message)


def identity(path):
    path = Path(path)
    details = path.lstat()
    need(stat.S_ISREG(details.st_mode) and not getattr(details, "st_file_attributes", 0) & 1024,
         "Owned configuration contains a nonordinary file")
    return {"sha256": sha256(path.read_bytes()).hexdigest(), "size_bytes": details.st_size,
            "device": details.st_dev, "inode": details.st_ino,
            "attributes": getattr(details, "st_file_attributes", 0)}


def configuration_snapshot(root):
    root = Path(root)
    details = root.lstat()
    need(root.is_dir() and not getattr(details, "st_file_attributes", 0) & 1024,
         "Owned configuration directory changed to a reparse object")
    files, directories = {}, {}
    for path in sorted(root.rglob("*")):
        details = path.lstat()
        need(not getattr(details, "st_file_attributes", 0) & 1024, "Owned configuration contains a reparse object")
        relative = path.relative_to(root).as_posix()
        need(len(files) + len(directories) < 4096, "Owned configuration exceeds bounded fixture scope")
        if path.is_dir():
            directories[relative] = {"device": details.st_dev, "inode": details.st_ino}
        else:
            need(details.st_size <= 32 * 1024 * 1024, "Owned configuration file exceeds bounded fixture scope")
            files[relative] = identity(path)
    return {"path": str(root), "device": root.stat().st_dev, "inode": root.stat().st_ino,
            "directories": directories, "files": files}


class CalibreEnvironment:
    """Select new process-local directories without editing installed settings."""
    def __init__(self, directory):
        self.directory = Path(directory)
        self.directory.mkdir()
        self.values = {}
        for variable, name in zip(VARIABLES, ("config", "cache", "temp")):
            destination = self.directory / name
            destination.mkdir()
            self.values[variable] = str(destination)
        self.before = self.snapshot()
        self.after = None
        self.restored = False
        self.original_environment = {key: os.environ.get(key) for key in VARIABLES}
        self.restored_environment = None

    def snapshot(self):
        return {key: configuration_snapshot(Path(value)) for key, value in self.values.items()}

    def environment(self, ordinary):
        return {**ordinary, **self.values}

    @contextmanager
    def inherited(self):
        previous = {key: os.environ.get(key) for key in VARIABLES}
        os.environ.update(self.values)
        try:
            yield
        finally:
            for key, value in previous.items():
                if value is None:
                    os.environ.pop(key, None)
                else:
                    os.environ[key] = value
            self.restored = all(os.environ.get(key) == value for key, value in previous.items())
            self.restored_environment = {key: os.environ.get(key) for key in VARIABLES}
            self.after = self.snapshot()

    def receipt(self):
        return {"environment": self.values, "before": self.before,
                "after": self.after if self.after is not None else self.snapshot(),
                "new_empty_process_directories": all(not item["files"] and not item["directories"] for item in self.before.values()),
                "inherited_environment_restored": self.restored,
                "original_environment": self.original_environment,
                "restored_environment": self.restored_environment,
                "scope": "Process-local config/cache/temp only; no installed or user configuration edited."}


class NativeFailure(RuntimeError):
    def __init__(self, message, process):
        super().__init__(message)
        self.process = process
        self.cleanup_safe = False


def native(command, cwd, environment, timeout=90):
    """Retain both streams and real native status even on a bounded timeout."""
    started, clock = datetime.now(timezone.utc).isoformat(), time.monotonic()
    child = subprocess.Popen(list(map(str, command)), cwd=cwd, env=environment,
                             stdin=subprocess.DEVNULL, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    observation = {"actual_process": True, "command": list(map(str, command)), "cwd": str(cwd),
                   "stdin": "DEVNULL", "stdin_utf8": None, "pid": child.pid,
                   "started_at": started, "timeout_seconds": timeout}
    NATIVE_OPERATIONS.append(observation)
    timed_out = False
    try:
        stdout, stderr = child.communicate(timeout=timeout)
    except subprocess.TimeoutExpired as error:
        timed_out = True
        stdout, stderr = error.output or b"", error.stderr or b""
        if child.poll() is None:
            child.terminate()  # Retained process handle; no inferred descendant-stop proof.
        observation["shutdown_unproved"] = True
    observation.update(finished_at=datetime.now(timezone.utc).isoformat(), elapsed_seconds=time.monotonic() - clock,
                       timed_out=timed_out, exit_code=child.poll(),
                       stdout=stdout.decode("utf-8", "strict"), stderr=stderr.decode("utf-8", "strict"),
                       stdout_sha256=sha256(stdout).hexdigest(), stderr_sha256=sha256(stderr).hexdigest())
    if timed_out:
        raise NativeFailure("Owned native deadline expired; descendant stop unproved, retain workspace", observation)
    return observation


def local_help(calibre, sources, work, environment):
    """Actual selected-format help, not an unversioned online manual."""
    result = {}
    for fmt in ("epub", "azw3"):
        target = Path(work) / ("help-only-" + fmt + ".pdf")
        before = identity(sources[fmt])
        process = native([calibre, sources[fmt], target, "--help"], work, environment, timeout=60)
        need(process["exit_code"] == 0 and not target.exists() and before == identity(sources[fmt]),
             "Selected-format local help created output or changed authored source")
        need("--output-profile" in process["stdout"] and "tablet" in process["stdout"],
             "Pinned converter help omitted the selected supported profile")
        result[fmt] = {"process": process, "input_path": str(sources[fmt]), "input_before": before,
                       "input_after": identity(sources[fmt]), "output_absent": True,
                       "selected_profile": "tablet", "profile_option_observed": "--output-profile"}
    return result


class LoopbackCanary:
    """One finite observer namespace; positive controls use the same endpoints."""
    def __init__(self):
        self.nonce = uuid.uuid4().hex
        self.phase = "created"
        self.events = []
        self.lock = threading.Lock()
        self.positive_controls = []
        owner = self
        class Handler(BaseHTTPRequestHandler):
            def setup(self):
                super().setup()
                self.connection.settimeout(5)

            def log_message(self, *args):
                pass

            def observe(self, head=False):
                length = self.headers.get("Content-Length", "0")
                bounded_length = int(length) if length.isdigit() and len(length) < 8 else -1
                # GET/HEAD never consume a body. Unexpected bodies are not accepted.
                with owner.lock:
                    need(len(owner.events) < 64, "Owned loopback event bound exceeded")
                    owner.events.append({"sequence": len(owner.events) + 1, "phase": owner.phase,
                        "method": self.command, "path": self.path,
                        "body_length_declared": bounded_length,
                        "observed_at": datetime.now(timezone.utc).isoformat()})
                data, content_type = (PNG, "image/png") if self.path == owner.paths[0] else (CSS, "text/css")
                accepted = self.command in {"GET", "HEAD"} and self.path in owner.paths and bounded_length == 0
                self.send_response(200 if accepted else 404)
                self.send_header("Content-Type", content_type if accepted else "text/plain")
                self.send_header("Content-Length", str(len(data) if accepted else 0))
                self.end_headers()
                if accepted and not head:
                    self.wfile.write(data)

            def do_GET(self): self.observe()
            def do_HEAD(self): self.observe(head=True)
            def do_POST(self): self.observe()
            def do_PUT(self): self.observe()
            def do_DELETE(self): self.observe()
        self.paths = [f"/{self.nonce}/image.png", f"/{self.nonce}/style.css"]
        self.server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        self.server.daemon_threads = False
        self.server.block_on_close = True
        self.port = self.server.server_address[1]
        self.base_url = f"http://127.0.0.1:{self.port}/{self.nonce}/"
        self.urls = [self.base_url + "image.png", self.base_url + "style.css"]
        self.thread = threading.Thread(target=self.server.serve_forever, name="WBS-authored-loopback", daemon=True)
        self.thread.start()
        self.stopped = False

    def positive(self, phase):
        need(phase in {"before", "after"}, "Unexpected loopback control phase")
        self.phase = "control-" + phase
        opener = build_opener(ProxyHandler({}))  # Do not use machine or external proxy settings.
        for url, expected in zip(self.urls, (PNG, CSS)):
            started = datetime.now(timezone.utc).isoformat()
            with opener.open(url, timeout=5) as response:
                data, status = response.read(4096), response.status
            need(status == 200 and data == expected, "Actual same-endpoint positive GET failed")
            self.positive_controls.append({"phase": phase, "url": url, "method": "GET", "status": status,
                "body_sha256": sha256(data).hexdigest(), "size_bytes": len(data), "started_at": started,
                "finished_at": datetime.now(timezone.utc).isoformat()})
        self.phase = "idle"

    @contextmanager
    def conversion(self, case_id):
        self.phase = "conversion:" + case_id
        try:
            yield
        finally:
            self.phase = "idle"

    def stop(self):
        self.server.shutdown()
        self.server.server_close()
        self.thread.join(timeout=5)
        self.stopped = not self.thread.is_alive()
        need(self.stopped, "Owned loopback listener did not stop")

    def receipt(self):
        with self.lock:
            events = list(self.events)
        conversion_events = [item for item in events if item["phase"].startswith("conversion:")]
        return {"binding": {"host": "127.0.0.1", "port": self.port, "nonce": self.nonce},
                "urls": self.urls, "events": events, "positive_controls": list(self.positive_controls),
                "conversion_requests": conversion_events, "listener_stopped": self.stopped,
                "scope": "Observed requests only to two authored loopback endpoints; no OS-wide network denial or sandbox claim.",
                "system_network_settings_changed": False}
