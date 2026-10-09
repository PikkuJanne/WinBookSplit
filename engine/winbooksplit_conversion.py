"""Bounded Calibre conversion inside an existing flat, ledger-owned workspace."""

from collections.abc import Mapping
from dataclasses import dataclass
from hashlib import sha256
import importlib.util
from io import BytesIO
import math
import os
from pathlib import Path
import stat
import threading
import time
from types import MappingProxyType

from pypdf import PdfReader


CONVERTED_FILENAME = "WinBookSplit_Converted.pdf"
STREAM_TAIL_BYTES = 65536
TREE_STOP_SECONDS = 10


def _freeze(value):
    if isinstance(value, Mapping):
        return MappingProxyType({key: _freeze(item) for key, item in value.items()})
    if isinstance(value, (tuple, list)):
        return tuple(_freeze(item) for item in value)
    return value


class ConversionError(ValueError):
    def __init__(self, code, message, *, diagnostic=None, conversion=None, tree_stopped=True):
        super().__init__(message)
        self.code = code
        self.diagnostic = diagnostic
        self.conversion = _freeze(conversion or {})
        self.tree_stopped = tree_stopped


@dataclass(frozen=True)
class ConversionResult:
    pdf_bytes: bytes
    original_source_identity: Mapping
    generated_pdf_identity: Mapping
    conversion: Mapping


def _default_process_factory(argv, *, cwd, env):
    # Resolve only the shipped sibling; neither CWD nor the ebook directory
    # supplies a process helper or an executable discovery candidate.
    path = Path(__file__).resolve().with_name("winbooksplit_job.py")
    spec = importlib.util.spec_from_file_location("_winbooksplit_conversion_job", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module.launch_job(argv, cwd=cwd, env=env)


class _StreamTail:
    def __init__(self, stream):
        self.stream = stream
        self.tail = bytearray()
        self.total = 0
        self.error = None

    def drain(self):
        try:
            while True:
                chunk = self.stream.read(8192)
                if not chunk:
                    break
                self.total += len(chunk)
                self.tail.extend(chunk)
                if len(self.tail) > STREAM_TAIL_BYTES:
                    del self.tail[:-STREAM_TAIL_BYTES]
        except Exception as error:
            self.error = error


def _run_converter(argv, *, cwd, env, timeout_seconds, process_factory):
    """Drain both pipes from launch until the tracked tree stops and reaches EOF."""
    started = time.monotonic()
    job, tails, threads, code = None, {}, [], None
    failure, stopped = None, False
    try:
        try:
            job = process_factory(argv, cwd=cwd, env=env)
        except BaseException as error:
            stopped = bool(getattr(error, "converter_tree_stopped", False))
            error_code = "conversion_cancelled" if isinstance(error, KeyboardInterrupt) else "conversion_start_failed"
            raise ConversionError(error_code, "Calibre startup was cancelled." if isinstance(error, KeyboardInterrupt)
                                  else "Cannot start Calibre: " + str(error), tree_stopped=stopped) from error
        stopped = False
        for name, stream in (("stdout", job.stdout), ("stderr", job.stderr)):
            tail = _StreamTail(stream)
            tails[name] = tail
            thread = threading.Thread(target=tail.drain, name="WinBookSplit-converter-" + name, daemon=True)
            threads.append(thread)
            thread.start()
        deadline = started + timeout_seconds
        while True:
            code = job.poll()
            if code is not None and job.active_process_count() == 0:
                job.wait_tree(timeout=TREE_STOP_SECONDS)
                stopped = True
                break
            if time.monotonic() >= deadline:
                raise ConversionError("conversion_timeout", "Calibre exceeded the conversion timeout.")
            time.sleep(0.02)
        if code != 0:
            raise ConversionError("conversion_failed", f"Calibre exited with code {code}.")
    except KeyboardInterrupt:
        failure = ConversionError("conversion_cancelled", "Ebook conversion was cancelled.")
    except ConversionError as error:
        failure = error
    except Exception as error:
        failure = ConversionError("conversion_failed", "Cannot supervise Calibre: " + str(error))
    finally:
        if job is not None:
            if not stopped:
                try:
                    job.terminate_tree(exit_code=1)
                    job.wait_tree(timeout=TREE_STOP_SECONDS)
                    code = job.wait(timeout=TREE_STOP_SECONDS)
                    stopped = True
                except (Exception, KeyboardInterrupt) as error:
                    original = str(failure) if failure is not None else "Calibre process supervision failed."
                    failure = ConversionError("conversion_cleanup_failed", original + "; cannot confirm stopped converter tree: " + str(error),
                                              diagnostic={"primary_code": failure.code if failure is not None else None,
                                                          "primary_message": original}, tree_stopped=False)
            # Never wait for pipe EOF from a tree whose termination is unproved.
            # The native job helper uses unbuffered pipes so its all-handle
            # shutdown cannot deadlock against a drain thread's BufferedReader.
            if stopped:
                deadline = time.monotonic() + TREE_STOP_SECONDS
                for thread in threads:
                    thread.join(max(0, deadline - time.monotonic()))
                if any(thread.is_alive() for thread in threads):
                    failure = failure or ConversionError("conversion_failed", "Cannot finish both Calibre output streams.")
            for tail in tails.values():
                if tail.error is not None:
                    failure = failure or ConversionError("conversion_failed", "Cannot read Calibre output: " + str(tail.error))
            try:
                job.close()
            except Exception as error:
                original = str(failure) if failure is not None else "Calibre completed."
                failure = ConversionError("conversion_cleanup_failed", original + "; cannot close converter handles: " + str(error),
                                          diagnostic={"primary_code": (failure.diagnostic or {}).get("primary_code", failure.code)
                                                      if failure is not None else None,
                                                      "primary_message": original}, tree_stopped=stopped)
    record = {"argv": tuple(argv), "elapsed_seconds": time.monotonic() - started,
              "exit_code": code}
    for name in ("stdout", "stderr"):
        tail = tails.get(name)
        record[name + "_tail"] = bytes(tail.tail).decode("utf-8", errors="replace") if tail is not None else ""
        record[name + "_total_bytes"] = tail.total if tail is not None else 0
        record[name + "_truncated"] = tail is not None and tail.total > STREAM_TAIL_BYTES
    if failure is not None:
        failure.conversion = _freeze(record)
        failure.tree_stopped = stopped
        raise failure
    return record


def _source_snapshot(path):
    details = os.lstat(path)
    if not stat.S_ISREG(details.st_mode) or details.st_file_attributes & stat.FILE_ATTRIBUTE_REPARSE_POINT:
        raise ConversionError("conversion_ownership_failed", "The ebook must be an ordinary source file.")
    digest, size = sha256(), 0
    with open(path, "rb") as stream:
        while chunk := stream.read(1024 * 1024):
            digest.update(chunk)
            size += len(chunk)
    identity = {"path": path, "resolved_path": os.path.realpath(path), "sha256": digest.hexdigest(),
                "size_bytes": size, "binding": "ebook_snapshot"}
    fingerprint = (details.st_dev, details.st_ino, details.st_size, details.st_mtime_ns, details.st_file_attributes)
    return identity, fingerprint


def _validate_pdf(data, path):
    if not data:
        raise ConversionError("conversion_output_invalid", "Calibre produced no nonempty PDF.")
    try:
        reader = PdfReader(BytesIO(data))
        if reader.is_encrypted:
            raise ValueError("Encrypted converted PDFs are not supported.")
        count = len(reader.pages)
        if count < 1:
            raise ValueError("The converted PDF contains no pages.")
        contents = []
        for page in reader.pages:
            page.get_object()
            stream = page.get_contents()
            contents.append(sha256(stream.get_data() if stream is not None else b"").hexdigest())
        return {"path": path, "resolved_path": os.path.realpath(path), "sha256": sha256(data).hexdigest(),
                "size_bytes": len(data), "page_count": count, "binding": "reader_snapshot",
                "page_content_sha256": contents}
    except Exception as error:
        raise ConversionError("conversion_output_invalid", "Cannot validate the converted PDF: " + str(error)) from error


def _cleanup_workspace(run):
    run.base_guard.restrict_writes(run.stage_guard)
    if not run.cleanup():
        raise OSError("The conversion workspace remains after cleanup.")


def convert_ebook(source_path, converter_path, output_base, *, new_run, timeout_seconds=1800,
                  process_factory=None, record_failure=None):
    """Return immutable validated bytes after removing the conversion workspace.

    new_run is the existing OutputRun constructor. record_failure may delegate
    to its existing bounded failure-record routine; absent that callback this
    helper reports cleanup state without claiming a diagnostic file was saved.
    The process factory must start a suspended child in a kill-on-close job,
    assign it before resume and leave no running child when construction fails.
    """
    if isinstance(timeout_seconds, bool) or not isinstance(timeout_seconds, (int, float)) \
            or not math.isfinite(timeout_seconds) or timeout_seconds <= 0:
        raise ConversionError("conversion_failed", "Conversion timeout must be a positive finite number of seconds.")
    source_path = os.path.abspath(os.fspath(source_path))
    try:
        converter_path = os.fspath(converter_path)
        if not isinstance(converter_path, str) or not os.path.isabs(converter_path):
            raise ValueError("Provide an absolute converter executable path.")
        converter_path = os.path.abspath(converter_path)
        details = os.lstat(converter_path)
        if not stat.S_ISREG(details.st_mode) or details.st_file_attributes & stat.FILE_ATTRIBUTE_REPARSE_POINT:
            raise ValueError("The converter must be an ordinary executable file, not a reparse point.")
    except (TypeError, ValueError, OSError) as error:
        raise ConversionError("converter_not_found", "Cannot use the selected Calibre converter: " + str(error)) from error
    if Path(source_path).suffix.casefold() not in {".epub", ".azw3"}:
        raise ConversionError("conversion_failed", "Convert an EPUB or AZW3 source.")
    factory = process_factory or _default_process_factory
    run, source_guard, failure, result = None, None, None, None
    conversion = {"converter_path": converter_path, "output_profile": "tablet", "timeout_seconds": timeout_seconds}
    diagnostic = None
    try:
        run = new_run(output_base, source_path)
        source_guard = run.windows.FileGuard(source_path)
        original, fingerprint = _source_snapshot(source_path)
        conversion["original_source_identity"] = original
        source_guard.assert_unchanged()
        run.write_owned(CONVERTED_FILENAME, b"")
        path = os.path.join(run.stage, CONVERTED_FILENAME)
        # Only this parent-created identity is writable by the converter. The
        # stage/ancestors remain held; replacement and unknown members reject.
        run.file_guards.pop(path).close()
        for guard in (*run.guards, run.stage_guard):
            guard.assert_unchanged()
        argv = (converter_path, source_path, path, "--output-profile", "tablet")
        try:
            conversion.update(_run_converter(argv, cwd=run.stage,
                env={**os.environ, "PYTHONUTF8": "1", "PYTHONIOENCODING": "utf-8:backslashreplace"},
                timeout_seconds=timeout_seconds, process_factory=factory))
        except ConversionError as error:
            error.conversion = _freeze({**conversion, **error.conversion})
            raise
        run.seal_file(path)
        run.assert_owned()
        data = Path(path).read_bytes()
        generated = _validate_pdf(data, path)
        source_guard.assert_unchanged()
        after, after_fingerprint = _source_snapshot(source_path)
        if after != original or after_fingerprint != fingerprint:
            raise ConversionError("conversion_source_changed", "The original ebook changed during conversion.")
        result = ConversionResult(data, _freeze(original), _freeze(generated), _freeze(conversion))
    except KeyboardInterrupt:
        failure = ConversionError("conversion_cancelled", "Ebook conversion was cancelled.", conversion=conversion)
    except ConversionError as error:
        error.conversion = _freeze({**conversion, **error.conversion})
        failure = error
    except Exception as error:
        code = {"output_ownership_failed": "conversion_ownership_failed", "output_cancelled": "conversion_cancelled"}.get(
            getattr(error, "code", None), "conversion_failed")
        failure = ConversionError(code, str(error), diagnostic=getattr(error, "diagnostic", None), conversion=conversion)
    finally:
        cleanup_error, close_errors = None, []
        prior_diagnostic = failure.diagnostic or {} if failure is not None else {}
        primary_code = prior_diagnostic.get("primary_code", failure.code) if failure is not None else None
        primary_message = prior_diagnostic.get("primary_message", str(failure)) if failure is not None else None
        if run is not None:
            diagnostic = {"run_id": run.run_id, "cleanup_complete": False, "retained_staging": run.stage,
                          "record_path": None, "cleanup_error": None}
            if failure is not None and not failure.tree_stopped:
                cleanup_error = "Cleanup refused because the converter tree is not confirmed stopped."
            else:
                try:
                    if failure is not None and record_failure is not None:
                        diagnostic = dict(record_failure(run, failure))
                        if not diagnostic.get("cleanup_complete"):
                            cleanup_error = diagnostic.get("cleanup_error") or "Conversion cleanup was refused."
                    else:
                        _cleanup_workspace(run)
                        diagnostic.update(cleanup_complete=True, retained_staging=None)
                except Exception as error:
                    cleanup_error = str(error)
            try:
                run.close()
            except Exception as error:
                close_errors.append(str(error))
        if source_guard is not None:
            try:
                source_guard.close()
            except Exception as error:
                close_errors.append(str(error))
        if cleanup_error is not None or close_errors:
            diagnostic = diagnostic or {}
            diagnostic.update(cleanup_error=cleanup_error, close_error="; ".join(close_errors) or None,
                              primary_code=primary_code, primary_message=primary_message)
            failure = ConversionError("conversion_cleanup_failed", "Cannot safely finish conversion cleanup: " +
                                      "; ".join(item for item in [cleanup_error, *close_errors] if item),
                                      diagnostic=diagnostic, conversion=failure.conversion if failure is not None else conversion,
                                      tree_stopped=failure.tree_stopped if failure is not None else True)
    if failure is not None:
        if diagnostic is not None:
            if primary_code is not None:
                diagnostic.update(primary_code=primary_code, primary_message=primary_message)
            failure.diagnostic = diagnostic
        raise failure
    return ConversionResult(result.pdf_bytes, result.original_source_identity, result.generated_pdf_identity,
        _freeze({**result.conversion, "original_source_identity": result.original_source_identity,
                 "generated_pdf_identity": result.generated_pdf_identity,
                 "workspace_cleanup": {"cleanup_complete": True, "retained_staging": None}}))
