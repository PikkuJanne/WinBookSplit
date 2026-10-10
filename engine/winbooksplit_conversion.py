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
from pypdf.generic import ArrayObject, DictionaryObject, IndirectObject, NumberObject


MAX_PDF_SNAPSHOT_BYTES = 512 * 1024 * 1024
MAX_PDF_PAGE_TREE_DEPTH = 64
MAX_PDF_PAGE_TREE_NODES = 200000
MAX_PDF_PAGES = 100000
MAX_PDF_FEATURE_DEPTH = 128
MAX_PDF_FEATURE_NODES = 200000
INHERITED_PAGE_FIELDS = frozenset({"/Resources", "/MediaBox", "/CropBox", "/Rotate"})
PDF_ACTION_TYPES = frozenset({"/GoTo", "/GoToR", "/GoToE", "/Launch", "/Thread", "/URI", "/Sound", "/Movie", "/Hide",
    "/Named", "/SubmitForm", "/ResetForm", "/ImportData", "/JavaScript", "/SetOCGState", "/Rendition", "/Trans", "/GoTo3DView"})
STATIC_ANNOTATION_FIELDS = frozenset({"/Type", "/Subtype", "/Rect", "/Contents", "/NM", "/M", "/F", "/C", "/CA", "/BS",
    "/Border", "/AP", "/AS", "/T", "/Open", "/Name", "/Subj", "/QuadPoints", "/InkList", "/L", "/LE", "/IC",
    "/RD", "/IT", "/CL", "/Rotate", "/DA", "/Q", "/DS", "/RC", "/CreationDate", "/State", "/StateModel", "/Vertices"})


class PdfPolicyError(ValueError):
    """Fixed local policy failures; never interpolate document-controlled text."""
    def __init__(self, code, message):
        super().__init__(message)
        self.code = code


def _unsupported_pdf(message):
    raise PdfPolicyError("unsupported_document", message)


def _invalid_pdf(message):
    raise PdfPolicyError("invalid_document", message)


def _pdf_value(value):
    return value.get_object() if isinstance(value, IndirectObject) else value


def _bounded_pdf_snapshot(stream):
    if isinstance(stream, (str, os.PathLike)):
        with open(stream, "rb") as source:
            data = source.read(MAX_PDF_SNAPSHOT_BYTES + 1)
    else:
        position = stream.tell()
        try:
            stream.seek(0)
            data = stream.read(MAX_PDF_SNAPSHOT_BYTES + 1)
        finally:
            stream.seek(position)
    if not isinstance(data, bytes):
        _invalid_pdf("A PDF input must be a binary byte snapshot.")
    if len(data) > MAX_PDF_SNAPSHOT_BYTES:
        _invalid_pdf("The PDF exceeds the supported 512 MiB input limit.")
    return BytesIO(data)


def _preflight_page_tree(catalog):
    """Validate every raw child before pypdf can flatten or skip a damaged one."""
    if "/Pages" not in catalog:
        _invalid_pdf("The PDF catalog has no page tree.")
    pending = [(catalog.get("/Pages"), 0, False)]
    visited, counts, pages = set(), {}, []
    while pending:
        value, depth, finish = pending.pop()
        node = _pdf_value(value)
        if not isinstance(node, DictionaryObject):
            _invalid_pdf("Every PDF page-tree child must be a dictionary.")
        identity = id(node)
        if finish:
            kids = _pdf_value(node.get("/Kids"))
            count = sum(counts[id(_pdf_value(kid))] for kid in kids)
            if _pdf_value(node.get("/Count")) != count:
                _invalid_pdf("The PDF page-tree count disagrees with its physical children.")
            counts[identity] = count
            continue
        if identity in visited:
            _invalid_pdf("The PDF page tree repeats a page or contains a cycle.")
        visited.add(identity)
        if depth > MAX_PDF_PAGE_TREE_DEPTH or len(visited) > MAX_PDF_PAGE_TREE_NODES:
            _invalid_pdf("The PDF page tree exceeds the supported traversal limit.")
        kind = node.get("/Type")
        if kind == "/Page":
            if "/Kids" in node:
                _invalid_pdf("A physical PDF page cannot have page-tree children.")
            pages.append(node)
            if len(pages) > MAX_PDF_PAGES:
                _invalid_pdf("The PDF exceeds the supported 100000 physical-page limit.")
            counts[identity] = 1
        elif kind == "/Pages":
            kids, count = _pdf_value(node.get("/Kids")), _pdf_value(node.get("/Count"))
            if not isinstance(kids, ArrayObject) or not isinstance(count, (int, NumberObject)) or isinstance(count, bool) or count < 0:
                _invalid_pdf("The PDF page tree has invalid children or a noninteger page count.")
            if len(kids) > MAX_PDF_PAGE_TREE_NODES:
                _invalid_pdf("The PDF page tree exceeds the supported traversal limit.")
            pending.append((node, depth, True))
            pending.extend((kid, depth + 1, False) for kid in reversed(kids))
        else:
            _invalid_pdf("The PDF page tree contains an object that is not a physical page or page branch.")
    if not pages:
        _invalid_pdf("The PDF must contain at least one physical page.")
    return pages, visited


def _preflight_pdf_features(reader):
    catalog = reader.root_object
    if not isinstance(catalog, DictionaryObject) or catalog.get("/Type") != "/Catalog":
        _invalid_pdf("The PDF catalog is not a valid document dictionary.")
    for key, message in (
        ("/AcroForm", "Interactive AcroForm/XFA documents are not supported."),
        ("/Perms", "Signed or permission-signature documents are not supported."),
        ("/Collection", "PDF portfolios are not supported."),
        ("/OpenAction", "Document opening actions or destinations are not supported."),
        ("/AA", "Document-level additional actions are not supported."),
        ("/AF", "Associated or embedded document attachments are not supported."),
    ):
        if key in catalog:
            _unsupported_pdf(message)
    names = _pdf_value(catalog.get("/Names"))
    if isinstance(names, DictionaryObject) and any(key in names for key in ("/JavaScript", "/EmbeddedFiles")):
        _unsupported_pdf("Document scripts and embedded-file name trees are not supported.")
    pages, page_tree_ids = _preflight_page_tree(catalog)
    physical_page_ids = {id(page) for page in pages}
    forbidden_annotations = ("/Widget", "/FileAttachment", "/RichMedia", "/Screen", "/Movie", "/Sound", "/3D")
    projected_annotations, examined_annotations = [], 0
    for page in pages:
        if "/AF" in page:
            _unsupported_pdf("Associated page attachments are not supported.")
        annotations = _pdf_value(page.get("/Annots"))
        if isinstance(annotations, ArrayObject):
            if len(annotations) > MAX_PDF_FEATURE_NODES:
                _invalid_pdf("PDF annotations exceed the supported inspection limit.")
            for reference in annotations:
                examined_annotations += 1
                if examined_annotations > MAX_PDF_FEATURE_NODES:
                    _invalid_pdf("PDF annotations exceed the supported inspection limit.")
                annotation = _pdf_value(reference)
                if isinstance(annotation, DictionaryObject) and (
                    _pdf_value(annotation.get("/Subtype")) in forbidden_annotations or _pdf_value(annotation.get("/FT")) == "/Sig"
                    or _pdf_value(annotation.get("/Type")) == "/Sig" or "/AF" in annotation or "/EF" in annotation
                ):
                    _unsupported_pdf("Interactive widgets, signatures, attachments and multimedia annotations are not supported.")
                if isinstance(annotation, DictionaryObject):
                    projected_annotations.extend(value for key, value in annotation.items() if key in STATIC_ANNOTATION_FIELDS)
    # Ordinary PDF graphs have Parent/P cycles. Repeated graph nodes are visited
    # once; only the physical page tree above rejects repeats. Do not decode
    # content/image streams or execute actions to identify these structures.
    pending = [(catalog, 0, False, False)] + [(value, 0, True, True) for value in projected_annotations]
    visited = set()
    while pending:
        value, depth, annotation_projection, copied_page_value = pending.pop()
        node = _pdf_value(value)
        if not isinstance(node, (DictionaryObject, ArrayObject)):
            continue
        identity = (id(node), annotation_projection, copied_page_value)
        if identity in visited:
            continue
        visited.add(identity)
        if len(visited) > MAX_PDF_FEATURE_NODES or depth > MAX_PDF_FEATURE_DEPTH:
            _invalid_pdf("The PDF object graph exceeds the supported inspection limit.")
        if isinstance(node, DictionaryObject):
            kind = _pdf_value(node.get("/Type"))
            # These projected values are dropped wholesale by the existing inert
            # annotation guard; do not traverse their excluded document backlinks.
            if annotation_projection and kind in ("/Page", "/Pages", "/Catalog", "/Action"):
                continue
            if copied_page_value and kind in ("/Page", "/Pages", "/Catalog"):
                _unsupported_pdf("Copied page resources cannot reference document catalogs or physical page trees.")
            if kind in ("/Sig", "/EmbeddedFile") or _pdf_value(node.get("/FT")) == "/Sig" or "/EF" in node or "/AF" in node:
                _unsupported_pdf("Signatures and embedded or associated attachments are not supported.")
            action = _pdf_value(node.get("/S"))
            if not annotation_projection and (kind == "/Action" or isinstance(action, str)
                    and action in PDF_ACTION_TYPES or "/JS" in node):
                _unsupported_pdf("Active document or resource actions are not supported.")
            # The existing planner bounds outline/destination retrieval. The
            # existing writer rebuilds annotations and excludes page AA/articles;
            # their tested warnings remain separate from catalog policy.
            skipped = set()
            if node is catalog:
                skipped.update({"/Outlines", "/Dests"})
            if node is names:
                skipped.add("/Dests")
            if id(node) in physical_page_ids:
                skipped.update({"/Parent", "/Annots", "/AA", "/B"})
            if len(node) > MAX_PDF_FEATURE_NODES:
                _invalid_pdf("A PDF dictionary exceeds the supported inspection limit.")
            pending.extend((child, depth + 1, annotation_projection, copied_page_value or id(node) in physical_page_ids
                            or id(node) in page_tree_ids and key in INHERITED_PAGE_FIELDS)
                           for key, child in node.items() if key not in skipped)
        else:
            if len(node) > MAX_PDF_FEATURE_NODES:
                _invalid_pdf("A PDF array exceeds the supported inspection limit.")
            pending.extend((child, depth + 1, annotation_projection, copied_page_value) for child in node)


class PolicyPdfReader(PdfReader):
    """Pinned pypdf reader guard shared by direct and converted PDF inputs."""
    def _handle_encryption(self, password):
        # pypdf6.19 otherwise reads encryption and tries verify(b'') in __init__.
        # Reject at its hook before Encryption.read/verify/decrypt is reached.
        _unsupported_pdf("Encrypted PDFs are not supported. No decryption is attempted.")

    def __init__(self, stream, strict=False):
        super().__init__(_bounded_pdf_snapshot(stream), strict=strict, root_object_recovery_limit=4096)
        _preflight_pdf_features(self)


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
        reader = PolicyPdfReader(BytesIO(data))
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
    except PdfPolicyError as error:
        code = "unsupported_document" if error.code == "unsupported_document" else "conversion_output_invalid"
        raise ConversionError(code, str(error)) from error
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
        try:
            data = _bounded_pdf_snapshot(path).getvalue()
        except PdfPolicyError as error:
            raise ConversionError("conversion_output_invalid", str(error)) from error
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
