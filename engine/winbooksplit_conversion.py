"""Bounded Calibre conversion inside an existing flat, ledger-owned workspace."""

from collections.abc import Mapping
from copy import copy
from dataclasses import dataclass
from hashlib import sha256
import importlib.util
from io import BytesIO
from itertools import islice
import math
import os
from pathlib import Path
import stat
import threading
import time
from types import MappingProxyType

from pypdf import PdfReader
from pypdf.errors import DependencyError, FileNotDecryptedError, PdfReadError
from pypdf.generic import (ArrayObject, BooleanObject, ByteStringObject, ContentStream, DictionaryObject, IndirectObject,
                          NameObject, NullObject, NumberObject, StreamObject, TextStringObject)


MAX_PDF_SNAPSHOT_BYTES = 512 * 1024 * 1024
MAX_PDF_PAGE_TREE_DEPTH = 64
MAX_PDF_PAGE_TREE_NODES = 200000
MAX_PDF_PAGES = 100000
MAX_PDF_FEATURE_DEPTH = 128
MAX_PDF_FEATURE_NODES = 200000
MAX_PDF_FORM_NODES = 20000
MAX_OMITTED_FEATURE_NODES = 20000
PDF_ACTION_TYPES = frozenset({"/GoTo", "/GoToR", "/GoToE", "/Launch", "/Thread", "/URI", "/Sound", "/Movie", "/Hide",
    "/Named", "/SubmitForm", "/ResetForm", "/ImportData", "/JavaScript", "/SetOCGState", "/Rendition", "/Trans",
    "/GoTo3DView", "/GoToDp", "/RichMediaExecute"})
STATIC_ANNOTATION_FIELDS = frozenset({"/Type", "/Subtype", "/Rect", "/Contents", "/NM", "/M", "/F", "/C", "/CA", "/BS",
    "/Border", "/AP", "/AS", "/T", "/Open", "/Name", "/Subj", "/QuadPoints", "/InkList", "/L", "/LE", "/IC",
    "/RD", "/IT", "/CL", "/Rotate", "/DA", "/Q", "/DS", "/RC", "/CreationDate", "/State", "/StateModel", "/Vertices", "/OC"})
STATIC_ANNOTATION_SUBTYPES = frozenset({"/Text", "/Highlight", "/Square", "/Stamp"})
CHAPTER_PAGE_FIELDS = frozenset({"/Type", "/Resources", "/MediaBox", "/CropBox", "/BleedBox", "/TrimBox", "/ArtBox",
    "/Contents", "/Rotate", "/Group", "/UserUnit", "/SeparationInfo", "/OutputIntents", "/OC"})
OMITTED_PAGE_FIELDS = frozenset({"/A", "/AA", "/AF", "/Annots", "/B", "/Collection", "/Dur", "/EF", "/FS", "/JS",
    "/Metadata", "/Names", "/OpenAction", "/Parent", "/Perms", "/PieceInfo", "/PresSteps", "/StructParents", "/Thumb", "/Trans"})
MEDIA_ANNOTATIONS = frozenset({"/RichMedia", "/Screen", "/Movie", "/Sound", "/3D"})
RESOURCE_NAME_MAPS = frozenset({"/ExtGState", "/ColorSpace", "/Pattern", "/Shading", "/XObject", "/Font", "/Properties"})
RENDERING_RESOURCE_OPERATORS = {b"Do": ("/XObject", 0), b"Tf": ("/Font", 0), b"gs": ("/ExtGState", 0),
    b"CS": ("/ColorSpace", 0), b"cs": ("/ColorSpace", 0), b"SCN": ("/Pattern", -1), b"scn": ("/Pattern", -1),
    b"sh": ("/Shading", 0), b"BDC": ("/Properties", 1), b"DP": ("/Properties", 1)}
OPENING_DESTINATION_FITS = {"/XYZ": 3, "/Fit": 0, "/FitH": 1, "/FitV": 1,
                            "/FitR": 4, "/FitB": 0, "/FitBH": 1, "/FitBV": 1}
MAX_OPENING_DESTINATION_NODES = 10000


class PdfPolicyError(ValueError):
    """Fixed local policy failures; never interpolate document-controlled text."""
    def __init__(self, code, message, *, reason=None, forms=None):
        super().__init__(message)
        self.code = code
        self.diagnostic = None if reason is None else {
            "reason": reason, **({"pdf_forms": forms} if forms is not None else {})}


def _unsupported_pdf(message):
    raise PdfPolicyError("unsupported_document", message)


def _invalid_pdf(message):
    raise PdfPolicyError("invalid_document", message)


def _pdf_value(value):
    return value.get_object() if isinstance(value, IndirectObject) else value


def _pdf_access_failure(error):
    if isinstance(error, FileNotDecryptedError):
        raise PdfPolicyError("password_required", "This PDF requires a nonempty password. Password entry is not supported.") from error
    if isinstance(error, DependencyError):
        raise PdfPolicyError("encryption_backend_missing", "The PDF encryption backend is unavailable in this runtime.") from error
    if isinstance(error, NotImplementedError):
        raise PdfPolicyError("unsupported_encryption", "The PDF security handler or encryption algorithm is not supported.") from error
    raise PdfPolicyError("invalid_document", "The PDF encryption data or document structure is invalid.") from error


def _validate_pdf_encryption(reader):
    """Do not misreport a malformed Standard dictionary as a secret password.

    pypdf can return NOT_DECRYPTED for truncated credentials or invalid key
    lengths. Check their public PDF representation, without managing keys or
    repeating authentication. Other security handlers remain pypdf's decision.
    """
    if not reader.is_encrypted:
        return
    encryption = reader.trailer.get("/Encrypt").get_object()
    if not isinstance(encryption, DictionaryObject) or isinstance(encryption, StreamObject):
        _invalid_pdf("The PDF encryption dictionary is invalid.")
    try:
        revision = _opening_value(reader, encryption.get("/R"))
        length = _opening_value(reader, encryption.get("/Length", 40))
    except PdfPolicyError as error:
        raise PdfPolicyError("invalid_document", "The PDF encryption parameters contain an invalid local reference.") from error
    if not isinstance(revision, int) or isinstance(revision, bool) or revision < 2 \
            or not isinstance(length, int) or isinstance(length, bool) or length < 40 or length > 256 or length % 8:
        _invalid_pdf("The PDF encryption revision or key length is invalid.")
    if revision > 6:
        raise PdfPolicyError("unsupported_encryption", "The PDF encryption revision is not supported.")
    if revision == 2 and length != 40 or revision in (3, 4) and length > 128 or revision in (5, 6) and length != 256:
        _invalid_pdf("The PDF encryption key length disagrees with its revision.")
    required = {"/O": 32, "/U": 32} if revision < 5 else {"/O": 48, "/U": 48, "/OE": 32, "/UE": 32, "/Perms": 16}
    for field, size in required.items():
        value = encryption.get(field)
        if not isinstance(value, (ByteStringObject, TextStringObject)) or len(value.original_bytes) != size:
            _invalid_pdf("The PDF encryption credentials have an invalid representation or length.")


def _accessible_pdf_catalog(reader):
    """Use actual access, never is_encrypted, to recognize authentication.

    The supported PdfReader constructor already tries the empty password once.
    Reading its catalog preserves that result, including on captured readers;
    another decrypt call could reset an already-authenticated reader's key.
    """
    try:
        _validate_pdf_encryption(reader)
        return reader.root_object
    except (FileNotDecryptedError, DependencyError, NotImplementedError, PdfReadError) as error:
        _pdf_access_failure(error)


def _opening_value(reader, value):
    """Resolve only local references with a finite chain and no cycles."""
    visited = set()
    while isinstance(value, IndirectObject):
        identity = (value.idnum, value.generation)
        if value.pdf is not reader or identity in visited or len(visited) >= MAX_PDF_FEATURE_DEPTH:
            _unsupported_pdf("The opening destination contains an invalid or excessive reference chain.")
        visited.add(identity)
        try:
            value = value.get_object()
        except (AttributeError, RecursionError, PdfReadError) as error:
            raise PdfPolicyError("unsupported_document", "The opening destination contains an unreadable local reference.") from error
    return value


def _optional_pdf_value(reader, value):
    try:
        return _opening_value(reader, value)
    except (ValueError, TypeError, IndexError, KeyError, OSError, NotImplementedError, RecursionError):
        return None


def _opening_named_destination(reader, catalog, target):
    """Look up a local name without invoking pypdf's recursive name-tree API."""
    names = _opening_value(reader, catalog.get("/Names"))
    modern = names.get("/Dests") if isinstance(names, DictionaryObject) else None
    legacy = catalog.get("/Dests")
    if legacy is not None and modern is not None:
        _unsupported_pdf("The opening destination has ambiguous local name definitions.")
    root = legacy if legacy is not None else modern
    pending, visited, destinations, scheduled = [(root, 0)], set(), {}, 1
    while pending:
        value, depth = pending.pop()
        node = _opening_value(reader, value)
        if not isinstance(node, DictionaryObject) or isinstance(node, StreamObject) or id(node) in visited:
            _unsupported_pdf("The opening destination has a malformed or cyclic local name tree.")
        visited.add(id(node))
        if depth > MAX_PDF_PAGE_TREE_DEPTH or len(visited) > MAX_OPENING_DESTINATION_NODES:
            _unsupported_pdf("The opening destination name tree exceeds the inspection limit.")
        if legacy is not None:
            pairs = node.items()
        else:
            if set(node) - {"/Kids", "/Names", "/Limits"} or ("/Kids" in node) == ("/Names" in node):
                _unsupported_pdf("The opening destination has ambiguous local name-tree children.")
            if "/Limits" in node:
                limits = _opening_value(reader, node.get("/Limits"))
                if not isinstance(limits, ArrayObject) or len(limits) != 2:
                    _unsupported_pdf("The opening destination has invalid local name-tree limits.")
                limits = [_opening_value(reader, item) for item in limits]
                if not all(isinstance(item, (TextStringObject, ByteStringObject)) for item in limits) \
                        or limits[0].original_bytes > limits[1].original_bytes:
                    _unsupported_pdf("The opening destination has invalid local name-tree limits.")
            if "/Kids" in node:
                kids = _opening_value(reader, node.get("/Kids"))
                if not isinstance(kids, ArrayObject) or len(kids) > MAX_OPENING_DESTINATION_NODES:
                    _unsupported_pdf("The opening destination has an invalid local name-tree child list.")
                scheduled += len(kids)
                if scheduled > MAX_OPENING_DESTINATION_NODES:
                    _unsupported_pdf("The opening destination name tree exceeds the inspection limit.")
                pending.extend((kid, depth + 1) for kid in kids)
                continue
            entries = _opening_value(reader, node.get("/Names"))
            if not isinstance(entries, ArrayObject) or len(entries) % 2 or len(entries) > MAX_OPENING_DESTINATION_NODES * 2:
                _unsupported_pdf("The opening destination has an invalid local name list.")
            pairs = zip(entries[::2], entries[1::2])
        if len(node) > MAX_OPENING_DESTINATION_NODES:
            _unsupported_pdf("The opening destination has too many local names.")
        for key, destination in pairs:
            key = _opening_value(reader, key)
            key_types = (NameObject,) if legacy is not None else (TextStringObject, ByteStringObject)
            if not isinstance(key, key_types) or key in destinations:
                _unsupported_pdf("The opening destination has invalid or duplicate local names.")
            destinations[key] = destination
            if len(destinations) > MAX_OPENING_DESTINATION_NODES:
                _unsupported_pdf("The opening destination has too many local names.")
    if target not in destinations:
        _unsupported_pdf("The opening destination does not resolve to a local name.")
    value = _opening_value(reader, destinations[target])
    if isinstance(value, DictionaryObject) and not isinstance(value, StreamObject) and set(value) == {"/D"}:
        value = _opening_value(reader, value.get("/D"))
    return value


def _validate_opening_destination(reader, catalog, physical_page_ids):
    value = _opening_value(reader, catalog.get("/OpenAction"))
    if isinstance(value, DictionaryObject):
        action = _opening_value(reader, value.get("/S"))
        kind = _opening_value(reader, value.get("/Type"))
        if isinstance(value, StreamObject) or set(value) - {"/Type", "/S", "/D"} or not {"/S", "/D"} <= set(value) \
                or not isinstance(action, NameObject) or action != "/GoTo" \
                or "/Type" in value and (not isinstance(kind, NameObject) or kind != "/Action"):
            _unsupported_pdf("Only a local GoTo opening destination without additional actions is supported.")
        value = _opening_value(reader, value.get("/D"))
        if isinstance(value, (NameObject, TextStringObject, ByteStringObject)):
            value = _opening_named_destination(reader, catalog, value)
    if not isinstance(value, ArrayObject) or len(value) < 2:
        _unsupported_pdf("The opening destination must be a valid local destination array.")
    reference, fit = value[:2]
    fit = _opening_value(reader, fit)
    if not isinstance(reference, IndirectObject) or reference.pdf is not reader \
            or id(_opening_value(reader, reference)) not in physical_page_ids \
            or not isinstance(fit, NameObject) or fit not in OPENING_DESTINATION_FITS \
            or len(value) != OPENING_DESTINATION_FITS[fit] + 2:
        _unsupported_pdf("The opening destination has an invalid local page or view type.")
    arguments = []
    for raw in value[2:]:
        argument = _opening_value(reader, raw)
        if isinstance(argument, NullObject) and fit != "/FitR":
            arguments.append(None)
            continue
        try:
            finite = isinstance(argument, (int, float)) and not isinstance(argument, bool) and math.isfinite(argument)
        except OverflowError:
            finite = False
        if not finite:
            _unsupported_pdf("The opening destination has invalid view parameters.")
        arguments.append(argument)
    if fit == "/XYZ" and arguments[2] is not None and arguments[2] < 0 \
            or fit == "/FitR" and (arguments[2] <= arguments[0] or arguments[3] <= arguments[1]):
        _unsupported_pdf("The opening destination has an invalid zoom or view rectangle.")


def _form_error(message):
    raise PdfPolicyError("invalid_document", message, reason="malformed_form")


def _form_value(reader, value):
    """Resolve PDF form references without using pypdf's field-tree API."""
    seen = set()
    while isinstance(value, IndirectObject):
        identity = (value.idnum, value.generation)
        if value.pdf is not reader or identity in seen or len(seen) >= MAX_PDF_FEATURE_DEPTH:
            _form_error("A PDF form contains an invalid, cyclic or excessive reference chain.")
        seen.add(identity)
        try:
            value = value.get_object()
        except (PdfReadError, ValueError, TypeError, KeyError, OSError, RecursionError) as error:
            raise PdfPolicyError("invalid_document", "A PDF form contains an unreadable reference.",
                                 reason="malformed_form") from error
        if value is None:
            _form_error("A PDF form references an object that is missing.")
    return value


def _plain_dictionary(value):
    return isinstance(value, DictionaryObject) and not isinstance(value, StreamObject)


def _classify_pdf_forms(reader, catalog, page_widgets):
    """Classify only structural metadata; never read or log current field values."""
    form = _form_value(reader, catalog.get("/AcroForm"))
    metadata = form is not None and not isinstance(form, NullObject)
    roots, xfa, action_edges, form_refs = (), "absent", [], []
    if metadata:
        if not _plain_dictionary(form) or form is catalog:
            _form_error("The PDF AcroForm container is not a dictionary.")
        fields = _form_value(reader, form.get("/Fields"))
        resources = _form_value(reader, form.get("/DR"))
        appearance = _form_value(reader, form.get("/DA"))
        need_appearances = _form_value(reader, form.get("/NeedAppearances"))
        if fields is not None and not isinstance(fields, (ArrayObject, NullObject)):
            _form_error("The PDF AcroForm field list is not an array.")
        if resources is not None and not isinstance(resources, NullObject) and not _plain_dictionary(resources):
            _form_error("The PDF AcroForm default resources are not a dictionary.")
        if appearance is not None and not isinstance(appearance, (ByteStringObject, TextStringObject, NullObject)):
            _form_error("The PDF AcroForm default appearance is not a string.")
        if need_appearances is not None and not isinstance(need_appearances, (BooleanObject, NullObject)):
            _form_error("The PDF AcroForm appearance flag is not a Boolean.")
        roots = fields if isinstance(fields, ArrayObject) else ()
        form_refs.extend((catalog.get("/AcroForm"), form.get("/Fields")))
        if "/XFA" in form:
            payload = _form_value(reader, form.get("/XFA"))
            if not isinstance(payload, (ArrayObject, NullObject, StreamObject)):
                _form_error("The PDF XFA entry is not an array, stream or null marker.")
            # The pinned reader bounds decompression. No XML is interpreted, and
            # nonempty payloads cannot establish complete static pages.
            try:
                nonempty = ("/F" in payload or bool(payload.get_data())) if isinstance(payload, StreamObject) else (
                    isinstance(payload, ArrayObject) and bool(payload))
            except (PdfReadError, ValueError, NotImplementedError, RecursionError) as error:
                raise PdfPolicyError("unsupported_document", "This PDF requires verified XFA rendering.",
                                     reason="xfa_rendering_required") from error
            xfa = "rendering_required" if nonempty else "empty"
        if "/A" in form:
            action_edges.append(("action", form.get("/A")))
        if "/AA" in form:
            action_edges.append(("events", form.get("/AA")))

    # Walk only downward Kids edges. Parent backlinks are validated but don't
    # count as a cycle. The same widget can be named here and in a page array.
    if len(roots) + len(page_widgets) > MAX_PDF_FORM_NODES:
        _form_error("The PDF form exceeds the supported field and widget limit.")
    pending = [(value, None, 0, False) for value in reversed(roots)]
    active, owners, widgets, fields, signatures = set(), {}, set(), set(), set()
    scheduled = len(pending)
    while pending:
        raw, owner, depth, finish = pending.pop()
        node = _form_value(reader, raw)
        if not _plain_dictionary(node) or node is form or node is catalog:
            _form_error("A PDF form field or widget is not a dictionary.")
        identity = id(node)
        if finish:
            active.remove(identity)
            continue
        form_refs.append(raw)
        if identity in active or depth > MAX_PDF_FEATURE_DEPTH:
            _form_error("The PDF form field tree contains a cycle or exceeds the supported depth.")
        if identity in owners:
            if owners[identity] is not owner:
                _form_error("A PDF form field belongs to more than one field parent.")
            continue
        owners[identity] = owner
        active.add(identity)
        form_refs.extend(node.get(key) for key in ("/V", "/DV", "/RV", "/Opt") if key in node)
        parent = _form_value(reader, node.get("/Parent"))
        if parent is not None and not isinstance(parent, NullObject) and parent is not owner:
            _form_error("A PDF form field has an inconsistent parent reference.")
        subtype = _form_value(reader, node.get("/Subtype"))
        object_type = _form_value(reader, node.get("/Type"))
        if object_type in ("/Page", "/Pages", "/Catalog", "/Action") or object_type == "/Annot" and subtype != "/Widget":
            _form_error("The PDF form field list references an unrelated PDF object.")
        if subtype == "/Widget":
            widgets.add(identity)
            # A pure widget may inherit its field from a separate parent.
            if any(key in node for key in ("/FT", "/T", "/V", "/Kids")) or owner is None:
                fields.add(identity)
        else:
            fields.add(identity)
        if _form_value(reader, node.get("/FT")) == "/Sig":
            signatures.add(identity)
            if "/V" in node:
                action_edges.append(("signature", node.get("/V")))
        if "/A" in node:
            action_edges.append(("action", node.get("/A")))
        if "/AA" in node:
            action_edges.append(("events", node.get("/AA")))
        kids = _form_value(reader, node.get("/Kids"))
        if kids is not None and not isinstance(kids, (ArrayObject, NullObject)):
            _form_error("A PDF form field's children are not an array.")
        pending.append((node, owner, depth, True))
        if isinstance(kids, ArrayObject):
            scheduled += len(kids)
            if scheduled > MAX_PDF_FORM_NODES:
                _form_error("The PDF form exceeds the supported field and widget limit.")
            pending.extend((child, node, depth + 1, False) for child in reversed(kids))

    orphan_count = 0
    for widget in page_widgets:
        identity = id(widget)
        if identity in widgets:
            continue
        orphan_count += 1
        # Follow only orphan parent links here. Don't infer field values or
        # silently treat detached widgets as harmless AcroForm metadata.
        chain, node = set(), widget
        while node is not None and not isinstance(node, NullObject):
            if (not _plain_dictionary(node) or node is form or node is catalog
                    or id(node) in chain or len(chain) >= MAX_PDF_FEATURE_DEPTH):
                _form_error("A PDF widget has a malformed or cyclic field parent chain.")
            chain.add(id(node))
            if id(node) != identity or any(key in node for key in ("/FT", "/T", "/V")):
                fields.add(id(node))
            if _form_value(reader, node.get("/FT")) == "/Sig":
                signatures.add(id(node))
                if "/V" in node:
                    action_edges.append(("signature", node.get("/V")))
            if "/A" in node:
                action_edges.append(("action", node.get("/A")))
            if "/AA" in node:
                action_edges.append(("events", node.get("/AA")))
            if id(node) in owners:
                break
            node = _form_value(reader, node.get("/Parent"))
        scheduled += len(chain)
        if scheduled > MAX_PDF_FORM_NODES:
            _form_error("The PDF form exceeds the supported field and widget limit.")

    widget_count = len(widgets) + orphan_count
    if xfa == "rendering_required":
        kind = "xfa"
    elif widget_count:
        kind = "acroform"
    elif fields:
        kind = "data_only"
    elif metadata or "/AcroForm" in catalog:
        kind = "metadata_only"
    else:
        kind = "none"
    return {"kind": kind, "field_count": len(fields), "widget_count": widget_count,
            "orphan_widget_count": orphan_count, "signature_field_count": len(signatures), "xfa": xfa}, action_edges, form_refs


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


class ChapterGraph:
    """Inspect or copy only the page graph used by a fresh chapter catalog."""

    def __init__(self, reader, *, writer=None, excluded=None):
        self.reader = reader
        self.writer = writer
        self.excluded = {} if excluded is None else dict(excluded)
        self.warnings = set()
        self._projected = {}
        self.omitted_inspector = None

    def note_omitted(self, value, code):
        self.warnings.add(code)
        if isinstance(value, IndirectObject) and value.pdf is self.reader:
            self.excluded.setdefault((value.idnum, value.generation), code)

    def _resolve(self, value, *, omit=True):
        aliases = set()
        while isinstance(value, IndirectObject):
            key = (value.idnum, value.generation)
            if value.pdf is not self.reader or key in aliases or len(aliases) >= MAX_PDF_FEATURE_DEPTH:
                _invalid_pdf("Copied PDF content contains an invalid or excessive reference chain.")
            if omit and key in self.excluded:
                self.warnings.add(self.excluded[key])
                return None
            aliases.add(key)
            value = self.reader.get_object(value)
            if value is None:
                _invalid_pdf("Copied PDF content contains a missing local reference.")
        return value

    def optional_value(self, value, *, omit=True):
        try:
            return self._resolve(value, omit=omit)
        except (PdfReadError, ValueError, TypeError, AttributeError, KeyError, OSError, RecursionError):
            return None

    def _omission(self, node, kind, *, named_entries=False):
        action = self.optional_value(node.get("/S"), omit=False)
        subtype = self.optional_value(node.get("/Subtype"), omit=False)
        field_kind = self.optional_value(node.get("/FT"), omit=False)
        if isinstance(action, NameObject) and action == "/JavaScript" or not named_entries and "/JS" in node \
                and set(node) <= {"/Type", "/S", "/JS", "/Next"}:
            return "pdf_javascript_omitted"
        if kind == "/Action" or isinstance(action, NameObject) and action in PDF_ACTION_TYPES:
            return "pdf_action_omitted"
        if kind in ("/EmbeddedFile", "/Filespec") or subtype == "/FileAttachment" or not named_entries and "/EF" in node:
            return "pdf_attachment_omitted"
        if kind in ("/Sig", "/DocTimeStamp") or field_kind == "/Sig" or {"/ByteRange", "/Contents"} <= node.keys():
            return "pdf_signature_omitted"
        if isinstance(subtype, NameObject) and subtype in MEDIA_ANNOTATIONS:
            return "pdf_media_omitted"
        return None

    def _context(self, node, context, key):
        if context == "resources" and key in RESOURCE_NAME_MAPS:
            return "names"
        if context == "appearances":
            return "appearance_states"
        if context != "names" and key == "/Resources" and isinstance(node, DictionaryObject):
            return "resources"
        return "value"

    def _project(self, raw, depth, context, annotation):
        if depth > MAX_PDF_FEATURE_DEPTH:
            _invalid_pdf("The copied PDF object graph exceeds the supported depth.")
        node = self._resolve(raw)
        if not isinstance(node, (DictionaryObject, ArrayObject)):
            if self.writer is None or node is None:
                return node
            value = copy(node)
            if getattr(value, "indirect_reference", None) is not None:
                value.indirect_reference = None
            return value

        kind = self.optional_value(node.get("/Type"), omit=False) if isinstance(node, DictionaryObject) else None
        if kind in ("/OCG", "/OCMD"):
            # An annotation and a page must reference the same copied layer
            # object for catalog default visibility to apply to both.
            annotation = False
        identity = (id(node), context, annotation)
        if identity in self._projected:
            return self._projected[identity]
        if len(self._projected) >= MAX_PDF_FEATURE_NODES or len(node) > MAX_PDF_FEATURE_NODES:
            _invalid_pdf("The copied PDF object graph exceeds the supported inspection limit.")
        named_entries = context in {"names", "appearance_states"} and not isinstance(node, StreamObject)
        if isinstance(node, DictionaryObject):
            if kind in ("/Page", "/Pages", "/Catalog"):
                if annotation:
                    return None
                _unsupported_pdf("Copied page resources cannot reference document catalogs or physical page trees.")
            omission = self._omission(node, kind, named_entries=named_entries)
            if omission is not None:
                self.note_omitted(raw, omission)
                if omission == "pdf_javascript_omitted":
                    self.note_omitted(node.get("/JS"), omission)
                if self.omitted_inspector is not None:
                    if omission in {"pdf_action_omitted", "pdf_javascript_omitted"}:
                        self.omitted_inspector.edge(node, "/Next", "action_chain")
                    if omission == "pdf_attachment_omitted":
                        self.omitted_inspector.edge(node, "/EF", "file_streams")
                        self.omitted_inspector.edge(node, "/AF", "associated_files")
                    if omission == "pdf_signature_omitted":
                        self.omitted_inspector.edge(node, "/Contents", "signature_payload")
                return None
            if isinstance(node, StreamObject) and "/F" in node and not isinstance(self._resolve(node.get("/F")), NullObject):
                _unsupported_pdf("A PDF rendering stream refers to external file data and cannot be reproduced locally.")

        target = None
        if self.writer is not None:
            target = type(node)() if isinstance(node, StreamObject) else \
                DictionaryObject() if isinstance(node, DictionaryObject) else ArrayObject()
            if isinstance(node, StreamObject):
                target._data = node._data
            result = self.writer._add_object(target) if isinstance(raw, IndirectObject) or isinstance(node, StreamObject) else target
        else:
            result = node
        self._projected[identity] = result

        if isinstance(node, DictionaryObject):
            for key, child in node.items():
                if key == "/AF" and not named_entries:
                    self.note_omitted(child, "pdf_attachment_omitted")
                    if self.omitted_inspector is not None:
                        self.omitted_inspector.edge(node, key, "associated_files")
                    continue
                if context == "optional_content" and key not in {"/OCGs", "/D", "/Configs"}:
                    continue
                if context == "appearances" and key not in {"/N", "/R", "/D"}:
                    continue
                value = self._project(child, depth + 1, self._context(node, context, key), annotation)
                if target is not None and value is not None:
                    target[key] = value
        else:
            for child in node:
                value = self._project(child, depth + 1, "value", annotation)
                if target is not None:
                    target.append(NullObject() if value is None else value)
        if isinstance(node, StreamObject) and self._resolve(node.get("/Subtype")) == "/Form":
            self._validate_graphics_resources(node, node.get("/Resources"))
        return result

    def project(self, value, *, annotation=False, context="value"):
        return self._project(value, 0, context, annotation)

    def _validate_graphics_resources(self, contents, resource_reference):
        resources = self._resolve(resource_reference)
        if not isinstance(resources, DictionaryObject) or not isinstance(contents, (ArrayObject, StreamObject)):
            return
        removed = set()
        for category in RESOURCE_NAME_MAPS:
            entries = self._resolve(resources.get(category))
            if category in resources and entries is None:
                removed.add((category, None))
            if not isinstance(entries, DictionaryObject):
                continue
            for name, raw in entries.items():
                entry = self._resolve(raw)
                if entry is None or isinstance(entry, DictionaryObject) and self._omission(entry, self._resolve(entry.get("/Type"))):
                    removed.add((category, name))
        if not removed:
            return

        # Only parse drawing instructions when omission affects a named
        # graphics slot. Parsing never invokes a viewer or document action.
        xobjects = self._resolve(resources.get("/XObject"))
        pending, visited, instructions = [contents], set(), 0
        while pending:
            stream = pending.pop()
            if id(stream) in visited:
                continue
            visited.add(id(stream))
            operations = ContentStream(stream, self.reader).operations
            instructions += len(operations)
            if instructions > MAX_PDF_FEATURE_NODES:
                _invalid_pdf("PDF graphics dependencies exceed the supported inspection limit.")
            for operands, operator in operations:
                resource = RENDERING_RESOURCE_OPERATORS.get(operator)
                if resource is None:
                    if operator == b"INLINE IMAGE":
                        settings = operands.get("settings", {})
                        colorspace = settings.get("/CS", settings.get("/ColorSpace"))
                        if isinstance(colorspace, NameObject) and ({("/ColorSpace", colorspace), ("/ColorSpace", None)} & removed):
                            _invalid_pdf("An omitted PDF object is required to render an inline image.")
                    continue
                category, index = resource
                if not operands or index >= len(operands) or not isinstance(operands[index], NameObject):
                    continue
                name = operands[index]
                if (category, name) in removed or (category, None) in removed:
                    _invalid_pdf("An omitted PDF object is required to reproduce the page graphics.")
                if operator == b"Do" and isinstance(xobjects, DictionaryObject):
                    child = self._resolve(xobjects.get(name))
                    if isinstance(child, StreamObject) and self._resolve(child.get("/Subtype")) == "/Form" \
                            and "/Resources" not in child:
                        pending.append(child)

    def page_values(self, page):
        values = {}
        for key, value in page.items():
            if key not in CHAPTER_PAGE_FIELDS:
                # Retain inert extension scalars without following arbitrary
                # page-level graph edges or copying known omitted page keys.
                if key in OMITTED_PAGE_FIELDS or isinstance(value, (DictionaryObject, ArrayObject, IndirectObject)):
                    continue
            if key == "/Contents":
                contents = self._resolve(value)
                if not isinstance(contents, (StreamObject, ArrayObject, NullObject)):
                    _invalid_pdf("The PDF page contents are not an available stream or stream array.")
                if isinstance(contents, ArrayObject):
                    if len(contents) > MAX_PDF_FEATURE_NODES:
                        _invalid_pdf("The PDF page contains too many content streams.")
                    for item in contents:
                        stream = self._resolve(item)
                        if not isinstance(stream, StreamObject) or self._omission(stream, self._resolve(stream.get("/Type"))):
                            _invalid_pdf("The PDF page content array contains a missing or non-rendering stream.")
            projected = self.project(value, context="resources" if key == "/Resources" else "value")
            if projected is None:
                _invalid_pdf("Required PDF page data refers to omitted document functionality.")
            values[key] = projected
        self._validate_graphics_resources(self._resolve(page.get("/Contents")), page.get("/Resources"))
        return values

    def optional_content(self, catalog):
        values = {}
        for key, context in (("/OCProperties", "optional_content"), ("/OutputIntents", "value")):
            if key in catalog:
                value = self.project(catalog.get(key), context=context)
                if value is not None:
                    values[NameObject(key)] = value
        return values


class _OmittedPdfFeatures:
    """Identify optional branches and alias references without decoding payloads."""

    def __init__(self, reader, graph, protected):
        self.reader = reader
        self.graph = graph
        self.protected = protected
        self.pending = []
        self.truncated = False

    def value(self, raw):
        return _optional_pdf_value(self.reader, raw)

    def edge(self, node, key, role):
        if key in node:
            self.children(role, (node.get(key),))

    def exclude(self, raw, code):
        # Name trees may themselves be empty. A generic warning is added only
        # for malformed metadata or when the copying path encounters its alias.
        if isinstance(raw, IndirectObject) and raw.pdf is self.reader:
            self.graph.excluded.setdefault((raw.idnum, raw.generation), code)

    def children(self, role, values):
        for item in values:
            if len(self.pending) >= MAX_OMITTED_FEATURE_NODES:
                self.truncated = True
                break
            self.pending.append((role, item))

    def inspect(self):
        seen = set()
        payloads = {"script_payload": "pdf_javascript_omitted", "attachment_payload": "pdf_attachment_omitted",
                    "signature_payload": "pdf_signature_omitted"}
        self.pending.reverse()
        for _ in range(MAX_OMITTED_FEATURE_NODES):
            if not self.pending:
                break
            role, raw = self.pending.pop()
            if raw is None or isinstance(raw, NullObject):
                continue
            # /JS and /EF payloads are not needed for inspection. Record
            # their raw references so a copy via a different key is omitted.
            if role in payloads:
                self.graph.note_omitted(raw, payloads[role])
                continue
            node = self.value(raw)
            if id(node) in self.protected or isinstance(node, NullObject):
                continue
            identity = (role, id(node) if node is not None else id(raw))
            if identity in seen:
                if role in {"javascript_names", "attachment_names"}:
                    self.graph.warnings.add("pdf_name_tree_omitted")
                continue
            seen.add(identity)
            if role in {"javascript_names", "attachment_names"}:
                self._name_tree(role, raw, node)
            elif role in {"action", "local_action", "events", "action_chain", "named_script"}:
                self._action(role, raw, node)
            elif role in {"associated_files", "file_spec", "file_streams"}:
                self._attachment(role, raw, node)
            elif role in {"outline_root", "outline_node"}:
                if _plain_dictionary(node):
                    self.edge(node, "/First", "outline_node")
                    self.edge(node, "/AA", "events")
                    if role == "outline_node":
                        self.edge(node, "/Next", "outline_node")
                        self.edge(node, "/A", "local_action")
            else:  # Permissions, signature dictionaries and portfolio metadata.
                code = "pdf_portfolio_omitted" if role == "portfolio" else "pdf_signature_omitted"
                self.graph.warnings.add(code)
                if _plain_dictionary(node):
                    self.exclude(raw, code)
                    if role == "permissions":
                        self.children("signature", node.values())
                    elif role == "signature":
                        self.edge(node, "/Contents", "signature_payload")
        # A large optional branch is still omitted. Do not claim types or
        # counts for entries beyond the bounded metadata inspection.
        if self.pending or self.truncated:
            self.graph.warnings.add("pdf_optional_metadata_omitted")

    def _name_tree(self, role, raw, node):
        generic = "pdf_name_tree_omitted"
        if not _plain_dictionary(node):
            self.graph.warnings.add(generic)
            return
        self.exclude(raw, generic)
        if "/Kids" in node:
            kids = self.value(node.get("/Kids"))
            if isinstance(kids, ArrayObject):
                self.exclude(node.get("/Kids"), generic)
                self.children(role, kids)
            elif not isinstance(kids, NullObject):
                self.graph.warnings.add(generic)
        if "/Names" in node:
            pairs = self.value(node.get("/Names"))
            if isinstance(pairs, ArrayObject):
                self.exclude(node.get("/Names"), generic)
                if len(pairs) % 2:
                    self.graph.warnings.add(generic)
                self.children("named_script" if role == "javascript_names" else "file_spec",
                              islice(pairs, 1, len(pairs), 2))
            elif not isinstance(pairs, NullObject):
                self.graph.warnings.add(generic)

    def _action(self, role, raw, node):
        code = "pdf_action_omitted"
        if isinstance(node, ArrayObject) and role in {"action_chain", "events"}:
            self.exclude(raw, code)
            self.children("action", node)
            return
        if not _plain_dictionary(node):
            self.graph.warnings.add("pdf_name_tree_omitted" if role == "named_script" else code)
            return
        if role == "events":
            if node:
                self.graph.warnings.add(code)
            self.exclude(raw, code)
            self.children("action", node.values())
            return
        kind = self.value(node.get("/S"))
        if kind == "/JavaScript" or "/JS" in node:
            self.graph.warnings.add("pdf_javascript_omitted")
            self.exclude(raw, "pdf_javascript_omitted")
            self.edge(node, "/JS", "script_payload")
        elif role != "local_action" or kind != "/GoTo":
            self.graph.warnings.add(code)
            self.exclude(raw, code)
        self.edge(node, "/Next", "action_chain")

    def _attachment(self, role, raw, node):
        code = "pdf_attachment_omitted"
        if role == "associated_files" and isinstance(node, ArrayObject):
            if node:
                self.graph.warnings.add(code)
            self.exclude(raw, code)
            self.children("file_spec", node)
        elif _plain_dictionary(node):
            self.graph.warnings.add(code)
            self.exclude(raw, code)
            if role == "file_streams":
                self.children("attachment_payload", node.values())
            else:
                self.edge(node, "/EF", "file_streams")
                self.edge(node, "/AF", "associated_files")
        elif role != "file_streams":
            self.graph.warnings.add(code if role == "associated_files" or isinstance(node, (TextStringObject, ByteStringObject))
                                    else "pdf_name_tree_omitted")


def _preflight_pdf_features(reader):
    catalog = _accessible_pdf_catalog(reader)
    if not _plain_dictionary(catalog) or catalog.get("/Type") != "/Catalog":
        _invalid_pdf("The PDF catalog is not a valid document dictionary.")
    pages, page_tree_ids = _preflight_page_tree(catalog)
    graph = ChapterGraph(reader)
    omitted = _OmittedPdfFeatures(reader, graph, page_tree_ids | {id(catalog)})
    graph.omitted_inspector = omitted
    names = omitted.value(catalog.get("/Names"))
    if _plain_dictionary(names):
        omitted.edge(names, "/JavaScript", "javascript_names")
        omitted.edge(names, "/EmbeddedFiles", "attachment_names")
    omitted.edge(catalog, "/A", "action")
    omitted.edge(catalog, "/AA", "events")
    omitted.edge(catalog, "/AF", "associated_files")
    omitted.edge(catalog, "/EF", "file_streams")
    omitted.edge(catalog, "/JS", "script_payload")
    omitted.edge(catalog, "/Perms", "permissions")
    omitted.edge(catalog, "/Collection", "portfolio")
    omitted.edge(catalog, "/Outlines", "outline_root")
    opening_destination = False
    if "/OpenAction" in catalog and not isinstance(catalog.get("/OpenAction"), NullObject):
        try:
            _validate_opening_destination(reader, catalog, {id(page) for page in pages})
        except (PdfPolicyError, ValueError, TypeError, IndexError, KeyError, OSError, NotImplementedError, RecursionError):
            opening = catalog.get("/OpenAction")
            if not isinstance(omitted.value(opening), NullObject):
                graph.warnings.add("pdf_opening_action_omitted")
                omitted.edge(catalog, "/OpenAction", "action")
        else:
            opening_destination = True

    projected_annotations, page_widgets, widget_ids = [], [], set()
    examined_annotations = 0
    for page in pages:
        omitted.edge(page, "/A", "action")
        omitted.edge(page, "/AA", "events")
        omitted.edge(page, "/AF", "associated_files")
        omitted.edge(page, "/EF", "file_streams")
        omitted.edge(page, "/JS", "script_payload")
        annotations = graph._resolve(page.get("/Annots"), omit=False)
        if not isinstance(annotations, ArrayObject):
            continue
        examined_annotations += len(annotations)
        if examined_annotations > MAX_PDF_FEATURE_NODES:
            _invalid_pdf("PDF annotations exceed the supported inspection limit.")
        for reference in annotations:
            annotation = graph._resolve(reference, omit=False)
            if not _plain_dictionary(annotation):
                continue
            subtype = graph._resolve(annotation.get("/Subtype"), omit=False)
            if subtype == "/Widget":
                identity = id(annotation)
                if identity in widget_ids:
                    _form_error("A PDF widget is listed more than once in the physical page annotations.")
                widget_ids.add(identity)
                page_widgets.append(annotation)
                continue
            omitted.edge(annotation, "/A", "local_action" if subtype == "/Link" else "action")
            omitted.edge(annotation, "/AA", "events")
            omitted.edge(annotation, "/AF", "associated_files")
            omitted.edge(annotation, "/EF", "file_streams")
            if subtype == "/FileAttachment":
                graph.warnings.add("pdf_attachment_omitted")
                omitted.edge(annotation, "/FS", "file_spec")
            if isinstance(subtype, NameObject) and subtype in MEDIA_ANNOTATIONS:
                graph.warnings.add("pdf_media_omitted")
            if _pdf_value(annotation.get("/Type")) == "/Sig" or _pdf_value(annotation.get("/FT")) == "/Sig":
                graph.warnings.add("pdf_signature_omitted")
                continue
            if isinstance(subtype, NameObject) and subtype in STATIC_ANNOTATION_SUBTYPES | MEDIA_ANNOTATIONS | {"/Link", "/FileAttachment"}:
                projected_annotations.extend((key, value) for key, value in annotation.items() if key in STATIC_ANNOTATION_FIELDS)

    forms, action_edges, form_refs = _classify_pdf_forms(reader, catalog, page_widgets)
    if forms["kind"] == "xfa":
        raise PdfPolicyError("unsupported_document",
                             "This PDF requires verified XFA rendering before its pages can be split.",
                             reason="xfa_rendering_required", forms=forms)
    if forms["widget_count"]:
        raise PdfPolicyError("unsupported_document",
                             "The visible appearance of this PDF's form widgets must be preserved before splitting.",
                             reason="form_appearance_required", forms=forms)
    if forms["kind"] == "data_only":
        graph.warnings.add("pdf_form_data_omitted")
    if forms["signature_field_count"]:
        graph.warnings.add("pdf_signature_omitted")
    for reference in form_refs:
        omitted.exclude(reference, "pdf_form_data_omitted")
    for role, value in action_edges:
        omitted.children(role, (value,))
    omitted.inspect()

    # Inspect only graphs that can contribute to a chapter. Catalog entries
    # such as name trees, permissions and portfolios aren't page resources.
    alias_count = len(graph.excluded)
    graph.optional_content(catalog)
    for page in reader.pages:
        graph.page_values(page)
    for key, value in projected_annotations:
        graph.project(value, annotation=True, context="appearances" if key == "/AP" else "value")
    omitted.inspect()
    if len(graph.excluded) != alias_count:
        # Some payload identities become known only inside a rendering
        # resource. Recheck dependencies with the complete exclusion map.
        graph._projected.clear()
        graph.omitted_inspector = None
        graph.optional_content(catalog)
        for page in reader.pages:
            graph.page_values(page)
        for key, value in projected_annotations:
            graph.project(value, annotation=True, context="appearances" if key == "/AP" else "value")
    reader._pdf_forms = MappingProxyType(forms)
    reader._chapter_graph_excluded = MappingProxyType(dict(graph.excluded))
    reader._chapter_graph_warnings = frozenset(graph.warnings)
    reader._chapter_opening_destination = opening_destination


class PolicyPdfReader(PdfReader):
    """Pinned pypdf reader guard shared by direct and converted PDF inputs."""
    def __init__(self, stream, strict=False):
        snapshot = _bounded_pdf_snapshot(stream)
        try:
            super().__init__(snapshot, strict=strict, root_object_recovery_limit=4096)
        except (DependencyError, NotImplementedError) as error:
            _pdf_access_failure(error)
        except PdfReadError as error:
            # Preserve ordinary damaged-file handling; a parsed /Encrypt entry
            # specifically identifies a corrupt encryption setup.
            if "/Encrypt" in getattr(self, "trailer", {}):
                _pdf_access_failure(error)
            raise
        _preflight_pdf_features(self)

    def _build_outline_item(self, node):
        if "/A" not in node:
            return super()._build_outline_item(node)
        # pypdf assumes a bookmark action is a dictionary. Provide only a
        # resolved local GoTo to its destination builder, leaving the original
        # node available for the planner's malformed/external-action warning.
        record = DictionaryObject({key: value for key, value in node.items() if key in {"/Title", "/C", "/F", "/Count"}})
        action = _optional_pdf_value(self, node.get("/A"))
        kind = _optional_pdf_value(self, action.get("/S")) if _plain_dictionary(action) else None
        if isinstance(kind, NameObject) and kind == "/GoTo":
            record[NameObject("/A")] = DictionaryObject({
                key: kind if key == "/S" else value for key, value in action.items() if key in {"/S", "/D"}})
        item = super()._build_outline_item(record)
        if item is not None:
            item.node = node
            reference = getattr(node, "indirect_reference", None)
            if reference is not None:
                item.indirect_reference = reference
        return item


def validate_chapter_graph(reader):
    """Verify the reopened chapter has static edges and no unused objects."""
    if not hasattr(reader, "_chapter_graph_warnings"):
        _preflight_pdf_features(reader)
    catalog = reader.root_object
    if reader.is_encrypted or reader._chapter_graph_warnings or any(key in catalog for key in
            ("/AcroForm", "/OpenAction", "/AA", "/AF", "/Collection", "/Perms", "/Names")):
        _invalid_pdf("A staged chapter retained excluded PDF functionality.")
    for page in reader.pages:
        if any(key in page for key in ("/A", "/AA", "/AF", "/B")):
            _invalid_pdf("A staged chapter retained an excluded page relationship.")
        for raw in page.get("/Annots", ()):
            annotation = _pdf_value(raw)
            if not isinstance(annotation, DictionaryObject) or any(key in annotation for key in ("/A", "/AA", "/AF", "/EF", "/FS")):
                _invalid_pdf("A staged chapter retained an excluded annotation relationship.")
    pending = [(reader.trailer, 0)]
    reachable, seen = set(), set()
    while pending:
        raw, depth = pending.pop()
        if isinstance(raw, IndirectObject):
            reference = (raw.idnum, raw.generation)
            if raw.pdf is not reader:
                _invalid_pdf("A staged chapter contains a nonlocal object reference.")
            if reference in reachable:
                continue
            reachable.add(reference)
            pending.append((reader.get_object(raw), depth + 1))
            continue
        if raw is None:
            _invalid_pdf("A staged chapter contains an unresolved object reference.")
        if not isinstance(raw, (DictionaryObject, ArrayObject)) or id(raw) in seen:
            continue
        seen.add(id(raw))
        if depth > MAX_PDF_FEATURE_DEPTH * 2 or len(seen) > MAX_PDF_FEATURE_NODES + MAX_PDF_PAGE_TREE_NODES \
                or len(raw) > MAX_PDF_FEATURE_NODES:
            _invalid_pdf("A staged chapter exceeds the supported object-graph inspection limit.")
        children = raw.values() if isinstance(raw, DictionaryObject) else raw
        pending.extend((child, depth + 1) for child in children)
    serialized = {(identity, generation) for generation, entries in reader.xref.items() for identity in entries
                  if identity and not reader.xref_free_entry.get(generation, {}).get(identity)}
    serialized.update((identity, 0) for identity in reader.xref_objStm)
    if reachable != serialized:
        _invalid_pdf("A staged chapter contains missing or unreferenced serialized objects.")


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
        code = error.code if error.code in {"unsupported_document", "password_required", "unsupported_encryption",
                                           "encryption_backend_missing"} else "conversion_output_invalid"
        raise ConversionError(code, str(error), diagnostic=error.diagnostic) from error
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
            failure.diagnostic = {**(failure.diagnostic or {}), **diagnostic}
        raise failure
    return ConversionResult(result.pdf_bytes, result.original_source_identity, result.generated_pdf_identity,
        _freeze({**result.conversion, "original_source_identity": result.original_source_identity,
                 "generated_pdf_identity": result.generated_pdf_identity,
                 "workspace_cleanup": {"cleanup_complete": True, "retained_staging": None}}))
