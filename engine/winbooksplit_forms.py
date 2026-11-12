"""Select saved widget artwork or regenerate ordinary variable text locally."""

from collections import Counter
from math import isfinite
import re
import unicodedata

from pypdf._text_extraction import mult
from pypdf.errors import PdfReadError
from pypdf.generic import (
    ArrayObject, ByteStringObject, ContentStream, DecodedStreamObject,
    DictionaryObject, FloatObject, NameObject, NullObject, NumberObject, StreamObject,
    TextStringObject,
)
from pypdf.generic._font import Font


MAX_APPEARANCE_BYTES = 1024 * 1024
MAX_APPEARANCE_OPERATIONS = 20000
MAX_VALUE_CHARACTERS = 16384
MAX_OPTIONS = 4096
MAX_RADIO_OPTIONS = 256
_MULTILINE, _PASSWORD, _RADIO, _PUSHBUTTON = 1 << 12, 1 << 13, 1 << 15, 1 << 16
_COMBO, _EDIT, _MULTISELECT, _COMB, _UNISON = 1 << 17, 1 << 18, 1 << 21, 1 << 24, 1 << 25
_PAINT = {b"Do", b"sh", b"INLINE IMAGE", b"S", b"s", b"f", b"F", b"f*", b"B", b"B*", b"b", b"b*"}
_COLOR_SIZE = {b"g": 1, b"rg": 3, b"k": 4}
_CORE_ALIASES = {"/Helv": "/Helvetica", "/Cour": "/Courier", "/TiRo": "/Times-Roman", "/ZaDb": "/ZapfDingbats"}
_UNREADABLE = (PdfReadError, TypeError, ValueError, KeyError, AttributeError, IndexError,
               OverflowError, NotImplementedError, OSError, RecursionError)
_IDENTITY = [1, 0, 0, 1, 0, 0]


class AppearanceError(ValueError):
    """A fixed, public-safe explanation for an unavailable appearance."""


def _dictionary(value):
    return isinstance(value, DictionaryObject) and not isinstance(value, StreamObject)


def _present(value):
    return value is not None and not isinstance(value, NullObject)


def _integer(value, default, message):
    if not _present(value):
        return default
    if not isinstance(value, int) or isinstance(value, bool):
        raise AppearanceError(message)
    return int(value)


def _finite(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool) and isfinite(value)


def _numeric_array(value, count, resolve, message):
    if not isinstance(value, ArrayObject) or len(value) != count:
        raise AppearanceError(message)
    values = [resolve(entry) for entry in value]
    if not all(_finite(entry) for entry in values):
        raise AppearanceError(message)
    return values


def _text(value):
    if isinstance(value, TextStringObject) and len(value) <= MAX_VALUE_CHARACTERS:
        return str(value)
    raise AppearanceError("A saved PDF form value is not a supported bounded text string.")


def _operations(reader, stream):
    if "/F" in stream or len(stream.get_data()) > MAX_APPEARANCE_BYTES:
        raise AppearanceError("A PDF form appearance uses external or excessive rendering data.")
    operations = ContentStream(stream, reader, "bytes").operations
    if len(operations) > MAX_APPEARANCE_OPERATIONS:
        raise AppearanceError("A PDF form appearance exceeds the supported rendering limit.")
    return operations


def _font_reference(resources, name, resolve):
    fonts = resolve(resources.get("/Font")) if _dictionary(resources) else None
    return fonts.get(name) if _dictionary(fonts) else None


def _font_metrics(raw, resolve):
    source = resolve(raw)
    if not _dictionary(source) or resolve(source.get("/Subtype")) not in (
            "/Type1", "/MMType1", "/TrueType", "/Type3", "/Type0"):
        raise AppearanceError("A PDF form appearance needs a supported local font resource.")
    # pypdf fills resolved entries into /DescendantFonts while reading metrics.
    # Keep that scratch array private and preserve source dictionaries.
    font_dict = DictionaryObject(source)
    descendants = resolve(source.get("/DescendantFonts"))
    if isinstance(descendants, ArrayObject):
        font_dict[NameObject("/DescendantFonts")] = ArrayObject(descendants)
    return Font.from_font_resource(font_dict)


def _decode(data, font):
    if isinstance(font.encoding, dict):
        return "".join(font.character_map.get(chr(code), font.character_map.get(
            font.encoding.get(code, ""), font.encoding.get(code, "\ufffd"))) for code in data)
    return "".join(font.character_map.get(glyph, glyph) for glyph in data.decode(font.encoding, "strict"))


def _normal(widget, resolve):
    appearances = resolve(widget.get("/AP"))
    return appearances.get("/N") if _dictionary(appearances) else None


def _normal_stream(normal, state, resolve):
    source = resolve(normal)
    if _dictionary(source):
        if not isinstance(state, NameObject):
            return None, None
        normal = source.get(state)
        source = resolve(normal)
    return (normal, source) if isinstance(source, StreamObject) else (None, None)


def _appearance_text(reader, source, page_resources, resolve, list_art=None):
    """Read visible text operators with the appearance's own font encoding."""
    resources = resolve(source.get("/Resources")) if "/Resources" in source else page_resources
    box = _numeric_array(resolve(source.get("/BBox")), 4, resolve, "A PDF form appearance has an invalid bounding box.")
    bounds = (min(box[0], box[2]), min(box[1], box[3]), max(box[0], box[2]), max(box[1], box[3]))
    fonts, visible, variable, scope, stack = {}, [], [], [], []
    font, size, mode, active = None, 0, 0, False
    graphics, line, current, leading = _IDENTITY, _IDENTITY, _IDENTITY, 0
    fill_color, rectangles = (b"g", 0), []
    for operands, op in _operations(reader, source):
        if op == b"q":
            stack.append((font, size, mode, graphics, leading, fill_color))
        elif op == b"Q":
            if not stack:
                return None
            font, size, mode, graphics, leading, fill_color = stack.pop()
        elif op == b"cm" and len(operands) == 6 and all(_finite(item) for item in operands):
            graphics = mult(operands, graphics)
        elif (list_art is not None and op in _COLOR_SIZE and len(operands) == _COLOR_SIZE[op]
              and all(_finite(item) for item in operands)):
            fill_color = (op, *operands)
        elif list_art is not None and op == b"re" and len(operands) == 4 and all(_finite(item) for item in operands):
            a, b, c, d, e, f = graphics
            x, y, width, height = operands
            if b or c:
                return None
            xs, ys = (a * x + e, a * (x + width) + e), (d * y + f, d * (y + height) + f)
            rectangles.append((min(xs), min(ys), max(xs), max(ys)))
        elif list_art is not None and op in (_PAINT | {b"n"}) and op != b"Do":
            if op in (b"f", b"F", b"f*"):
                list_art["fills"].extend((rectangle, fill_color) for rectangle in rectangles)
            rectangles = []
        elif op in (b"BMC", b"BDC"):
            scope.append(bool(operands and operands[0] == "/Tx"))
        elif op == b"EMC":
            if not scope:
                return None
            scope.pop()
        elif op == b"BT":
            active = True
            line = current = _IDENTITY
        elif op == b"ET":
            active = False
        elif op == b"Tm" and len(operands) == 6 and all(_finite(item) for item in operands):
            line = current = operands
        elif op in (b"Td", b"TD") and len(operands) == 2 and all(_finite(item) for item in operands):
            x, y = operands
            if op == b"TD":
                leading = -y
            line = current = mult([1, 0, 0, 1, x, y], line)
        elif op == b"TL" and len(operands) == 1 and _finite(operands[0]):
            leading = operands[0]
        elif op == b"T*":
            line = current = mult([1, 0, 0, 1, 0, -leading], line)
        elif op == b"Tr" and len(operands) == 1:
            mode = operands[0]
        elif op in (b"Tc", b"Tw", b"Tz", b"Ts") and (
                len(operands) != 1 or operands[0] != (100 if op == b"Tz" else 0)):
            # A changed text state needs more than the simple run metrics
            # below to establish that the complete saved value is visible.
            return None
        elif op == b"Tf" and len(operands) == 2:
            key, size = operands
            if not isinstance(key, NameObject):
                return None
            if key not in fonts:
                reference = _font_reference(resources, key, resolve)
                fonts[key] = _font_metrics(reference, resolve) if _present(reference) else None
            font = fonts[key]
        elif active and op in (b"Tj", b"TJ", b"'", b'"'):
            if not operands or font is None or not _finite(size) or size == 0:
                return None
            if op in (b"'", b'"'):
                line = current = mult([1, 0, 0, 1, 0, -leading], line)
            values = operands[-1] if op == b"TJ" else operands[-1:]
            strings = [item.original_bytes for item in values if isinstance(item, (TextStringObject, ByteStringObject))]
            text = "".join(_decode(item, font) for item in strings)
            width = _run_width(font, strings, text) * size
            if op == b"TJ":
                width -= sum(item for item in values if _finite(item)) * size / 1000
            position = mult(current, graphics)
            if text.strip() and (mode not in (0, 1, 2, 4, 5, 6)
                                or not _intersects_appearance(bounds, position, size, width,
                                                              check_horizontal=isinstance(font.encoding, dict))):
                return None
            if list_art is not None and text:
                list_art["rows"].append((text, position[4], position[5]))
            current = mult([1, 0, 0, 1, width, 0], current)
            visible.append(text)
            if any(scope):
                variable.append(text)
        elif op == b"Do":
            xobjects = resolve(resources.get("/XObject")) if _dictionary(resources) else None
            artwork = resolve(xobjects.get(operands[0])) if _dictionary(xobjects) and operands else None
            if not isinstance(artwork, StreamObject) or resolve(artwork.get("/Subtype")) != "/Image":
                # Text in nested artwork cannot establish a current value.
                return None
    return (variable or visible) if not stack and not scope and not active else None


def _matches_choice_artwork(source, artwork, options, selected, inherited, resolve):
    rows = artwork["rows"]
    top = _integer(inherited.get("/TI"), 0, "A PDF choice field has an invalid first visible item.")
    if not 0 <= top < len(options) or not 0 < len(rows) <= 512 or top + len(rows) > len(options):
        return False
    if any(not _matches_value([row[0]], label) for row, label in zip(rows, options[top:])):
        return False
    if any(not top <= index < top + len(rows) for index in selected):
        return False
    x0, y0, x1, y1 = _numeric_array(resolve(source.get("/BBox")), 4, resolve,
                                      "A PDF form appearance has an invalid bounding box.")
    width, height = abs(x1 - x0), abs(y1 - y0)
    fills = [(box, color) for box, color in artwork["fills"] if box[2] - box[0] >= width / 2]
    background = {color for box, color in fills if box[3] - box[1] >= height * 0.8}
    marks = [box for box, color in fills if 0 < box[3] - box[1] < height / 2 and color not in background]
    if len(marks) > len(rows) * 2:
        return False
    observed = {top + index for index, (_, x, y) in enumerate(rows)
                if any(left <= x <= right and bottom <= y <= upper for left, bottom, right, upper in marks)}
    return observed == set(selected)


def _run_width(font, strings, text):
    if isinstance(font.encoding, dict):
        return sum(font.get_text_width(chr(glyph)) for string in strings for glyph in string) / 1000
    return len(text) / 2


def _intersects_appearance(bounds, matrix, size, width, *, check_horizontal):
    a, b, c, d, e, f = matrix
    corners = [(a * x + c * y + e, b * x + d * y + f)
               for x in (0, width or size / 2) for y in (-size / 4, size)]
    xs, ys = zip(*corners)
    if not (all(map(isfinite, (*xs, *ys))) and max(xs) > bounds[0] and min(xs) < bounds[2]
            and max(ys) > bounds[1] and min(ys) < bounds[3]):
        return False
    return not check_horizontal or c != 0 or b != 0 or min(xs) >= bounds[0] - 1 and max(xs) <= bounds[2] + 1


def _matches_value(parts, expected):
    if parts is None:
        return False
    def normalized(text):
        return " ".join(unicodedata.normalize("NFC", text).split())
    target = normalized(expected)
    return any(normalized(separator.join(parts)) == target for separator in ("", " "))


def _button(reader, widget, inherited, flags, normal, resolve):
    state = resolve(widget.get("/AS"))
    value = inherited.get("/V")
    source = resolve(normal)
    if flags & _PUSHBUTTON:
        raw, stream = _normal_stream(normal, state, resolve)
        if stream is None:
            return None
    elif _dictionary(source):
        if _present(value) and not isinstance(value, NameObject):
            raise AppearanceError("A saved PDF button selection is not a valid state name.")
        if _present(state) and (not isinstance(state, NameObject) or state not in source):
            raise AppearanceError("A PDF button references an unknown appearance state.")
        selected = NameObject("/Off")
        if _present(value) and value != "/Off":
            if flags & _RADIO:
                siblings = inherited.get("/Kids")
                if isinstance(siblings, ArrayObject) and len(siblings) > MAX_RADIO_OPTIONS:
                    raise AppearanceError("A PDF radio group exceeds the supported selection limit.")
                candidates = {id(widget): widget} if value in source else {}
                for sibling_raw in siblings if isinstance(siblings, ArrayObject) else ():
                    sibling = resolve(sibling_raw)
                    sibling_normal = resolve(_normal(sibling, resolve)) if _dictionary(sibling) else None
                    if _dictionary(sibling_normal) and value in sibling_normal:
                        candidates[id(sibling)] = sibling
                if not candidates:
                    raise AppearanceError("A PDF radio group lacks artwork for its saved selected state.")
                if len(candidates) > 1 and not flags & _UNISON:
                    on = [sibling for sibling in candidates.values() if resolve(sibling.get("/AS")) == value]
                    if len(on) != 1:
                        raise AppearanceError("The saved PDF radio group selection is ambiguous.")
                    selected = value if on[0] is widget else selected
                elif id(widget) in candidates:
                    selected = value
            elif value in source:
                selected = value
            else:
                raise AppearanceError("A PDF checkbox lacks artwork for its saved selected state.")
        elif not _present(value) and _present(state):
            selected = state
        raw, stream = _normal_stream(normal, selected, resolve)
        if stream is None:
            raise AppearanceError("A PDF button lacks artwork for its saved state.")
        if selected == "/Off":
            _operations(reader, stream)
            return raw
        off = resolve(source.get("/Off"))
        if isinstance(off, StreamObject) and off.get("/Resources") == stream.get("/Resources"):
            if "/F" not in off and len(off.get_data()) <= MAX_APPEARANCE_BYTES and off.get_data() == stream.get_data():
                raise AppearanceError("A PDF button's saved selected and unselected appearances are identical.")
    else:
        if _present(value) and value != "/Off" or _present(state) and state != "/Off":
            raise AppearanceError("A PDF button has no verifiable artwork for its saved state.")
        raw, stream = _normal_stream(normal, None, resolve)
        if stream is None:
            raise AppearanceError("A visible PDF button lacks its static appearance.")
    if not _has_artwork(reader, stream):
        raise AppearanceError("A visible PDF button appearance has no reproducible artwork.")
    return raw


def _has_artwork(reader, stream):
    mode = 0
    for operands, op in _operations(reader, stream):
        if op == b"Tr" and len(operands) == 1:
            mode = operands[0]
        if op in _PAINT or op in (b"Tj", b"TJ", b"'", b'"') and operands and mode not in (3, 7):
            return True
    return False


def _choice(inherited, flags, resolve):
    value = inherited.get("/V")
    if not _present(value):
        saved = []
    elif isinstance(value, ArrayObject) and flags & _MULTISELECT and not flags & _COMBO and len(value) <= MAX_OPTIONS:
        saved = [_text(resolve(entry)) for entry in value]
    else:
        saved = [_text(value)]
    raw_options = inherited.get("/Opt")
    if not _present(raw_options):
        raw_options = ()
    elif not isinstance(raw_options, ArrayObject) or len(raw_options) > MAX_OPTIONS:
        raise AppearanceError("A PDF choice field has an unsupported option list.")
    options = []
    characters = 0
    for raw in raw_options:
        option = resolve(raw)
        if isinstance(option, ArrayObject) and len(option) == 2:
            options.append(tuple(_text(resolve(item)) for item in option))
        else:
            display = _text(option)
            options.append((display, display))
        characters += sum(map(len, options[-1]))
        if characters > 4 * MAX_VALUE_CHARACTERS:
            raise AppearanceError("A PDF choice field exceeds the supported display-label limit.")
    indices = inherited.get("/I")
    if _present(indices):
        if not isinstance(indices, ArrayObject) or len(indices) > MAX_OPTIONS:
            raise AppearanceError("A PDF choice field has invalid saved selection indices.")
        chosen = [_integer(resolve(index), -1, "A PDF choice field has an invalid selection index.") for index in indices]
        if any(not 0 <= index < len(options) for index in chosen) or sorted(set(chosen)) != chosen:
            raise AppearanceError("A PDF choice field has invalid saved selection indices.")
        if _present(value) and Counter(options[index][0] for index in chosen) != Counter(saved):
            raise AppearanceError("A PDF choice field has conflicting saved values and selection indices.")
        saved = [options[index][0] for index in chosen]
    else:
        chosen = []
        for item in saved:
            matching = [index for index, option in enumerate(options) if option[0] == item]
            if len(matching) != 1:
                if flags & _COMBO and (flags & _EDIT or not options) and not matching:
                    continue
                raise AppearanceError("A saved PDF choice cannot be matched unambiguously to a display label.")
            chosen.append(matching[0])
    if flags & _COMBO:
        if len(saved) > 1:
            raise AppearanceError("A PDF combo box has multiple saved selections.")
        return options[chosen[0]][1] if chosen else saved[0] if saved else "", None, ()
    if len(chosen) > 1 and not flags & _MULTISELECT:
        raise AppearanceError("A PDF single-select list has multiple saved selections.")
    return "", [display for _, display in options], chosen


def _pdf_number(value):
    return f"{float(value):.6f}".rstrip("0").rstrip(".") or "0"


def _color(values, operators):
    if len(values) not in operators or not all(_finite(value) and 0 <= value <= 1 for value in values):
        raise AppearanceError("A PDF form appearance has an unsupported color.")
    return " ".join([*map(_pdf_number, values), operators[len(values)]])


def _widget_color(characteristics, key, default, resolve, *, stroke=False):
    if key not in characteristics:
        values = default
    else:
        values = resolve(characteristics.get(key))
        if not isinstance(values, ArrayObject) or len(values) not in (0, 1, 3, 4):
            raise AppearanceError("A PDF form widget has an unsupported decoration color.")
        values = [resolve(value) for value in values]
    return _color(values, {1: "G", 3: "RG", 4: "K"} if stroke else {1: "g", 3: "rg", 4: "k"}) if values else None


def _default_appearance(reader, inherited, form, resolve):
    da = inherited.get("/DA")
    if not _present(da):
        da = resolve(form.get("/DA"))
    name, size, color = NameObject("/Helv"), 12.0, "0 g"
    if _present(da):
        if not isinstance(da, (TextStringObject, ByteStringObject)):
            raise AppearanceError("The saved PDF form default appearance is not a string.")
        data = da.original_bytes
        if len(data) > 8192:
            raise AppearanceError("The saved PDF form default appearance exceeds the supported size.")
        stream = DecodedStreamObject()
        stream.set_data(data)
        for operands, op in _operations(reader, stream):
            if op == b"Tf" and len(operands) == 2 and isinstance(operands[0], NameObject) and _finite(operands[1]):
                name, size = operands
            elif op in _COLOR_SIZE and len(operands) == _COLOR_SIZE[op]:
                color = _color(operands, {_COLOR_SIZE[op]: op.decode("ascii")})
            else:
                raise AppearanceError("The PDF form's saved text styling cannot be reproduced locally.")
    if not 0 <= size <= 512:
        raise AppearanceError("The PDF form's saved font size cannot be reproduced locally.")
    resources = inherited.get("/DR")
    if not _present(resources):
        resources = resolve(form.get("/DR"))
    raw_font = _font_reference(resources, name, resolve)
    if _present(raw_font):
        return raw_font, _font_metrics(raw_font, resolve), float(size), color
    try:
        font = Font.from_core_font_name(_CORE_ALIASES.get(name, name))
    except (ValueError, KeyError):
        raise AppearanceError("The PDF form's saved font resource is unavailable.") from None
    return font.as_font_resource(), font, float(size), color


def _decoration(widget, resolve):
    rect = _numeric_array(resolve(widget.get("/Rect")), 4, resolve, "A PDF form widget has an invalid rectangle.")
    width, height = abs(rect[2] - rect[0]), abs(rect[3] - rect[1])
    if not 0 < min(width, height) <= max(width, height) <= 1_000_000:
        raise AppearanceError("A PDF form widget has an unusable rendering rectangle.")
    mk = resolve(widget.get("/MK"))
    if not _present(mk):
        mk = DictionaryObject()
    if not _dictionary(mk):
        raise AppearanceError("A PDF form widget has malformed appearance characteristics.")
    rotation = _integer(resolve(mk.get("/R")), 0, "A PDF form widget has invalid appearance rotation.") % 360
    if rotation not in (0, 90, 180, 270):
        raise AppearanceError("A PDF form widget has unsupported appearance rotation.")
    if rotation in (90, 270):
        width, height = height, width
    matrix = {90: (0, 1, -1, 0, height, 0), 180: (-1, 0, 0, -1, width, height),
              270: (0, -1, 1, 0, 0, width)}.get(rotation)
    border = resolve(widget.get("/Border"))
    border_width, dash = 1, None
    if _present(border):
        if not isinstance(border, ArrayObject) or len(border) not in (3, 4):
            raise AppearanceError("A PDF form widget has an unsupported border.")
        horizontal, vertical, border_width = (resolve(item) for item in border[:3])
        if horizontal != 0 or vertical != 0:
            raise AppearanceError("A rounded PDF form widget border cannot be regenerated.")
        if len(border) == 4:
            dash = resolve(border[3])
    style = resolve(widget.get("/BS"))
    if _present(style):
        if not _dictionary(style):
            raise AppearanceError("A PDF form widget has malformed border styling.")
        border_width = resolve(style.get("/W", NumberObject(1)))
        kind = resolve(style.get("/S", NameObject("/S")))
        if kind not in ("/S", "/D", "/U"):
            raise AppearanceError("The PDF form widget border style cannot be regenerated.")
        dash = resolve(style.get("/D", ArrayObject([NumberObject(3)]))) if kind == "/D" else None
    else:
        kind = "/S"
    if not _finite(border_width) or not 0 <= border_width < min(width, height) / 2:
        raise AppearanceError("A PDF form widget has an invalid border width.")
    lines = ["q"]
    background = _widget_color(mk, "/BG", (1,), resolve)
    if background:
        lines.append(f"{background} 0 0 {_pdf_number(width)} {_pdf_number(height)} re f")
    foreground = _widget_color(mk, "/BC", (0,), resolve, stroke=True)
    if foreground and border_width:
        lines.append(f"{foreground} {_pdf_number(border_width)} w")
        if dash is not None:
            if not isinstance(dash, ArrayObject) or not 0 < len(dash) <= 8:
                raise AppearanceError("A PDF form widget has an invalid dashed border.")
            pattern = [resolve(item) for item in dash]
            if not all(_finite(item) and item >= 0 for item in pattern) or not any(pattern):
                raise AppearanceError("A PDF form widget has an invalid dashed border.")
            lines.append(f"[{' '.join(map(_pdf_number, pattern))}] 0 d")
        half = _pdf_number(border_width / 2)
        if kind == "/U":
            lines.append(f"0 {half} m {_pdf_number(width)} {half} l S")
        else:
            lines.append(f"{half} {half} {_pdf_number(width - border_width)} {_pdf_number(height - border_width)} re S")
    lines.append("Q")
    return width, height, matrix, max(0, border_width if foreground else 0), lines


def _glyphs(font):
    if not font.interpretable or font.sub_type == "Type3":
        raise AppearanceError("The PDF form's font cannot reliably reproduce the saved characters.")
    reverse, encoding = font._get_typographic_maps()

    def encode(text):
        parts, width = [], 0.0
        for char in text:
            glyph = reverse.get(char)
            data = encoding.get(glyph)
            if not data or _decode(data, font) != char:
                raise AppearanceError("The PDF form's font cannot display every saved character.")
            parts.append(data)
            width += font.get_text_width(glyph)
        if not isfinite(width) or width < 0:
            raise AppearanceError("The PDF form's font has invalid display metrics.")
        return b"".join(parts), width / 1000

    return encode


def _wrap_text(text, fit_width, size, encode):
    """Wrap without losing current-value characters or overflowing a field."""
    rows = []
    for paragraph in text.split("\n"):
        line, width = [], 0.0
        for token in re.findall(r"\S+|[^\S\n]+", paragraph):
            token_width = encode(token)[1] * size
            if token_width > fit_width:
                units = token
            else:
                units = (token,)
            for unit in units:
                unit_width = encode(unit)[1] * size
                if unit_width > fit_width:
                    raise AppearanceError("A saved PDF form character does not fit its widget.")
                if line and width + unit_width > fit_width:
                    rows.append("".join(line))
                    line, width = [], 0.0
                line.append(unit)
                width += unit_width
        rows.append("".join(line))
    return rows


def _text_pdf(data):
    # Literal syntax keeps ordinary text readable in PDF inspection tools.
    escaped = bytearray()
    for char in data:
        if char in (40, 41, 92):
            escaped.extend((92, char))
        elif not 32 <= char <= 126:
            escaped.extend(f"\\{char:03o}".encode("ascii"))
        else:
            escaped.append(char)
    return f"({escaped.decode('ascii')})"


def _generated_stream(width, height, matrix, lines, font=None):
    stream = DecodedStreamObject()
    stream.set_data(("\n".join(lines) + "\n").encode("ascii"))
    if len(stream.get_data()) > MAX_APPEARANCE_BYTES:
        raise AppearanceError("Generated PDF form artwork exceeds the supported size.")
    stream.update({
        NameObject("/Type"): NameObject("/XObject"),
        NameObject("/Subtype"): NameObject("/Form"),
        NameObject("/FormType"): NumberObject(1),
        NameObject("/BBox"): ArrayObject(FloatObject(value) for value in (0, 0, width, height)),
        NameObject("/Resources"): DictionaryObject(),
    })
    if matrix:
        stream[NameObject("/Matrix")] = ArrayObject(FloatObject(value) for value in matrix)
    if font is not None:
        stream["/Resources"][NameObject("/Font")] = DictionaryObject({NameObject("/WBSFont"): font})
    return stream


def _regenerate(reader, widget, inherited, form, resolve, text, flags, *, options=None, selected=()):
    if _present(inherited.get("/RV")):
        raise AppearanceError("This PDF form needs a saved rich-text appearance that cannot be regenerated.")
    width, height, matrix, border, lines = _decoration(widget, resolve)
    if not text and not options:
        return _generated_stream(width, height, matrix, lines)
    raw_font, font, size, color = _default_appearance(reader, inherited, form, resolve)
    encode = _glyphs(font)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    if any((unicodedata.category(char).startswith("C") and char != "\n")
           or unicodedata.bidirectional(char) in ("R", "AL", "AN") for char in text):
        raise AppearanceError("This PDF form value requires text layout that cannot be regenerated locally.")
    left = max(border + 1, min(8, width / 10))
    vertical = max(border + 1, min(8, height / 5))
    available_width = width - 2 * left
    multiline = bool(flags & _MULTILINE)
    alignment = inherited.get("/Q")
    if not _present(alignment):
        alignment = resolve(form.get("/Q"))
    alignment = _integer(alignment, 0, "A PDF form widget has invalid text alignment.")
    if alignment not in (0, 1, 2):
        raise AppearanceError("A PDF form widget has unsupported text alignment.")
    comb = 0
    if flags & _COMB:
        comb = _integer(inherited.get("/MaxLen"), 0, "A comb PDF form has an invalid maximum length.")
        if multiline or not 0 < comb <= MAX_VALUE_CHARACTERS or len(text) > comb:
            raise AppearanceError("A saved PDF comb value does not fit its declared cells.")
    if available_width <= 0 or not multiline and "\n" in text:
        raise AppearanceError("A saved PDF form value does not fit its declared text layout.")
    if size == 0:
        width_estimate = max((encode(row)[1] for row in text.split("\n")), default=0)
        size = min(12.0, height - 2 * vertical)
        if not multiline and not options and width_estimate:
            size = min(size, available_width / width_estimate)
    if not 2 <= size <= height - 2 * border:
        raise AppearanceError("The PDF form text would be clipped at the saved font size.")

    # Draw each run at an explicit baseline and validate its typographic extents
    # before writing, avoiding implicit renderer clipping or placeholder glyphs.
    leading = max(size * 1.2, size + 2)
    if options is not None:
        top = _integer(inherited.get("/TI"), 0, "A PDF choice field has an invalid first visible item.")
        if not 0 <= top < len(options):
            raise AppearanceError("A PDF choice field has an invalid first visible item.")
        capacity = max(0, int((height - 2 * vertical - size) / leading) + 1)
        rows = options[top:top + capacity]
        if any(not top <= index < top + len(rows) for index in selected):
            raise AppearanceError("A saved PDF choice would be hidden outside its display area.")
        baseline = height - vertical - size
        selected = {index - top for index in selected}
    elif multiline:
        rows = _wrap_text(text, available_width, size, encode)
        baseline = height - vertical - size
    else:
        rows = list(text) if comb else [text]
        baseline = (height - size) / 2

    descent = float(font.font_descriptor.descent) / 1000
    ascent = float(font.font_descriptor.ascent) / 1000
    if not isfinite(descent) or not isfinite(ascent):
        raise AppearanceError("The PDF form's font has invalid vertical display metrics.")
    for index, row in enumerate(rows):
        raw, unscaled = encode(row)
        row_width = unscaled * size
        y = baseline if comb else baseline - index * leading
        cell_width = available_width / comb if comb else available_width
        if row_width > cell_width + 1e-6 or y + descent * size < border - 1e-6 or y + ascent * size > height - border + 1e-6:
            raise AppearanceError("A saved PDF form value would be clipped in its widget.")
        if comb:
            x = left + index * cell_width + (cell_width - row_width) / 2
        else:
            x = (left, (width - row_width) / 2, width - left - row_width)[alignment]
        if index in selected:
            lines.append(f"q 0.75 0.85 1 rg 2 {_pdf_number(y - 2)} {_pdf_number(width - 4)} {_pdf_number(leading)} re f Q")
        lines.append(f"BT /WBSFont {_pdf_number(size)} Tf {color} {_pdf_number(x)} {_pdf_number(y)} Td {_text_pdf(raw)} Tj ET")
    return _generated_stream(width, height, matrix, lines, raw_font)


def build_widget_appearance(reader, widget, inherited, form, resolve, *, page_resources=None):
    """Return only the selected normal stream or freshly composed static artwork."""
    try:
        kind = inherited.get("/FT")
        flags = _integer(inherited.get("/Ff"), 0, "A PDF form widget has invalid field flags.")
        if not 0 <= flags < 1 << 32:
            raise AppearanceError("A PDF form widget has invalid field flags.")
        normal = _normal(widget, resolve)
        if kind == "/Btn":
            result = _button(reader, widget, inherited, flags, normal, resolve)
            if result is not None:
                return result
            mk = resolve(widget.get("/MK"))
            caption = resolve(mk.get("/CA")) if _dictionary(mk) else None
            if not _present(caption):
                raise AppearanceError("A visible PDF push button lacks its static appearance.")
            return _regenerate(reader, widget, inherited, form, resolve, _text(caption), 0)

        raw, source = _normal_stream(normal, resolve(widget.get("/AS")), resolve)
        if kind == "/Sig":
            if source is not None and _has_artwork(reader, source):
                return raw
            if not _present(inherited.get("/V")) and source is None and "/MK" not in widget and "/BS" not in widget:
                return None
            raise AppearanceError("A visible PDF signature widget lacks reproducible static artwork.")
        if kind not in ("/Tx", "/Ch"):
            if not _present(kind) and source is not None and _has_artwork(reader, source):
                return raw
            raise AppearanceError("A visible PDF form widget uses an unsupported field type.")

        value = inherited.get("/V")
        options, selected = None, ()
        if kind == "/Ch":
            text, options, selected = _choice(inherited, flags, resolve)
            flags = 0
        else:
            text = _text(value) if _present(value) else ""
            if flags & _PASSWORD:
                text = "*" * len(text)
        if source is not None:
            artwork = {"rows": [], "fills": []} if options else None
            try:
                parts = _appearance_text(reader, source, page_resources, resolve, artwork)
            except _UNREADABLE:
                parts = None
            if (options is None and _matches_value(parts, text)
                    or artwork is not None and parts is not None
                    and _matches_choice_artwork(source, artwork, options, selected, inherited, resolve)):
                return raw
        return _regenerate(reader, widget, inherited, form, resolve, text, flags, options=options, selected=selected)
    except AppearanceError:
        raise
    except _UNREADABLE as error:
        raise AppearanceError("A PDF form appearance or its local rendering resources cannot be read reliably.") from error
