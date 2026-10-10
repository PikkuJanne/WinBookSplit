"""Original deterministic page-fidelity fixtures; no source document is read.

Generated PDF/image bytes use the repository MIT license. This developer-only
module uses pinned ReportLab, Pillow and pypdf, never application dependencies.
"""
from __future__ import annotations

from hashlib import sha256
from importlib.metadata import version
from io import BytesIO
import json
from pathlib import Path
import platform
import re

from PIL import Image, ImageDraw
from pypdf import PdfReader, PdfWriter
from pypdf.generic import (ArrayObject, BooleanObject, DecodedStreamObject,
                           DictionaryObject, FloatObject, IndirectObject,
                           NameObject, NullObject, NumberObject, TextStringObject)
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas

ROOT = Path(__file__).resolve().parents[2]
AUTHOR = 'Authored Å café 日本 source author'
TITLE = 'Original Å café 日本 page fidelity fixture'
SIZES = [(420, 560), (500, 400), (360, 520), (610, 390), (400, 600), (520, 520)]
ROTATIONS = [0, 90, 180, 270, 0, 90]
RANGES = {'manual': [[0, 3], [3, 6]], '1': [[0, 3], [3, 6]],
          '2': [[0, 1], [1, 3], [3, 4], [4, 6]]}
STATIC_KEYS = ('/Subtype', '/Rect', '/Contents', '/F', '/C', '/QuadPoints',
               '/Border', '/BS', '/Open', '/Name', '/AP')


def array(values):
    return ArrayObject([FloatObject(value) for value in values])


def authored_image(page, scanned=False):
    image = Image.new('RGB', (300, 380) if scanned else (120, 90), 'white')
    draw = ImageDraw.Draw(image)
    for y in range(0, image.height, 20):
        for x in range(0, image.width, 20):
            color = ((page * 31 + x) % 230, (page * 47 + y) % 230, (page * 73 + x + y) % 230)
            draw.rectangle((x, y, x + 18, y + 18), fill=color)
    draw.rectangle((5, 5, image.width - 6, 42), fill='white', outline='black', width=2)
    draw.text((12, 15), ('IMAGE-ONLY-' if scanned else 'EMBEDDED-') + f'PAGE-{page:03d}', fill='black')
    payload = BytesIO()
    image.save(payload, format='PNG', optimize=False)
    return payload.getvalue()


def appearance(writer, subtype, width, height):
    stream = DecodedStreamObject()
    if subtype == '/Highlight':
        paint = f'q 1 0.85 0 rg 0 0 {width} {height} re f Q\n'
    elif subtype == '/Square':
        paint = f'q 0.1 0.25 0.8 RG 2 w 1 1 {width - 2} {height - 2} re S Q\n'
    else:
        paint = f'q 1 0.9 0.1 rg 0 0 {width} {height} re f 0 0 0 RG 1 w 1 1 {width - 2} {height - 2} re S Q\n'
    stream.set_data(paint.encode('ascii'))
    stream.update({NameObject('/Type'): NameObject('/XObject'), NameObject('/Subtype'): NameObject('/Form'),
                   NameObject('/BBox'): array([0, 0, width, height]), NameObject('/Resources'): DictionaryObject()})
    return DictionaryObject({NameObject('/N'): writer._add_object(stream)})


def add_static(writer, page_index):
    subtype = ('/Text', '/Highlight', '/Square')[page_index % 3]
    rect = [55, 55, 175, 73] if subtype == '/Highlight' else [55, 55, 145, 105] if subtype == '/Square' else [55, 55, 75, 75]
    annotation = DictionaryObject({NameObject('/Type'): NameObject('/Annot'), NameObject('/Subtype'): NameObject(subtype),
                                  NameObject('/Rect'): array(rect), NameObject('/Contents'): TextStringObject(f'Authored static {subtype[1:]} {page_index + 1}'),
                                  NameObject('/NM'): TextStringObject(f'static-{page_index}'), NameObject('/F'): NumberObject(4),
                                  NameObject('/C'): array([1, 0.85, 0]), NameObject('/AP'): appearance(writer, subtype, rect[2] - rect[0], rect[3] - rect[1])})
    if subtype == '/Highlight':
        annotation[NameObject('/QuadPoints')] = array([55, 73, 175, 73, 55, 55, 175, 55])
    if subtype == '/Text':
        annotation.update({NameObject('/Open'): BooleanObject(False), NameObject('/Name'): NameObject('/Comment')})
    # Deliberately authored cross-page /P on page0 reproduces clone-by-reference
    # risk. The visible static appearance remains normal and inert.
    annotation[NameObject('/P')] = writer.pages[5 if page_index == 0 else page_index].indirect_reference
    writer.add_annotation(page_index, annotation)
    if page_index == 0:
        # add_annotation replaces /P; restore the authored reference afterward.
        writer.pages[page_index]['/Annots'][-1].get_object()[NameObject('/P')] = writer.pages[5].indirect_reference


def add_link(writer, source, name, destination, action=False):
    annotation = DictionaryObject({NameObject('/Type'): NameObject('/Annot'), NameObject('/Subtype'): NameObject('/Link'),
                                  NameObject('/Rect'): array([200, 55, 220, 75]), NameObject('/Border'): array([0, 0, 0]),
                                  NameObject('/NM'): TextStringObject(name), NameObject('/F'): NumberObject(4)})
    if action:
        annotation[NameObject('/A')] = DictionaryObject({NameObject('/S'): NameObject('/GoTo'), NameObject('/D'): destination})
    else:
        annotation[NameObject('/Dest')] = destination
    writer.add_annotation(source, annotation)


def destination(writer, page):
    return ArrayObject([writer.pages[page].indirect_reference, NameObject('/Fit')])


def fixture_bytes(kind):
    scanned = kind == 'scanned'
    count = 4 if scanned else 6
    stream = BytesIO()
    document = canvas.Canvas(stream, invariant=1, pageCompression=1)
    for index in range(count):
        width, height = SIZES[index]
        document.setPageSize((width, height))
        image = ImageReader(BytesIO(authored_image(index + 1, scanned)))
        if scanned:
            document.drawImage(image, 20, 20, width=width - 40, height=height - 40)
        else:
            document.setFont('Helvetica-Bold', 16)
            document.drawString(30, height - 42, 'Original WinBookSplit fidelity fixture')
            document.setFont('Courier-Bold', 22)
            document.drawString(30, height - 75, f'WBS-FID-PAGE-{index + 1:03d}')
            document.setFont('Helvetica', 11)
            document.drawString(30, height - 96, f'Physical source page {index + 1}; original MIT content')
            document.setStrokeColorRGB(0.2, 0.3, 0.75)
            document.setLineWidth(2)
            document.rect(30, height - 135, width - 60, 20, fill=0, stroke=1)
            document.line(30, 130, width - 30, 160 + index * 3)
            document.circle(width - 65, 100, 24, fill=0, stroke=1)
            document.drawImage(image, 35, 175, width=120, height=90)
        document.showPage()
    document.save()
    writer = PdfWriter()
    for index, page in enumerate(PdfReader(BytesIO(stream.getvalue())).pages):
        added = writer.add_page(page)
        width, height = SIZES[index]
        added[NameObject('/Rotate')] = NumberObject(ROTATIONS[index])
        added[NameObject('/CropBox')] = array([10, 10, width - 10, height - 10])
        added[NameObject('/TrimBox')] = array([15, 15, width - 15, height - 15])
        added[NameObject('/BleedBox')] = array([12, 12, width - 12, height - 12])
        added[NameObject('/ArtBox')] = array([20, 20, width - 20, height - 20])
        added[NameObject('/WBSFixturePage')] = NumberObject(index + 1)
    writer.metadata = None
    writer.add_metadata({'/Title': TITLE + (' image-only' if scanned else ''), '/Author': AUTHOR,
                         '/Subject': 'Original generated developer fixture; no private document',
                         '/Creator': 'WinBookSplit original fidelity fixture generator',
                         '/Producer': 'Authored source producer; do not falsely retain as output tool',
                         '/CreationDate': 'D:20000101000000Z', '/ModDate': 'D:20000101000000Z',
                         '/AuthoredPrivateField': 'MOCK_LOCAL_METADATA_NOT_FOR_OUTPUT_CATALOG'})
    if not scanned:
        first = writer.add_outline_item('Chapter Å café 日本 One', 0)
        writer.add_outline_item('Child One', 1, parent=first)
        second = writer.add_outline_item('Chapter Two', 3)
        writer.add_outline_item('Child Two', 4, parent=second)
        for index in range(count):
            add_static(writer, index)
            add_link(writer, index, f'forward-{index}', destination(writer, min(index + 1, count - 1)), True)
            add_link(writer, index, f'backward-{index}', destination(writer, max(index - 1, 0)))
            add_link(writer, index, f'cross-{index}', destination(writer, 5 if index < 3 else 0), index % 2 == 0)
            named = f'authored-target-{index}'
            writer.add_named_destination(named, (index + 1) % count)
            add_link(writer, index, f'named-{index}', TextStringObject(named), index % 2 == 1)
        add_link(writer, 0, 'null-target', ArrayObject([NullObject(), NameObject('/Fit')]))
        add_link(writer, 0, 'missing-named', TextStringObject('authored-missing-target'), True)
    result = BytesIO()
    writer.write(result)
    return result.getvalue()


def resolved(value):
    return value.get_object() if hasattr(value, 'get_object') else value


def normalized(value, depth=0):
    if depth > 16:
        raise ValueError('Authored structural snapshot exceeded depth')
    value = resolved(value)
    if isinstance(value, DictionaryObject):
        result = {str(key): normalized(item, depth + 1) for key, item in sorted(value.items()) if key not in ('/Length', '/Filter', '/DecodeParms')}
        if hasattr(value, 'get_data'):
            result['decoded_stream_sha256'] = sha256(value.get_data()).hexdigest()
        return result
    if isinstance(value, (list, tuple, ArrayObject)):
        return [normalized(item, depth + 1) for item in value]
    if isinstance(value, (NullObject, type(None))):
        return None
    if isinstance(value, BooleanObject):
        return bool(value.value)
    if isinstance(value, (NumberObject, FloatObject, int, float)):
        return float(value)
    if isinstance(value, bytes):
        return {'bytes_sha256': sha256(value).hexdigest()}
    return str(value)


def destination_page(reader, value):
    value = resolved(value)
    if isinstance(value, (str, NameObject, TextStringObject)):
        name = str(value)
        entry = reader.named_destinations.get(name)
        if entry is None:
            return None
        return reader.get_destination_page_number(entry)
    if isinstance(value, DictionaryObject):
        value = resolved(value.get('/D'))
    if not isinstance(value, ArrayObject) or len(value) < 2 or not isinstance(value[0], IndirectObject):
        return None
    return next((index for index, page in enumerate(reader.pages) if page.indirect_reference == value[0]), None)


def destination_view(reader, value):
    value = resolved(value)
    if isinstance(value, (str, NameObject, TextStringObject)):
        entry = reader.named_destinations.get(str(value))
        value = None if entry is None else entry.dest_array
    if isinstance(value, DictionaryObject):
        value = resolved(value.get('/D'))
    if not isinstance(value, ArrayObject) or len(value) < 2:
        return {'fit': None, 'args': []}
    return {'fit': normalized(value[1]), 'args': normalized(value[2:])}


def page_snapshot(reader, index):
    page = reader.pages[index]
    resources = resolved(page.get('/Resources', {}))
    images = []
    for name, item in sorted(resolved(resources.get('/XObject', {})).items()):
        item = resolved(item)
        if item.get('/Subtype') == '/Image':
            images.append({'name': str(name), 'width': int(item['/Width']), 'height': int(item['/Height']),
                           'bits_per_component': int(item['/BitsPerComponent']), 'color_space': normalized(item['/ColorSpace']),
                           'decoded_sha256': sha256(item.get_data()).hexdigest()})
    static, links = [], []
    for entry in page.get('/Annots', []):
        entry = resolved(entry)
        if entry.get('/Subtype') == '/Link':
            value = entry.get('/Dest')
            action = resolved(entry.get('/A', {}))
            if value is None and action.get('/S') == '/GoTo':
                value = action.get('/D')
            links.append({'name': str(entry.get('/NM', '')), 'target_page': destination_page(reader, value),
                          'form': 'action' if '/A' in entry else 'destination',
                          **destination_view(reader, value),
                          'rect': normalized(entry.get('/Rect')), 'border': normalized(entry.get('/Border'))})
        else:
            static.append({'name': str(entry.get('/NM', '')), 'structure': {key: normalized(entry[key]) for key in STATIC_KEYS if key in entry}})
    contents = page.get_contents()
    return {'page_id': int(page['/WBSFixturePage']), 'text': page.extract_text() or '',
            'content_sha256': sha256(b'' if contents is None else contents.get_data()).hexdigest(),
            'images': images, 'geometry': {'media': list(map(float, page.mediabox)), 'crop': list(map(float, page.cropbox)),
                                         'trim': list(map(float, page.trimbox)), 'bleed': list(map(float, page.bleedbox)),
                                         'art': list(map(float, page.artbox)), 'rotation': int(page.get('/Rotate', 0))},
            'static_annotations': static, 'links': links}


def all_page_object_count(reader):
    ids = {(generation, number) for generation, rows in reader.xref.items() for number in rows if number != 0}
    ids.update((0, number) for number in reader.xref_objStm)
    count = 0
    for generation, number in sorted(ids):
        value = reader.get_object(IndirectObject(number, generation, reader))
        if isinstance(value, DictionaryObject) and value.get('/Type') == '/Page':
            count += 1
    return count


def serialized_inventory(reader):
    """Inspect all serialized objects, including unreachable cloned resources."""
    ids = {(generation, number) for generation, rows in reader.xref.items() for number in rows if number != 0}
    ids.update((0, number) for number in reader.xref_objStm)
    markers, images = set(), set()
    for generation, number in sorted(ids):
        value = reader.get_object(IndirectObject(number, generation, reader))
        if isinstance(value, DictionaryObject) and hasattr(value, 'get_data'):
            raw = value.get_data()
            markers.update(int(match) for match in re.findall(rb'WBS-FID-PAGE-([0-9]{3})', raw))
            if value.get('/Subtype') == '/Image':
                images.add(sha256(raw).hexdigest())
    return {'all_page_object_count': all_page_object_count(reader), 'serialized_content_marker_ids': sorted(markers),
            'all_image_decoded_sha256': sorted(images)}


def generate_fixtures(directory):
    directory = Path(directory)
    directory.mkdir(parents=True, exist_ok=False)
    result, records = {}, {}
    for kind in ('rich', 'scanned'):
        path = directory / (kind + '-original.pdf')
        payload = fixture_bytes(kind)
        with path.open('xb') as handle:
            handle.write(payload)
        reader = PdfReader(path)
        snapshots = [page_snapshot(reader, index) for index in range(len(reader.pages))]
        assert [row['page_id'] for row in snapshots] == list(range(1, len(reader.pages) + 1))
        assert all(bool(row['text']) == (kind == 'rich') and len(row['images']) == 1 for row in snapshots)
        result[kind] = path
        records[kind] = {'filename': path.name, 'bytes': len(payload), 'sha256': sha256(payload).hexdigest(),
                         'page_count': len(reader.pages), 'snapshots': snapshots, 'metadata': dict(reader.metadata),
                         'all_page_object_count': all_page_object_count(reader)}
    provenance = {'schema_version': 1, 'origin': 'Original programmatically generated text/vectors/images/annotations/navigation; no source document',
                  'license': 'MIT - repository LICENSE', 'private_data': False,
                  'versions': {'python': platform.python_version(), 'pypdf': version('pypdf'), 'reportlab': version('reportlab'), 'Pillow': version('Pillow')},
                  'generator_sha256': sha256(Path(__file__).read_bytes()).hexdigest(), 'fixtures': records}
    with (directory / 'provenance.json').open('x', encoding='utf-8', newline='\n') as handle:
        json.dump(provenance, handle, indent=2, sort_keys=True, ensure_ascii=True)
        handle.write('\n')
    return result, provenance
