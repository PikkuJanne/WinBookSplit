"""Original, synthetic PDF regression seeds. Not a splitter or runtime component.

Requires developer-only pypdf and ReportLab. Writes only a NEW output directory.
Run: python tests/fixtures/generate_form_fixtures.py --output fixtures-new
The XFA sample is a detection sentinel, NOT an XFA-renderer acceptance document.
"""
from __future__ import annotations
import argparse
from hashlib import sha256
from io import BytesIO
import json
from pathlib import Path
import sys

from reportlab.pdfgen import canvas
from pypdf import PdfReader, PdfWriter
from pypdf.generic import (ArrayObject, BooleanObject, DecodedStreamObject,
    DictionaryObject, FloatObject, NameObject, NullObject, NumberObject,
    TextStringObject)

FEATURE_CASES = (
    'empty_feature_trees', 'document_javascript', 'document_attachment',
    'script_and_attachment', 'associated_attachment', 'page_and_annotation_actions',
    'named_navigation_and_features', 'broken_omitted_branches', 'payload_aliases',
    'form_and_features', 'encrypted_metadata_and_features', 'encrypted_form_and_features',
)
SCRIPT_MARKER = 'WBS-SYNTHETIC-OMITTED-SCRIPT-407'
ATTACHMENT_MARKER = 'WBS-SYNTHETIC-OMITTED-PAYLOAD-618'
ATTACHMENT_NAME = 'authored-attachment-never-extract.txt'
FEATURE_MARKERS = (SCRIPT_MARKER, ATTACHMENT_MARKER, ATTACHMENT_NAME)
FORM_ACTION_MARKER = 'WBS_SYNTHETIC_ACTION_DO_NOT_EXECUTE'
PRIVATE_FORM_MARKERS = {
    'hidden_fields': ('HIDDEN-FORM-SECRET-61', 'NOVIEW-FORM-SECRET-29', 'DETACHED-FORM-SECRET-73'),
    'signature_visible': ('SYNTHETIC-NOT-A-REAL-SIGNATURE',),
    'stale_appearance': ('STALE-VALUE',),
    'form_action': (FORM_ACTION_MARKER,),
}
FORM_CASES = (
    'text', 'controls', 'shared_field', 'missing_appearance', 'stale_appearance',
    'orphan_widget', 'rotated_crop', 'form_action', 'encrypted_form',
    'multiline_readonly', 'multiselect', 'multiselect_missing_appearance', 'combo_missing_appearance',
    'winansi_missing_appearance', 'hierarchy', 'transformed_appearance',
    'hidden_fields', 'signature_visible', 'button_stale_state',
    'form_and_features', 'encrypted_form_and_features',
)

CASES = (
    'plain', 'acroform_null', 'acroform_empty', 'fields_empty', 'resource_only',
    'text', 'controls', 'shared_field', 'missing_appearance', 'stale_appearance',
    'orphan_widget', 'rotated_crop', 'form_action', 'xfa_null', 'xfa_empty_array',
    'xfa_detection_only', 'fields_malformed', 'field_cycle',
    'signature_structure', 'encrypted_form', *FEATURE_CASES, 'multiline_readonly',
    'multiselect', 'multiselect_missing_appearance', 'combo_missing_appearance', 'winansi_missing_appearance',
    'hierarchy', 'transformed_appearance', 'hidden_fields',
    'signature_visible', 'button_stale_state', 'missing_appearance_font',
)
NEGATIVE = {'xfa_detection_only', 'fields_malformed', 'field_cycle', 'signature_structure',
            'missing_appearance_font'}


def name(s: str) -> NameObject:
    return NameObject(s)


def arr(values) -> ArrayObject:
    return ArrayObject([FloatObject(x) for x in values])


def literal(value: str) -> str:
    return value.replace('\\', '\\\\').replace('(', '\\(').replace(')', '\\)')


def winansi_literal(value: str) -> str:
    return ''.join(literal(chr(byte)) if 32 <= byte < 127 else f'\\{byte:03o}' for byte in value.encode('cp1252'))


def base_writer() -> PdfWriter:
    raw = BytesIO()
    c = canvas.Canvas(raw, pagesize=(612, 792), invariant=1, pageCompression=0)
    c.setTitle('Original WinBookSplit form-support fixture')
    c.setAuthor('WinBookSplit synthetic test fixture')
    for i in range(4):
        c.setFont('Helvetica-Bold', 18)
        c.drawString(48, 744, 'WinBookSplit form regression')
        c.setFont('Helvetica', 12)
        c.drawString(48, 712, f'SOURCE-PAGE-{i + 1}')
        c.drawString(48, 684, ['Front matter', 'Chapter A', 'Chapter B', 'Appendix'][i])
        c.setFont('Helvetica', 10)
        c.drawString(48, 572, 'Original synthetic data only. No personal or book content.')
        c.drawString(48, 550, 'Field region below is intentionally blank in the base pages.')
        c.showPage()
    c.save()
    raw.seek(0)
    w = PdfWriter()
    for page in PdfReader(raw).pages:
        w.add_page(page)
    a = w.add_outline_item('A', 1)
    w.add_outline_item('A.1', 1, parent=a)
    b = w.add_outline_item('B', 2)
    w.add_outline_item('B.1', 2, parent=b)
    w.add_outline_item('Appendix', 3)
    return w


def font_ref(w: PdfWriter):
    return w._add_object(DictionaryObject({
        name('/Type'): name('/Font'), name('/Subtype'): name('/Type1'),
        name('/BaseFont'): name('/Helvetica'),
    }))


def appearance(w: PdfWriter, label: str = '', width: float = 260,
               height: float = 32, mark: str | None = None):
    s = DecodedStreamObject()
    if mark is not None:
        # Intentionally obvious graphic state so a dropped checkbox cannot pass.
        drawing = f'q 1 1 1 rg 0 0 {width} {height} re f 0 0 0 RG 1 w 1 1 {width-2} {height-2} re S\n'
        if mark == 'on':
            drawing += f'2 w 4 4 m {width-4} {height-4} l S 4 {height-4} m {width-4} 4 l S\n'
    else:
        drawing = (f'q 1 1 1 rg 0 0 {width} {height} re f 0 0 0 RG 1 w '
                   f'0.5 0.5 {width-1} {height-1} re S '
                   f'BT /F0 12 Tf 0 0 0 rg 8 10 Td ({literal(label)}) Tj ET\n')
    drawing += 'Q\n'
    s.set_data(drawing.encode('ascii'))
    s.update({name('/Type'): name('/XObject'), name('/Subtype'): name('/Form'),
              name('/FormType'): NumberObject(1), name('/BBox'): arr([0, 0, width, height]),
              name('/Matrix'): arr([1, 0, 0, 1, 0, 0]),
              name('/Resources'): DictionaryObject({name('/Font'): DictionaryObject({name('/F0'): font_ref(w)})})})
    return w._add_object(s)


def form(w: PdfWriter) -> DictionaryObject:
    if '/AcroForm' not in w.root_object or isinstance(w.root_object['/AcroForm'], NullObject):
        obj = DictionaryObject({name('/Fields'): ArrayObject(),
            name('/NeedAppearances'): BooleanObject(False),
            name('/DA'): TextStringObject('/F0 12 Tf 0 g'),
            name('/DR'): DictionaryObject({name('/Font'): DictionaryObject({name('/F0'): font_ref(w)})})})
        w.root_object[name('/AcroForm')] = w._add_object(obj)
    return w.root_object['/AcroForm']


def add_widget(w: PdfWriter, page_index: int, field_type: str | None,
               title: str | None, value, rect, ap=None, parent=None, state=None,
               register=True):
    widget = DictionaryObject({name('/Type'): name('/Annot'), name('/Subtype'): name('/Widget'),
        name('/Rect'): arr(rect), name('/P'): w.pages[page_index].indirect_reference,
        name('/F'): NumberObject(4)})
    if field_type is not None:
        widget[name('/FT')] = name(field_type)
    if title is not None:
        widget[name('/T')] = TextStringObject(title)
    if value is not None:
        widget[name('/V')] = value
    if parent is not None:
        widget[name('/Parent')] = parent
    if ap is not None:
        widget[name('/AP')] = DictionaryObject({name('/N'): ap})
    if state is not None:
        widget[name('/AS')] = name(state)
    ref = w._add_object(widget)
    page = w.pages[page_index]
    if '/Annots' not in page:
        page[name('/Annots')] = ArrayObject()
    page['/Annots'].append(ref)
    if register:
        form(w)['/Fields'].append(ref)
    return ref


def paint_reference(w: PdfWriter, page_index: int, label: str, rect,
                    mark: str | None = None, commands: str | None = None):
    # Independent expected static representation: direct invocation of original
    # artwork at a known fixture coordinate, no field reader or flattener used.
    page = w.pages[page_index]
    resources = page['/Resources']
    if '/XObject' not in resources:
        resources[name('/XObject')] = DictionaryObject()
    objects = resources['/XObject']
    key = name(f'/Expected{len(objects) + 1}')
    x0, y0, x1, y1 = rect
    artwork = appearance(w, label, x1-x0, y1-y0, mark)
    if commands is not None:
        artwork.get_object().set_data(commands.encode('ascii'))
    objects[key] = artwork
    content = DecodedStreamObject()
    content.set_data(f'q 1 0 0 1 {x0} {y0} cm {key} Do Q\n'.encode('ascii'))
    extra = w._add_object(content)
    old = page.get('/Contents')
    # Retain all old streams without changing source page artwork.
    page[name('/Contents')] = ArrayObject(list(old) + [extra] if isinstance(old, ArrayObject) else [old, extra])


def add_static_drawing(writer, page_index, content):
    page = writer.pages[page_index]
    resources = page['/Resources']
    font_map = resources['/Font']
    if '/FStatic' not in font_map:
        font_map[name('/FStatic')] = font_ref(writer)
    drawing = DecodedStreamObject()
    drawing.set_data(content.encode('ascii'))
    old = page.get('/Contents')
    page[name('/Contents')] = ArrayObject([*(list(old) if isinstance(old, ArrayObject) else [old]),
                                           writer._add_object(drawing)])


def pdf_dict(**entries):
    return DictionaryObject({name('/' + key): value for key, value in entries.items()})


def add_image_and_layer(writer):
    """Use a real image resource inside a visible optional-content group."""
    layer = writer._add_object(pdf_dict(Type=name('/OCG'), Name=TextStringObject('Fixture artwork')))
    layers = ArrayObject([layer])
    writer.root_object[name('/OCProperties')] = pdf_dict(OCGs=layers, D=pdf_dict(Order=layers, ON=layers))
    image = DecodedStreamObject()
    image.set_data(bytes([23, 101, 211, 242, 160, 24, 215, 55, 29, 13, 144, 68]))
    image.update(pdf_dict(Type=name('/XObject'), Subtype=name('/Image'), Width=NumberObject(2),
                          Height=NumberObject(2), ColorSpace=name('/DeviceRGB'), BitsPerComponent=NumberObject(8)))
    page = writer.pages[1]
    resources = page['/Resources']
    resources.setdefault(name('/XObject'), DictionaryObject()).get_object()[name('/FixtureImage')] = writer._add_object(image)
    resources[name('/Properties')] = pdf_dict(FixtureLayer=layer)
    graphics = DecodedStreamObject()
    graphics.set_data(b'/OC /FixtureLayer BDC q 32 0 0 32 72 430 cm /FixtureImage Do Q EMC\n')
    original = page.get('/Contents')
    streams = list(original) if isinstance(original, ArrayObject) else [original]
    page[name('/Contents')] = ArrayObject([*streams, writer._add_object(graphics)])


def attach_annotation(writer, page, subtype, rect, **entries):
    obj = pdf_dict(Type=name('/Annot'), Subtype=name(subtype), Rect=arr(rect),
                   P=writer.pages[page].indirect_reference, **entries)
    writer.pages[page].setdefault(name('/Annots'), ArrayObject()).append(writer._add_object(obj))


def add_optional_features(writer, expected, kind, widget=None):
    script, file_spec = None, None
    add_image_and_layer(writer)
    add_image_and_layer(expected)
    if kind in {'named_navigation_and_features', 'form_and_features', 'encrypted_form_and_features'}:
        target = ArrayObject([writer.pages[1].indirect_reference, name('/Fit')])
        writer.add_named_destination_array(TextStringObject('LocalTarget'), target)
        writer.root_object['/Outlines']['/First'][name('/A')] = pdf_dict(
            S=name('/GoTo'), D=TextStringObject('LocalTarget'))
        writer.root_object['/Outlines']['/First']['/First'][name('/A')] = pdf_dict(
            S=name('/GoTo'), D=TextStringObject('LocalTarget'))
        # Both links stay inside the same chapter for Manual, Auto1 and Auto2.
        attach_annotation(writer, 1, '/Link', [250, 600, 320, 620], Dest=TextStringObject('LocalTarget'))
        attach_annotation(expected, 1, '/Link', [250, 600, 320, 620],
                          Dest=ArrayObject([expected.pages[1].indirect_reference, name('/Fit')]))
        if widget is not None:
            writer.root_object[name('/OpenAction')] = pdf_dict(S=name('/GoTo'), D=TextStringObject('LocalTarget'))
    names = writer.root_object.setdefault(name('/Names'), DictionaryObject()).get_object()
    combined = {'script_and_attachment', 'named_navigation_and_features', 'payload_aliases',
                'form_and_features', 'encrypted_metadata_and_features', 'encrypted_form_and_features'}
    if kind in combined | {'document_javascript', 'page_and_annotation_actions'}:
        script_text = TextStringObject(f'/* {SCRIPT_MARKER} */')
        if kind == 'payload_aliases':
            script_data = DecodedStreamObject()
            script_data.set_data(script_text.encode('ascii'))
            script_text = writer._add_object(script_data)
        script = writer._add_object(pdf_dict(Type=name('/Action'), S=name('/JavaScript'),
                                             JS=script_text))
        names[name('/JavaScript')] = pdf_dict(Names=ArrayObject([TextStringObject('AuthoredScript'), script]))
    if kind in combined | {'document_attachment', 'associated_attachment'}:
        payload = DecodedStreamObject()
        payload[name('/Type')] = name('/EmbeddedFile')
        payload.set_data((ATTACHMENT_MARKER + '\n').encode('ascii'))
        file_spec = writer._add_object(pdf_dict(Type=name('/Filespec'), F=TextStringObject(ATTACHMENT_NAME),
            EF=pdf_dict(F=writer._add_object(payload)), AFRelationship=name('/Supplement')))
        names[name('/EmbeddedFiles')] = pdf_dict(Names=ArrayObject([TextStringObject(ATTACHMENT_NAME), file_spec]))
    if kind == 'empty_feature_trees':
        names[name('/JavaScript')] = pdf_dict(Names=ArrayObject())
        names[name('/EmbeddedFiles')] = DictionaryObject()
    if kind == 'associated_attachment':
        writer.root_object[name('/AF')] = ArrayObject([file_spec])
        writer.pages[1][name('/AF')] = ArrayObject([file_spec])
        rect = [72, 370, 232, 402]
        attach_annotation(writer, 1, '/FileAttachment', rect, F=NumberObject(4), FS=file_spec,
                          AP=pdf_dict(N=appearance(writer, 'Attachment artwork', 160, 32)))
        paint_reference(expected, 1, 'Attachment artwork', rect)
    if kind == 'page_and_annotation_actions':
        writer.root_object[name('/OpenAction')] = script
        writer.pages[0][name('/AA')] = pdf_dict(O=script)
        # This annotation's static artwork remains visible after removing the action.
        rect = [72, 370, 232, 402]
        attach_annotation(writer, 1, '/Square', rect, AP=pdf_dict(N=appearance(writer, 'Annotation artwork', 160, 32)),
                          AA=pdf_dict(E=script))
        attach_annotation(expected, 1, '/Square', rect,
                          AP=pdf_dict(N=appearance(expected, 'Annotation artwork', 160, 32)))
        # Ordinary printed words can look like PDF keys and still must remain visible.
        for current in (writer, expected):
            paint_reference(current, 3, '/JavaScript /EmbeddedFiles', [72, 370, 332, 402])
    if kind == 'broken_omitted_branches':
        branch = writer._add_object(DictionaryObject())
        branch.get_object()[name('/Kids')] = ArrayObject([branch])
        names[name('/JavaScript')] = branch
        names[name('/EmbeddedFiles')] = TextStringObject('Unreadable attachment name tree')
        writer.pages[0][name('/AA')] = branch
    if kind == 'payload_aliases':
        writer.pages[0]['/Resources'][name('/FixtureAlias')] = ArrayObject([script, script_text, file_spec])
        rect = [72, 370, 232, 402]
        source_ap = appearance(writer, 'Annotation artwork', 160, 32)
        source_ap.get_object()['/Resources'][name('/FixtureAlias')] = file_spec.get_object()['/EF'].raw_get('/F')
        attach_annotation(writer, 2, '/Square', rect, AP=pdf_dict(N=source_ap))
        attach_annotation(expected, 2, '/Square', rect,
                          AP=pdf_dict(N=appearance(expected, 'Annotation artwork', 160, 32)))
    if widget is not None:
        widget.get_object()[name('/AA')] = pdf_dict(K=script, V=pdf_dict(S=name('/SubmitForm')),
                                                    F=pdf_dict(S=name('/ResetForm')))
        form(writer)[name('/CO')] = ArrayObject([widget])
        writer.root_object[name('/AF')] = ArrayObject([file_spec])


def make_case(kind: str):
    w, expected = base_writer(), base_writer()
    texts, regions = {}, []
    def text(page=1, title='value', value='FORM-VALUE-42', apvalue=None, missing=False, orphan=False):
        rect = [72, 500, 332, 532]
        ref = add_widget(w, page, '/Tx', title, TextStringObject(value), rect,
            None if missing else appearance(w, value if apvalue is None else apvalue), register=not orphan)
        if orphan:
            ref.get_object()[name('/DA')] = TextStringObject('/F0 12 Tf 0 g')
        paint_reference(expected, page, value, rect)
        texts.setdefault(str(page + 1), []).append(value)
        regions.append({'page': page+1, 'rect_pdf': rect})
        return ref

    if kind == 'acroform_null':
        w.root_object[name('/AcroForm')] = NullObject()
    elif kind == 'acroform_empty':
        w.root_object[name('/AcroForm')] = w._add_object(DictionaryObject())
    elif kind == 'fields_empty':
        w.root_object[name('/AcroForm')] = w._add_object(DictionaryObject({name('/Fields'): ArrayObject()}))
    elif kind == 'resource_only':
        f = form(w); del f['/Fields']
    elif kind in ('text', 'missing_appearance', 'stale_appearance', 'orphan_widget',
                  'rotated_crop', 'form_action', 'encrypted_form'):
        ref = text(value='CURRENT-VALUE' if kind == 'stale_appearance' else 'FORM-VALUE-42',
                   apvalue='STALE-VALUE' if kind == 'stale_appearance' else None,
                   missing=kind == 'missing_appearance', orphan=kind == 'orphan_widget')
        if kind in ('missing_appearance', 'stale_appearance'):
            form(w)[name('/NeedAppearances')] = BooleanObject(True)
        if kind == 'rotated_crop':
            for writer in (w, expected):
                writer.pages[1][name('/CropBox')] = arr([20, 30, 592, 762])
                writer.pages[1].rotate(90)
        if kind == 'form_action':
            ref.get_object()[name('/AA')] = DictionaryObject({name('/K'): DictionaryObject({
                name('/S'): name('/JavaScript'), name('/JS'): TextStringObject(f'/* {FORM_ACTION_MARKER} */')})})
        if kind == 'encrypted_form':
            w.encrypt(user_password='', owner_password='synthetic-owner', algorithm='RC4-40')
    elif kind == 'controls':
        for page, checked in [(1, True), (2, False)]:
            rect = [72, 492, 94, 514]
            states = DictionaryObject({name('/Off'): appearance(w, width=22, height=22, mark='off'),
                                      name('/Chosen'): appearance(w, width=22, height=22, mark='on')})
            state = '/Chosen' if checked else '/Off'
            add_widget(w, page, '/Btn', f'check{page}', name(state), rect, states, state=state)
            paint_reference(expected, page, '', rect, 'on' if checked else 'off')
            regions.append({'page':page+1, 'rect_pdf':rect})
        parent_obj = DictionaryObject({name('/FT'): name('/Btn'), name('/T'): TextStringObject('radio'),
            name('/Ff'): NumberObject(1 << 15), name('/V'): name('/StateB'), name('/Kids'): ArrayObject()})
        parent = w._add_object(parent_obj); form(w)['/Fields'].append(parent)
        for page, state in [(1, '/StateA'), (2, '/StateB')]:
            rect = [120, 492, 142, 514]
            aps = DictionaryObject({name('/Off'): appearance(w, width=22, height=22, mark='off'),
                                   name(state): appearance(w, width=22, height=22, mark='on')})
            ref=add_widget(w, page, None, None, None, rect, aps, parent=parent,
                           state=state if state == '/StateB' else '/Off', register=False)
            parent_obj['/Kids'].append(ref)
            paint_reference(expected, page, '', rect, 'on' if state == '/StateB' else 'off')
            regions.append({'page':page+1, 'rect_pdf':rect})
        rect = [72, 420, 332, 452]
        ref=add_widget(w, 1, '/Ch', 'city', TextStringObject('export-berlin'), rect, appearance(w,'Berlin'))
        ref.get_object()[name('/Ff')]=NumberObject(1 << 17)
        ref.get_object()[name('/Opt')]=ArrayObject([ArrayObject([TextStringObject('export-berlin'),TextStringObject('Berlin')]),ArrayObject([TextStringObject('export-turku'),TextStringObject('Turku')])])
        paint_reference(expected,1,'Berlin',rect); texts['2']=['Berlin']; regions.append({'page':2,'rect_pdf':rect})
    elif kind == 'shared_field':
        obj=DictionaryObject({name('/FT'):name('/Tx'),name('/T'):TextStringObject('shared'),
            name('/V'):TextStringObject('SHARED-VALUE'),name('/Kids'):ArrayObject(),
            name('/DA'):TextStringObject('/F0 12 Tf 0 g')})
        parent=w._add_object(obj);form(w)['/Fields'].append(parent)
        for page in (1,2):
            rect=[72,500,332,532]
            ref=add_widget(w,page,None,None,None,rect,appearance(w,'SHARED-VALUE'),parent=parent,register=False)
            obj['/Kids'].append(ref);paint_reference(expected,page,'SHARED-VALUE',rect)
            texts[str(page+1)]=['SHARED-VALUE'];regions.append({'page':page+1,'rect_pdf':rect})
    elif kind == 'multiline_readonly':
        rect = [72, 474, 332, 536]
        lines = ['MULTILINE-FIRST-57', 'MULTILINE-SECOND-38']
        ref = add_widget(w, 1, '/Tx', 'read-only-lines', TextStringObject('\n'.join(lines)), rect)
        ref.get_object()[name('/Ff')] = NumberObject((1 << 12) | 1)
        paint_reference(expected, 1, '', rect, commands=
            f'q 1 g 0 0 260 62 re f 0 G 1 w 0.5 0.5 259 61 re S Q\n'
            f'BT /F0 12 Tf 0 g 8 42 Td ({lines[0]}) Tj ET\n'
            f'BT /F0 12 Tf 0 g 8 27.6 Td ({lines[1]}) Tj ET\n')
        texts['2'] = lines
        regions.append({'page': 2, 'rect_pdf': rect})
        blank = [72, 420, 332, 452]
        ref = add_widget(w, 1, '/Tx', 'read-only-blank', TextStringObject(''), blank, appearance(w))
        ref.get_object()[name('/Ff')] = NumberObject(1)
        paint_reference(expected, 1, '', blank)
        regions.append({'page': 2, 'rect_pdf': blank})
    elif kind in ('multiselect', 'multiselect_missing_appearance'):
        rect = [72, 462, 332, 532]
        labels = ['List Alpha', 'List Beta', 'List Gamma']
        exports = ['export-alpha', 'export-beta', 'export-gamma']
        if kind == 'multiselect':
            marks = ('0.15 0.3 0.7 rg 2 48 256 17 re f 2 14 256 17 re f '
                     '1 g BT /F0 12 Tf 8 52 Td (List Alpha) Tj ET '
                     '0 g BT /F0 12 Tf 8 35 Td (List Beta) Tj ET '
                     '1 g BT /F0 12 Tf 8 18 Td (List Gamma) Tj ET')
        else:
            marks = ('0.75 0.85 1 rg 2 48 256 14.4 re f 2 19.2 256 14.4 re f '
                     '0 g BT /F0 12 Tf 8 50 Td (List Alpha) Tj '
                     '0 -14.4 Td (List Beta) Tj 0 -14.4 Td (List Gamma) Tj ET')
        drawing = 'q 1 g 0 0 260 70 re f 0 G 0.5 0.5 259 69 re S ' + marks + ' Q\n'
        art = None
        if kind == 'multiselect':
            art = appearance(w, width=260, height=70)
            art.get_object().set_data(drawing.encode('ascii'))
        ref = add_widget(w, 1, '/Ch', 'select-many', ArrayObject([TextStringObject(exports[index]) for index in (0, 2)]),
                         rect, art)
        ref.get_object().update(pdf_dict(Ff=NumberObject(1 << 21), I=ArrayObject([NumberObject(0), NumberObject(2)]),
            Opt=ArrayObject([ArrayObject([TextStringObject(key), TextStringObject(label)])
                             for key, label in zip(exports, labels)])))
        add_static_drawing(expected, 1, 'q 1 0 0 1 72 462 cm ' + drawing.replace('/F0', '/FStatic') + 'Q\n')
        texts['2'] = labels
        regions.append({'page': 2, 'rect_pdf': rect})
    elif kind == 'combo_missing_appearance':
        rect = [72, 500, 332, 532]
        ref = add_widget(w, 1, '/Ch', 'city', TextStringObject('export-turku'), rect)
        ref.get_object().update(pdf_dict(Ff=NumberObject(1 << 17),
            Opt=ArrayObject([ArrayObject([TextStringObject('export-berlin'), TextStringObject('Berlin')]),
                             ArrayObject([TextStringObject('export-turku'), TextStringObject('Turku')])])))
        paint_reference(expected, 1, 'Turku', rect)
        texts['2'] = ['Turku']
        regions.append({'page': 2, 'rect_pdf': rect})
    elif kind == 'winansi_missing_appearance':
        rect = [72, 500, 332, 532]
        label = 'Mikä Café'
        add_widget(w, 1, '/Tx', 'winansi', TextStringObject(label), rect)
        form(w)['/DR']['/Font']['/F0'][name('/Encoding')] = name('/WinAnsiEncoding')
        add_static_drawing(expected, 1, f'q 1 g 72 500 260 32 re f 0 G 1 w 72.5 500.5 259 31 re S '
                                       f'BT /FStatic 12 Tf 0 g 80 510 Td ({winansi_literal(label)}) Tj ET Q\n')
        expected.pages[1]['/Resources']['/Font']['/FStatic'][name('/Encoding')] = name('/WinAnsiEncoding')
        texts['2'] = [label]
        regions.append({'page': 2, 'rect_pdf': rect})
    elif kind == 'hierarchy':
        for page, branch in ((1, 'north'), (2, 'south')):
            label = f'BRANCH-{branch.upper()}-19'
            rect = [72, 500, 332, 532]
            node = pdf_dict(FT=name('/Tx'), T=TextStringObject(branch), DA=TextStringObject('/F0 12 Tf 0 g'),
                            Kids=ArrayObject())
            parent = w._add_object(node)
            form(w)['/Fields'].append(parent)
            widget = add_widget(w, page, None, 'same-leaf', TextStringObject(label), rect,
                                appearance(w, label), parent=parent, register=False)
            node['/Kids'].append(widget)
            paint_reference(expected, page, label, rect)
            texts[str(page+1)] = [label]
            regions.append({'page': page+1, 'rect_pdf': rect})
    elif kind == 'transformed_appearance':
        label = 'TRANSFORMED-VALUE-84'
        rect = [72, 500, 332, 532]
        art = appearance(w, label)
        stream = art.get_object()
        stream.set_data(b'q 0.5 0 0 0.5 10 20 cm\n' + stream.get_data() + b'Q\n')
        stream.update(pdf_dict(BBox=arr([10, 20, 140, 36]), Matrix=arr([2, 0, 0, 2, -10, 15])))
        add_widget(w, 1, '/Tx', 'transformed', TextStringObject(label), rect, art)
        paint_reference(expected, 1, label, rect)
        for writer in (w, expected):
            page = writer.pages[1]
            page['/Resources']['/Font'][name('/F0')] = writer._add_object(
                pdf_dict(Type=name('/Font'), Subtype=name('/Type1'), BaseFont=name('/Courier')))
            add_static_drawing(writer, 1, 'BT /F0 10 Tf 72 448 Td (PAGE-RESOURCE-COLLISION) Tj ET\n')
        texts['2'] = [label, 'PAGE-RESOURCE-COLLISION']
        regions.append({'page': 2, 'rect_pdf': rect})
    elif kind == 'hidden_fields':
        for page, flag, value in ((1, 2, PRIVATE_FORM_MARKERS[kind][0]), (2, 32, PRIVATE_FORM_MARKERS[kind][1])):
            saved = w._add_object(TextStringObject(value))
            ap = appearance(w, value)
            ref = add_widget(w, page, '/Tx', f'hidden-{page}', saved, [72, 500, 332, 532], ap)
            ref.get_object()[name('/F')] = NumberObject(flag)
            w.pages[0]['/Resources'][name(f'/HiddenValueAlias{page}')] = saved
            w.pages[0]['/Resources'][name(f'/HiddenAppearanceAlias{page}')] = ap
        detached = w._add_object(TextStringObject(PRIVATE_FORM_MARKERS[kind][2]))
        form(w)['/Fields'].append(w._add_object(pdf_dict(FT=name('/Tx'), T=TextStringObject('detached'), V=detached)))
        w.pages[0]['/Resources'][name('/DetachedValueAlias')] = detached
        text(page=3, title='visible-public', value='VISIBLE-FORM-VALUE-11')
    elif kind == 'button_stale_state':
        rect = [72, 492, 94, 514]
        states = pdf_dict(Off=appearance(w, width=22, height=22, mark='off'),
                          Chosen=appearance(w, width=22, height=22, mark='on'))
        add_widget(w, 1, '/Btn', 'stale-state', name('/Chosen'), rect, states, state='/Off')
        paint_reference(expected, 1, '', rect, 'on')
        regions.append({'page': 2, 'rect_pdf': rect})
    elif kind in ('xfa_null','xfa_empty_array','xfa_detection_only'):
        f=form(w)
        if kind == 'xfa_null':
            f[name('/XFA')]=NullObject()
        elif kind == 'xfa_empty_array':
            f[name('/XFA')]=ArrayObject()
        else:
            s=DecodedStreamObject();s.set_data(b'<?xml version="1.0"?><xdp:xdp xmlns:xdp="http://ns.adobe.com/xdp/"><test>DETECTION SENTINEL ONLY - NOT A RENDERER FIXTURE</test></xdp:xdp>')
            f[name('/XFA')]=w._add_object(s)
    elif kind == 'fields_malformed':
        form(w)[name('/Fields')]=TextStringObject('not-an-array')
    elif kind == 'field_cycle':
        obj=DictionaryObject({name('/T'):TextStringObject('cycle'),name('/Kids'):ArrayObject()})
        ref=w._add_object(obj);obj['/Kids'].append(ref);form(w)['/Fields'].append(ref)
    elif kind in ('signature_structure', 'signature_visible'):
        # These are structural fixtures, without a cryptographically valid signature.
        sig=DictionaryObject({name('/Type'):name('/Sig'),name('/Filter'):name('/Adobe.PPKLite'),
            name('/SubFilter'):name('/adbe.pkcs7.detached'),name('/ByteRange'):arr([0,1,2,3]),
            name('/Contents'):TextStringObject('SYNTHETIC-NOT-A-REAL-SIGNATURE')})
        rect = [72, 500, 332, 532]
        label = 'VISIBLE-SIGNATURE-ARTWORK'
        add_widget(w, 1, '/Sig', 'signature', w._add_object(sig), rect,
                   appearance(w, label) if kind == 'signature_visible' else None)
        if kind == 'signature_visible':
            paint_reference(expected, 1, label, rect)
            texts['2'] = [label]
            regions.append({'page': 2, 'rect_pdf': rect})
    elif kind == 'missing_appearance_font':
        ref = text(page=3, missing=True)
        ref.get_object()[name('/DA')] = TextStringObject('/UnavailableFont 12 Tf 0 g')
    elif kind in FEATURE_CASES:
        widget = text() if kind in ('form_and_features', 'encrypted_form_and_features') else None
        if kind == 'encrypted_metadata_and_features':
            form(w)
        add_optional_features(w, expected, kind, widget)
        if kind in ('encrypted_metadata_and_features', 'encrypted_form_and_features'):
            w.encrypt(user_password='', owner_password='synthetic-owner', algorithm='RC4-40')
    elif kind != 'plain':
        raise ValueError(kind)
    return w, expected, texts, regions


def generate(output: Path) -> dict:
    output=output.absolute()
    output.mkdir(parents=True, exist_ok=False)
    cases=[]
    for number,kind in enumerate(CASES,1):
        w, ref, texts, regions=make_case(kind)
        basename=f'{number:02d}-{kind}'
        source=output/(basename+'.pdf')
        with source.open('xb') as stream:
            w.write(stream)
        golden=None
        if kind not in NEGATIVE:
            golden=output/(basename+'.expected-static.pdf')
            with golden.open('xb') as stream:
                ref.write(stream)
        cases.append({'id':f'F{number:02d}', 'kind':kind, 'source':source.name,
            'source_sha256':sha256(source.read_bytes()).hexdigest(),
            'expected_static':None if golden is None else golden.name,
            'expected_sha256':None if golden is None else sha256(golden.read_bytes()).hexdigest(),
            'expected_behavior':'specific_failure' if kind in NEGATIVE else 'static_split_success',
            'page_count':4,'auto1_ranges_zero_based':[[0,1],[1,2],[2,3],[3,4]],
            'visible_text_by_page':texts, 'field_regions':regions,
            'notes':'Detection sentinel only; not evidence of genuine XFA rendering.' if kind=='xfa_detection_only' else
                    'Structural signature sentinel; not cryptographically signed.' if kind=='signature_structure' else ''})
    manifest={'schema_version':1,'purpose':'test-input seeds, not application acceptance',
        'original_content':True,'no_private_documents':True,'cases':cases}
    (output/'cases.json').write_text(json.dumps(manifest,indent=2)+'\n',encoding='utf-8')
    return manifest


def main() -> int:
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output',type=Path,required=True)
    args=parser.parse_args()
    try:
        manifest=generate(args.output)
    except (OSError,ValueError) as error:
        parser.exit(2,f'Fixture generation failed: {error}\n')
    print(f'Created {len(manifest["cases"])} original input cases in {args.output}')
    return 0

if __name__=='__main__':
    raise SystemExit(main())
