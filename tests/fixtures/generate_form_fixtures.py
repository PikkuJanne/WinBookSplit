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

CASES = (
    'plain', 'acroform_null', 'acroform_empty', 'fields_empty', 'resource_only',
    'text', 'controls', 'shared_field', 'missing_appearance', 'stale_appearance',
    'orphan_widget', 'rotated_crop', 'form_action', 'xfa_null', 'xfa_empty_array',
    'xfa_detection_only', 'fields_malformed', 'field_cycle',
    'signature_structure', 'encrypted_form',
)
NEGATIVE = {'xfa_detection_only', 'fields_malformed', 'field_cycle', 'signature_structure'}


def name(s: str) -> NameObject:
    return NameObject(s)


def arr(values) -> ArrayObject:
    return ArrayObject([FloatObject(x) for x in values])


def literal(value: str) -> str:
    return value.replace('\\', '\\\\').replace('(', '\\(').replace(')', '\\)')


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
                    mark: str | None = None):
    # Independent expected static representation: direct invocation of original
    # artwork at a known fixture coordinate, no field reader or flattener used.
    page = w.pages[page_index]
    resources = page['/Resources']
    if '/XObject' not in resources:
        resources[name('/XObject')] = DictionaryObject()
    objects = resources['/XObject']
    key = name(f'/Expected{len(objects) + 1}')
    x0, y0, x1, y1 = rect
    objects[key] = appearance(w, label, x1-x0, y1-y0, mark)
    content = DecodedStreamObject()
    content.set_data(f'q 1 0 0 1 {x0} {y0} cm {key} Do Q\n'.encode('ascii'))
    extra = w._add_object(content)
    old = page.get('/Contents')
    # Retain all old streams without changing source page artwork.
    page[name('/Contents')] = ArrayObject(list(old) + [extra] if isinstance(old, ArrayObject) else [old, extra])


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
                name('/S'): name('/JavaScript'), name('/JS'): TextStringObject('/* WBS_SYNTHETIC_ACTION_DO_NOT_EXECUTE */')})})
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
    elif kind == 'signature_structure':
        # This is structural rejection coverage, not a valid cryptographic signature.
        sig=DictionaryObject({name('/Type'):name('/Sig'),name('/Filter'):name('/Adobe.PPKLite'),
            name('/SubFilter'):name('/adbe.pkcs7.detached'),name('/ByteRange'):arr([0,1,2,3]),
            name('/Contents'):TextStringObject('SYNTHETIC-NOT-A-REAL-SIGNATURE')})
        add_widget(w,1,'/Sig','signature',w._add_object(sig),[72,500,332,532])
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
