"""Original inert PDF structures for unsupported-document policy testing.

Signature cases contain authored signature dictionaries, not valid signatures.
Actions and attachments contain harmless original markers and are never opened.
"""
from hashlib import sha256
from io import BytesIO
import json
from pathlib import Path

from pypdf import PdfWriter
from pypdf.generic import (ArrayObject, BooleanObject, ByteStringObject, DecodedStreamObject,
                          DictionaryObject, NameObject, NullObject, NumberObject, TextStringObject)

KINDS = ('ordinary', 'page-aa-exclusion', 'encrypted-user', 'encrypted-owner-empty',
         'acroform', 'xfa', 'signature-widget', 'signature-perms', 'orphan-widget', 'orphan-signature',
         'catalog-openaction', 'catalog-aa', 'catalog-javascript', 'embedded-attachment', 'portfolio',
         'page-fileattachment', 'page-3d', 'rich-media', 'catalog-associated-file', 'page-associated-file',
         'malformed-header', 'truncated', 'zero-pages', 'deep-outline', 'page-tree-cycle')
UNSUPPORTED = KINDS[2:20]
MALFORMED = KINDS[20:]
# Authored fixture credentials are generation inputs only, never app arguments.
USER_PASSWORD = 'WBS-authored-test-open-731'
OWNER_PASSWORD = 'WBS-authored-test-owner-946'


def name(value):
    return NameObject(value)


def text(value):
    return TextStringObject(value)


def array(values):
    return ArrayObject([NumberObject(value) for value in values])


def dictionary(values):
    return DictionaryObject({name(key): value for key, value in values.items()})


def stream(writer, data, fields=None):
    value = DecodedStreamObject()
    value.set_data(data)
    for key, item in (fields or {}).items():
        value[name(key)] = item
    return writer._add_object(value)


def base_writer(pages=4):
    writer = PdfWriter()
    font = writer._add_object(dictionary({'/Type': name('/Font'), '/Subtype': name('/Type1'), '/BaseFont': name('/Helvetica')}))
    for index in range(pages):
        page = writer.add_blank_page(width=120, height=120)
        page[name('/WBSFixturePage')] = NumberObject(index + 1)
        page[name('/Resources')] = dictionary({'/Font': dictionary({'/F1': font})})
        page[name('/Contents')] = stream(writer, f'BT /F1 8 Tf 12 96 Td (WBS-POLICY-PAGE-{index + 1:03d}) Tj ET\n'.encode())
    if pages:
        first = writer.add_outline_item('Authored first policy parent', 0)
        writer.add_outline_item('Authored first child', 1, parent=first)
        second = writer.add_outline_item('Authored second policy parent', 2)
        writer.add_outline_item('Authored second child', 3, parent=second)
    writer.add_metadata({'/Title': 'Original generated document policy fixture', '/Author': 'Original authored fixture author'})
    return writer


def inert_action():
    return dictionary({'/S': name('/JavaScript'), '/JS': text('/* ORIGINAL_INERT_POLICY_ACTION */ void 0;')})


def signature(writer):
    return writer._add_object(dictionary({'/Type': name('/Sig'), '/Filter': name('/Adobe.PPKLite'),
        '/SubFilter': name('/adbe.pkcs7.detached'), '/ByteRange': array([0, 0, 0, 0]),
        '/Contents': ByteStringObject(b'\0' * 32), '/Reason': text('Authored structural test; not a cryptographic signature')}))


def widget(writer, signed=False):
    value = dictionary({'/Type': name('/Annot'), '/Subtype': name('/Widget'), '/Rect': array([10, 10, 80, 30]),
        '/FT': name('/Sig' if signed else '/Tx'), '/T': text('Authored policy field'), '/P': writer.pages[0].indirect_reference})
    value[name('/V')] = signature(writer) if signed else text('Authored field value')
    reference = writer._add_object(value)
    writer.pages[0][name('/Annots')] = ArrayObject([reference])
    return reference


def file_spec(writer):
    embedded = stream(writer, b'ORIGINAL_INERT_POLICY_ATTACHMENT\n', {'/Type': name('/EmbeddedFile')})
    return writer._add_object(dictionary({'/Type': name('/Filespec'), '/F': text('authored-attachment.txt'),
        '/UF': text('authored-attachment.txt'), '/EF': dictionary({'/F': embedded}), '/AFRelationship': name('/Data')}))


def fixture_bytes(kind):
    if kind not in KINDS:
        raise ValueError('Unknown authored policy fixture')
    if kind == 'malformed-header':
        return b'Authored invalid file, no PDF header/trailer/pages.\n'
    writer = base_writer(0 if kind == 'zero-pages' else 4)
    catalog = writer._root_object
    if kind in {'encrypted-user', 'encrypted-owner-empty'}:
        writer.encrypt(USER_PASSWORD if kind == 'encrypted-user' else '', OWNER_PASSWORD, algorithm='RC4-128')
    elif kind in {'acroform', 'signature-widget'}:
        field = widget(writer, kind == 'signature-widget')
        catalog[name('/AcroForm')] = writer._add_object(dictionary({'/Fields': ArrayObject([field]), '/NeedAppearances': BooleanObject(False)}))
    elif kind == 'xfa':
        xfa = stream(writer, b'<xfa>ORIGINAL_INERT_POLICY_XFA</xfa>')
        catalog[name('/AcroForm')] = writer._add_object(dictionary({'/Fields': ArrayObject(), '/XFA': xfa}))
    elif kind == 'signature-perms':
        catalog[name('/Perms')] = dictionary({'/DocMDP': signature(writer)})
    elif kind in {'orphan-widget', 'orphan-signature'}:
        widget(writer, kind == 'orphan-signature')
    elif kind == 'catalog-openaction':
        catalog[name('/OpenAction')] = inert_action()
    elif kind == 'catalog-aa':
        catalog[name('/AA')] = dictionary({'/WC': inert_action()})
    elif kind == 'catalog-javascript':
        writer.add_js('/* ORIGINAL_INERT_POLICY_ACTION */ void 0;')
    elif kind == 'embedded-attachment':
        writer.add_attachment('authored-attachment.txt', b'ORIGINAL_INERT_POLICY_ATTACHMENT\n')
    elif kind == 'portfolio':
        catalog[name('/Collection')] = dictionary({'/Type': name('/Collection'), '/View': name('/D')})
    elif kind == 'catalog-associated-file':
        catalog[name('/AF')] = ArrayObject([file_spec(writer)])
    elif kind == 'page-associated-file':
        writer.pages[0][name('/AF')] = ArrayObject([file_spec(writer)])
    elif kind == 'page-fileattachment':
        annotation = dictionary({'/Type': name('/Annot'), '/Subtype': name('/FileAttachment'),
            '/Rect': array([10, 10, 30, 30]), '/FS': file_spec(writer), '/P': writer.pages[0].indirect_reference})
        writer.pages[0][name('/Annots')] = ArrayObject([writer._add_object(annotation)])
    elif kind in {'page-3d', 'rich-media'}:
        annotation = dictionary({'/Type': name('/Annot'), '/Subtype': name('/3D' if kind == 'page-3d' else '/RichMedia'),
            '/Rect': array([10, 10, 30, 30]), '/P': writer.pages[0].indirect_reference})
        if kind == 'page-3d':
            annotation[name('/3DD')] = stream(writer, b'ORIGINAL_INERT_POLICY_3D')
        else:
            annotation[name('/RichMediaContent')] = dictionary({'/Assets': dictionary({'/Names': ArrayObject()})})
        writer.pages[0][name('/Annots')] = ArrayObject([writer._add_object(annotation)])
    elif kind == 'page-aa-exclusion':
        writer.pages[0][name('/AA')] = dictionary({'/O': inert_action()})
    elif kind == 'deep-outline':
        outline = dictionary({'/Type': name('/Outlines'), '/Count': NumberObject(1100)})
        root = previous = writer._add_object(outline)
        for index in range(1100):
            node = dictionary({'/Title': text('Authored deep outline ' + str(index)), '/Parent': previous,
                '/Dest': ArrayObject([writer.pages[0].indirect_reference, name('/Fit')]), '/Count': NumberObject(1)})
            current = writer._add_object(node)
            previous.get_object()[name('/First')] = current
            previous.get_object()[name('/Last')] = current
            previous = current
        catalog[name('/Outlines')] = root
    elif kind == 'page-tree-cycle':
        writer._pages.get_object()[name('/Kids')] = ArrayObject([writer._pages])
        writer._pages.get_object()[name('/Count')] = NumberObject(1)
    output = BytesIO()
    writer.write(output)
    data = output.getvalue()
    return data[:64] if kind == 'truncated' else data


def generate_fixtures(directory):
    directory = Path(directory)
    directory.mkdir(exist_ok=False)
    observations, files = {}, {}
    for kind in KINDS:
        data = fixture_bytes(kind)
        path = directory / (kind + '.pdf')
        with path.open('xb') as stream_file:
            stream_file.write(data)
        files[kind] = path
        observations[kind] = {'filename': path.name, 'sha256': sha256(data).hexdigest(), 'bytes': len(data),
            'expected': 'unsupported_document' if kind in UNSUPPORTED else 'bounded_invalid' if kind in MALFORMED else 'ordinary_or_defined_exclusion',
            'encrypted': kind.startswith('encrypted-'), 'structural_signature_only': kind in {'signature-widget', 'signature-perms', 'orphan-signature'},
            'source_actions_inert': True, 'fixture_private_data': False}
    provenance = {'protocol': 'winbooksplit.document-policy-fixtures', 'version': 1, 'private_data': False,
        'license': 'Original MIT-authored generated test data', 'generator_sha256': sha256(Path(__file__).read_bytes()).hexdigest(),
        'pdf_file_count': len(files), 'fixtures': observations,
        'limits': ['Signature structure detection only, no cryptographic validity proof.',
                   'No fixture action/attachment/3D/XFA/media is executed or opened.']}
    with (directory / 'provenance.json').open('x', encoding='utf-8', newline='\n') as stream_file:
        json.dump(provenance, stream_file, ensure_ascii=True, indent=2)
        stream_file.write('\n')
    return files, provenance
