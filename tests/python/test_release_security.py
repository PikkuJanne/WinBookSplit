"""Bounded dependency guards and actual BAT shell-shadowing regression."""
from io import BytesIO
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest

from pypdf import PdfReader, PdfWriter, apply_configuration, get_configuration
from pypdf.errors import LimitReachedError, PdfReadError
from pypdf.filters import decode_stream_data
from pypdf.generic import (ArrayObject, ContentStream, DecodedStreamObject,
                           DictionaryObject, EncodedStreamObject, NameObject, TextStringObject)

ROOT = Path(__file__).resolve().parents[2]


class BoundedReadStream(BytesIO):
    """Stop a broken parser after finite tiny reads, without a hanging child."""
    def __init__(self, data):
        if len(data) > 2048:
            raise ValueError('Only tiny authored parser inputs are allowed')
        super().__init__(data)
        self.remaining_reads = 4096

    def read(self, size=-1):
        self.remaining_reads -= 1
        if self.remaining_reads < 0:
            raise AssertionError('Authored parser read budget exhausted')
        return super().read(size)


class BoundedCloneObjects(list):
    """Bound destination growth while exercising the real dictionary cloner."""
    def append(self, value):
        if len(self) >= 32:
            raise AssertionError('Authored dictionary clone object budget exhausted')
        super().append(value)


def tiny_shared_page_pdf():
    """Three physical references copy five entries each, with no large resource."""
    objects = (
        b'<< /Type /Catalog /Pages 2 0 R >>',
        b'<< /Type /Pages /Count 3 /Kids [3 0 R 3 0 R 3 0 R] /MediaBox [0 0 20 20] /Resources << >> >>',
        b'<< /Type /Page /Parent 2 0 R /Marker (authored) >>',
    )
    result = bytearray(b'%PDF-1.4\n')
    offsets = []
    for number, body in enumerate(objects, 1):
        offsets.append(len(result))
        result.extend(str(number).encode('ascii') + b' 0 obj\n' + body + b'\nendobj\n')
    start = len(result)
    result.extend(b'xref\n0 4\n0000000000 65535 f \n')
    for offset in offsets:
        result.extend(f'{offset:010} 00000 n \n'.encode('ascii'))
    result.extend(f'trailer\n<< /Root 1 0 R /Size 4 >>\nstartxref\n{start}\n%%EOF\n'.encode('ascii'))
    if len(result) > 1024:
        raise AssertionError('Authored shared page fixture must remain below 1 KiB')
    return bytes(result)


def small_xref_stream_pdf(counts):
    """Authored parser fixtures only: at most two streams/four tiny entries each."""
    if not 1 <= len(counts) <= 2 or any(type(count) is not int or not 1 <= count <= 4 for count in counts):
        raise ValueError('Only bounded positive xref entry counts are test inputs')
    data = bytearray(b'%PDF-1.5\n1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n'
                     b'2 0 obj\n<< /Type /Pages /Kids [] /Count 0 >>\nendobj\n')
    previous = None
    for number, count in enumerate(counts, 1):
        offset = len(data)
        first = number * 10
        fields = (f'/Type /XRef /Size {first + count} /Root 1 0 R '
                  f'/Index [{first} {count}] /W [1 0 0] /Length {count}')
        if previous is not None:
            fields += f' /Prev {previous}'
        data.extend(f'{90 + number} 0 obj\n<< {fields} >>\nstream\n'.encode('ascii'))
        data.extend(b'\x02' * count)
        data.extend(b'\nendstream\nendobj\n')
        previous = offset
    data.extend(f'startxref\n{previous}\n%%EOF\n'.encode('ascii'))
    if len(data) >= 1024:
        raise AssertionError('The authored xref fixture must remain smaller than 1 KiB')
    return bytes(data), previous


class StreamResourceGuardTests(unittest.TestCase):
    def test_standard_xref_missing_trailer_rejects_before_finite_read_watchdog(self):
        # Upstream security fix: https://github.com/py-pdf/pypdf/pull/4289
        # This tiny incomplete section reproduces the EOF loop, while the
        # watchdog makes the old dependency fail without consuming unbounded work.
        prefix = b'%PDF-1.4\nstartxref\n28\n%%EOF\n'
        self.assertEqual(len(prefix), 28)
        stream = BoundedReadStream(prefix + b'xref\n1 0\nend')
        with self.assertRaisesRegex(PdfReadError, r'Xref parsing stuck at position'):
            PdfReader(stream)
        self.assertGreater(stream.remaining_reads, 0)
        control = PdfReader(BoundedReadStream(tiny_shared_page_pdf()))
        self.assertEqual(len(control.pages), 3)

    def test_page_tree_copied_entries_have_a_small_document_budget_boundary(self):
        # Upstream security fix: https://github.com/py-pdf/pypdf/pull/4290
        original = getattr(get_configuration(), 'page_tree_maximum_page_entries', None)
        self.assertEqual(original, 2_000_000)
        document = tiny_shared_page_pdf()
        with apply_configuration(page_tree_maximum_page_entries=15):
            pages = list(PdfReader(BytesIO(document)).pages)
            self.assertEqual(len(pages), 3)
            self.assertTrue(all(set(page) == {'/Type', '/Parent', '/Marker', '/MediaBox', '/Resources'}
                                and page['/Marker'] == 'authored' for page in pages))
        with apply_configuration(page_tree_maximum_page_entries=14):
            reader = PdfReader(BytesIO(document))
            with self.assertRaisesRegex(LimitReachedError, 'Maximum page tree page entry limit exceeded'):
                len(reader.pages)
        self.assertEqual(get_configuration().page_tree_maximum_page_entries, original)
        self.assertEqual(len(PdfReader(BytesIO(document)).pages), 3)

    def test_dictionary_clone_tracks_source_cycles_with_bounded_destination(self):
        # Upstream security fix: https://github.com/py-pdf/pypdf/pull/4288
        source = PdfWriter()
        nodes = [DictionaryObject({NameObject('/T'): TextStringObject(label)})
                 for label in ('entry', 'second', 'third')]
        references = [source._add_object(node) for node in nodes]
        nodes[0][NameObject('/N')] = references[1]
        nodes[1][NameObject('/N')] = references[2]
        nodes[2][NameObject('/N')] = references[1]
        destination = PdfWriter()
        destination._objects = BoundedCloneObjects(destination._objects)
        cloned = nodes[0].clone(destination, force_duplicate=True)
        self.assertIsInstance(cloned, DictionaryObject)
        self.assertEqual(cloned['/T'], 'entry')
        self.assertEqual(cloned['/N']['/T'], 'second')
        self.assertLess(len(destination._objects), 32)
        self.assertEqual([node['/T'] for node in nodes], ['entry', 'second', 'third'])
        self.assertEqual([node.raw_get('/N') for node in nodes],
                         [references[1], references[2], references[1]])

    def test_array_parser_bounds_elements_and_restores_configuration(self):
        original = get_configuration().array_maximum_element_count
        self.assertEqual(original, 100_000)
        with apply_configuration(array_maximum_element_count=3):
            self.assertEqual(len(ArrayObject.read_from_stream(BytesIO(b'[0 1 2]'), None)), 3)
            with self.assertRaisesRegex(LimitReachedError, 'maximum array element count'):
                ArrayObject.read_from_stream(BytesIO(b'[0 1 2 3]'), None)
        self.assertEqual(get_configuration().array_maximum_element_count, original)

    def test_dictionary_parser_bounds_entries_and_restores_configuration(self):
        original = get_configuration().dictionary_maximum_element_count
        self.assertEqual(original, 100_000)
        with apply_configuration(dictionary_maximum_element_count=3):
            self.assertEqual(len(DictionaryObject.read_from_stream(
                BytesIO(b'<</A 0 /B 1 /C 2>>'), None)), 3)
            with self.assertRaisesRegex(LimitReachedError, 'maximum number of dictionary entries'):
                DictionaryObject.read_from_stream(BytesIO(b'<</A 0 /B 1 /C 2 /D 3>>'), None)
        self.assertEqual(get_configuration().dictionary_maximum_element_count, original)

    def test_content_parser_bounds_tokens_and_restores_configuration(self):
        original = get_configuration().stream_parsing_work_maximum_count
        self.assertEqual(original, 2_000_000)
        def operations(data):
            stream = DecodedStreamObject()
            stream.set_data(data)
            return ContentStream(stream, None).operations
        with apply_configuration(stream_parsing_work_maximum_count=3):
            self.assertEqual(len(operations(b'q Q q')), 3)
            with self.assertRaisesRegex(LimitReachedError, 'maximum number of tokens per stream'):
                operations(b'q Q q Q')
        self.assertEqual(get_configuration().stream_parsing_work_maximum_count, original)

    def test_xref_stream_entry_budget_has_small_constructor_boundary_controls(self):
        at_limit, _ = small_xref_stream_pdf((3,))
        over_limit, _ = small_xref_stream_pdf((4,))
        for strict in (False, True):
            with self.subTest(strict=strict):
                reader = PdfReader(BytesIO(at_limit), strict=strict, xref_stream_maximum_entries=3)
                self.assertEqual(sorted(reader.xref_objStm), [10, 11, 12])
                self.assertEqual(reader._xref_stream_entry_count, 3)
                with self.assertRaisesRegex(PdfReadError, r'Maximum XRef stream entry limit reached: 4 > 3'):
                    PdfReader(BytesIO(over_limit), strict=strict, xref_stream_maximum_entries=3)

    def test_xref_stream_document_budget_accumulates_before_inserting_entries(self):
        first, _ = small_xref_stream_pdf((2,))
        second, offset = small_xref_stream_pdf((2, 2))
        reader = PdfReader(BytesIO(first))
        self.assertEqual(sorted(reader.xref_objStm), [10, 11])
        # Reduce only this reader's budget. Setting the same tiny state on the
        # old dependency proves its parser ignores the missing document guard.
        reader._xref_stream_maximum_entries = 3
        reader._xref_stream_entry_count = 2
        stream = BytesIO(second)
        stream.seek(offset + 1)
        with self.assertRaisesRegex(LimitReachedError, r'Maximum XRef stream entry limit reached: 4 > 3'):
            reader._read_pdf15_xref_stream(stream)
        self.assertEqual(sorted(reader.xref_objStm), [10, 11])

    def test_xref_stream_prev_budget_and_defaults_reset_for_each_document(self):
        document, _ = small_xref_stream_pdf((2, 2))
        with self.assertLogs('pypdf', level='WARNING') as logged:
            limited = PdfReader(BytesIO(document), xref_stream_maximum_entries=3)
        self.assertEqual(sorted(limited.xref_objStm), [20, 21])
        self.assertTrue(any('Maximum XRef stream entry limit reached: 4 > 3' in line for line in logged.output))
        boundary = PdfReader(BytesIO(document), xref_stream_maximum_entries=4)
        self.assertEqual(sorted(boundary.xref_objStm), [10, 11, 20, 21])
        self.assertEqual(boundary._xref_stream_entry_count, 4)
        with self.assertLogs('pypdf', level='WARNING') as reread:
            boundary.read(BytesIO(document))
        self.assertTrue(all('Overwriting cache for' in line for line in reread.output))
        self.assertEqual(boundary._xref_stream_entry_count, 4)
        restored = PdfReader(BytesIO(document))
        self.assertEqual(restored._xref_stream_maximum_entries, 2_000_000)
        self.assertEqual(restored._xref_stream_entry_count, 4)

    def test_many_filters_reject_without_large_allocation(self):
        stream = EncodedStreamObject()
        stream._data = b"x"
        stream[NameObject('/Filter')] = ArrayObject([NameObject('/DCTDecode')] * 17)
        with self.assertRaisesRegex(LimitReachedError, 'Maximum filter count'):
            decode_stream_data(stream)

    def test_accumulated_decoding_work_limit_stays_enabled(self):
        configuration = get_configuration()
        self.assertEqual(getattr(configuration, 'stream_filters_maximum_length', None), 16)
        self.assertEqual(getattr(configuration, 'stream_decoding_work_maximum_length', None), 300_000_000)
        stream = EncodedStreamObject()
        stream._data = b"xxx"
        stream[NameObject('/Filter')] = NameObject('/DCTDecode')
        with apply_configuration(stream_decoding_work_maximum_length=5):
            with self.assertRaisesRegex(LimitReachedError, 'accumulated'):
                decode_stream_data(stream)


@unittest.skipUnless(sys.platform == 'win32', 'Actual Windows BAT execution required')
class BatchShellTrustTests(unittest.TestCase):
    def setUp(self):
        self.owned = tempfile.TemporaryDirectory(prefix='wbs-shell-trust-')
        self.addCleanup(self.owned.cleanup)
        self.work = Path(self.owned.name)
        self.app = self.work / 'app'
        self.app.mkdir()
        shutil.copyfile(ROOT / 'WinBookSplit.bat', self.app / 'WinBookSplit.bat')
        self.marker = self.work / 'called.txt'
        self.decoy = self.work / 'decoy.txt'
        (self.app / 'WinBookSplit.ps1').write_text(
            "param([string]$InputFile)\n"
            "[IO.File]::WriteAllText($env:WBS_SHELL_TRUST_MARKER, 'inbox:' + $InputFile)\n"
            "exit 37\n", encoding='utf-8')
        self.system = Path(os.environ['SystemRoot']) / 'System32'
        self.wrapper = self.work / 'run.cmd'
        self.wrapper.write_bytes(b'@echo off\r\nsetlocal DisableDelayedExpansion\r\n'
            b'"%WBS_SHELL_TRUST_BAT%" "%WBS_SHELL_TRUST_INPUT%"\r\n')
        self.environment = os.environ.copy()
        self.environment.update(WBS_SHELL_TRUST_BAT=str(self.app / 'WinBookSplit.bat'),
            WBS_SHELL_TRUST_MARKER=str(self.marker), WBS_SHELL_TRUST_DECOY=str(self.decoy),
            WBS_SHELL_TRUST_INPUT="Literal [%UNUSED%!] O'Brien & (book) 日本語.pdf",
            PATHEXT='.COM;.EXE;.BAT;.CMD')

    def invoke(self, cwd):
        return subprocess.run([str(self.system / 'cmd.exe'), '/d', '/c', str(self.wrapper)],
            cwd=cwd, env=self.environment, stdin=subprocess.DEVNULL, capture_output=True,
            timeout=20, check=False)

    def plant_decoy(self, directory):
        directory.mkdir()
        for extension in ('cmd', 'bat'):
            (directory / ('powershell.' + extension)).write_bytes(
                b'@echo off\r\necho decoy>"%WBS_SHELL_TRUST_DECOY%"\r\nexit /b 73\r\n')

    def test_cwd_shell_decoys_do_not_execute_and_literal_argv_exit_survive(self):
        cwd = self.work / 'cwd'
        self.plant_decoy(cwd)
        self.environment['PATH'] = str(self.system)
        result = self.invoke(cwd)
        self.assertEqual(result.returncode, 37, result.stderr.decode(errors='replace'))
        self.assertFalse(self.decoy.exists())
        self.assertEqual(self.marker.read_text(encoding='utf-8'), 'inbox:' + self.environment['WBS_SHELL_TRUST_INPUT'])

    def test_path_shell_decoys_do_not_execute_for_no_input(self):
        directory = self.work / 'path'
        self.plant_decoy(directory)
        self.environment['PATH'] = str(directory) + os.pathsep + str(self.system)
        self.environment['WBS_SHELL_TRUST_INPUT'] = ''
        result = self.invoke(self.work)
        self.assertEqual(result.returncode, 37, result.stderr.decode(errors='replace'))
        self.assertFalse(self.decoy.exists())
        self.assertEqual(self.marker.read_text(encoding='utf-8'), 'inbox:')

    def test_missing_system_shell_fails_without_path_fallback(self):
        directory = self.work / 'path'
        self.plant_decoy(directory)
        self.environment['PATH'] = str(directory)
        self.environment['SystemRoot'] = str(self.work / 'absent-windows')
        result = self.invoke(self.work)
        self.assertEqual(result.returncode, 3)
        self.assertIn(b'Windows PowerShell', result.stdout)
        self.assertFalse(self.decoy.exists())
        self.assertFalse(self.marker.exists())
