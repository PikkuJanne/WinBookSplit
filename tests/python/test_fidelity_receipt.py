"""Synthetic oracle mutations only; these do not represent native test runs."""
from copy import deepcopy
from hashlib import sha256
import importlib.util
import json
from pathlib import Path
import sys
import unittest

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location('fidelity_receipt_guard', ROOT / 'tests/fidelity/validate_fidelity_report.py')
guard = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = guard
spec.loader.exec_module(guard)


def synthetic_output():
    empty_hash = sha256(b'').hexdigest()
    process = {'command': [r'C:\synthetic\renderer.exe'], 'cwd': r'C:\synthetic', 'stdin': 'DEVNULL',
               'actual_process': True, 'exit_code': 0, 'timed_out': False, 'timeout_seconds': 30, 'pid': 123,
               'elapsed_seconds': 0.1, 'stdout': '', 'stderr': '', 'stdout_sha256': empty_hash, 'stderr_sha256': empty_hash}
    png = {'page': 1, 'path': str(ROOT / 'synthetic-does-not-exist.png'), 'filename': 'synthetic-does-not-exist.png',
           'bytes': 100, 'width': 100, 'height': 100, 'png_sha256': '4' * 64, 'pixels_sha256': '5' * 64}
    rendered = {name: {'process': deepcopy(process), 'pages': [deepcopy(png)]} for name in ('Poppler', 'PDFium')}
    source = {'page_id': 1, 'text': 'WBS-FID-PAGE-001\n', 'content_sha256': '1' * 64,
              'images': [{'name': '/Image0', 'width': 100, 'height': 100, 'bits_per_component': 8,
                          'color_space': '/DeviceRGB', 'decoded_sha256': '2' * 64}],
              'geometry': {'media': [0., 0., 100., 100.], 'crop': [0., 0., 100., 100.], 'trim': [0., 0., 100., 100.],
                           'bleed': [0., 0., 100., 100.], 'art': [0., 0., 100., 100.], 'rotation': 0},
              'static_annotations': [], 'links': [{'name': 'forward-0', 'target_page': 0, 'form': 'action',
                                                   'fit': '/XYZ', 'args': [None, 70., 1.25],
                                                   'rect': [0., 0., 10., 10.], 'border': [0., 0., 0.]}]}
    actual = deepcopy(source)
    actual['links'][0]['form'] = 'destination'
    entry = {'filename': '01 - Chapter.pdf', 'start': 0, 'end': 1, 'sha256': '3' * 64, 'size_bytes': 200}
    output = {**{key: entry[key] for key in ('filename', 'start', 'end', 'sha256')}, 'bytes': 200,
              'pages': [actual], 'expected_pages': [deepcopy(source)],
              'inventory': {'all_page_object_count': 1, 'serialized_content_marker_ids': [1], 'all_image_decoded_sha256': ['2' * 64]},
              'metadata': {'/Title': 'Chapter', '/Author': guard.AUTHOR, '/Creator': 'WinBookSplit',
                           '/Producer': 'WinBookSplit 1.0.0-dev (pypdf 6.19.0)', '/Subject': 'Source: ' + guard.TITLE},
              'outline': [{'title': 'Chapter', 'page': 0}], 'catalog_keys': ['/Type', '/Pages', '/Outlines'],
              'annotations': [{'page': 0, 'name': 'forward-0', 'P_matches_page': True, 'source_relations_absent': True, 'unsupported_action_absent': True}],
              'renders': deepcopy(rendered)}
    return {'kind': 'rich', 'source_renders': rendered}, output, entry, [source]


class FidelityReceiptTests(unittest.TestCase):
    def check(self, changed=None):
        row, output, entry, source = synthetic_output()
        if changed:
            changed(row, output, entry, source)
        guard.validate_output(row, output, entry, source)

    def test_canonical_direct_link_is_valid(self):
        self.check()

    def test_hidden_page_objects_fail(self):
        with self.assertRaisesRegex(ValueError, 'hidden off-segment'):
            self.check(lambda row, out, entry, source: out['inventory'].update(all_page_object_count=2))

    def test_offsegment_plaintext_marker_fails(self):
        with self.assertRaisesRegex(ValueError, 'hidden off-segment'):
            self.check(lambda row, out, entry, source: out['inventory'].update(serialized_content_marker_ids=[1, 2]))

    def test_offsegment_image_object_fails(self):
        with self.assertRaisesRegex(ValueError, 'hidden off-segment'):
            self.check(lambda row, out, entry, source: out['inventory']['all_image_decoded_sha256'].append('6' * 64))

    def test_renderer_pixel_contradiction_fails(self):
        with self.assertRaisesRegex(ValueError, 'pixels or PNG'):
            self.check(lambda row, out, entry, source: out['renders']['PDFium']['pages'][0].update(pixels_sha256='6' * 64))

    def test_rebased_link_to_another_page_fails(self):
        with self.assertRaisesRegex(ValueError, 'not exactly remapped|Rebased destination'):
            self.check(lambda row, out, entry, source: out['pages'][0]['links'][0].update(target_page=1))

    def test_rebased_link_view_contradiction_fails(self):
        for changed in ({'fit': '/FitH'}, {'args': [None, 71., 1.25]}, {'args': [None, 70., True]}):
            with self.subTest(changed=changed), self.assertRaises(ValueError):
                self.check(lambda row, out, entry, source: out['pages'][0]['links'][0].update(changed))

    def test_ordinary_transport_without_interaction_counters_is_valid_but_activity_fails(self):
        process = {'JobAssigned': True, 'ParentStopped': True, 'DescendantsStopped': True,
                   'StreamsComplete': True, 'ExitCode': 0, 'ResultRecordCount': 1}
        guard.validate_transport([process], '', '')
        for field, value in (('InteractionCount', 1), ('ReplyCount', False)):
            with self.subTest(field=field), self.assertRaisesRegex(ValueError, 'interaction activity'):
                guard.validate_transport([{**process, field: value}], '', '')
        with self.assertRaisesRegex(ValueError, 'request or reply'):
            guard.validate_transport([process], '', '[INTERACTION] {}')

    def test_actual_exporter_argv_cannot_be_substituted(self):
        directory = ROOT / 'synthetic-export-control'
        row = {'input_path': str(directory / 'b/source.pdf'), 'output_base': str(directory / 'o'),
               'cwd': str(directory / 'c'), 'shell_executable': str(directory / 'powershell.exe'),
               'console_evidence': {'owner': {'run_id': 'a' * 32}}}
        process = synthetic_output()[0]['source_renders']['Poppler']['process']
        process['cwd'] = row['cwd']
        process['command'] = [row['shell_executable'], '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'RemoteSigned', '-File',
                              str(directory / 'a/Export-WinBookSplitDiagnostics.ps1'), '-ManifestPath',
                              str(directory / 'o' / ('.WinBookSplit-console-' + 'a' * 32) / 'WinBookSplit_Run.json'),
                              '-OutputPath', str(directory / 'c/redacted-summary.json')]
        guard.validate_export_process(row, process)
        process['command'][6] = str(directory / 'unrelated.ps1')
        with self.assertRaisesRegex(ValueError, 'Actual native argv disagrees'):
            guard.validate_export_process(row, process)

    def test_export_frame_requires_version_one_and_typed_integer_fields(self):
        frame = {'protocol': 'winbooksplit.support-export', 'version': 1, 'status': 'success',
                 'code': 'support_export_complete', 'exit_code': 0, 'written_count': 1}
        guard.validate_export_frame('[SUPPORT-EXPORT] ' + json.dumps(frame) + '\n')
        for field, value in (('version', True), ('version', 2), ('exit_code', False), ('written_count', True)):
            bad = {**frame, field: value}
            # Re-serialize the contradictory frame so raw JSON agrees with it.
            with self.subTest(field=field, value=value), self.assertRaisesRegex(ValueError, 'native/frame mismatch'):
                guard.validate_export_frame('[SUPPORT-EXPORT] ' + json.dumps(bad) + '\n')

    def test_boolean_geometry_alias_fails(self):
        with self.assertRaisesRegex(ValueError, 'Boolean numeric'):
            self.check(lambda row, out, entry, source: out['pages'][0]['geometry'].update(rotation=False))

    def test_chapter_metadata_or_start_bookmark_contradiction_fails(self):
        for key in ('metadata', 'outline'):
            with self.subTest(key=key), self.assertRaises(ValueError):
                self.check(lambda row, out, entry, source: out['metadata'].update({'/Title': 'Wrong'}) if key == 'metadata' else out['outline'][0].update(page=1))

    def test_source_annotation_backreference_fails(self):
        with self.assertRaisesRegex(ValueError, 'another source page'):
            self.check(lambda row, out, entry, source: out['annotations'][0].update(P_matches_page=False))


if __name__ == '__main__':
    unittest.main()
