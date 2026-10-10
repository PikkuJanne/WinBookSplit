"""AC-074 seeded page-owner oracle, callable engine checks and real writer samples.

Expected ownership is generated before the outline is arranged. This module does
not use engine planning/normalization helpers to construct expected partitions.
It launches no shell, child process, viewer, converter or network request.
"""
from __future__ import annotations

from collections.abc import Mapping
from contextlib import redirect_stdout
from hashlib import sha256
import importlib.util
from importlib.metadata import version
from io import StringIO
import json
from pathlib import Path
import platform
import random
import re
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[2]
SEED = 20261010074
MANUAL_COUNT = OUTLINE_COUNT = 240
INVALID_MANUAL_COUNT = 160
MALFORMED_COUNT = 48
WRITER_COUNT = 12
SOURCE_PATHS = (
    'VERSION',
    'engine/winbooksplit_engine.py', 'engine/winbooksplit_conversion.py',
    'engine/winbooksplit_job.py', 'engine/winbooksplit_windows.py',
    'engine/WinBookSplit.Outcomes.json', 'tests/fixtures/generate_pdf_fixtures.py',
    'tests/faults/invariants.py', 'tests/python/test_fault_invariants.py',
)


def require(value, message):
    if not value:
        raise ValueError(message)


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def source_hashes():
    return {name: sha256((ROOT / name).read_bytes()).hexdigest() for name in SOURCE_PATHS}


def full_source_hashes():
    runner = load('wbs_fault_invariant_source_manifest', ROOT / 'tests/run_tests.py')
    return runner.source_manifest()


class Destination(dict):
    def __init__(self, title, page):
        super().__init__({'/Title': title})
        self.title, self.page, self.node = title, page, None


class Reader:
    def __init__(self, outline, pages):
        self.outline, self.pages = outline, [None] * pages

    def get_destination_page_number(self, node):
        return node.page


class RawReader:
    def __init__(self, catalog, pages):
        self.root_object, self.pages = catalog, [None] * pages
        self.outline_read = False

    @property
    def outline(self):
        self.outline_read = True
        raise AssertionError('Malformed raw tree reached recursive outline retrieval')


def runs(owners):
    """Run-length encode independently assigned (title, reason, parent) labels."""
    result = []
    for page, owner in enumerate(owners):
        if not result or result[-1]['owner'] != owner:
            result.append({'start': page, 'end': page + 1, 'owner': owner})
        else:
            result[-1]['end'] = page + 1
    return result


def manual_cases(*, count=MANUAL_COUNT, maximum_pages=96, seed=SEED ^ 0x1100):
    rng = random.Random(seed)
    result = []
    for number in range(count):
        pages = rng.randint(1, maximum_pages)
        starts = [rng.randint(1, pages) for _ in range(rng.randint(1, 20))]
        rng.shuffle(starts)
        text = ', '.join(('0' * rng.randint(0, 3)) + str(page) for page in starts)
        selected = set(starts) | {1}
        owners = [('manual-' + str(max(start for start in selected if start <= page + 1)), 'manual', None)
                  for page in range(pages)]
        result.append({'id': f'manual-{number:03}', 'pages': pages, 'text': text,
                       'starts': sorted(selected), 'owners': owners})
    return result


def outline_cases(*, count=OUTLINE_COUNT, maximum_pages=96, seed=SEED ^ 0x2200):
    rng = random.Random(seed)
    result = []
    for number in range(count):
        pages = rng.randint(1, maximum_pages)
        starts = sorted(rng.sample(range(pages), rng.randint(1, min(6, pages))))
        parents = []
        for index, start in enumerate(starts):
            end = next((candidate for candidate in starts if candidate > start), pages)
            child_starts = sorted(rng.sample(range(start, end), rng.randint(0, min(4, end - start))))
            parents.append({'title': f'P{number}-{index}', 'start': start, 'end': end,
                            'children': [{'title': f'C{number}-{index}-{child}', 'page': page}
                                         for child, page in enumerate(child_starts)]})
        # Every accepted L2 arrangement has a genuine direct child. Other parents
        # may fall back intact; equal-start children and opening pages both occur.
        if not any(parent['children'] for parent in parents):
            parents[0]['children'] = [{'title': f'C{number}-forced', 'page': parents[0]['start']}]
        owners1, owners2 = [], []
        for page in range(pages):
            parent = next((p for p in parents if p['start'] <= page < p['end']), None)
            if parent is None:
                owners1.append(('Front matter', 'front_matter', None))
                owners2.append(('Front matter', 'front_matter', None))
                continue
            owners1.append((parent['title'], 'bookmark', None))
            candidates = [child for child in parent['children'] if child['page'] <= page]
            if candidates:
                child = max(candidates, key=lambda value: value['page'])
                owners2.append((child['title'], 'bookmark', parent['title']))
            elif parent['children']:
                owners2.append((parent['title'] + ' - Opening pages', 'parent_opening', parent['title']))
            else:
                owners2.append((parent['title'], 'parent_fallback', parent['title']))
        # Only now arrange the outline. This cannot alter the independent labels.
        arranged = []
        for parent in parents:
            children = [{'title': c['title'], 'page': c['page'], 'children': []} for c in parent['children']]
            rng.shuffle(children)
            if children:
                children[0]['children'] = [{'title': f'Depth3-{number}-{parent["title"]}',
                                            'page': rng.randrange(parent['start'], parent['end']), 'children': []}]
                children.append({'title': f'Alias-child-{parent["title"]}',
                                 'page': children[0]['page'], 'children': []})
            outside = parent['start'] - 1 if parent['start'] else parent['end']
            if outside < pages:
                children.append({'title': f'Outside-{parent["title"]}', 'page': outside, 'children': []})
            children.append({'title': f'Invalid-{parent["title"]}', 'page': None, 'children': []})
            arranged.append({'title': parent['title'], 'page': parent['start'], 'children': children})
        rng.shuffle(arranged)
        arranged.append({'title': f'Alias-parent-{number}', 'page': arranged[0]['page'],
                         'children': [{'title': f'Alias-subtree-{number}', 'page': 0, 'children': []}]})
        arranged.append({'title': f'Unusable-parent-{number}', 'page': None,
                         'children': [{'title': f'Orphan-subtree-{number}', 'page': 0, 'children': []}]})
        result.append({'id': f'outline-{number:03}', 'pages': pages, 'parents': parents,
                       'outline': arranged, 'owners1': owners1, 'owners2': owners2})
    return result


def as_outline(nodes):
    result = []
    for node in nodes:
        result.append(Destination(node['title'], node['page']))
        if node['children']:
            result.append(as_outline(node['children']))
    return result


def plan_projection(plan):
    names = {node['id']: node['title'] for node in plan.get('bookmarks', ())}
    return {'pages': plan['total_pages'], 'mode': plan['mode'],
            'ranges': [list(pair) for pair in plan['ranges']],
            'entries': [{'sequence': entry['sequence'], 'start': entry['start'], 'end': entry['end'],
                         'title': entry['title'], 'reason': entry['reason'],
                         'parent': names.get(entry['parent_id']) if entry['parent_id'] is not None else None}
                        for entry in plan['entries']],
            'coverage': dict(plan['coverage'])}


def check_partition(record, owners, *, mode, parents=()):
    pages = len(owners)
    expected = runs(owners)
    ranges = record['ranges']
    require(type(record['pages']) is int and record['pages'] == pages and record['mode'] == mode, 'Partition mode/pages differ')
    require(isinstance(ranges, list) and ranges and all(isinstance(pair, list) and len(pair) == 2
            and all(type(value) is int for value in pair) and 0 <= pair[0] < pair[1] <= pages for pair in ranges),
            'Partition ranges must be finite nonempty physical intervals')
    require([page for start, end in ranges for page in range(start, end)] == list(range(pages)),
            'Partition omits, repeats or reorders a physical page')
    require(ranges == [[entry['start'], entry['end']] for entry in expected], 'Independent ownership boundaries differ')
    entries = record['entries']
    require(len(entries) == len(expected), 'Entry count differs from independent owner runs')
    intervals = {p['title']: (p['start'], p['end']) for p in parents}
    for sequence, (entry, model) in enumerate(zip(entries, expected), 1):
        require(type(entry['sequence']) is int and entry['sequence'] == sequence
                and type(entry['start']) is int and type(entry['end']) is int
                and (entry['start'], entry['end']) == (model['start'], model['end'])
                and (entry['title'], entry['reason'], entry['parent']) == tuple(model['owner']),
                'Partition entry/sequence/title/parent differs')
        if mode == '2' and entry['parent'] is not None:
            start, end = intervals[entry['parent']]
            require(start <= entry['start'] < entry['end'] <= end, 'Level2 section crosses its known physical parent')
    require(record['coverage'] == {'complete': True, 'covered_pages': pages, 'section_count': len(expected)}
            and record['coverage']['complete'] is True
            and type(record['coverage']['covered_pages']) is int
            and type(record['coverage']['section_count']) is int, 'Coverage metadata differs')


def manual_projection(engine, case):
    raw = engine.plan_manual_starts(case['text'], case['pages'])
    entries = [{'sequence': number, 'title': f'manual-{start + 1}', 'start': start, 'end': end,
                'parent_id': None, 'reason': 'manual', 'filename': f'{number:03}.pdf', 'warnings': []}
               for number, (start, end) in enumerate(raw['ranges'], 1)]
    result = plan_projection(engine.validate_plan({'mode': 'manual', 'total_pages': case['pages'], 'entries': entries}))
    result['normalized_starts'] = raw['starts']
    return result


def invalid_manual_cases():
    rng = random.Random(SEED ^ 0x3300)
    result = []
    for number in range(INVALID_MANUAL_COUNT):
        pages = rng.randint(1, 96)
        bad = ['', ' ', '0', str(pages + 1), '+1', '-1', '1.0', '1-2', 'x', '１', '١', '1 2',
               '1\n2', '0x1', '1e0', '1;2'][number % 16]
        valid = str(rng.randint(1, pages))
        text = bad if number % 3 == 0 else (valid + ',' + bad if number % 3 == 1 else bad + ',' + valid)
        result.append({'id': f'invalid-manual-{number:03}', 'pages': pages, 'text': text})
    return result


MALFORMED_KINDS = ('self-list', 'ancestor-list', 'shared-list', 'orphan-list', 'root-dict', 'root-tuple',
                   'shared-destination', 'depth-limit', 'raw-catalog', 'raw-root', 'raw-next-cycle',
                   'raw-reused-node', 'raw-names-cycle', 'raw-odd-names', 'raw-duplicate-names', 'raw-kids-type')


def malformed_reader(number):
    rng = random.Random((SEED ^ 0x4400) + number)
    pages = rng.randint(1, 96)
    kind = MALFORMED_KINDS[number % len(MALFORMED_KINDS)]
    node = Destination(f'Fault-{number}', rng.randrange(pages))
    code = 'malformed_outline'
    if kind == 'self-list':
        outline = [node]; outline.append(outline); code = 'outline_cycle'
    elif kind == 'ancestor-list':
        outline = [node]; child = [Destination('Child', 0), outline]; outline.append(child); code = 'outline_cycle'
    elif kind == 'shared-list':
        child = [Destination('Shared', 0)]; outline = [node, child, Destination('Other', 0), child]; code = 'outline_cycle'
    elif kind == 'orphan-list':
        outline = [[node], Destination('Later parent', 0)]
    elif kind == 'root-dict':
        outline = {'unexpected': node}
    elif kind == 'root-tuple':
        outline = (node,)
    elif kind == 'shared-destination':
        outline = [node, [node]]; code = 'outline_cycle'
    elif kind == 'depth-limit':
        outline = [node]
        for depth in range(64):
            outline = [Destination(f'Depth-{depth}', 0), outline]
        code = 'outline_limit'
    else:
        raw = {'/Title': f'Raw-{number}'}
        if kind == 'raw-catalog':
            catalog = []
        elif kind == 'raw-root':
            catalog = {'/Outlines': 17}
        elif kind == 'raw-next-cycle':
            raw['/Next'] = raw; catalog = {'/Outlines': {'/First': raw}}; code = 'outline_cycle'
        elif kind == 'raw-reused-node':
            shared = {}; raw.update({'/Next': shared, '/First': shared})
            catalog = {'/Outlines': {'/First': raw}}; code = 'outline_cycle'
        elif kind == 'raw-names-cycle':
            raw['/Kids'] = [raw]; catalog = {'/Names': {'/Dests': raw}}; code = 'outline_cycle'
        elif kind == 'raw-odd-names':
            catalog = {'/Names': {'/Dests': {'/Names': ['A', [], 'Unpaired']}}}
        elif kind == 'raw-duplicate-names':
            catalog = {'/Names': {'/Dests': {'/Names': ['A', [], 'A', []]}}}
        else:
            catalog = {'/Names': {'/Dests': {'/Kids': 'invalid'}}}
        return RawReader(catalog, pages), pages, kind, code
    return Reader(outline, pages), pages, kind, code


def characterize_pure(engine=None):
    engine = engine or load('wbs_fault_invariant_engine', ROOT / 'engine/winbooksplit_engine.py')
    before = source_hashes()
    accepted, rejected_manual, malformed = [], [], []
    for case in manual_cases():
        observed = manual_projection(engine, case)
        require(observed['normalized_starts'] == case['starts'], 'Manual normalization differs')
        check_partition(observed, case['owners'], mode='manual')
        accepted.append({'id': case['id'], **observed})
    for case in outline_cases():
        for mode in ('1', '2'):
            planner = engine.plan_level1 if mode == '1' else engine.plan_level2
            plan = engine.validate_plan(planner(Reader(as_outline(case['outline']), case['pages']), case['pages']))
            observed = plan_projection(plan)
            check_partition(observed, case['owners' + mode], mode=mode, parents=case['parents'])
            accepted.append({'id': case['id'] + '-L' + mode, **observed})
    for case in invalid_manual_cases():
        try:
            engine.plan_manual_starts(case['text'], case['pages'])
        except engine.ManualPlanError as error:
            require(error.code == 'invalid_start_pages' and str(error), 'Manual rejection lost its diagnostic')
            rejected_manual.append({'id': case['id'], 'code': error.code, 'message_nonempty': bool(str(error))})
        else:
            raise ValueError('Invalid manual case was accepted: ' + case['id'])
    for number in range(MALFORMED_COUNT):
        for mode in ('1', '2'):
            reader, pages, kind, warning = malformed_reader(number)
            planner = engine.plan_level1 if mode == '1' else engine.plan_level2
            try:
                planner(reader, pages)
            except engine.BookmarkPlanError as error:
                require(error.code == 'invalid_outline' and str(error) and len(error.warnings) == 1
                        and error.warnings[0]['code'] == warning, 'Malformed graph diagnostic differs')
            else:
                raise ValueError('Malformed graph was accepted: ' + kind)
            reader, _, _, _ = malformed_reader(number)
            with patch.object(engine, 'PdfReader', return_value=reader), patch.object(engine, 'OutputRun') as reservation, \
                    patch.object(engine, 'write_slice') as writer:
                result = engine.run_split('authored-unused-input.pdf', 'authored-unused-output', mode)
            require(result['exit_code'] == 6 and result['status'] == 'invalid_input'
                    and result['code'] == 'invalid_outline' and result['written_count'] == 0
                    and result['execution'] is None and 'plan' not in result and not result['fallback_modes'],
                    'Malformed outline did not reject before plan/publication')
            require(not writer.called and not reservation.called, 'Malformed outline reached writer/reservation')
            require(not isinstance(reader, RawReader) or reader.outline_read is False, 'Raw guard ran too late')
            malformed.append({'id': f'malformed-{number:03}-L{mode}', 'kind': kind,
                'code': result['code'], 'status': result['status'], 'exit_code': result['exit_code'],
                'warning_code': warning, 'written_count': result['written_count'], 'execution': None,
                'plan_present': False, 'fallback_modes': [], 'writer_calls': writer.call_count,
                'reservation_calls': reservation.call_count,
                'raw_outline_read': reader.outline_read if isinstance(reader, RawReader) else None})
    require(source_hashes() == before, 'Invariant checks changed source')
    return {'protocol': 'winbooksplit.fault-invariants', 'version': 1, 'task_id': 'M4-T03', 'acceptance_id': 'AC-074',
            'seed': SEED, 'source_sha256': before, 'source_unchanged': True,
            'runtime': {'python': platform.python_version(), 'pypdf': version('pypdf')},
            'accepted_plan_count': len(accepted), 'invalid_manual_count': len(rejected_manual),
            'malformed_outline_count': len(malformed), 'accepted_plans': accepted,
            'invalid_manual': rejected_manual, 'malformed_outlines': malformed}


def file_identity(path):
    stat = path.stat()
    return {'sha256': sha256(path.read_bytes()).hexdigest(), 'size_bytes': stat.st_size,
            'device': stat.st_dev, 'inode': stat.st_ino}


def fixture_definition(case, mode):
    def nodes(items):
        return [{'title': n['title'], 'start_page': n['page'] + 1, 'children': nodes(n['children'])}
                for n in items if type(n['page']) is int and 0 <= n['page'] < case['pages']]
    return {'title': 'WinBookSplit original AC074 invariant sample', 'pages': case['pages'],
            'outline': [] if mode == 'manual' else nodes(case['outline'])}


def writer_models():
    manual = manual_cases(count=4, maximum_pages=24, seed=SEED ^ 0x5500)
    outlines = outline_cases(count=4, maximum_pages=24, seed=SEED ^ 0x6600)
    return [(f'writer-manual-{number}', 'manual', case, case['owners']) for number, case in enumerate(manual)] + \
           [(f'writer-L{mode}-{number}', mode, case, case['owners' + mode])
            for mode in ('1', '2') for number, case in enumerate(outlines)]


def characterize(work: Path) -> dict:
    """Author samples only in a new owned path; caller owns later cleanup."""
    work = Path(work).resolve()
    work.mkdir(exist_ok=False)
    tested_sources = full_source_hashes()
    report = characterize_pure()
    engine = load('wbs_fault_invariant_writer', ROOT / 'engine/winbooksplit_engine.py')
    fixtures = load('wbs_fault_invariant_fixtures', ROOT / 'tests/fixtures/generate_pdf_fixtures.py')
    samples = []
    for case_id, mode, case, owners in writer_models():
        owned = work / case_id; owned.mkdir()
        source = owned / 'source.pdf'
        with source.open('xb') as stream:
            stream.write(fixtures._pdf_bytes(fixture_definition(case, mode)))
        output = owned / 'output'; output.mkdir()
        prior, neighbor = output / 'prior-output.pdf', output / 'neighbor.txt'
        with prior.open('xb') as stream:
            stream.write(fixtures._pdf_bytes({'title': 'Original prior-run sentinel', 'pages': 1, 'outline': []}))
        with neighbor.open('xb') as stream:
            stream.write(b'Original AC074 neighbor must remain unchanged\n')
        protected = {'source': source, 'prior': prior, 'neighbor': neighbor}
        before = {key: file_identity(path) for key, path in protected.items()}
        prepared = engine.prepare_split(source, mode, case.get('text'), output_base=output)
        observed = plan_projection(engine.preview_plan(prepared))
        if mode == 'manual':
            # Actual manual titles describe the physical interval. The model
            # labels are start-page ownership keys; ranges remain independent.
            for entry in observed['entries']:
                entry['title'] = 'manual-' + str(entry['start'] + 1)
        check_partition(observed, owners, mode=mode, parents=case.get('parents', ()))
        with redirect_stdout(StringIO()):
            result = engine.execute_split(prepared, output)
        require(result['written_count'] == len(runs(owners)) and not result.get('post_publication_warnings'),
                'Actual writer did not publish exactly the complete plan')
        final = Path(result['final_directory'])
        manifest_path = final / result['manifest_filename']
        manifest = json.loads(manifest_path.read_text(encoding='utf-8'))
        chapters = []
        for entry, actual in zip(prepared.plan['entries'], result['outputs']):
            path = final / entry['filename']
            ids = fixtures.page_ids(path)
            require(ids == list(range(entry['start'] + 1, entry['end'] + 1)), 'Writer omitted/repeated/reordered marked pages')
            require(actual['start'] == entry['start'] and actual['end'] == entry['end'], 'Writer used a different plan')
            chapters.append({'filename': path.name, 'start': entry['start'], 'end': entry['end'],
                             'page_ids': ids, 'file_identity': file_identity(path),
                             'manifest_sha256': actual['sha256'], 'manifest_size_bytes': actual['size_bytes'],
                             'manifest_page_count': actual['page_count']})
        expected_members = {entry['filename'] for entry in prepared.plan['entries']} | \
                           {'.WinBookSplit-owner.json', result['manifest_filename']}
        require({p.name for p in final.iterdir()} == expected_members and
                {p.name for p in output.iterdir()} == {'prior-output.pdf', 'neighbor.txt', final.name},
                'Writer left unexpected output/staging members')
        after = {key: file_identity(path) for key, path in protected.items()}
        require(before == after, 'Writer changed input/neighbor/prior identity or bytes')
        samples.append({'id': case_id, 'mode': mode, 'plan': observed, 'protected_before': before,
                        'protected_after': after, 'captured_source_sha256': prepared.plan['source_identity']['sha256'],
                        'source_fixture_sha256': sha256(fixtures._pdf_bytes(fixture_definition(case, mode))).hexdigest(),
                        'written_count': result['written_count'], 'chapters': chapters,
                        'manifest_identity': file_identity(manifest_path), 'manifest_status': manifest['status'],
                        'manifest_written_count': manifest['written_count'],
                        'manifest_outputs': [{key: output[key] for key in
                            ('filename', 'start', 'end', 'sha256', 'size_bytes', 'page_count')}
                            for output in manifest['outputs']],
                        'manifest_source_sha256': manifest['source_identity']['sha256'],
                        'manifest_coverage': manifest['coverage'],
                        'final_members': sorted(p.name for p in final.iterdir()),
                        'base_members': sorted(p.name for p in output.iterdir()), 'final_basename': final.name,
                        'scope': 'actual_callable_engine_writer_reopened_authored_PDFs'})
    report.update(writer_sample_count=len(samples), writer_samples=samples, success=True, exit_code=0,
                  tested_path_sha256=tested_sources,
                  tested_paths_digest=sha256(json.dumps(tested_sources, sort_keys=True, separators=(',', ':')).encode()).hexdigest(),
                  result='FAULT_INVARIANTS_PASSED', engine_native_exit_claimed=False,
                  scope='seeded_callable_plans_and_real_callable_writer_not_PS_BAT_GUI_or_human')
    require(source_hashes() == report['source_sha256'], 'Sources changed while authoring samples')
    require(full_source_hashes() == tested_sources, 'Full runner sources changed while authoring samples')
    validate_invariant_report(report)
    return report


def validate_invariant_report(report, *, expected_source_sha256=None):
    """Rebuild ownership from the fixed seed; never trust totals/pass flags alone."""
    require(isinstance(report, dict) and report.get('protocol') == 'winbooksplit.fault-invariants'
            and type(report.get('version')) is int and report['version'] == 1 and report.get('task_id') == 'M4-T03'
            and report.get('acceptance_id') == 'AC-074' and type(report.get('seed')) is int and report['seed'] == SEED,
            'Invariant protocol/task/seed differs')
    expected_source_sha256 = source_hashes() if expected_source_sha256 is None else expected_source_sha256
    require(set(report['source_sha256']) == set(SOURCE_PATHS) and
            report['source_sha256'] == {key: expected_source_sha256[key] for key in SOURCE_PATHS}
            and report['source_unchanged'] is True, 'Invariant source binding differs')
    require(report['runtime'] == {'python': platform.python_version(), 'pypdf': version('pypdf')},
            'Invariant observed runtime versions differ')
    for key, value in (('accepted_plan_count', 720), ('invalid_manual_count', 160), ('malformed_outline_count', 96)):
        require(type(report[key]) is int and report[key] == value, 'Invariant case count differs: ' + key)
    expected_ids = [case['id'] for case in manual_cases()] + \
                   [case['id'] + '-L' + mode for case in outline_cases() for mode in ('1', '2')]
    require([row['id'] for row in report['accepted_plans']] == expected_ids, 'Accepted deterministic case IDs differ')
    for observed, case in zip(report['accepted_plans'][:MANUAL_COUNT], manual_cases()):
        check_partition(observed, case['owners'], mode='manual')
        require(observed['normalized_starts'] == case['starts'] and all(type(v) is int for v in observed['normalized_starts']),
                'Manual normalized starts differ')
    for index, case in enumerate(outline_cases()):
        for offset, mode in enumerate(('1', '2')):
            check_partition(report['accepted_plans'][MANUAL_COUNT + index * 2 + offset],
                            case['owners' + mode], mode=mode, parents=case['parents'])
    require(len(report['invalid_manual']) == INVALID_MANUAL_COUNT and
            [row['id'] for row in report['invalid_manual']] == [case['id'] for case in invalid_manual_cases()],
            'Invalid manual deterministic cases differ')
    require(all(row['code'] == 'invalid_start_pages' and row['message_nonempty'] is True for row in report['invalid_manual']),
            'Manual invalid token silently accepted/lost diagnostic')
    require(len(report['malformed_outlines']) == MALFORMED_COUNT * 2, 'Malformed case rows missing')
    for number in range(MALFORMED_COUNT):
        reader, _, kind, warning = malformed_reader(number)
        for offset, mode in enumerate(('1', '2')):
            row = report['malformed_outlines'][number * 2 + offset]
            require(row == {'id': f'malformed-{number:03}-L{mode}', 'kind': kind, 'code': 'invalid_outline',
                    'status': 'invalid_input', 'exit_code': 6, 'warning_code': warning, 'written_count': 0,
                    'execution': None, 'plan_present': False, 'fallback_modes': [], 'writer_calls': 0,
                    'reservation_calls': 0, 'raw_outline_read': False if isinstance(reader, RawReader) else None},
                    'Malformed rejection/early guard evidence differs')
            require(all(type(row[key]) is int for key in ('exit_code', 'written_count', 'writer_calls', 'reservation_calls')),
                    'Malformed numeric proof is not an integer')
            require(row['plan_present'] is False and row['raw_outline_read'] is
                    (False if isinstance(reader, RawReader) else None), 'Malformed Boolean guard proof differs')
    # Pure tests may validate their narrower report before fixture authoring.
    if 'writer_samples' not in report:
        require('success' not in report and 'result' not in report and 'writer_sample_count' not in report,
                'Pure report falsely claims complete writer acceptance')
        return
    require(report.get('success') is True and report.get('result') == 'FAULT_INVARIANTS_PASSED'
            and report.get('engine_native_exit_claimed') is False
            and type(report.get('exit_code')) is int and report['exit_code'] == 0
            and type(report.get('writer_sample_count')) is int and report['writer_sample_count'] == WRITER_COUNT,
            'Writer scope/result/count differs')
    expected_full = full_source_hashes() if set(expected_source_sha256) == set(SOURCE_PATHS) else expected_source_sha256
    require(report['tested_path_sha256'] == expected_full
            and report['tested_paths_digest'] == sha256(json.dumps(expected_full, sort_keys=True, separators=(',', ':')).encode()).hexdigest(),
            'Full source map/digest differs')
    samples = report['writer_samples']
    require([row['id'] for row in samples] == [model[0] for model in writer_models()], 'Writer case IDs differ')
    fixtures = load('wbs_fault_invariant_validation_fixtures', ROOT / 'tests/fixtures/generate_pdf_fixtures.py')
    for row, (case_id, mode, case, owners) in zip(samples, writer_models()):
        require(row['mode'] == mode and row['scope'] == 'actual_callable_engine_writer_reopened_authored_PDFs', 'Writer mode/scope differs')
        check_partition(row['plan'], owners, mode=mode, parents=case.get('parents', ()))
        require(set(row['protected_before']) == {'source', 'prior', 'neighbor'}
                and row['protected_before'] == row['protected_after'], 'Protected identity/bytes changed')
        for identity in (*row['protected_before'].values(), row['manifest_identity']):
            check_identity(identity)
        source_digest = sha256(fixtures._pdf_bytes(fixture_definition(case, mode))).hexdigest()
        require(row['source_fixture_sha256'] == row['captured_source_sha256'] == row['manifest_source_sha256']
                == row['protected_before']['source']['sha256'] == source_digest, 'Actual writer/source capture differs')
        prior_bytes = fixtures._pdf_bytes({'title': 'Original prior-run sentinel', 'pages': 1, 'outline': []})
        neighbor_bytes = b'Original AC074 neighbor must remain unchanged\n'
        for role, known_bytes in (('prior', prior_bytes), ('neighbor', neighbor_bytes)):
            require(row['protected_before'][role]['sha256'] == sha256(known_bytes).hexdigest()
                    and row['protected_before'][role]['size_bytes'] == len(known_bytes), 'Protected sentinel differs')
        require(type(row['written_count']) is int and row['written_count'] == len(runs(owners))
                and type(row['manifest_written_count']) is int and row['manifest_written_count'] == row['written_count']
                and row['manifest_status'] == 'complete' and row['manifest_coverage'] == row['plan']['coverage'],
                'Writer zero/incomplete/manifest count differs')
        require(len(row['chapters']) == row['written_count'], 'Writer chapter evidence missing')
        require(len(row['manifest_outputs']) == len(row['chapters']), 'Actual completion manifest outputs missing')
        for sequence, (chapter, expected, manifest_entry) in enumerate(zip(row['chapters'], runs(owners), row['manifest_outputs']), 1):
            require((chapter['start'], chapter['end']) == (expected['start'], expected['end'])
                    and type(chapter['start']) is int and type(chapter['end']) is int
                    and chapter['page_ids'] == list(range(expected['start'] + 1, expected['end'] + 1))
                    and all(type(page) is int for page in chapter['page_ids']), 'Reopened writer pages/ranges differ')
            check_identity(chapter['file_identity'])
            expected_title = f'Section (Page {expected["start"] + 1}-{expected["end"]})' if mode == 'manual' else expected['owner'][0]
            expected_filename = f'{sequence:02d} - {expected_title}.pdf'
            require(chapter['filename'] == expected_filename and manifest_entry ==
                    {'filename': expected_filename, 'start': chapter['start'], 'end': chapter['end'],
                     'sha256': chapter['manifest_sha256'], 'size_bytes': chapter['manifest_size_bytes'],
                     'page_count': chapter['manifest_page_count']}, 'Manifest/name projection differs from actual chapter proof')
            require(safe_leaf(chapter['filename']) and chapter['filename'].endswith('.pdf')
                    and chapter['manifest_sha256'] == chapter['file_identity']['sha256']
                    and type(chapter['manifest_size_bytes']) is int
                    and chapter['manifest_size_bytes'] == chapter['file_identity']['size_bytes']
                    and type(chapter['manifest_page_count']) is int
                    and chapter['manifest_page_count'] == expected['end'] - expected['start'],
                    'Writer chapter manifest/hash/size/count differs')
        names = [chapter['filename'] for chapter in row['chapters']]
        require(len(set(name.casefold() for name in names)) == len(names)
                and row['final_members'] == sorted(names + ['.WinBookSplit-owner.json', 'WinBookSplit_Manifest.json'])
                and safe_leaf(row['final_basename'])
                and row['base_members'] == sorted(['prior-output.pdf', 'neighbor.txt', row['final_basename']]),
                'Writer leaked staging/unexpected members or overwrote a prior run')


def safe_leaf(value):
    return isinstance(value, str) and bool(value) and value not in {'.', '..'} and not re.search(r'[<>:"/\\|?*\x00-\x1f]', value)


def check_identity(identity):
    require(isinstance(identity, dict) and set(identity) == {'sha256', 'size_bytes', 'device', 'inode'}
            and isinstance(identity['sha256'], str) and re.fullmatch('[0-9a-f]{64}', identity['sha256'])
            and all(type(identity[key]) is int and identity[key] >= 0 for key in ('size_bytes', 'device', 'inode'))
            and identity['size_bytes'] > 0 and identity['inode'] > 0, 'File hash/native identity proof malformed')
