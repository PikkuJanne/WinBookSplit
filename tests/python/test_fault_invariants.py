"""Independent AC074 ownership, structured rejection and authored writer proof."""
import copy
from pathlib import Path
import tempfile
import unittest

import importlib.util

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location('wbs_fault_invariants', ROOT / 'tests/faults/invariants.py')
invariants = importlib.util.module_from_spec(spec)
spec.loader.exec_module(invariants)
engine = invariants.load('wbs_fault_invariant_unit_engine', ROOT / 'engine/winbooksplit_engine.py')


class IndependentOwnerTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.report = invariants.characterize_pure(engine)

    def test_seed_ids_and_generated_parent_owners_are_deterministic(self):
        self.assertEqual(invariants.manual_cases(), invariants.manual_cases())
        self.assertEqual(invariants.outline_cases(), invariants.outline_cases())
        for case in invariants.outline_cases():
            self.assertTrue(1 <= case['pages'] <= 96)
            self.assertEqual(len(case['owners1']), case['pages'])
            self.assertEqual(len(case['owners2']), case['pages'])
        self.assertEqual(len(self.report['accepted_plans']), 720)

    def test_all_240_seeded_manual_plans_preserve_independent_page_owners(self):
        for case, actual in zip(invariants.manual_cases(), self.report['accepted_plans'][:240]):
            with self.subTest(case=case['id']):
                invariants.check_partition(actual, case['owners'], mode='manual')
                self.assertEqual(actual['normalized_starts'], case['starts'])

    def test_all_240_level1_arrangements_preserve_frontmatter_and_first_parent(self):
        for index, case in enumerate(invariants.outline_cases()):
            with self.subTest(case=case['id']):
                invariants.check_partition(self.report['accepted_plans'][240 + index * 2],
                                           case['owners1'], mode='1', parents=case['parents'])
        self.assertTrue(any(case['owners1'][0][1] == 'front_matter' for case in invariants.outline_cases()))

    def test_all_240_level2_arrangements_preserve_each_known_parent_interval(self):
        reasons = set()
        for index, case in enumerate(invariants.outline_cases()):
            with self.subTest(case=case['id']):
                invariants.check_partition(self.report['accepted_plans'][241 + index * 2],
                                           case['owners2'], mode='2', parents=case['parents'])
                reasons.update(owner[1] for owner in case['owners2'])
        self.assertEqual(reasons, {'front_matter', 'parent_opening', 'parent_fallback', 'bookmark'})

    def test_160_invalid_manual_tokens_fail_entirely_with_structured_diagnostic(self):
        self.assertEqual(len(self.report['invalid_manual']), 160)
        self.assertTrue(all(row['code'] == 'invalid_start_pages' and row['message_nonempty'] is True
                            for row in self.report['invalid_manual']))

    def test_96_malformed_graphs_reject_before_reservation_writer_or_recursive_raw_read(self):
        self.assertEqual(len(self.report['malformed_outlines']), 96)
        self.assertEqual({row['kind'] for row in self.report['malformed_outlines']}, set(invariants.MALFORMED_KINDS))
        self.assertTrue(all(row['code'] == 'invalid_outline' and row['exit_code'] == 6
                            and row['writer_calls'] == row['reservation_calls'] == 0
                            and row['raw_outline_read'] is not True for row in self.report['malformed_outlines']))

    def test_invalid_page_count_and_documented_no_plan_are_distinct_from_malformed_graphs(self):
        for pages in (0, -1, None, True, 1.0, '1'):
            with self.subTest(pages=pages), self.assertRaises(engine.ManualPlanError) as raised:
                engine.plan_manual_starts('1', pages)
            self.assertEqual(raised.exception.code, 'invalid_document')
        for mode, outline, code in [('1', [], 'no_bookmarks'),
                                    ('1', [invariants.Destination('Unusable', None)], 'no_usable_bookmarks'),
                                    ('2', [invariants.Destination('Parent', 0)], 'no_bookmarks_at_level')]:
            planner = engine.plan_level1 if mode == '1' else engine.plan_level2
            with self.subTest(code=code), self.assertRaises(engine.BookmarkPlanError) as raised:
                planner(invariants.Reader(outline, 3), 3)
            self.assertEqual(raised.exception.code, code)

    def test_unusable_destination_is_visible_warning_with_valid_complete_plan(self):
        reader = invariants.Reader([invariants.Destination('Parent', 1),
                                    [invariants.Destination('Child', 2), invariants.Destination('Invalid', None)]], 4)
        plan = engine.validate_plan(engine.plan_level2(reader, 4))
        self.assertEqual(plan['ranges'], ((0, 1), (1, 2), (2, 4)))
        self.assertTrue(any(warning['code'] == 'invalid_destination' for warning in plan['warnings']))

    def test_oracle_rejects_equal_total_omission_duplicate_order_and_parent_crossing(self):
        actual = next(row for row in self.report['accepted_plans'] if len(row['ranges']) >= 3)
        case = next(case for case in invariants.manual_cases() if case['id'] == actual['id'])
        for label in ('duplicate', 'reorder', 'first_page_loss', 'empty', 'tail_loss'):
            damaged = copy.deepcopy(actual)
            if label == 'duplicate':
                damaged['ranges'][1] = list(damaged['ranges'][0])
            elif label == 'reorder':
                damaged['ranges'].reverse()
            elif label == 'first_page_loss':
                damaged['ranges'][0][0] += 1
            elif label == 'empty':
                damaged['ranges'][0][1] = damaged['ranges'][0][0]
            else:
                damaged['ranges'][-1][1] -= 1
            with self.subTest(label=label), self.assertRaises(ValueError):
                invariants.check_partition(damaged, case['owners'], mode='manual')
        case = next(case for case in invariants.outline_cases() if len(case['parents']) >= 2)
        actual = copy.deepcopy(self.report['accepted_plans'][241 + int(case['id'].split('-')[1]) * 2])
        child = next(entry for entry in actual['entries'] if entry['parent'] is not None)
        child['parent'] = next(p['title'] for p in case['parents'] if p['title'] != child['parent'])
        with self.assertRaises(ValueError):
            invariants.check_partition(actual, case['owners2'], mode='2', parents=case['parents'])

    def test_pure_report_never_claims_native_or_complete_writer_acceptance(self):
        invariants.validate_invariant_report(self.report)
        self.assertNotIn('success', self.report)
        self.assertNotIn('writer_samples', self.report)
        forged = copy.deepcopy(self.report); forged['success'] = True
        with self.assertRaises(ValueError):
            invariants.validate_invariant_report(forged)

    def test_receipt_rejects_missing_ids_wrong_seed_boolean_count_and_fabricated_rejections(self):
        mutations = [('seed', lambda r: r.update(seed=0)),
                     ('count', lambda r: r.update(accepted_plan_count=True)),
                     ('missing', lambda r: r['accepted_plans'].pop()),
                     ('diagnostic', lambda r: r['invalid_manual'][0].update(code='split_complete')),
                     ('writer', lambda r: r['malformed_outlines'][0].update(writer_calls=1)),
                     ('raw_read', lambda r: next(row for row in r['malformed_outlines']
                                                if row['raw_outline_read'] is False).update(raw_outline_read=True)),
                     ('boolean_guard', lambda r: r['malformed_outlines'][0].update(plan_present=0)),
                     ('parent', lambda r: r['accepted_plans'][241]['entries'][0].update(parent='Wrong parent'))]
        for name, mutate in mutations:
            damaged = copy.deepcopy(self.report); mutate(damaged)
            with self.subTest(mutation=name), self.assertRaises(ValueError):
                invariants.validate_invariant_report(damaged)


class ActualWriterInvariantTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.owned = tempfile.TemporaryDirectory(prefix='wbs-fault-invariant-unit-')
        cls.addClassCleanup(cls.owned.cleanup)
        cls.report = invariants.characterize(Path(cls.owned.name) / 'authored')

    def test_twelve_actual_writer_samples_reopen_every_page_and_preserve_protected_identities(self):
        invariants.validate_invariant_report(self.report)
        self.assertEqual(len(self.report['writer_samples']), 12)
        self.assertEqual({mode: sum(row['mode'] == mode for row in self.report['writer_samples'])
                          for mode in ('manual', '1', '2')}, {'manual': 4, '1': 4, '2': 4})
        self.assertTrue(all(row['protected_before'] == row['protected_after'] for row in self.report['writer_samples']))

    def test_reopened_same_count_duplicate_marker_is_rejected(self):
        damaged = copy.deepcopy(self.report)
        chapter = next(chapter for row in damaged['writer_samples'] for chapter in row['chapters'] if len(chapter['page_ids']) > 1)
        chapter['page_ids'][-1] = chapter['page_ids'][0]
        with self.assertRaises(ValueError):
            invariants.validate_invariant_report(damaged)

    def test_changed_source_or_neighbor_identity_cannot_be_accepted_as_preservation(self):
        for role, field in [('source', 'sha256'), ('neighbor', 'inode'), ('prior', 'size_bytes')]:
            damaged = copy.deepcopy(self.report)
            value = damaged['writer_samples'][0]['protected_after'][role]
            value[field] = '0' * 64 if field == 'sha256' else value[field] + 1
            with self.subTest(role=role), self.assertRaises(ValueError):
                invariants.validate_invariant_report(damaged)

    def test_publication_manifest_hash_and_missing_output_evidence_fail_closed(self):
        for field in ('manifest_sha256', 'manifest_page_count', 'manifest_size_bytes'):
            damaged = copy.deepcopy(self.report)
            chapter = damaged['writer_samples'][0]['chapters'][0]
            chapter[field] = '0' * 64 if field == 'manifest_sha256' else chapter[field] + 1
            with self.subTest(field=field), self.assertRaises(ValueError):
                invariants.validate_invariant_report(damaged)
        damaged = copy.deepcopy(self.report)
        damaged['writer_samples'][0]['manifest_outputs'][0]['start'] += 1
        with self.assertRaises(ValueError):
            invariants.validate_invariant_report(damaged)


if __name__ == '__main__':
    unittest.main(verbosity=2)
