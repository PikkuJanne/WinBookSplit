"""Offline handoff tests; no app code imports or Windows execution."""
import importlib.util
import json
from pathlib import Path
import shutil
import tempfile
import unittest

TOOLS=Path(__file__).resolve().parents[1]
ROOT=TOOLS.parents[1]
spec=importlib.util.spec_from_file_location('handoff_check',TOOLS/'check_handoff.py')
helper=importlib.util.module_from_spec(spec);spec.loader.exec_module(helper)

class HandoffTests(unittest.TestCase):
    def test_auto_mode_cannot_skip_bundle_manifest_when_agents_is_present(self):
        with tempfile.TemporaryDirectory(prefix='wbs-SYNTHETIC-handoff-') as directory:
            root = Path(directory)
            for relative in ('docs/codex-v1.0.3','tools/codex-v1.0.3'):
                shutil.copytree(ROOT/relative, root/relative)
            for name in ('START_HERE.md','CODEX_START_PROMPT.md','CODEX_CONTINUE_PROMPT.md','CODEX_RELEASE_PROMPT.md'):
                (root/name).write_text('SYNTHETIC integrity-test placeholder.', encoding='utf-8')
            (root/'AGENTS.md').write_text('docs/codex-v1.0.3/SCOPE.md', encoding='utf-8')
            (root/'HANDOFF_MANIFEST.json').write_text('{"files":[]}', encoding='utf-8')
            with self.assertRaisesRegex(AssertionError, 'manifest member set mismatch'):
                helper.check(root)

    def test_structure_and_references(self):
        self.assertEqual(helper.check(ROOT)['tasks'],8)
    def test_acceptance_statuses_and_checkpoint_evidence(self):
        data=json.loads((ROOT/'docs/codex-v1.0.3/ACCEPTANCE_CASES.json').read_text())
        if not (ROOT/'AGENTS.md').is_file():
            self.assertTrue(all(c['status']=='NOT_RUN' for c in data['cases']))
            return
        for case in data['cases']:
            self.assertIn(case['status'], ('NOT_RUN','PASS','FAIL','BLOCKED','DEFERRED_NATIVE'))
            if case['status'] != 'NOT_RUN':
                self.assertTrue(case.get('evidence'), case['id'])
                self.assertTrue(all((ROOT/'docs/codex-v1.0.3'/p).is_file() for p in case['evidence']))
    def test_task_statuses(self):
        data=json.loads((ROOT/'docs/codex-v1.0.3/TASKS.json').read_text())
        if not (ROOT/'AGENTS.md').is_file():
            self.assertTrue(all(c['status']=='NOT_STARTED' for c in data['tasks']))
        else:
            for task in data['tasks']:
                self.assertIn(task['status'], data['task_statuses'])
    def test_preserve_report_failures(self):
        text=(ROOT/'docs/codex-v1.0.3/fixtures/discovery-report.sanitized.txt').read_text()
        for term in ('SignatureKind     : Store','runtime_not_found','Cannot process the XML',
                     'hard-linked executable aliases','packages\\":[]'):
            self.assertIn(term,text)
    def test_raw_owner_report_not_bundled(self):
        self.assertFalse(any(p.name=='WinBookSplit-discovery-20261017-122739.txt' for p in ROOT.rglob('*')))

if __name__=='__main__':unittest.main()
