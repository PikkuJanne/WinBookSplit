"""Synthetic in-memory report tests. Never native Windows evidence."""
import copy
import hashlib
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest

TOOLS=Path(__file__).resolve().parents[1]
spec=importlib.util.spec_from_file_location('native_gate_under_test',TOOLS/'check_native_store_report.py')
gate=importlib.util.module_from_spec(spec);spec.loader.exec_module(gate)
COMMIT='a'*40
ZIP='b'*64
RUNTIME='c'*64

class GateTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.addCleanup(self.temp.cleanup)
        self.root=Path(self.temp.name);self.path=self.root/'report.json'
        self.report={
          'schema_version':1,'scenario':'existing_owner_windows10_store_installation',
          'evidence_kind':'native_execution','status':'passed','application_version':'1.0.3',
          'run_id':'SYNTHETIC_UNIT_ONLY','source_commit':COMMIT,'zip_sha256':ZIP,
          'host':{'execution_kind':'native','owner_desktop':True,'os':'Windows 10','build':19045,
                  'architecture':'x64','powershell_version':'5.1.19041.7417'},
          'registration':{'family':gate.FAMILY,'publisher':gate.PUBLISHER,'signature_kind':'Store',
                          'architecture':'X64','status':'Ok','version':'26.4.240.0'},
          'forbidden_substitutions':dict.fromkeys(gate.FORBIDDEN,False),
          'before':{'application_venv_exists':False,'runtimes':[{'runtime_ref':'OWNER_PYTHON_315',
                     'sha256':RUNTIME,'version':'3.15.0','managed_by_store':True}]},
          'after':{'application_venv_exists':True,'runtimes':[{'runtime_ref':'OWNER_PYTHON_315',
                     'sha256':RUNTIME,'version':'3.15.0','managed_by_store':True}]},
          'discovery':{'mode':'automatic','route':'registered_store_manager',
                       'provenance':'pre_existing_owner_install','runtime_ref':'OWNER_PYTHON_315',
                       'version':'3.15.0','interpreter_probe_exit':0,'package_state':'setup_required'},
          'setup':{'entry':'Setup-WinBookSplit.bat','runtime_selection':'automatic','exit_code':0,
                   'base_runtime_ref':'OWNER_PYTHON_315','venv_base_runtime_ref':'OWNER_PYTHON_315'},
          'checks':dict.fromkeys(gate.CHECKS,True),'artifacts':[],
          'notes':'Entire report is a synthetic unit-test fixture, NOT observed native evidence.'}
        self.write()
    def write(self):
        self.report['artifacts']=[]
        for kind,summary in gate.expected_summaries(self.report).items():
            path=self.root/(kind+'.json')
            data={'schema_version':1,'kind':kind,'run_id':self.report['run_id'],
                  'source_commit':self.report['source_commit'],'zip_sha256':self.report['zip_sha256'],
                  'summary':summary,'records':[{'test_only':'synthetic unit input, not execution'}]}
            raw=(json.dumps(data,ensure_ascii=False)+'\n').encode('utf8');path.write_bytes(raw)
            self.report['artifacts'].append({'kind':kind,'path':path.name,'sha256':hashlib.sha256(raw).hexdigest()})
        self.save()
    def save(self):self.path.write_text(json.dumps(self.report),encoding='utf8')
    def validate(self):return gate.validate(self.path,COMMIT,ZIP)
    def rejected(self):
        with self.assertRaises((gate.GateError,OSError,ValueError,KeyError,TypeError)):
            self.validate()
    def test_consistent_synthetic_report_contract(self):
        self.assertEqual(self.validate()['status'],'CONSISTENT_NATIVE_REPORT')
    def test_existing_global_package_ready_also_allowed(self):
        self.report['discovery']['package_state']='ready';self.write();self.validate()
    def test_unexecuted_template_rejected(self):
        self.report['status']='not_run';self.save();self.rejected()
    def test_synthetic_evidence_label_rejected(self):
        self.report['evidence_kind']='synthetic';self.save();self.rejected()
    def test_private_fallback_rejected(self):
        self.report['forbidden_substitutions']['private_runtime_fallback_used']=True;self.write();self.rejected()
    def test_every_substitution_rejected(self):
        for key in gate.FORBIDDEN:
            with self.subTest(key=key):
                self.report['forbidden_substitutions'][key]=True;self.write();self.rejected()
                self.report['forbidden_substitutions'][key]=False
    def test_unknown_substitution_not_false(self):
        self.report['forbidden_substitutions']['child_path_injected']=None;self.save();self.rejected()
    def test_string_false_not_bool(self):
        self.report['forbidden_substitutions']['child_path_injected']='false';self.save();self.rejected()
    def test_missing_substitution_observation(self):
        del self.report['forbidden_substitutions']['runtime_copied_or_cloned'];self.save();self.rejected()
    def test_explicit_selection_even_without_flag(self):
        self.report['discovery']['mode']='explicit';self.write();self.rejected()
    def test_old_private_baseline_even_if_pdf_passed(self):
        self.report['discovery']['provenance']='reviewed_private3150_archive';self.write();self.rejected()
    def test_precreated_venv_rejected(self):
        self.report['before']['application_venv_exists']=True;self.write();self.rejected()
    def test_windows_server_not_windows10(self):
        self.report['host']['os']='Windows Server 2025';self.save();self.rejected()
    def test_wrong_windows_build(self):
        self.report['host']['build']=26300;self.save();self.rejected()
    def test_ps7_not_inbox_native_case(self):
        self.report['host']['powershell_version']='7.6.7';self.save();self.rejected()
    def test_mocked_host_rejected(self):
        self.report['host']['execution_kind']='mocked';self.save();self.rejected()
    def test_wrong_manager_channel(self):
        self.report['registration']['family']='PythonSoftwareFoundation.PythonManager_3847v3x7pw1km';self.write();self.rejected()
    def test_crosspaired_publisher_rejected(self):
        self.report['registration']['publisher']='CN=Python Software Foundation';self.write();self.rejected()
    def test_different_manager_version_not_exact_pin(self):
        self.report['registration']['version']='26.5.100.0';self.write();self.validate()
    def test_new_runtime_absent_from_before_inventory(self):
        self.report['discovery']['runtime_ref']='NEW_PRIVATE_PYTHON';self.write();self.rejected()
    def test_runtime_hash_changed(self):
        self.report['after']['runtimes'][0]['sha256']='d'*64;self.write();self.rejected()
    def test_runtime_duplicate_inventory(self):
        self.report['before']['runtimes'].append(copy.deepcopy(self.report['before']['runtimes'][0]));self.write();self.rejected()
    def test_venv_base_changed(self):
        self.report['setup']['venv_base_runtime_ref']='OTHER_RUNTIME';self.write();self.rejected()
    def test_setup_failure(self):
        self.report['setup']['exit_code']=3;self.write();self.rejected()
    def test_false_cannot_be_zero_exit_code(self):
        self.report['setup']['exit_code']=False;self.write();self.rejected()
    def test_missing_native_check(self):
        self.report['checks']['normal_setup_bat_passed']=False;self.write();self.rejected()
    def test_wrong_artifact_source(self):
        self.report['source_commit']='e'*40;self.write();self.rejected()
    def test_wrong_zip(self):
        self.report['zip_sha256']='f'*64;self.write();self.rejected()
    def test_receipt_hash_mismatch(self):
        (self.root/'bat.json').write_text('{}');self.rejected()
    def test_missing_receipt_file(self):
        (self.root/'setup.json').unlink();self.rejected()
    def test_missing_receipt_kind(self):
        self.report['artifacts'].pop();self.save();self.rejected()
    def test_duplicate_receipt_kind(self):
        self.report['artifacts'][-1]['kind']='bat';self.save();self.rejected()
    def test_receipt_wrong_run_id_even_with_new_hash(self):
        item=self.report['artifacts'][0];p=self.root/item['path']
        raw=json.loads(p.read_text());raw['run_id']='WRONG_RECEIPT_RUN'
        data=json.dumps(raw).encode();p.write_bytes(data);item['sha256']=hashlib.sha256(data).hexdigest()
        self.save();self.rejected()
    def test_receipt_summary_mismatch_even_with_new_hash(self):
        item=self.report['artifacts'][1];p=self.root/item['path']
        raw=json.loads(p.read_text());raw['summary']['route']='explicit'
        data=json.dumps(raw).encode();p.write_bytes(data);item['sha256']=hashlib.sha256(data).hexdigest()
        self.save();self.rejected()
    def test_receipt_records_required(self):
        item=self.report['artifacts'][0];p=self.root/item['path']
        raw=json.loads(p.read_text());raw['records']=[]
        data=json.dumps(raw).encode();p.write_bytes(data);item['sha256']=hashlib.sha256(data).hexdigest()
        self.save();self.rejected()
    def test_path_traversal_rejected(self):
        self.report['artifacts'][0]['path']='../outside.json';self.save();self.rejected()
    def test_backslash_path_rejected(self):
        self.report['artifacts'][0]['path']='..\\outside.json';self.save();self.rejected()
    def test_absolute_path_rejected(self):
        self.report['artifacts'][0]['path']='/tmp/outside.json';self.save();self.rejected()
    def test_symlink_receipt_rejected(self):
        p=self.root/'registration.json';p.rename(self.root/'original.json')
        try:p.symlink_to(self.root/'original.json')
        except OSError:self.skipTest('symlinks unavailable in this test environment')
        self.rejected()
    def test_duplicate_json_key_rejected(self):
        text=self.path.read_text();self.path.write_text(text.replace('"schema_version": 1', '"schema_version": 1, "schema_version": 1',1));self.rejected()
    def test_nonfinite_json_rejected(self):
        text=self.path.read_text();self.path.write_text(text[:-1]+',"bad":NaN}');self.rejected()
    def test_oversized_report_rejected(self):
        self.path.write_bytes(b' '*(2*1024*1024+1));self.rejected()
    def test_pending_shipped_template_fails(self):
        path=TOOLS.parents[1]/'docs/codex-v1.0.3/templates/native-store-report.json'
        with self.assertRaises(gate.GateError):gate.validate(path,COMMIT,ZIP)

if __name__=='__main__':unittest.main()
