"""Synthetic receipt controls; no native Server or Windows10 execution claim."""
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

ROOT=Path(__file__).resolve().parents[2]
MODULE_PATH=ROOT/'tools/ci/validate_reports.py'
SPEC=importlib.util.spec_from_file_location('wbs_ci_reports',MODULE_PATH)
checker=importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(checker)

class RequiredReportControls(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory(prefix='wbs-ci-report-controls-')
        self.addCleanup(self.temp.cleanup)
        self.root=Path(self.temp.name)
        self.commit='a'*40
        source={'base_commit':self.commit,'source_unchanged':True}
        self.records={
            'os.json':{'source_commit':self.commit,'caption':'Synthetic Windows Server','version':'10.0.26100','build':'26100','architecture':'64-bit','product_type':3,'process_bits':64,'powershell':'7.6.6','clr':'10.0.12'},
            'package-inputs.json':{'success':True,'exit_code':0,'source_commit':self.commit,'working_tree_inputs_used':False,'package_built':False},
            'tools-bootstrap.json':{'success':True,'exit_code':0,'result':'CI_TOOL_BOOTSTRAP_PASSED'},
            'python.json':{'success':True,'source':source,'synthetic_input_unchanged':True,'owned_work_directory_removed':True,'layer':'python','steps':[{'name':'stdlib-unittest','exit_code':0,'timed_out':False,'stderr':'Ran 575 tests in 300.0s\n\nOK\n'}]}}
        for name in ('shell-ps51.json','shell-ps7.json'):
            self.records[name]={'success':True,'layer':'shell','source':dict(source),'synthetic_input_unchanged':True,'owned_work_directory_removed':True,
                'steps':[{'name':'setup-test-wheel','exit_code':0}, {'name':'shell-0','exit_code':0,'timed_out':False,'evidence':{'host':{'version':'5.1.26100.33438' if name=='shell-ps51.json' else '7.6.6'},'pester':{'total':182,'passed':182,'failed':0,'skipped':0,'not_run':0,'failed_containers':0},'syntax':[{'path':'WinBookSplit.ps1','errors':[]}],'scaffold_analysis':[{'path':'tests/Invoke-ShellTests.ps1','findings':[]}]}}]}
        self.write()

    def write(self):
        for name,record in self.records.items(): (self.root/name).write_text(json.dumps(record),encoding='utf-8')

    def check(self): return checker.validate('reference',self.commit,self.root)

    def test_reference_requires_every_declared_report(self):
        self.assertEqual(len(self.check()),6)
        for name in self.records:
            (self.root/name).unlink()
            with self.assertRaises(ValueError): self.check()
            self.write()

    def test_old_source_and_workstation_cannot_supply_server_pass(self):
        for name,key,value in [('os.json','source_commit','b'*40),('os.json','product_type',1),('os.json','product_type',True),('package-inputs.json','source_commit','b'*40),('package-inputs.json','exit_code',False)]:
            before=self.records[name][key]
            self.records[name][key]=value
            self.write()
            with self.assertRaises(ValueError): self.check()
            self.records[name][key]=before
        self.records['python.json']['source']['base_commit']='b'*40
        self.write()
        with self.assertRaises(ValueError): self.check()

    def test_partial_failed_skipped_or_timed_out_suites_fail(self):
        probes=[('python.json','success',False),('tools-bootstrap.json','success',False)]
        for name,key,value in probes:
            before=self.records[name][key]
            self.records[name][key]=value
            self.write()
            with self.assertRaises(ValueError): self.check()
            self.records[name][key]=before
        for key in ('failed','skipped','not_run','failed_containers'):
            pester=self.records['shell-ps51.json']['steps'][1]['evidence']['pester']
            pester[key]=1
            self.write()
            with self.assertRaises(ValueError): self.check()
            pester[key]=0
        self.records['python.json']['steps'][0]['timed_out']=True
        self.write()
        with self.assertRaises(ValueError): self.check()

    def test_duplicate_keys_and_nonfinite_json_cannot_reseal_result(self):
        for text in ('{"success": false,"success":true}','{"value":NaN}'):
            (self.root/'python.json').write_text(text,encoding='utf-8')
            with self.assertRaises(ValueError): self.check()

    def test_application_target_partial_cases_and_workstation_substitution_fail(self):
        roster=('test_compatibility','test_filenames','test_interactive','test_pdf_policy','test_functional_pdf_repair','test_pdf_fidelity','test_conversion','test_outcomes','test_version','test_warnings')
        candidate=str(Path('C:/synthetic-runtime/python.exe'))
        units={'protocol':'winbooksplit.application-units','version':1,'success':True,'required_modules':list(roster),
               'modules':[{'name':name,'test_count':1} for name in roster],'python_executable':candidate,'isolated':True,'no_bytecode':True,'loading_error':None,
               'python_version':'3.13.16','tests_run':10,'skipped':[], 'failures':0,'errors':0,'expected_failures':0,'unexpected_successes':0}
        cases=[{'id':name,'status':'PASS'} for name in ('manual-front-matter','auto-level1','auto-level2','late-invalid-token','no-bookmarks','zero-page','wrapper-PS5','wrapper-PS7')]
        matrix={'success':True,'cleanup_safe':True,'timed_out':False,'application_runtime':{'candidate':candidate,'full_probe':{'version':'3.13.16'}},
                'runtime_only_units':units,'hosts':[{'product_type':3,'build':'26100','shell_version':'5.1.26100.33438'},{'product_type':3,'build':'26100','shell_version':'7.6.6'}],'cases':cases}
        self.records['application.json']={'success':True,'source':{'base_commit':self.commit,'source_unchanged':True},
            'synthetic_input_unchanged':True,'owned_work_directory_removed':True,
            'steps':[{'name':'application-matrix','exit_code':0,'timed_out':False,'evidence':matrix}]}
        self.write()
        for target in ('3.13.16','3.14.8','3.15.0'):
            units['python_version']=target
            matrix['application_runtime']['full_probe']['version']=target
            self.write()
            with self.subTest(target=target):
                checker.validate('application',self.commit,self.root,target)
            substitute='3.14.8' if target!='3.14.8' else '3.13.16'
            controls=[(units,'python_version',substitute),(units,'skipped',['missing module']),
                      (units,'tests_run',True),(units,'failures',False),
                      (matrix,'cleanup_safe',False),(matrix['application_runtime']['full_probe'],'version',substitute),
                      (matrix['hosts'][1],'product_type',1),(matrix['hosts'][1],'product_type',True),
                      (matrix['hosts'][0],'shell_version','7.6.6'),(units,'modules',units['modules'][:-1]),
                      (matrix,'cases',cases[:-1]),(matrix,'cases',cases+[cases[-1]])]
            for record,key,value in controls:
                before=record[key]
                record[key]=value
                self.write()
                with self.subTest(target=target,field=key),self.assertRaises(ValueError):
                    checker.validate('application',self.commit,self.root,target)
                record[key]=before
        self.write()
        for target in (None,True,'3.15.0rc1','3.16.0','3.15.99','3.14.1'):
            with self.subTest(unpinned_target=target),self.assertRaisesRegex(ValueError,'Exact matrix test target required'):
                checker.validate('application',self.commit,self.root,target)
        original_read=checker.read
        for targets in (None,'3.15.0',[],[True],['3.15.0','3.15.0'],['3.15.0rc1'],['03.15.0']):
            # SYNTHETIC source-lock corruption, not a change to tracked files.
            def controlled_read(path):
                if path==ROOT/'tests/compatibility/runtime-lock.json': return {'python_versions':targets}
                return original_read(path)
            with self.subTest(malformed_lock=targets),patch.object(checker,'read',side_effect=controlled_read),self.assertRaisesRegex(ValueError,'Exact source matrix lock required'):
                checker.validate('application',self.commit,self.root,'3.15.0')

    def test_boolean_counts_duplicate_cases_and_missing_native_deadline_fail(self):
        for name in ('shell-ps51.json','shell-ps7.json'):
            pester=self.records[name]['steps'][1]['evidence']['pester']
            before=pester.copy()
            pester.update(total=True,passed=True)
            self.write()
            with self.assertRaises(ValueError): self.check()
            pester.update(before)
        del self.records['python.json']['steps'][0]['timed_out']
        self.write()
        with self.assertRaises(ValueError): self.check()

    def test_empty_coverage_and_wrong_declared_shell_host_fail(self):
        evidence=self.records['shell-ps51.json']['steps'][1]['evidence']
        for key in ('syntax','scaffold_analysis'):
            before=evidence[key]
            evidence[key]=[]
            self.write()
            with self.assertRaises(ValueError): self.check()
            evidence[key]=before
        evidence['host']['version']='7.6.6'
        self.write()
        with self.assertRaises(ValueError): self.check()

if __name__=='__main__': unittest.main(verbosity=2)
