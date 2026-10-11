"""Synthetic final-ZIP binding controls; never native acceptance evidence."""
import copy
from hashlib import sha256
import importlib.util
import json
from pathlib import Path
from types import SimpleNamespace
import tempfile
import unittest
from unittest.mock import patch

from pypdf import PdfReader, PdfWriter

ROOT=Path(__file__).resolve().parents[2]
spec=importlib.util.spec_from_file_location('wbs_ps7_adapter_controls',ROOT/'tools/compatibility/final_zip_ps7.py')
adapter=importlib.util.module_from_spec(spec);spec.loader.exec_module(adapter)
spec=importlib.util.spec_from_file_location('wbs_ps7_existing_oracle_controls',ROOT/'tools/compatibility/native_smoke.py')
smoke=importlib.util.module_from_spec(spec);spec.loader.exec_module(smoke)
spec=importlib.util.spec_from_file_location('wbs_ps7_existing_path_guards',ROOT/'engine/winbooksplit_windows.py')
windows=importlib.util.module_from_spec(spec);spec.loader.exec_module(windows)


def synthetic_probe(stdout=''):
    return {'Pid':81,'ExitCode':0,'JobAssigned':True,'ParentStopped':True,'DescendantsStopped':True,
        'StreamsComplete':True,'InputWriterStopped':True,'TimedOut':False,'Cancelled':False,
        'StdoutTruncated':False,'StderrTruncated':False,'StartError':None,'StopError':None,'StreamError':None,
        'ResultError':None,'InteractionError':None,'InputError':None,'Stdout':stdout,'Stderr':'',
        'StdoutTotalBytes':len(stdout.encode('utf-8')),'StderrTotalBytes':0,'ElapsedSeconds':0.1}


class Ps7BindingControls(unittest.TestCase):
    def setUp(self):
        self.temporary=tempfile.TemporaryDirectory();self.addCleanup(self.temporary.cleanup)
        self.root=Path(self.temporary.name).resolve();self.app=self.root/'application';self.app.mkdir()
        self.kit=self.root/'kit';self.kit.mkdir();fixtures=self.kit/'fixtures';fixtures.mkdir()
        self.work=self.root/'work';self.work.mkdir()
        self.pwsh=self.root/'pwsh.exe';self.pwsh.write_bytes(b'SYNTHETIC file; never executed')
        self.python=self.app/'.venv/Scripts/python.exe';self.python.parent.mkdir(parents=True);self.python.write_bytes(b'SYNTHETIC file; never executed')
        self.package=self.app/'.venv/Lib/site-packages/pypdf/__init__.py';self.package.parent.mkdir(parents=True);self.package.write_bytes(b'# SYNTHETIC package origin; never imported')
        self.bindings={'production':{'version':'1.0.1','policy_sha256':'a'*64},'fixtures':{}}
        for name,count in (('simple-10.pdf',10),('nested-12.pdf',12)):
            writer=PdfWriter()
            for index in range(count):writer.add_blank_page(width=101+index,height=113)
            with (fixtures/name).open('xb') as stream:writer.write(stream)
            self.bindings['fixtures'][name]={'sha256':sha256((fixtures/name).read_bytes()).hexdigest()}
        self.runtime={'executable':str(self.python),'version':'3.14.8','pypdf_version':'6.20.0','pypdf_path':str(self.package)}
        self.original={'python_path':str(self.python),'runtime':self.runtime}
        self.prior={'os':{'Caption':'Microsoft Windows 11 Pro','Version':'10.0.26300','BuildNumber':'26300',
                    'ProductType':1,'processor_architectures':[9],'os_bits':64,'process_bits':64}}
        self.runtime_path=self.root/'runtime.json';self.runtime_path.write_text(json.dumps(synthetic_probe(json.dumps(self.runtime))),encoding='utf-8')
        self.args=SimpleNamespace(kit_root=self.kit,work_root=self.work,pwsh_path=self.pwsh,runtime_probe=self.runtime_path,
            bindings_sha256='b'*64,prior_report_sha256='c'*64,expected_pwsh_version='7.6.6')
        self.context=(self.bindings,self.original,self.app,[],smoke,windows,self.prior)

    def prepare(self):
        with patch.object(adapter,'context',return_value=self.context),patch.object(smoke,'compatibility',SimpleNamespace(validate_interpreter=lambda record,policy:{'ok':True})),patch.object(smoke,'RUNTIME_POLICY',{'pypdf':{'version':'6.20.0'}}):
            return adapter.prepare(self.args)

    def test_positive_controlled_prepare_preserves_prior_os_and_actual_fixture_bytes(self):
        plan=self.prepare()
        self.assertEqual(plan['prior_os'],self.prior['os'])
        self.assertEqual([case['id'] for case in plan['cases']],list(adapter.IDS))
        for case,count in ((plan['cases'][2],10),(plan['cases'][3],12)):
            source=Path(case['source_path'])
            fixture=self.kit/'fixtures'/('simple-10.pdf' if count==10 else 'nested-12.pdf')
            self.assertEqual(source.read_bytes(),fixture.read_bytes())
            self.assertEqual(len(PdfReader(source).pages),count)
            self.assertEqual((Path(case['output_base'])/'prior-output.pdf').read_bytes(),b'Authored prior; preserve exact bytes.\n')
        self.assertEqual(plan['cases'][2]['arguments'][-2:],['-StartPages','1,4,7'])
        self.assertIn('-PythonPath',plan['cases'][3]['arguments'])

    def test_prepare_rejects_changed_fixture_before_adoption(self):
        fixture=self.kit/'fixtures/simple-10.pdf';fixture.write_bytes(b'altered')
        with self.assertRaisesRegex(ValueError,'fixture_pin_differ'):self.prepare()
        self.assertFalse(any(self.work.rglob('Original*')))

    def test_prepare_rejects_changed_runtime_without_creating_case_directories(self):
        bad={**self.runtime,'version':'3.13.16'}
        self.runtime_path.write_text(json.dumps(synthetic_probe(json.dumps(bad))),encoding='utf-8')
        with self.assertRaisesRegex(ValueError,'runtime_changed'):self.prepare()
        self.assertEqual(list(self.work.iterdir()),[])

    def test_prepare_rejects_unproved_runtime_transport(self):
        row=synthetic_probe(json.dumps(self.runtime));del row['StopError']
        self.runtime_path.write_text(json.dumps(row),encoding='utf-8')
        with self.assertRaisesRegex(smoke.OracleError,'native_transport_error'):self.prepare()
        self.assertEqual(list(self.work.iterdir()),[])

    def test_prepare_rejects_foreign_work_content_without_writes(self):
        marker=self.work/'foreign.txt';marker.write_bytes(b'preserve')
        with self.assertRaisesRegex(ValueError,'fresh_external_work_required'):self.prepare()
        self.assertEqual(list(self.work.iterdir()),[marker]);self.assertEqual(marker.read_bytes(),b'preserve')

    def test_strict_json_duplicate_and_nonfinite_refused(self):
        for text in ('{"x":1,"x":2}','{"x":1,"X":2}','{"x":NaN}','{"x":Infinity}'):
            path=self.root/'bad.json';path.write_text(text,encoding='utf-8')
            with self.subTest(text=text),self.assertRaises(ValueError):adapter.read(path)

    def test_helper_digest_and_typed_length_required_before_any_execution(self):
        path=self.root/'never_execute.py';path.write_bytes(b'raise AssertionError("MUST NEVER EXECUTE")\n')
        row={'bytes':path.stat().st_size,'sha256':sha256(path.read_bytes()).hexdigest()}
        for bad in ({**row,'sha256':'0'*64},{**row,'bytes':True},{**row,'unexpected':True}):
            with self.subTest(bad=bad),self.assertRaisesRegex(ValueError,'helper_pin_differ'):
                adapter.load_pinned(path,bad,'wbs_must_never_execute')

    def test_bounded_input_refuses_relative_hardlinked_and_oversize_files(self):
        with self.assertRaisesRegex(ValueError,'ordinary_path_required'):adapter.bounded(Path('relative.json'))
        source=self.root/'source.json';source.write_bytes(b'{}');alias=self.root/'alias.json';alias.hardlink_to(source)
        with self.assertRaisesRegex(ValueError,'bounded_ordinary_file_required'):adapter.bounded(alias)
        big=self.root/'big.json'
        with big.open('wb') as stream:stream.truncate(adapter.MAXIMUM+1)
        with self.assertRaisesRegex(ValueError,'bounded_ordinary_file_required'):adapter.bounded(big)

    def test_fixture_is_consumed_once_under_real_read_only_guard(self):
        fixture=self.kit/'fixtures/simple-10.pdf';original=fixture.read_bytes();calls=[]
        bounded=adapter.bounded
        def guarded_read(path):
            if Path(path)==fixture:
                calls.append(path)
                with self.assertRaises(OSError):fixture.write_bytes(b'concurrent change')
            return bounded(path)
        with patch.object(adapter,'bounded',side_effect=guarded_read):plan=self.prepare()
        self.assertEqual(calls,[fixture]);self.assertEqual(fixture.read_bytes(),original)
        self.assertEqual(Path(plan['cases'][2]['source_path']).read_bytes(),original)

    def test_fixture_observation_is_taken_under_the_same_verified_read_guard(self):
        fixture=self.kit/'fixtures/simple-10.pdf';original=fixture.read_bytes();calls=[]
        observe=smoke.observation
        def guarded_observation(path):
            if Path(path)==fixture:
                calls.append(path)
                with self.assertRaises(OSError):fixture.write_bytes(b'late concurrent change')
            return observe(path)
        with patch.object(smoke,'observation',side_effect=guarded_observation):plan=self.prepare()
        self.assertEqual(calls,[fixture]);self.assertEqual(fixture.read_bytes(),original)
        recorded=next(item for item in plan['immutable_files'] if item['path']==str(fixture))
        self.assertEqual(recorded['size_bytes'],len(original));self.assertEqual(recorded['sha256'],sha256(original).hexdigest())

    def test_nonzero_runtime_probe_with_plausible_json_refused(self):
        row=synthetic_probe(json.dumps(self.runtime));row['ExitCode']=1
        self.runtime_path.write_text(json.dumps(row),encoding='utf-8')
        with self.assertRaisesRegex(ValueError,'runtime_probe_failed'):self.prepare()
        self.assertEqual(list(self.work.iterdir()),[])

    def test_prior_report_requires_complete_same_zip_win11_acceptance(self):
        digest='b'*64
        prior={'protocol':'winbooksplit.native-acceptance','schema_version':1,'success':True,'cleanup_safe':True,
            'identities':{'bindings_sha256':digest,'production':self.bindings['production']},
            'os':{'platform_scope':'NATIVE_WINDOWS11_DRY_RUN'},'oracle':{'cases':[{'id':name,'status':'PASS'} for name in smoke.CASE_IDS]}}
        adapter.accepted_prior(prior,self.bindings,digest,smoke.CASE_IDS)
        mutations=[lambda row:row.update(success=False),lambda row:row.update(schema_version=True),
            lambda row:row.update(cleanup_safe=False),lambda row:row['identities'].update(bindings_sha256='d'*64),
            lambda row:row['identities'].update(production={'version':'foreign'}),
            lambda row:row['os'].update(platform_scope='NATIVE_WINDOWS10'),
            lambda row:row['oracle']['cases'][0].update(status='NOT_RUN'),lambda row:row['oracle']['cases'].pop()]
        for change in mutations:
            bad=copy.deepcopy(prior);change(bad)
            with self.subTest(change=change),self.assertRaisesRegex(ValueError,'accepted_same_zip_win11_receipt_required'):
                adapter.accepted_prior(bad,self.bindings,digest,smoke.CASE_IDS)

    def test_actual_shell_version_path_and_coherent_workstation_facts_required(self):
        facts={**self.prior['os'],'shell_executable':str(self.pwsh),'powershell_version':'7.6.6','clr':'10.0.7'}
        plan={'prior_os':self.prior['os'],'pwsh_path':str(self.pwsh),'expected_pwsh_version':'7.6.6'}
        adapter.native_facts(facts,plan,self.prior,'7.6.6')
        deltas=({'powershell_version':'7.6.5'},{'powershell_version':'7.7.0-preview.1'},
                {'shell_executable':str(self.root/'different.exe')},{'Caption':'Microsoft Windows 11 Pro PRIVATE_HOST'},
                {'ProductType':3},{'processor_architectures':[12]},{'Version':'10.0.99999'},
                {'process_bits':True},{'os_bits':32},{'clr':'PRIVATE_PROFILE'})
        for delta in deltas:
            with self.subTest(delta=delta),self.assertRaises(ValueError):adapter.native_facts({**facts,**delta},plan,self.prior,'7.6.6')
        with self.assertRaisesRegex(ValueError,'same_native_workstation_required'):
            adapter.native_facts(facts,{**plan,'prior_os':{}},self.prior,'7.6.6')

    def test_literal_process_command_cwd_answers_and_stop_proof_cannot_be_rebound(self):
        case={'arguments':['-NoProfile','-File',str(self.app/'WinBookSplit.ps1'),'-Version'],'cwd':str(self.work)}
        plan={'pwsh_path':str(self.pwsh)}
        row={'probe':synthetic_probe(),'command':[str(self.pwsh),*case['arguments']],'cwd':str(self.work),'answers':[]}
        adapter.native_row(row,case,plan,smoke)
        mutations=[lambda item:item['command'].__setitem__(0,str(self.root/'other.exe')),
            lambda item:item['command'].append('-Preview'),lambda item:item.update(cwd=str(self.root)),
            lambda item:item.update(answers=['Y']),lambda item:item['probe'].update(ParentStopped=False),
            lambda item:item['probe'].pop('StopError'),lambda item:item['probe'].update(StdoutTotalBytes=True),
            lambda item:item['probe'].update(ExitCode=1),lambda item:item['probe'].update(Stderr='error',StderrTotalBytes=5)]
        for mutate in mutations:
            bad=copy.deepcopy(row);mutate(bad)
            with self.subTest(mutate=mutate),self.assertRaises(ValueError):adapter.native_row(bad,case,plan,smoke)

    def test_safe_report_omits_private_plan_runtime_and_os_extras_and_rejects_row_leakage(self):
        facts={**self.prior['os'],'powershell_version':'7.6.6','clr':'10.0.7','shell_executable':'C:/Users/SECRET_PROFILE/pwsh.exe','CSName':'SECRET_HOST'}
        plan={'bindings_sha256':'b'*64,'prior_report_sha256':'c'*64,'pwsh_sha256':'d'*64,'runtime':{**self.runtime,'private':'PRIVATE_DOCUMENT'},
            'pwsh_path':'C:/Users/SECRET_PROFILE/pwsh.exe','application_root':'PRIVATE_DOCUMENT'}
        rows=[{'id':name,'status':'PASS','exit_code':0,'probe_sha256':'e'*64} for name in adapter.IDS]
        result=adapter.safe_result(self.bindings,plan,facts,rows,'f'*64,'1'*64)
        text=json.dumps(result)
        for marker in ('SECRET_PROFILE','SECRET_HOST','PRIVATE_DOCUMENT',str(self.python),str(self.package)):
            self.assertNotIn(marker,text)
        for delta in ({'private':'PRIVATE_DOCUMENT'},{'status':'NOT_RUN'},{'exit_code':False},{'probe_sha256':'PRIVATE_PATH'}):
            bad=copy.deepcopy(rows);bad[0].update(delta)
            with self.subTest(delta=delta),self.assertRaisesRegex(ValueError,'safe_case_table_differ'):
                adapter.safe_result(self.bindings,plan,facts,bad,'f'*64,'1'*64)


if __name__=='__main__':unittest.main(verbosity=2)
