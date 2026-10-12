"""Diagnostic-only shadow of exact R's shell coordinator; never source acceptance."""
from datetime import datetime, timezone
from hashlib import sha256
import argparse
import importlib.util
import json
import os
from pathlib import Path, PurePosixPath
import re
import struct
import subprocess
import sys
import time
import unicodedata

R = '001669a051077536864b3f787ca0ec111fe06013'
DIGEST = 'a80aa2f1b08d97d9e13fff39f2b08b079d40977120efe72df32a86f2cd7c50c6'
ORIGINAL = 'b44333f809851edac7fd84d0c59bcda100c2a9e8528ac240ab06120e56568567'
PHASES = ('entry','host-policy-start','host-policy-end','roots-validated',
          'harness-import-start','harness-import-end','tool-import-start','tool-import-end',
          'parse-start','parse-end','scoped-analysis-start','scoped-analysis-end',
          'legacy-analysis-start','legacy-analysis-end','pester-config-start','pester-config-end',
          'pester-start','pester-end','later-native-start','later-native-end','serialize-start','serialize-end')

LOGGER = r'''# DIAGNOSTIC_ONLY: bounded observations; no test or assertion replacement.
$script:wbsDiagClock = [Diagnostics.Stopwatch]::StartNew()
$script:wbsDiagEvents = 0
$script:wbsDiagTrace = [IO.Path]::Combine([IO.Path]::GetFullPath($WorkRoot), 'phase.jsonl')
if ([IO.File]::Exists($script:wbsDiagTrace)) { throw 'Diagnostic trace already exists.' }
function ConvertTo-WbsDiagnosticJsonString {
    param([AllowEmptyString()][string]$Text)
    $builder = [Text.StringBuilder]::new()
    $null = $builder.Append('"')
    foreach ($character in $Text.ToCharArray()) {
        $code = [int]$character
        if ($code -eq 34) { $null = $builder.Append('\"') }
        elseif ($code -eq 92) { $null = $builder.Append('\\') }
        elseif ($code -lt 32) { $null = $builder.Append('\u').Append($code.ToString('x4',[Globalization.CultureInfo]::InvariantCulture)) }
        else { $null = $builder.Append($character) }
    }
    $null = $builder.Append('"')
    return $builder.ToString()
}
function Write-WbsDiagnosticPhase {
    param([ValidateSet('entry','host-policy-start','host-policy-end','roots-validated','harness-import-start','harness-import-end','tool-import-start','tool-import-end','parse-start','parse-end','scoped-analysis-start','scoped-analysis-end','legacy-analysis-start','legacy-analysis-end','pester-config-start','pester-config-end','pester-start','pester-end','later-native-start','later-native-end','serialize-start','serialize-end')][string]$Phase,
          [ValidateRange(0,4096)][int]$Index = 0, [AllowEmptyString()][string]$Label = '')
    if ($Label.Length -gt 256 -or $Label -match '[\x00-\x1f\x7f]') { throw 'Diagnostic label is out of bounds.' }
    $script:wbsDiagEvents += 1
    if ($script:wbsDiagEvents -gt 512) { throw 'Diagnostic phase event bound exceeded.' }
    # Pure .NET/literal JSON avoids priming PowerShell Utility/Security autoload.
    $invariant = [Globalization.CultureInfo]::InvariantCulture
    $json = '{"schema_version":1,"diagnostic_only":true,"phase":' + (ConvertTo-WbsDiagnosticJsonString $Phase) + ',"index":' + $Index.ToString($invariant) + ',"label":' + (ConvertTo-WbsDiagnosticJsonString $Label) + ',"utc":' + (ConvertTo-WbsDiagnosticJsonString ([DateTime]::UtcNow.ToString('o'))) + ',"elapsed_seconds":' + $script:wbsDiagClock.Elapsed.TotalSeconds.ToString('R',$invariant)
    if ($Phase -eq 'entry') {
        $json += ',"host":{"executable":' + (ConvertTo-WbsDiagnosticJsonString ([Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)) + ',"version":' + (ConvertTo-WbsDiagnosticJsonString $PSVersionTable.PSVersion.ToString()) + ',"edition":' + (ConvertTo-WbsDiagnosticJsonString $PSVersionTable.PSEdition) + ',"clr":' + (ConvertTo-WbsDiagnosticJsonString ([Environment]::Version.ToString())) + ',"process_bits":' + ([IntPtr]::Size*8).ToString($invariant) + '}'
    }
    $bytes = ([Text.UTF8Encoding]::new($false,$true)).GetBytes($json + "}`n")
    if ($bytes.Length -gt 2048) { throw 'Diagnostic phase record bound exceeded.' }
    $stream = [IO.File]::Open($script:wbsDiagTrace,[IO.FileMode]::OpenOrCreate,[IO.FileAccess]::Write,[IO.FileShare]::Read)
    try {
        if ($stream.Length + $bytes.Length -gt 262144) { throw 'Diagnostic phase trace bound exceeded.' }
        $null = $stream.Seek(0,[IO.SeekOrigin]::End)
        $stream.Write($bytes,0,$bytes.Length); $stream.Flush()
    } finally { $stream.Dispose() }
}
Write-WbsDiagnosticPhase 'entry'
'''

def shadow(source: bytes) -> bytes:
    if sha256(source).hexdigest() != ORIGINAL:
        raise ValueError('Only exact reviewed R shell coordinator bytes are accepted')
    text = source.decode('utf-8','strict')
    def replace(old, new):
        nonlocal text
        if text.count(old) != 1:
            raise ValueError('Diagnostic anchor does not occur exactly once')
        text = text.replace(old,new,1)
    replace("Set-StrictMode -Version Latest\n", LOGGER+"\nSet-StrictMode -Version Latest\n")
    replace("        execution_policy = @((Get-ExecutionPolicy -List) | ForEach-Object {\n            @{ scope = $_.Scope.ToString(); policy = $_.ExecutionPolicy.ToString() }\n        })",
            "        execution_policy = @(Write-WbsDiagnosticPhase 'host-policy-start'; (Get-ExecutionPolicy -List) | ForEach-Object {\n            @{ scope = $_.Scope.ToString(); policy = $_.ExecutionPolicy.ToString() }\n        }; Write-WbsDiagnosticPhase 'host-policy-end')")
    replace("    $reportValidated = $true\n", "    $reportValidated = $true\n    Write-WbsDiagnosticPhase 'roots-validated'\n")
    import_line="    Import-Module -Name (Join-Path -Path $repository -ChildPath 'tests\\powershell\\HarnessSupport.psm1') -Force -ErrorAction Stop\n"
    replace(import_line,"    Write-WbsDiagnosticPhase 'harness-import-start'\n"+import_line+"    Write-WbsDiagnosticPhase 'harness-import-end'\n")
    import_tool="        $module = Import-Module -Name $manifest -Force -PassThru -ErrorAction Stop\n"
    replace(import_tool,"        Write-WbsDiagnosticPhase 'tool-import-start' -Label $name\n"+import_tool+"        Write-WbsDiagnosticPhase 'tool-import-end' -Label $name\n")
    replace("    foreach ($path in @($application) + $scaffoldPaths) {\n", "    $wbsDiagParseIndex = 0\n    foreach ($path in @($application) + $scaffoldPaths) {\n        $wbsDiagParseIndex += 1\n        Write-WbsDiagnosticPhase 'parse-start' -Index $wbsDiagParseIndex -Label $path.Substring($repository.Length + 1)\n")
    syntax_line="        $outcomes.Add([pscustomobject]@{ name = 'syntax: ' + $path.Substring($repository.Length + 1); passed = (@($parseErrors).Count -eq 0) })\n"
    replace(syntax_line,syntax_line+"        Write-WbsDiagnosticPhase 'parse-end' -Index $wbsDiagParseIndex -Label $path.Substring($repository.Length + 1)\n")
    replace("    foreach ($path in $scaffoldPaths) {\n", "    $wbsDiagAnalysisIndex = 0\n    foreach ($path in $scaffoldPaths) {\n        $wbsDiagAnalysisIndex += 1\n        Write-WbsDiagnosticPhase 'scoped-analysis-start' -Index $wbsDiagAnalysisIndex -Label $path.Substring($repository.Length + 1)\n")
    analysis_line="        $outcomes.Add([pscustomobject]@{ name = 'scaffold analysis: ' + $path.Substring($repository.Length + 1); passed = ($findings.Count -eq 0) })\n"
    replace(analysis_line,analysis_line+"        Write-WbsDiagnosticPhase 'scoped-analysis-end' -Index $wbsDiagAnalysisIndex -Label $path.Substring($repository.Length + 1)\n")
    replace("    $report.legacy_application_analysis = @(Invoke-ScriptAnalyzer", "    Write-WbsDiagnosticPhase 'legacy-analysis-start'\n    $report.legacy_application_analysis = @(Invoke-ScriptAnalyzer")
    legacy_line="    $outcomes.Add([pscustomobject]@{ name = 'legacy application findings recorded (observation only)'; passed = $true })\n"
    replace(legacy_line,legacy_line+"    Write-WbsDiagnosticPhase 'legacy-analysis-end'\n    Write-WbsDiagnosticPhase 'pester-config-start'\n")
    replace("    $config.Output.Verbosity = 'None'\n", "    $config.Output.Verbosity = 'Detailed'\n    $config.Debug.ShowStartMarkers = $true\n    $config.Output.RenderMode = 'Plaintext'\n")
    invoke="    $pesterResult = Invoke-Pester -Configuration $config\n"
    replace(invoke,"    Write-WbsDiagnosticPhase 'pester-config-end' -Index $containers.Count\n    Write-WbsDiagnosticPhase 'pester-start'\n"+invoke+"    Write-WbsDiagnosticPhase 'pester-end' -Index $pesterResult.TotalCount\n")
    native="    & $env:ComSpec /d /c 'exit 0'\n"
    replace(native,"    Write-WbsDiagnosticPhase 'later-native-start'\n"+native)
    replace("    $nativeExitCode = $LASTEXITCODE\n", "    $nativeExitCode = $LASTEXITCODE\n    Write-WbsDiagnosticPhase 'later-native-end'\n")
    replace("        $json = $report | ConvertTo-Json -Depth 12\n", "        Write-WbsDiagnosticPhase 'serialize-start'\n        $json = $report | ConvertTo-Json -Depth 12\n")
    replace("        finally { $stream.Dispose() }\n", "        finally { $stream.Dispose() }\n        Write-WbsDiagnosticPhase 'serialize-end'\n")
    return text.encode('utf-8')

def main() -> int:
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--tool-root',type=Path,required=True)
    parser.add_argument('--reports',type=Path,required=True)
    args=parser.parse_args()
    repo=Path.cwd().resolve();reports=args.reports
    if sys.version_info[:3]!=(3,14,8) or struct.calcsize('P')!=8 or not sys.flags.isolated or not sys.dont_write_bytecode:
        raise ValueError('Diagnostic requires the same isolated regular x64 developer Python3.14.8')
    env={key:value for key,value in os.environ.items() if not key.upper().startswith('GIT_')}
    env.update(GIT_NO_REPLACE_OBJECTS='1',GIT_TERMINAL_PROMPT='0')
    def git(*arguments):
        result=subprocess.run(['git',*arguments],cwd=repo,env=env,stdin=subprocess.DEVNULL,capture_output=True,timeout=60,check=True)
        if len(result.stdout)>32*1024*1024:raise ValueError('Diagnostic Git read exceeds bound')
        return result.stdout
    if git('rev-parse','HEAD').decode().strip()!=R or git('status','--porcelain','--untracked-files=all')!=b'':
        raise ValueError('Diagnostic checkout must be clean exact R')
    if git('remote','get-url','origin').decode().strip().removesuffix('.git')!='https://github.com/PikkuJanne/WinBookSplit':
        raise ValueError('Diagnostic repository origin differs')
    runner_path=repo/'tests/run_tests.py'
    if runner_path.read_bytes()!=git('show',R+':tests/run_tests.py'):raise ValueError('Diagnostic runner bytes differ')
    spec=importlib.util.spec_from_file_location('exact_R_diagnostic_runner',runner_path)
    runner=importlib.util.module_from_spec(spec);spec.loader.exec_module(runner)
    source=runner.source_manifest()
    if len(source)!=207 or runner.manifest_digest(source)!=DIGEST:raise ValueError('Diagnostic207-path source differs')
    for name,digest in source.items():
        if type(name) is not str or any(c in name for c in '\\:<>"|?*') or any(unicodedata.category(c) in {'Cc','Cf','Cs'} for c in name) or PurePosixPath(name).as_posix()!=name or any(part in ('','.','..') for part in name.split('/')):
            raise ValueError('Diagnostic source name is not canonical')
        if sha256(git('show',R+':'+name)).hexdigest()!=digest:raise ValueError('Diagnostic source differs from R Git blob')
    if not reports.is_absolute() or not reports.is_dir() or reports.is_symlink() or reports.is_junction() or reports.is_relative_to(repo):raise ValueError('Diagnostic reports need owned external ordinary directory')
    work=reports.parent/'work';work.mkdir(exist_ok=False)
    original=(repo/'tests/Invoke-ShellTests.ps1').read_bytes();copy=shadow(original)
    script=work/'Invoke-ShellTests.DIAGNOSTIC_ONLY.ps1'
    with script.open('xb') as stream:stream.write(copy)
    inbox=Path(os.environ['SystemRoot'])/'System32/WindowsPowerShell/v1.0/powershell.exe'
    command,environment=runner.shell_command(inbox,work,work/'report.json',args.tool_root,False)
    environment['WBS_TEST_SCRIPT']=str(script)
    # Exact maintained fixture preparation; supplied pinned wheelhouse only.
    if not os.environ.get('WBS_TEST_WHEELHOUSE'):raise ValueError('Diagnostic requires the verified supplied offline wheelhouse')
    wheelhouse,wheel=runner.prepare_setup_test_wheel(work)
    if wheelhouse is None or wheel.get('exit_code')!=0:raise ValueError('Diagnostic fixture wheel validation failed')
    environment['WBS_TEST_WHEELHOUSE']=str(wheelhouse)
    started=datetime.now(timezone.utc).isoformat();clock=time.monotonic()
    result=runner.run_command(command,work,environment=environment,timeout=300)
    record={'schema_version':1,'kind':'DIAGNOSTIC_ONLY_SHADOW','acceptance_claimed':False,'source_commit':R,'source_digest':DIGEST,'source_sha256':source,'original_shell_coordinator_sha256':sha256(original).hexdigest(),'shadow_sha256':sha256(copy).hexdigest(),'helper_sha256':sha256(Path(__file__).read_bytes()).hexdigest(),'orchestration_commit':os.environ.get('GITHUB_SHA'),'run_id':os.environ.get('GITHUB_RUN_ID'),'run_attempt':os.environ.get('GITHUB_RUN_ATTEMPT'),'started_at_utc':started,'finished_at_utc':datetime.now(timezone.utc).isoformat(),'elapsed_seconds':time.monotonic()-clock,'maintained_timeout_seconds':300,'source_unchanged':runner.source_manifest()==source,'clean_R_after':git('rev-parse','HEAD').decode().strip()==R and git('status','--porcelain','--untracked-files=all')==b'','native_exit_code':None if result['timed_out'] else result['exit_code'],'transport':result,'fixture_wheel':wheel,'pester_diagnostic_configuration':{'verbosity':'Detailed','show_start_markers':True,'render_mode':'Plaintext','filters_added':False,'assertions_changed':False},'phase_trace':None,'completed_shadow_report':None,'meaning':'Timing/localization only; original required CI failures remain failures; no source or release acceptance'}
    phase=work/'phase.jsonl'
    if phase.is_file():
        data=phase.read_bytes()
        if len(data)>262144:raise ValueError('Diagnostic phase trace exceeds bound')
        events=[json.loads(line) for line in data.splitlines()]
        if len(events)>512 or any(event['phase'] not in PHASES for event in events):raise ValueError('Diagnostic phase events differ')
        with (reports/'phase.jsonl').open('xb') as stream:stream.write(data)
        record['phase_trace']={'bytes':len(data),'sha256':sha256(data).hexdigest(),'events':events}
    child=work/'report.json'
    if child.is_file():
        data=child.read_bytes()
        if len(data)>64*1024*1024:raise ValueError('Diagnostic child report exceeds bound')
        with (reports/'shadow-report.json').open('xb') as stream:stream.write(data)
        record['completed_shadow_report']={'bytes':len(data),'sha256':sha256(data).hexdigest(),'value':json.loads(data)}
    output=reports/'diagnostic.json'
    with output.open('x',encoding='utf-8',newline='\n') as stream:json.dump(record,stream,indent=2)
    print(json.dumps({'kind':record['kind'],'acceptance_claimed':False,'runner_exit_code':result['exit_code'],'timed_out':result['timed_out'],'native_exit_code':record['native_exit_code'],'phase_trace':None if record['phase_trace'] is None else {'events':len(record['phase_trace']['events']),'sha256':record['phase_trace']['sha256']},'report':str(output),'source_unchanged':record['source_unchanged']}))
    return 0 if result['exit_code']==0 and record['source_unchanged'] and record['clean_R_after'] else 1

if __name__=='__main__':raise SystemExit(main())
