[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$RepositoryRoot,[Parameter(Mandatory=$true)][string]$TestPython,
      [Parameter(Mandatory=$true)][string]$WorkRoot,[Parameter(Mandatory=$true)][string]$ReportPath)
$ErrorActionPreference='Stop'
foreach($name in @('Logging','Diagnostics','Process','Support','Paths','Runtime')){. ([IO.Path]::Combine($RepositoryRoot,'engine','WinBookSplit.'+$name+'.ps1'))}
foreach($name in @('Invoke-NativeAcceptance.ps1','Invoke-NativeDiscoveryCollector.ps1','Invoke-StoreExistingAcceptance.ps1')){. ([IO.Path]::Combine($RepositoryRoot,'tools','compatibility',$name))}
Initialize-WinBookSplitLogTypes;Initialize-WinBookSplitSupport;Initialize-WinBookSplitProcess
$null=New-NativeAcceptanceDirectory $WorkRoot
$cases=New-Object 'Collections.Generic.List[object]'
function Require($Value,[string]$Message){if(-not $Value){throw $Message}}
function Case([string]$Id,[scriptblock]$Body){$row=@{id=$Id;scope='SYNTHETIC_MANAGER_SEAMS_ACTUAL_WINDOWS_TOOLING_ONLY';status='FAIL';error=$null};try{&$Body;$row.status='PASS'}catch{$row.error=$_.Exception.Message};$cases.Add($row)}
$nativeInventory=${function:Get-NativeStoreExistingInventory}
$nativeQuery=${function:Get-WinBookSplitRegisteredManagerPaths}
$nativeListing=${function:Invoke-WinBookSplitManagerListing}
$nativeProcess=${function:Invoke-NativeAcceptanceProcessCase}
$nativeManagerState=${function:Get-NativeStoreExistingManagerState}
$nativePackages=${function:Get-NativeStoreExistingPackages}
$nativeCandidate=${function:Test-WinBookSplitPythonCandidate}
$nativeDownload=${function:Get-NativeAcceptanceDownload};$nativeExpand=${function:Expand-NativeAcceptancePython}
Case 'store-inputs-refuse-explicit-baseline-and-provisioning' {
    foreach($options in @(@{BarePythonPath=$TestPython},@{BaselinePythonPath=$TestPython},@{ProvisionPrerequisites=$true})){
        $failed=$false;try{Assert-NativeStoreExistingInputs @options}catch{$failed=$true};Require $failed 'Forbidden explicit/provisioned Store input accepted.'
    }
}
Case 'store-inputs-refuse-runtime-environment' {
    foreach($key in @('VIRTUAL_ENV','PYTHONHOME','PYTHONPATH','PYMANAGER_CONFIG','WBS_TEST_PYTHON')){
        $old=[Environment]::GetEnvironmentVariable($key,'Process');try{
            [Environment]::SetEnvironmentVariable($key,'AUTHORED_RUNTIME_INJECTION','Process')
            $failed=$false;try{Assert-NativeStoreExistingInputs}catch{$failed=$true};Require $failed 'Runtime injection was accepted.'
        }finally{[Environment]::SetEnvironmentVariable($key,$old,'Process')}
    }
}
Case 'redaction-parses-nested-json-case-and-key-paths' {
    $tokens=[ordered]@{};$tokens['C:\Users\Authored Å日']='<USER>'
    $value=@{'C:\USERS\AUTHORED Å日\dictionary'=@{probe=@{Stdout='{"path":"C:\\Users\\Authored \u00c5\u65e5\\runtime.exe"}'};frame='[SETUP] {"path":"c:\\users\\authored Å日\\app"}'}}
    $safe=ConvertTo-NativeStoreSafeValue $value $tokens;$text=ConvertTo-Json -InputObject $safe -Depth 20 -Compress
    Require (-not $text.Contains('Authored') -and -not $text.Contains('authored') -and $text.Contains('USER')) 'Returned redacted bytes retained a path spelling/JSON/key.'
}
Case 'actual-redacted-stream-lengths-pass-rejecting-validator' {
    $run=Invoke-NativeAcceptanceProcessCase $TestPython @('-I','-B','-c','import json,sys;print(json.dumps({"path":sys.executable}))') $WorkRoot 10
    Assert-NativeAcceptanceProcess $run.probe;Require ($run.probe.ExitCode-eq 0) 'Actual redaction positive child failed.'
    $tokens=[ordered]@{};$tokens[$TestPython]='<RUNTIME>'
    $safe=ConvertTo-NativeStoreSafeValue @{role='application';environment=@{};process=$run} $tokens
    Require ($safe.process.probe.StdoutTotalBytes-ne $safe.process.probe.redacted_stream_bytes.stdout) 'Actual redaction control did not change native stdout length.'
    $path=[IO.Path]::Combine($WorkRoot,'redacted-application-record.json');Write-NativeAcceptanceJson $path @{records=@($safe)}
    $validator=[IO.Path]::Combine($RepositoryRoot,'tools','codex-v1.0.3','check_native_store_report.py')
    $code='import importlib.util,json,sys; s=importlib.util.spec_from_file_location("gate",sys.argv[1]);m=importlib.util.module_from_spec(s);s.loader.exec_module(m);r=json.load(open(sys.argv[2],encoding="utf-8"))["records"];m.validate_receipt_records(r,"discovery");r[0]["process"]["probe"]["redacted_stream_bytes"]["stdout"]+=1;rejected=False' + "`n"+'try: m.validate_receipt_records(r,"discovery")' + "`n"+'except m.GateError: rejected=True' + "`n"+'assert rejected;print("ACTUAL_REDACTION_ACCEPTED_MUTATED_LENGTH_REJECTED")'
    $checked=Invoke-NativeAcceptanceProcessCase $TestPython @('-I','-B','-c',$code,$validator,$path) $WorkRoot 10
    Assert-NativeAcceptanceProcess $checked.probe
    Require ($checked.probe.ExitCode-eq 0 -and $checked.probe.Stdout.Contains('ACTUAL_REDACTION_ACCEPTED_MUTATED_LENGTH_REJECTED')) 'Redacted bytes did not roundtrip through rejecting Store validator.'
}
Case 'policy-action-refuses-runtime-expression-and-extra-command' {
    $inbox=[IO.Path]::Combine([Environment]::SystemDirectory,'WindowsPowerShell\v1.0\powershell.exe')
    $action=Get-WinBookSplitSetupCommand -ApplicationRoot $WorkRoot
    Assert-NativeStoreExistingSetupAction $action $WorkRoot $inbox
    foreach($bad in @($action+" -PythonPath 'AUTHORED_RUNTIME'",$action+'; Write-Host AUTHORED_EXTRA',
        $action.Replace("'"+$inbox+"'",'$(Write-Host AUTHORED_EXPRESSION)'))){
        $failed=$false;try{Assert-NativeStoreExistingSetupAction $bad $WorkRoot $inbox}catch{$failed=$true}
        Require $failed 'Automatic policy Action admitted supplied runtime, expression or extra command.'
    }
}
Case 'policy-actual-authored-motw-positive-and-controller-domain' {
    $app=New-NativeAcceptanceDirectory ([IO.Path]::Combine($WorkRoot,'authored-policy-app'))
    $work=New-NativeAcceptanceDirectory ([IO.Path]::Combine($WorkRoot,'authored-policy-work'))
    $venv=New-NativeAcceptanceDirectory ([IO.Path]::Combine($app,'.venv'))
    [IO.File]::WriteAllText([IO.Path]::Combine($venv,'AUTHORED_SENTINEL'),'No runtime is present; this is only a policy transport control.')
    $main=[IO.Path]::Combine($app,'WinBookSplit.ps1')
    [IO.File]::WriteAllText($main,'param([switch]$Version);Write-Host "WinBookSplit AUTHORED_POLICY_CONTROL"',[Text.UTF8Encoding]::new($true))
    [IO.File]::Copy([IO.Path]::Combine($RepositoryRoot,'Setup-WinBookSplit.bat'),[IO.Path]::Combine($app,'Setup-WinBookSplit.bat'),$false)
    $parseTokens=$parseErrors=$null
    $ast=[Management.Automation.Language.Parser]::ParseFile([IO.Path]::Combine($RepositoryRoot,'Setup-WinBookSplit.ps1'),[ref]$parseTokens,[ref]$parseErrors)
    $hostFunction=$ast.Find({param($node)$node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name-ceq 'Get-WinBookSplitSetupHostObservation'},$true).Extent.Text
    $source='param([switch]$Check,[switch]$NonInteractive);'+"`r`n"+$hostFunction+"`r`n"+@'
$record=@{status='success';exit_code=0;code='setup_ready';changed=$false;operations=@();host=(Get-WinBookSplitSetupHostObservation)}
Write-Host ('[SETUP] '+($record|ConvertTo-Json -Depth 10 -Compress));exit 0
'@
    [IO.File]::WriteAllText([IO.Path]::Combine($app,'Setup-WinBookSplit.ps1'),$source,[Text.UTF8Encoding]::new($true))
    $inbox=[IO.Path]::Combine([Environment]::SystemDirectory,'WindowsPowerShell\v1.0\powershell.exe')
    if($PSVersionTable.PSEdition-ceq 'Desktop'){
        $policyLeases=New-Object 'Collections.Generic.List[IDisposable]'
        try{
            foreach($path in @($main,[IO.Path]::Combine($app,'Setup-WinBookSplit.ps1'))){$policyLeases.Add((New-Object WinBookSplitLogging.PathLease($path,$false,$false,$true)))}
            $result=Invoke-NativeStoreExistingPolicyComponent $app $work $inbox
            foreach($lease in $policyLeases){$lease.AssertUnchanged()}
        }finally{foreach($lease in $policyLeases){$lease.Dispose()}}
        Require ($result.component.status-ceq 'PASS' -and $result.component.direct_remotesigned_denied -and $result.component.marked_setup_bat_noop -and
            $result.component.exact_formatter_action_noop -and $result.component.venv_unchanged -and $result.component.source_bytes_and_marks_preserved -and
            $result.component.organization_policy_denial.status-ceq 'NOT_RUN') 'Actual authored MOTW/BAT/Action control failed or invented organization denial.'
        Require (@($result.records|Where-Object{$_.stage-ceq 'direct_file_remotesigned'}).Count-eq 1 -and @($result.records|Where-Object{$_.role-ceq 'application'}).Count-eq 2) 'Policy controls did not retain native positive/negative process evidence.'
        $tokens=[ordered]@{};$tokens[$WorkRoot]='<WORK>'
        $safe=ConvertTo-NativeStoreSafeValue $result $tokens
        $path=[IO.Path]::Combine($work,'policy-returned-records.json');Write-NativeAcceptanceJson $path $safe
        $validator=[IO.Path]::Combine($RepositoryRoot,'tools','codex-v1.0.3','check_native_store_report.py')
        $code='import importlib.util,json,sys;s=importlib.util.spec_from_file_location("gate",sys.argv[1]);m=importlib.util.module_from_spec(s);s.loader.exec_module(m);r=json.load(open(sys.argv[2],encoding="utf-8"))["records"];m.validate_receipt_records(r,"preservation");m.validate_policy_records(r);print("ACTUAL_POLICY_RECEIPTS_ACCEPTED")'
        $checked=Invoke-NativeAcceptanceProcessCase $TestPython @('-I','-B','-c',$code,$validator,$path) $work 10
        Assert-NativeAcceptanceProcess $checked.probe
        Require ($checked.probe.ExitCode-eq 0 -and $checked.probe.Stdout.Contains('ACTUAL_POLICY_RECEIPTS_ACCEPTED')) 'Actual redacted policy operations were not accepted by the separate checker.'
    }else{
        $failed=$false;try{$null=Invoke-NativeStoreExistingPolicyComponent $app $work $inbox}catch{
            $failed=$_.Exception.Message.Contains('actual inbox controller policy domain') -and $_.Exception.Data.Contains('StorePolicyReceipt')
        }
        Require $failed 'Policy component compared the Core and inbox policy domains.'
        Require ($null-eq (Get-Item -LiteralPath $main -Stream Zone.Identifier -ErrorAction SilentlyContinue)) 'Rejected policy controller wrote marks.'
    }
}
Case 'no-store-registration-before-any-bootstrap-or-fallback' {
    $script:queryCalls=0;$script:processCalls=0;$script:downloadCalls=0;$script:expandCalls=0
    function Get-WinBookSplitRegisteredManagerPaths {$script:queryCalls++;return @()}
    function Invoke-NativeAcceptanceProcessCase {$script:processCalls++;throw 'UNEXPECTED_BOOTSTRAP'}
    function Get-NativeAcceptanceDownload {$script:downloadCalls++;throw 'UNEXPECTED_DOWNLOAD'}
    function Expand-NativeAcceptancePython {$script:expandCalls++;throw 'UNEXPECTED_CLONE'}
    $work=New-NativeAcceptanceDirectory ([IO.Path]::Combine($WorkRoot,'absent'))
    $bindings=@{production=@{version='1.0.3';source_commit=('a'*40);assets=@{'WinBookSplit-v1.0.3.zip'=@{sha256=('b'*64)}}}}
    $result=Invoke-NativeStoreExistingAcceptance -KitRoot $RepositoryRoot -BindingsSha256 ('c'*64) -Bindings $bindings -WorkRoot $work -Inbox ([IO.Path]::Combine($PSHOME,'powershell.exe')) -Os @{platform_scope='NATIVE_WINDOWS11_DRY_RUN';BuildNumber=22631} -Execution AGENT_LAUNCHED
    Require ($script:queryCalls-eq 1 -and $script:processCalls-eq 0 -and $script:downloadCalls-eq 0 -and $script:expandCalls-eq 0) 'Missing Store invoked bootstrap/acquisition or did not use controlled query.'
    Require (-not $result.success -and $result.report.status-ceq 'blocked' -and -not $result.report.fallback_acquisition.attempted -and -not $result.report.fallback_acquisition.used) 'Missing Store produced untruthful no-fallback result.'
    Require (-not [IO.Directory]::Exists([IO.Path]::Combine($work,'downloads')) -and -not [IO.Directory]::Exists([IO.Path]::Combine($work,'application'))) 'Missing Store created fallback/extraction work.'
    Require ([IO.File]::Exists($result.return_zip)) 'Missing Store did not emit the bounded redacted return ZIP.'
}
${function:Get-WinBookSplitRegisteredManagerPaths}=$nativeQuery;${function:Invoke-NativeAcceptanceProcessCase}=$nativeProcess;${function:Get-NativeAcceptanceDownload}=$nativeDownload;${function:Expand-NativeAcceptancePython}=$nativeExpand
Case 'inventory-positive-control-real-child-is-actually-used' {
    $script:listingCalls=0;$script:queryCalls=0
    function Get-WinBookSplitRegisteredManagerPaths {$script:queryCalls++;return @(@{Path='SYNTHETIC_MANAGER_ENTRY';Package=@{PackageFamilyName='PythonSoftwareFoundation.PythonManager_qbz5n2kfra8p0'}})}
    function Invoke-WinBookSplitManagerListing {
        $script:listingCalls++;$probe=Invoke-WinBookSplitProcess -Path $TestPython -Arguments @('-I','-B','-c','import sys;print(sys.executable)') -WorkingDirectory $WorkRoot -TimeoutSeconds 10
        $probe.PackageIdentityVerified=$true;return $probe
    }
    $inventory=Get-NativeStoreExistingInventory $RepositoryRoot
    Require ($script:queryCalls-eq 1 -and $script:listingCalls-eq 1 -and $inventory.runtimes.Count-eq 1 -and $inventory.runtimes[0].candidate.Details.isolated -and $inventory.runtimes[0].sha256-cmatch '^[a-f0-9]{64}$') 'Actual-child synthetic manager positive did not reach independent verified runtime.'
}
Case 'inventory-nonzero-listing-refuses-runtime-probe' {
    $script:listingCalls=0;$script:candidateCalls=0
    function Get-WinBookSplitRegisteredManagerPaths {return @(@{Path='SYNTHETIC_MANAGER_ENTRY';Package=@{PackageFamilyName='PythonSoftwareFoundation.PythonManager_qbz5n2kfra8p0'}})}
    function Invoke-WinBookSplitManagerListing {$script:listingCalls++;$probe=Invoke-WinBookSplitProcess -Path $TestPython -Arguments @('-I','-B','-c','import sys;print("controlled listing failure",file=sys.stderr);sys.exit(17)') -WorkingDirectory $WorkRoot -TimeoutSeconds 10;$probe.PackageIdentityVerified=$true;return $probe}
    function Test-WinBookSplitPythonCandidate {$script:candidateCalls++;throw 'UNEXPECTED_CANDIDATE'}
    $failed=$false;try{$null=Get-NativeStoreExistingInventory $RepositoryRoot}catch{$failed=$true}
    Require ($failed -and $script:listingCalls-eq 1 -and $script:candidateCalls-eq 0) 'Nonzero installed-only listing was accepted/ignored.'
}
Case 'rejected-runtime-probe-survives-returned-failure' {
    $script:listingCalls=0;$script:candidateCalls=0
    function Get-WinBookSplitRegisteredManagerPaths {return @(@{Path='SYNTHETIC_MANAGER_ENTRY';Package=@{PackageFamilyName='PythonSoftwareFoundation.PythonManager_qbz5n2kfra8p0'}})}
    function Invoke-WinBookSplitManagerListing {$script:listingCalls++;$probe=Invoke-WinBookSplitProcess -Path $TestPython -Arguments @('-I','-B','-c','import sys;print(sys.executable)') -WorkingDirectory $WorkRoot -TimeoutSeconds 10;$probe.PackageIdentityVerified=$true;return $probe}
    function Test-WinBookSplitPythonCandidate {param($Path,$Source,$ApplicationRoot,$Attempts,$TimeoutSeconds,[switch]$InterpreterOnly);$script:candidateCalls++;$Attempts.Add(@{Path=$Path;Accepted=$false;Reason='python_security_floor';ObservedVersion='3.14.7';IdentityVerified=$true;scope='SYNTHETIC_REJECTION_REASON_ACTUAL_LISTING_CHILD'});return $null}
    $work=New-NativeAcceptanceDirectory ([IO.Path]::Combine($WorkRoot,'rejected-runtime'))
    $bindings=@{production=@{version='1.0.3';source_commit=('a'*40);assets=@{'WinBookSplit-v1.0.3.zip'=@{sha256=('b'*64)}}}}
    $result=Invoke-NativeStoreExistingAcceptance -KitRoot $RepositoryRoot -BindingsSha256 ('c'*64) -Bindings $bindings -WorkRoot $work -Inbox ([IO.Path]::Combine($PSHOME,'powershell.exe')) -Os @{platform_scope='NATIVE_WINDOWS11_DRY_RUN';BuildNumber=22631} -Execution AGENT_LAUNCHED
    $receipt=(Read-NativeAcceptanceJson ([IO.Path]::Combine($result.report_directory,'registration.json'))).value
    $rejection=$receipt.records[0].StoreInventoryReceipt.runtime_probes[0]
    Require ($script:listingCalls-eq 1 -and $script:candidateCalls-eq 1 -and -not $result.success -and -not $rejection.accepted -and $rejection.attempts[0].Reason-ceq 'python_security_floor' -and $rejection.attempts[0].ObservedVersion-ceq '3.14.7') 'Rejected interpreter reason was lost from returned failed inventory.'
}
${function:Get-WinBookSplitRegisteredManagerPaths}=$nativeQuery;${function:Invoke-WinBookSplitManagerListing}=$nativeListing;${function:Test-WinBookSplitPythonCandidate}=$nativeCandidate
Case 'finite-return-zip-positive-and-foreign-member-refusal' {
    $directory=New-NativeAcceptanceDirectory ([IO.Path]::Combine($WorkRoot,'finite-return'))
    $names=@('native-store-report.json','registration.json','discovery.json','setup.json','bat.json','pdf_validation.json','preservation.json')
    foreach($name in $names){Write-NativeAcceptanceJson ([IO.Path]::Combine($directory,$name)) @{scope='AUTHORED_FINITE_PACK_CONTROL';name=$name}}
    $zip=[IO.Path]::Combine($WorkRoot,'finite-positive.zip');$packed=Write-NativeStoreExistingReturnZip $directory $zip
    Require ($packed.members.Count-eq 7 -and $packed.archive.sha256-cmatch '^[a-f0-9]{64}$') 'Finite seven-file positive archive was not created.'
    $stream=[IO.File]::OpenRead($zip);$archive=New-Object IO.Compression.ZipArchive($stream,[IO.Compression.ZipArchiveMode]::Read,$false)
    try{Require ($archive.Entries.Count-eq 7 -and @($archive.Entries|Where-Object{$_.FullName-cnotin $names}).Count-eq 0) 'Foreign member entered positive return ZIP.'}finally{$archive.Dispose();$stream.Dispose()}
    Write-NativeAcceptanceJson ([IO.Path]::Combine($directory,'foreign.json')) @{scope='AUTHORED_FOREIGN_FILE_MUST_NOT_EXPORT'}
    $blocked=[IO.Path]::Combine($WorkRoot,'finite-foreign-rejected.zip');$failed=$false
    try{$null=Write-NativeStoreExistingReturnZip $directory $blocked}catch{$failed=$true}
    Require ($failed -and -not [IO.File]::Exists($blocked) -and [IO.File]::Exists([IO.Path]::Combine($directory,'foreign.json'))) 'Foreign return member was included, removed, or reached archive creation.'
}
Case 'existing-inventory-positive-reaches-bootstrap-no-acquisition' {
    $script:inventoryCalls=0;$script:bootstrapCalls=0;$script:downloadCalls=0;$script:expandCalls=0
    function Get-NativeStoreExistingInventory {$script:inventoryCalls++;return @{registration=@{Path='SYNTHETIC_MANAGER_ENTRY';Package=@{PackageFamilyName='PythonSoftwareFoundation.PythonManager_qbz5n2kfra8p0';Publisher='SYNTHETIC';SignatureKind='Store';Status='Ok';Architecture='X64';Version='1.2.3.4'}};runtimes=@(@{path=$TestPython;version='3.14.8';sha256=('d'*64)})}}
    function Get-NativeStoreExistingManagerState {return @{scope='SYNTHETIC_MANAGER_STATE'}}
    function Get-NativeStoreExistingPackages {return @{scope='SYNTHETIC_PACKAGE_STATE'}}
    function Invoke-NativeAcceptanceProcessCase {$script:bootstrapCalls++;throw 'CONTROLLED_BOOTSTRAP_STOP'}
    function Get-NativeAcceptanceDownload {$script:downloadCalls++;throw 'UNEXPECTED_DOWNLOAD'}
    function Expand-NativeAcceptancePython {$script:expandCalls++;throw 'UNEXPECTED_CLONE'}
    $work=New-NativeAcceptanceDirectory ([IO.Path]::Combine($WorkRoot,'bootstrap-positive'))
    $bindings=@{production=@{version='1.0.3';source_commit=('a'*40);assets=@{'WinBookSplit-v1.0.3.zip'=@{sha256=('b'*64)}}}}
    $result=Invoke-NativeStoreExistingAcceptance -KitRoot $RepositoryRoot -BindingsSha256 ('c'*64) -Bindings $bindings -WorkRoot $work -Inbox ([IO.Path]::Combine($PSHOME,'powershell.exe')) -Os @{platform_scope='NATIVE_WINDOWS11_DRY_RUN';BuildNumber=22631} -Execution AGENT_LAUNCHED
    Require ($script:inventoryCalls-eq 1 -and $script:bootstrapCalls-eq 1 -and $script:downloadCalls-eq 0 -and $script:expandCalls-eq 0 -and -not $result.success) 'Positive existing-inventory fixture did not reach bounded bootstrap without fallback.'
}
${function:Get-NativeStoreExistingInventory}=$nativeInventory;${function:Get-NativeStoreExistingManagerState}=$nativeManagerState;${function:Get-NativeStoreExistingPackages}=$nativePackages;${function:Invoke-NativeAcceptanceProcessCase}=$nativeProcess;${function:Get-NativeAcceptanceDownload}=$nativeDownload;${function:Expand-NativeAcceptancePython}=$nativeExpand
$report=@{schema_version=1;kind='T05_STORE_SCENARIO_CONTROLLED_NATIVE_TOOLING';host=@{shell=$PSVersionTable.PSVersion.ToString();edition=$PSVersionTable.PSEdition};success=(@($cases|Where-Object{$_.status-cne 'PASS'}).Count-eq 0);cases=$cases.ToArray();limits=@('Synthetic manager seams and authored inputs only; no native Store registration/owner Windows10/exact-ZIP acceptance is claimed.','No Python install/update/download/copy/clone or PATH/alias change was performed.')}
Write-NativeAcceptanceJson $ReportPath $report
Write-Host ('STORE_PROBE_REPORT '+$ReportPath)
exit $(if($report.success){0}else{1})
