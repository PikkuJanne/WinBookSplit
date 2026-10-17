# T05's separate existing-installation route. No function in this file acquires,
# copies or clones a Python runtime. Legacy explicit-runtime recipes stay separate.
function Assert-NativeStoreExistingInputs {
    param([string]$BarePythonPath,[string]$BaselinePythonPath,[switch]$ProvisionPrerequisites)
    if(-not [string]::IsNullOrWhiteSpace($BarePythonPath) -or -not [string]::IsNullOrWhiteSpace($BaselinePythonPath) -or $ProvisionPrerequisites){
        throw 'Store-existing forbids supplied runtime paths and prerequisite provisioning before discovery.'
    }
    foreach($key in [Environment]::GetEnvironmentVariables('Process').Keys){
        if(([string]$key -cmatch '(?i)^(VIRTUAL_ENV|PYTHONHOME|PYTHONPATH|PYMANAGER_.*|PYTHON_MANAGER_.*|WBS_.*PYTHON.*|WINBOOKSPLIT_.*PYTHON.*)$') -and
            -not [string]::IsNullOrWhiteSpace([Environment]::GetEnvironmentVariable([string]$key,'Process'))){
            throw 'Store-existing refuses an activated or runtime-injected controller environment.'
        }
    }
}

function Get-NativeStoreExistingInventory {
    param([string]$ApplicationRoot)
    $attempts=New-Object 'Collections.Generic.List[object]'
    $runtimeProbes=New-Object 'Collections.Generic.List[object]'
    $registered=$listing=$null
    try{
    $registered=@(Get-WinBookSplitRegisteredManagerPaths -ApplicationRoot $ApplicationRoot -Attempts $attempts -TimeoutSeconds 5 -WithRegistration)
    $store=@($registered|Where-Object{$_.Package.PackageFamilyName -ceq 'PythonSoftwareFoundation.PythonManager_qbz5n2kfra8p0'})
    if($store.Count-ne 1){throw 'Existing Store registration is absent or ambiguous; Python fallback acquisition is disabled.'}
    $listing=Invoke-WinBookSplitManagerListing -Registration $store[0] -ApplicationRoot $ApplicationRoot -TimeoutSeconds 5
    Assert-NativeAcceptanceProcess $listing
    if($listing.ExitCode-ne 0 -or -not $listing.PackageIdentityVerified){throw 'Existing Store installed-only listing failed; Python fallback acquisition is disabled.'}
    $paths=@(ConvertFrom-WinBookSplitPythonListing -Probe $listing -Kind Manager)
    if($paths.Count-lt 1 -or $paths.Count-gt 32){throw 'Existing Store runtime inventory is empty or overbound; Python fallback acquisition is disabled.'}
    $runtimes=New-Object 'Collections.Generic.List[object]'
    foreach($path in $paths){
        $probes=New-Object 'Collections.Generic.List[object]'
        $candidate=$null
        try{$candidate=Test-WinBookSplitPythonCandidate $path 'manager' $ApplicationRoot $probes 5 -InterpreterOnly}
        finally{$runtimeProbes.Add(@{path=$path;accepted=($null-ne $candidate);attempts=$probes.ToArray()})}
        if($candidate){$runtimes.Add([ordered]@{path=$candidate.Path;version=$candidate.Version;sha256=(Get-NativeAcceptanceFileHash $candidate.Path 8388608).sha256;candidate=$candidate})}
    }
    if($runtimes.Count-eq 0){throw 'No compatible pre-existing Store interpreter was verified; Python fallback acquisition is disabled.'}
    return [pscustomobject]@{registration=$store[0];query_attempts=$attempts.ToArray();listing=$listing;runtime_probes=$runtimeProbes.ToArray();runtimes=$runtimes.ToArray()}
    }catch{
        $_.Exception.Data['StoreInventoryReceipt']=@{query_attempts=$attempts.ToArray();registrations=$registered;listing=$listing;runtime_probes=$runtimeProbes.ToArray()}
        throw
    }
}

function Add-NativeStoreRedactedProbeIdentity {
    param([Collections.IDictionary]$Safe,$Original)
    if('Stdout'-cin @($Safe.Keys) -and 'Stderr'-cin @($Safe.Keys) -and 'StdoutTotalBytes'-cin @($Safe.Keys)){
        $Safe.redacted_stream_bytes=@{stdout=[Text.Encoding]::UTF8.GetByteCount([string]$Safe.Stdout);stderr=[Text.Encoding]::UTF8.GetByteCount([string]$Safe.Stderr)}
        $hashes=@{};$algorithm=[Security.Cryptography.SHA256]::Create()
        try{foreach($stream in @('Stdout','Stderr')){$hashes[$stream.ToLowerInvariant()]=([BitConverter]::ToString($algorithm.ComputeHash([Text.Encoding]::UTF8.GetBytes([string]$Original.$stream)))).Replace('-','').ToLowerInvariant()}}finally{$algorithm.Dispose()}
        $Safe.raw_stream_sha256=$hashes
    }
}

function ConvertTo-NativeStoreSafeValue {
    param($Value,[Collections.IDictionary]$Tokens)
    if($null-eq $Value){return $null}
    if($Value -is [string]){
        if($Value.TrimStart().StartsWith('{') -or $Value.TrimStart().StartsWith('[')){
            try{$parsed=[WinBookSplitSupport.StrictJson]::Parse($Value);$clean=ConvertTo-NativeStoreSafeValue $parsed $Tokens;return (ConvertTo-Json -InputObject $clean -Depth 100 -Compress)}catch{$Value=[string]$Value} # Non-JSON text retains literal/escaped path replacement below.
        }
        $safe=$Value
        foreach($key in @($Tokens.Keys|Sort-Object Length -Descending)){
            $literal=[string]$key;$replacement=[string]$Tokens[$key]
            for($level=0;$level-lt 4;$level++){
                $safe=[regex]::Replace($safe,[regex]::Escape($literal),[Text.RegularExpressions.MatchEvaluator]{param($match)$replacement},[Text.RegularExpressions.RegexOptions]::IgnoreCase)
                $literal=$literal.Replace('\','\\')
            }
        }
        return $safe
    }
    if($Value -is [Collections.IDictionary]){
        $safe=[ordered]@{};foreach($key in $Value.Keys){$safe[(ConvertTo-NativeStoreSafeValue ([string]$key) $Tokens)]=ConvertTo-NativeStoreSafeValue $Value[$key] $Tokens};Add-NativeStoreRedactedProbeIdentity $safe $Value;return $safe
    }
    if($Value -is [Collections.IEnumerable]){return ,@($Value|ForEach-Object{ConvertTo-NativeStoreSafeValue $_ $Tokens})}
    if($Value -is [bool] -or $Value -is [ValueType]){return $Value}
    $safe=[ordered]@{};foreach($property in $Value.PSObject.Properties){$safe[$property.Name]=ConvertTo-NativeStoreSafeValue $property.Value $Tokens};Add-NativeStoreRedactedProbeIdentity $safe $Value;return $safe
}

function Get-NativeStoreExistingManagerState {
    param($Registration)
    if(-not ('WinBookSplitNative.StoreAliasSnapshot' -as [type])){
        Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
namespace WinBookSplitNative {
 public static class StoreAliasSnapshot {
  [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)] static extern IntPtr CreateFileW(string name,uint access,uint share,IntPtr security,uint creation,uint flags,IntPtr template);
  [DllImport("kernel32.dll", SetLastError=true)] static extern bool DeviceIoControl(IntPtr handle,uint code,IntPtr input,uint inputLength,byte[] output,uint outputLength,out uint returned,IntPtr overlapped);
  [DllImport("kernel32.dll", SetLastError=true)] static extern bool CloseHandle(IntPtr handle);
  public static string ReparseSha256(string path) {
   IntPtr handle=CreateFileW(path,0,7,IntPtr.Zero,3,0x02200000,IntPtr.Zero);
   if(handle==new IntPtr(-1))throw new Win32Exception(Marshal.GetLastWin32Error());
   try {byte[] data=new byte[16384];uint count;if(!DeviceIoControl(handle,0x000900A8,IntPtr.Zero,0,data,(uint)data.Length,out count,IntPtr.Zero))throw new Win32Exception(Marshal.GetLastWin32Error());
    if(count<8 || count>data.Length)throw new InvalidOperationException("Invalid alias reparse bound.");
    using(SHA256 algorithm=SHA256.Create())return BitConverter.ToString(algorithm.ComputeHash(data,0,(int)count)).Replace("-","").ToLowerInvariant();
   } finally {CloseHandle(handle);}
  }
 }
}
'@ -ErrorAction Stop
    }
    $aliases=[ordered]@{};$directory=[IO.Path]::GetDirectoryName([IO.Path]::GetDirectoryName($Registration.Path))
    foreach($path in @($Registration.Path)+@('python.exe','python3.exe','py.exe','pymanager.exe'|ForEach-Object{[IO.Path]::Combine($directory,$_)})){
        try{$attributes=[IO.File]::GetAttributes($path)}catch [IO.FileNotFoundException]{$aliases[$path]=@{exists=$false};continue}catch [IO.DirectoryNotFoundException]{$aliases[$path]=@{exists=$false};continue}
        $hash=if($attributes-band [IO.FileAttributes]::ReparsePoint){[WinBookSplitNative.StoreAliasSnapshot]::ReparseSha256($path)}else{(Get-NativeAcceptanceFileHash $path 8388608 -AllowEmpty).sha256}
        $aliases[$path]=@{exists=$true;attributes=[int]$attributes;last_write_ticks=[IO.File]::GetLastWriteTimeUtc($path).Ticks;sha256=$hash}
    }
    $algorithm=[Security.Cryptography.SHA256]::Create()
    try{$pathHash=([BitConverter]::ToString($algorithm.ComputeHash([Text.Encoding]::UTF8.GetBytes([string][Environment]::GetEnvironmentVariable('PATH','Process'))))).Replace('-','').ToLowerInvariant()}finally{$algorithm.Dispose()}
    return @{registration=$Registration;aliases=$aliases;parent_path_sha256=$pathHash;manager_image=(Get-NativeAcceptanceFileHash ([IO.Path]::Combine($Registration.Package.InstallLocation,'pymanager.exe')) 8388608)}
}

function Invoke-NativeStoreExistingChild {
    param([string]$Inbox,[string]$WorkingDirectory,[string]$Name,[string]$Body)
    $path=[IO.Path]::Combine($WorkingDirectory,$Name+'.ps1')
    # BOM is deliberate: inbox Windows PowerShell parses authored UTF-8 paths.
    $Body='$ErrorActionPreference="Stop";$ProgressPreference="SilentlyContinue";[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false);'+"`r`n"+$Body
    $bytes=[Text.UTF8Encoding]::new($true).GetPreamble()+[Text.Encoding]::UTF8.GetBytes($Body)
    $stream=[IO.File]::Open($path,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::Read)
    try{$stream.Write($bytes,0,$bytes.Length)}finally{$stream.Dispose()}
    $before=[Environment]::GetEnvironmentVariables('Process')
    $run=Invoke-NativeDiscoveryCollector -Path $Inbox -Arguments @('-NoProfile','-NonInteractive','-InputFormat','Text','-OutputFormat','Text','-ExecutionPolicy','Bypass','-File',$path) -WorkingDirectory $WorkingDirectory -TimeoutSeconds 120 -OutputPath ([IO.Path]::Combine($WorkingDirectory,$Name+'-collector.json'))
    $after=[Environment]::GetEnvironmentVariables('Process');$restored=$before.Count-eq $after.Count
    foreach($key in $before.Keys){if($before[$key]-cne $after[$key]){$restored=$false}}
    $run|Add-Member NoteProperty environment_before_after_restored $restored
    $run|Add-Member NoteProperty failure $run.collector_error
    if($run.status-cne 'PASS'){$error=New-Object InvalidOperationException('Store-existing child failed: '+$Name+' ('+$run.reason+').');$error.Data['StoreChildReceipt']=$run;throw $error}
    if(-not $restored){throw 'Store-existing child changed controller environment.'}
    return $run
}

function Get-NativeStoreExistingPackages {
    param([string]$Python,[string]$WorkingDirectory)
    $code='import importlib.metadata,json; print(json.dumps(sorted((d.metadata.get("Name",""),d.version,str(d.locate_file(""))) for d in importlib.metadata.distributions()),ensure_ascii=True))'
    $run=Invoke-NativeAcceptanceProcessCase $Python @('-I','-B','-c',$code) $WorkingDirectory 10
    Assert-NativeAcceptanceProcess $run.probe
    if($run.failure -or $run.probe.ExitCode-ne 0){throw 'Read-only existing-runtime package inventory failed.'}
    $null=[WinBookSplitSupport.StrictJson]::Parse($run.probe.Stdout.Trim())
    return $run
}

function Write-NativeStoreExistingReturnZip {
    param([Parameter(Mandatory=$true)][string]$ReportRoot,[Parameter(Mandatory=$true)][string]$Destination)
    $root=Get-NativeAcceptanceOrdinaryPath $ReportRoot -Directory
    $names=@('native-store-report.json','registration.json','discovery.json','setup.json','bat.json','pdf_validation.json','preservation.json')
    $leases=New-Object 'Collections.Generic.List[IDisposable]';$stream=$archive=$null;$identities=[ordered]@{}
    try{
        $leases.Add((New-Object WinBookSplitLogging.PathLease($root,$true)))
        $members=@([IO.Directory]::GetFileSystemEntries($root))
        if($members.Count-ne $names.Count){throw 'Store return directory differs from its finite seven-file set.'}
        foreach($path in $members){if([IO.Path]::GetFileName($path)-cnotin $names){throw 'Unexpected member in Store return directory.'}}
        foreach($name in $names){
            $path=[IO.Path]::Combine($root,$name);$null=Get-NativeAcceptanceOrdinaryPath $path
            $leases.Add((New-Object WinBookSplitLogging.PathLease($path,$false,$false,$true)))
            $identities[$name]=Get-NativeAcceptanceFileHash $path 4194304
        }
        Add-Type -AssemblyName System.IO.Compression -ErrorAction Stop
        $stream=[IO.File]::Open($Destination,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::Read)
        $archive=New-Object IO.Compression.ZipArchive($stream,[IO.Compression.ZipArchiveMode]::Create,$false)
        foreach($name in $names){
            foreach($lease in $leases){$lease.AssertUnchanged()}
            $entry=$archive.CreateEntry($name,[IO.Compression.CompressionLevel]::NoCompression);$member=$entry.Open()
            try{$bytes=[IO.File]::ReadAllBytes([IO.Path]::Combine($root,$name));$member.Write($bytes,0,$bytes.Length)}finally{$member.Dispose()}
        }
        foreach($lease in $leases){$lease.AssertUnchanged()}
    }finally{if($archive){$archive.Dispose()};if($stream){$stream.Dispose()};for($index=$leases.Count-1;$index-ge 0;$index--){$leases[$index].Dispose()}}
    return @{members=$identities;archive=(Get-NativeAcceptanceFileHash $Destination 33554432)}
}

function Get-NativeStoreExistingPolicyObservation {
    # The central controller is inbox PS5.1; do not compare a PS7 policy store.
    Import-Module -Name ([IO.Path]::Combine($PSHOME,'Modules','Microsoft.PowerShell.Security','Microsoft.PowerShell.Security.psd1')) -Scope Local -ErrorAction Stop
    $policies=@(Get-ExecutionPolicy -List -ErrorAction Stop)
    $scopes=@('MachinePolicy','UserPolicy','Process','CurrentUser','LocalMachine')
    $rows=@(foreach($scope in $scopes){
        $matches=@($policies|Where-Object{$_.Scope.ToString()-ceq $scope})
        if($matches.Count-ne 1){throw 'Incomplete inbox execution-policy observation.'}
        @{scope=$scope;policy=$matches[0].ExecutionPolicy.ToString()}
    })
    return @{effective=(Get-ExecutionPolicy -ErrorAction Stop).ToString();scopes=$rows}
}

function Assert-NativeStoreExistingSetupAction {
    param([string]$Action,[string]$ApplicationRoot,[string]$Inbox)
    # Reuse the emitted command unchanged only after proving its finite literal
    # AST. In particular no runtime, wheelhouse, repair, expression or redirect.
    $tokens=$errors=$null
    $ast=[Management.Automation.Language.Parser]::ParseInput($Action,[ref]$tokens,[ref]$errors)
    if($errors.Count-ne 0 -or $ast.EndBlock.Statements.Count-ne 1){throw 'Invalid automatic setup Action.'}
    $pipeline=$ast.EndBlock.Statements[0]
    if($pipeline -isnot [Management.Automation.Language.PipelineAst] -or $pipeline.PipelineElements.Count-ne 1){throw 'Automatic setup Action is not a single command.'}
    $command=$pipeline.PipelineElements[0]
    if($command -isnot [Management.Automation.Language.CommandAst] -or $command.InvocationOperator.ToString()-cne 'Ampersand' -or $command.Redirections.Count-ne 0){throw 'Automatic setup Action is not a literal invocation.'}
    $values=@(foreach($element in $command.CommandElements){
        if($element -is [Management.Automation.Language.StringConstantExpressionAst]){[string]$element.Value}
        elseif($element -is [Management.Automation.Language.CommandParameterAst] -and $null-eq $element.Argument){'-'+$element.ParameterName}
        else{throw 'Automatic setup Action contains a nonliteral element.'}
    })
    $expected=@($Inbox,'-NoProfile','-ExecutionPolicy','Bypass','-File',[IO.Path]::Combine($ApplicationRoot,'Setup-WinBookSplit.ps1'),'-NonInteractive')
    if(($values|ConvertTo-Json -Compress)-cne ($expected|ConvertTo-Json -Compress)){throw 'Automatic setup Action differs from the no-runtime command.'}
}

function Assert-NativeStoreExistingSetupPolicy {
    param($Frame,$Baseline,[string]$Inbox)
    if($Frame.status-cne 'success' -or $Frame.exit_code-ne 0 -or $Frame.code-cne 'setup_ready' -or $Frame.changed-isnot [bool] -or $Frame.changed -or @($Frame.operations).Count-ne 0 -or
        $Frame.host.edition-cne 'Desktop' -or $Frame.host.version-cnotmatch '^5\.1\.' -or
        -not [StringComparer]::OrdinalIgnoreCase.Equals($Frame.host.executable,$Inbox) -or
        $Frame.host.execution_policy.status-cne 'observed' -or $Frame.host.execution_policy.effective-cne 'Bypass'){
        throw 'Marked automatic setup was not an observed healthy inbox no-op.'
    }
    foreach($scope in @('MachinePolicy','UserPolicy','Process','CurrentUser','LocalMachine')){
        $actual=@($Frame.host.execution_policy.scopes|Where-Object{$_.scope-ceq $scope})
        $before=@($Baseline.scopes|Where-Object{$_.scope-ceq $scope})
        $expected=if($scope-ceq 'Process'){'Bypass'}else{$before[0].policy}
        if($actual.Count-ne 1 -or $before.Count-ne 1 -or $actual[0].policy-cne $expected){throw 'Marked setup changed or misreported persistent/process policy.'}
    }
}

function Invoke-NativeStoreExistingPolicyComponent {
    param([string]$ApplicationRoot,[string]$WorkRoot,[string]$Inbox)
    $component=[ordered]@{status='NOT_RUN';scope='OWNED_EXACT_EXTRACTION_HEALTHY_VENV';
        direct_remotesigned_denied=$false;marked_setup_bat_noop=$false;exact_formatter_action_noop=$false;
        source_bytes_and_marks_preserved=$false;venv_unchanged=$false;inbox_policy_unchanged=$false;
        organization_policy_denial=@{status='NOT_RUN';reason='No real enforced organization-policy denial has been demonstrated.'}}
    $records=New-Object 'Collections.Generic.List[object]'
    try{
        if(-not [StringComparer]::OrdinalIgnoreCase.Equals([IO.Path]::GetFullPath($PSHOME),[IO.Path]::GetDirectoryName($Inbox))){throw 'Policy component requires the actual inbox controller policy domain.'}
        $null=Get-NativeAcceptanceOrdinaryPath $ApplicationRoot -Directory
        if((New-Object IO.DriveInfo([IO.Path]::GetPathRoot($ApplicationRoot))).DriveFormat-cne 'NTFS'){throw 'Policy component requires owned NTFS alternate streams.'}
        $baseline=Get-NativeStoreExistingPolicyObservation
        $records.Add(@{role='policy_observation';stage='before';inbox_policy=$baseline})
        $venv=[IO.Path]::Combine($ApplicationRoot,'.venv');$venvBefore=Get-NativeAcceptanceEnvironmentSnapshot $venv
        $marks=New-Object 'Collections.Generic.List[object]'
        foreach($name in @('WinBookSplit.ps1','Setup-WinBookSplit.ps1')){
            $path=[IO.Path]::Combine($ApplicationRoot,$name);$hash=Get-NativeAcceptanceFileHash $path 8388608
            if($null-ne (Get-Item -LiteralPath $path -Stream Zone.Identifier -ErrorAction SilentlyContinue)){throw 'Fresh extraction unexpectedly already has Internet marks.'}
            $ads="[ZoneTransfer]`r`nZoneId=3`r`n"
            Set-Content -LiteralPath $path -Stream Zone.Identifier -Value $ads -Encoding ASCII -NoNewline -ErrorAction Stop
            $marks.Add(@{path=$path;sha256=$hash.sha256;ads=$ads})
        }
        $records.Add(@{role='policy_observation';stage='authored_motw';marks=$marks.ToArray();note='ZoneId=3 is authored only on this owned extraction; package data bytes are unchanged.'})
        $main=[IO.Path]::Combine($ApplicationRoot,'WinBookSplit.ps1')
        $run=Invoke-NativeAcceptanceProcessCase $Inbox @('-NoProfile','-NonInteractive','-InputFormat','Text','-OutputFormat','Text','-ExecutionPolicy','RemoteSigned','-File',$main,'-Version') $WorkRoot 30
        $records.Add(@{role='expected_policy_denial';environment=@{};stage='direct_file_remotesigned';process=$run})
        Assert-NativeAcceptanceProcess $run.probe
        if($run.failure -or $run.probe.ExitCode-eq 0 -or -not $run.environment_before_after_restored -or
            $run.probe.Stderr-cnotmatch 'PSSecurityException|UnauthorizedAccess' -or $run.probe.Stdout-cmatch '(?m)^WinBookSplit '){throw 'Marked direct -File execution did not prove policy denial.'}
        $component.direct_remotesigned_denied=$true
        $setupBat=[IO.Path]::Combine($ApplicationRoot,'Setup-WinBookSplit.bat')
        if($setupBat-match '[%"\r\n]'){throw 'Marked setup BAT requires a literal cmd route.'}
        $wrapper=[IO.Path]::Combine($WorkRoot,'policy-setup-check.cmd')
        $stream=[IO.File]::Open($wrapper,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::Read)
        try{$bytes=[Text.Encoding]::UTF8.GetBytes("@echo off`r`nsetlocal DisableDelayedExpansion`r`nchcp 65001 >nul`r`n"+'"'+$setupBat+'" -Check -NonInteractive'+"`r`n");$stream.Write($bytes,0,$bytes.Length)}finally{$stream.Dispose()}
        $cmd=[IO.Path]::Combine([Environment]::SystemDirectory,'cmd.exe')
        $run=Invoke-NativeAcceptanceProcessCase $cmd @('/d','/v:off','/c',$wrapper) $WorkRoot 120
        $records.Add(@{role='application';environment=@{};stage='marked_setup_bat_check';process=$run})
        Assert-NativeAcceptanceProcess $run.probe
        if($run.failure -or $run.probe.ExitCode-ne 0 -or -not $run.environment_before_after_restored){throw 'Marked setup BAT failed; no policy bypass retry is permitted.'}
        Assert-NativeStoreExistingSetupPolicy (Get-NativeAcceptanceFrame $run.probe.Stdout '[SETUP] ') $baseline $Inbox
        $component.marked_setup_bat_noop=$true
        $action=Get-WinBookSplitSetupCommand -ApplicationRoot $ApplicationRoot
        Assert-NativeStoreExistingSetupAction $action $ApplicationRoot $Inbox
        $literal=ConvertTo-WinBookSplitPowerShellLiteral $action
        $body='$action='+$literal+"`r`n"+'Write-Host (''[EXACT_ACTION] ''+$action); & ([scriptblock]::Create($action));exit $LASTEXITCODE'
        $run=Invoke-NativeStoreExistingChild $Inbox $WorkRoot 'policy-exact-action' $body
        $records.Add(@{role='application';environment=@{};stage='exact_automatic_formatter_action';action=$action;process=$run;note='Actual shared formatter output executed unchanged on the healthy venv; no fresh remediation or package installation is claimed.'})
        $text=$run.probe.Stdout
        Assert-NativeStoreExistingSetupPolicy (Get-NativeAcceptanceFrame $text '[SETUP] ') $baseline $Inbox
        if(-not $text.Contains('[EXACT_ACTION] '+$action)){throw 'Executed automatic Action differs from formatter output.'}
        $component.exact_formatter_action_noop=$true
        foreach($mark in $marks){
            if((Get-NativeAcceptanceFileHash $mark.path 8388608).sha256-cne $mark.sha256 -or
                (Get-Content -LiteralPath $mark.path -Stream Zone.Identifier -Raw -Encoding ASCII -ErrorAction Stop)-cne $mark.ads){throw 'Marked package bytes or authored Internet metadata changed.'}
        }
        $component.source_bytes_and_marks_preserved=$true
        if((Get-NativeAcceptanceEnvironmentSnapshot $venv)-cne $venvBefore){throw 'Marked healthy no-op changed the application environment.'}
        $component.venv_unchanged=$true
        $after=Get-NativeStoreExistingPolicyObservation
        $records.Add(@{role='policy_observation';stage='after';inbox_policy=$after})
        if(($baseline|ConvertTo-Json -Depth 8 -Compress)-cne ($after|ConvertTo-Json -Depth 8 -Compress)){throw 'Inbox controller policy changed.'}
        $component.inbox_policy_unchanged=$true;$component.status='PASS'
        return @{component=$component;records=$records.ToArray()}
    }catch{
        $component.status='FAIL'
        $_.Exception.Data['StorePolicyReceipt']=@{component=$component;records=$records.ToArray();failure=$_.Exception.Message}
        throw
    }
}

function Invoke-NativeStoreExistingAcceptance {
    [CmdletBinding()]
    param([string]$KitRoot,[string]$BindingsSha256,$Bindings,[string]$WorkRoot,[string]$Inbox,$Os,[string]$Execution,
        [string]$BarePythonPath,[string]$BaselinePythonPath,[switch]$ProvisionPrerequisites,[string]$Wheelhouse)
    $report=[ordered]@{schema_version=2;scenario='existing_owner_windows10_store_installation';evidence_kind='native_execution';status='blocked';
        application_version=$Bindings.production.version;run_id=('store_'+[Guid]::NewGuid().ToString('N'));source_commit=$Bindings.production.source_commit;
        zip_sha256=$Bindings.production.assets[('WinBookSplit-v'+$Bindings.production.version+'.zip')].sha256;
        host=[ordered]@{execution_kind='native';owner_desktop=($Execution-ceq 'USER_LAUNCHED_AUTOMATED');os=$(if($Os.platform_scope-ceq 'NATIVE_WINDOWS10'){'Windows 10'}else{'Windows 11'});build=[int]$Os.BuildNumber;architecture='x64';powershell_version=$PSVersionTable.PSVersion.ToString()};
        registration=$null;forbidden_substitutions=[ordered]@{};fallback_acquisition=[ordered]@{policy='disabled_before_discovery';attempted=$false;used=$false};
        inventory=[ordered]@{purpose='verification_only';supplied_to_application=$false};before=@{application_venv_exists=$false;runtimes=@()};after=$null;
        discovery=$null;setup=$null;policy_component=@{status='NOT_RUN';organization_policy_denial=@{status='NOT_RUN';reason='Native policy component has not executed.'}};checks=[ordered]@{};artifacts=@();notes='NOT_RUN until every native stage actually succeeds; no runtime acquisition exists in this route.'}
    foreach($name in @('python_installed_or_updated_by_kit','manager_installed_or_updated_by_kit','private_runtime_fallback_attempted','private_runtime_fallback_used','runtime_supplied_by_kit','runtime_copied_or_cloned','explicit_python_path_supplied','hidden_runtime_selection_supplied','parent_path_changed','child_path_injected','localappdata_overridden','virtual_env_overridden','runtime_environment_injected','application_venv_prebuilt','alias_settings_changed')){$report.forbidden_substitutions[$name]=$false}
    foreach($name in @('registration_binding_verified','installed_only_listing','bounded_probe_completed','isolated_interpreter_verified','normal_setup_bat_passed','second_setup_noop','readiness_readonly','ordinary_application_bat_passed','manual_pdf_valid','auto1_pdf_valid','auto2_pdf_valid','full_page_coverage','source_and_neighbors_unchanged','unique_output_no_overwrite','unrelated_cwd_passed','manager_state_unchanged','global_packages_unchanged','parent_environment_unchanged')){$report.checks[$name]=$false}
    $records=[ordered]@{};foreach($kind in @('registration','discovery','setup','bat','pdf_validation','preservation')){$records[$kind]=New-Object 'Collections.Generic.List[object]'}
    $tokens=[ordered]@{};$tokens[$WorkRoot]='<WORK>';$tokens[$KitRoot]='<KIT>';$tokens[[Environment]::GetFolderPath('UserProfile')]='<USER>'
    $stage='inputs';$failure=$null;$inventoryBefore=$null;$environmentBefore=[Environment]::GetEnvironmentVariables('Process');$leases=New-Object 'Collections.Generic.List[IDisposable]'
    $report.forbidden_substitutions.explicit_python_path_supplied=(-not [string]::IsNullOrWhiteSpace($BarePythonPath) -or -not [string]::IsNullOrWhiteSpace($BaselinePythonPath))
    $report.forbidden_substitutions.runtime_supplied_by_kit=$report.forbidden_substitutions.explicit_python_path_supplied
    foreach($key in $environmentBefore.Keys){if(([string]$key -match '^(?i:VIRTUAL_ENV|PYTHONHOME|PYTHONPATH|PYMANAGER_.*|PYTHON_MANAGER_.*|WBS_.*PYTHON.*|WINBOOKSPLIT_.*PYTHON.*)$') -and -not [string]::IsNullOrWhiteSpace([string]$environmentBefore[$key])){$report.forbidden_substitutions.runtime_environment_injected=$true}}
    try{
        Assert-NativeStoreExistingInputs $BarePythonPath $BaselinePythonPath -ProvisionPrerequisites:$ProvisionPrerequisites
        if([IO.Directory]::Exists([IO.Path]::Combine($KitRoot,'.venv'))){$report.forbidden_substitutions.application_venv_prebuilt=$true;throw 'Store-existing kit root must not contain a prebuilt application venv.'}
        $stage='independent_existing_inventory'
        # This inventory is evidence and a tooling bootstrap only. Neither its
        # chosen path nor the resulting variable is passed to setup or the app.
        $inventoryBefore=Get-NativeStoreExistingInventory $KitRoot
        $package=$inventoryBefore.registration.Package
        $report.registration=[ordered]@{family=$package.PackageFamilyName;publisher=$package.Publisher;signature_kind=$package.SignatureKind;status=$package.Status;architecture=$package.Architecture;version=$package.Version}
        $records.registration.Add(@{role='independent_inventory';result=$inventoryBefore})
        $managerBefore=Get-NativeStoreExistingManagerState $inventoryBefore.registration
        $index=0;foreach($runtime in $inventoryBefore.runtimes){$index++;$tokens[$runtime.path]='<STORE_RUNTIME_'+$index+'>';$report.before.runtimes+=@(@{runtime_ref=$tokens[$runtime.path];version=$runtime.version;sha256=$runtime.sha256;managed_by_store=$true})}
        $report.checks.registration_binding_verified=$true;$report.checks.installed_only_listing=$true
        $packageBefore=@{};foreach($runtime in $inventoryBefore.runtimes){$packageBefore[$runtime.path]=Get-NativeStoreExistingPackages $runtime.path $WorkRoot}
        $stage='verify_extract';$app=[IO.Path]::Combine($WorkRoot,'application');$builder=[IO.Path]::Combine($KitRoot,'tools\compatibility\kit_builder.py');$smoke=[IO.Path]::Combine($KitRoot,'tools\compatibility\native_smoke.py')
        $extraction=[IO.Path]::Combine($WorkRoot,'store-extraction.json');$bootstrap=$inventoryBefore.runtimes[0].path
        $run=Invoke-NativeAcceptanceProcessCase $bootstrap @('-I','-B',$builder,'verify-extract','--kit-root',$KitRoot,'--bindings-sha256',$BindingsSha256,'--destination',$app,'--output',$extraction) $WorkRoot 120
        Assert-NativeAcceptanceProcess $run.probe
        if($run.failure -or $run.probe.ExitCode-ne 0){throw 'Store-existing exact asset verification/extraction failed.'}
        $verified=(Read-NativeAcceptanceJson $extraction).value
        if(-not $verified.success -or -not(Test-NativeAcceptanceProductionIdentity $verified.production $Bindings.production)){throw 'Store extraction identity mismatch.'}
        $leases.Add((New-Object WinBookSplitLogging.PathLease($app,$true)))
        foreach($relative in $verified.files.Keys){$path=[IO.Path]::Combine($app,$relative.Replace('/','\'));$null=Get-NativeAcceptanceOrdinaryPath $path;$leases.Add((New-Object WinBookSplitLogging.PathLease($path,$false,$false,$true)))}
        $records.discovery.Add(@{role='tooling_bootstrap';stage='exact_extraction';process=$run;receipt=$verified})
        if([IO.Directory]::Exists([IO.Path]::Combine($app,'.venv'))){$report.forbidden_substitutions.application_venv_prebuilt=$true;throw 'Fresh Store application extraction contains a venv.'}
        $appLiteral="'"+$app.Replace("'","''")+"'";$resultPath=[IO.Path]::Combine($WorkRoot,'store-discovery-local.json');$resultLiteral="'"+$resultPath.Replace("'","''")+"'"
        $body=@'
$ErrorActionPreference='Stop';$app=__APP__;foreach($name in @('Logging','Diagnostics','Process','Support','Paths','Runtime')){. ([IO.Path]::Combine($app,'engine','WinBookSplit.'+$name+'.ps1'))}
$result=@{base=$null;processing=$null;failure=$null}
try{$result.base=Resolve-WinBookSplitRuntime -ApplicationRoot $app -InterpreterOnly;$result.processing=try{Resolve-WinBookSplitRuntime -ApplicationRoot $app}catch{@{failure=$_.Exception.Data['Code'];attempts=$_.Exception.Data['Attempts']}}}catch{$result.failure=@{code=$_.Exception.Data['Code'];attempts=$_.Exception.Data['Attempts']}}
[IO.File]::WriteAllText(__RESULT__,($result|ConvertTo-Json -Depth 32 -Compress),[Text.UTF8Encoding]::new($false));if($null-eq $result.base){exit 3};exit 0
'@
        $stage='automatic_application_discovery';$child=Invoke-NativeStoreExistingChild $Inbox $WorkRoot 'store-discovery' ($body.Replace('__APP__',$appLiteral).Replace('__RESULT__',$resultLiteral))
        $discovery=(Read-NativeAcceptanceJson $resultPath 8388608).value;$records.discovery.Add(@{role='application';environment=@{};process=$child;result=$discovery})
        $base=$discovery.base;$matching=@($inventoryBefore.runtimes|Where-Object{[StringComparer]::OrdinalIgnoreCase.Equals($_.path,$base.Path) -and $_.version-ceq $base.Version})
        if($matching.Count-ne 1 -or $base.Source-cne 'manager' -or -not $base.Details.isolated -or -not $base.Details.dont_write_bytecode){throw 'Automatic application discovery did not choose a verified existing Store runtime.'}
        $storeAttempts=@($base.Attempts|Where-Object{$_.Accepted -and $_.Source-ceq 'manager_listing' -and [StringComparer]::OrdinalIgnoreCase.Equals($_.Path,$inventoryBefore.registration.Path) -and $_.Probe.PackageIdentityVerified})
        if($storeAttempts.Count-ne 1 -or $base.Path -cnotin @(ConvertFrom-WinBookSplitPythonListing $storeAttempts[0].Probe Manager)){throw 'Application did not itself complete the registered Store manager listing.'}
        $report.discovery=[ordered]@{mode='automatic';route='registered_store_manager';provenance='pre_existing_owner_install';runtime_ref=$tokens[$base.Path];version=$base.Version;interpreter_probe_exit=$base.Probe.ExitCode;package_state=$(if($discovery.processing.Path){'ready'}else{'setup_required'})}
        $report.checks.bounded_probe_completed=$true;$report.checks.isolated_interpreter_verified=$true
        $stage='automatic_setup_bat';$setupBat=[IO.Path]::Combine($app,'Setup-WinBookSplit.bat');$cmd=[IO.Path]::Combine([Environment]::SystemDirectory,'cmd.exe')
        if(($setupBat+$Wheelhouse)-match '[%"\r\n]'){throw 'Store setup path requires a literal cmd route.'}
        # A UTF8 .cmd wrapper switches its own console code page before reading
        # path-bearing lines. It never selects or configures an interpreter.
        $setupCommand='"'+$setupBat+'" -NonInteractive';if($Wheelhouse){$setupCommand+=' -Wheelhouse "'+(Get-NativeAcceptanceOrdinaryPath $Wheelhouse -Directory)+'"'}
        $setupWrapper=[IO.Path]::Combine($WorkRoot,'automatic-setup.cmd');[IO.File]::WriteAllText($setupWrapper,("@echo off`r`nsetlocal DisableDelayedExpansion`r`nchcp 65001 >nul`r`n"+$setupCommand+"`r`n"),[Text.UTF8Encoding]::new($false))
        foreach($id in @('install','check','noop')){
            $venv=[IO.Path]::Combine($app,'.venv');$snapshot=if($id-cne 'install'){Get-NativeAcceptanceEnvironmentSnapshot $venv}else{$null}
            if($id-ceq 'check'){$run=Invoke-NativeAcceptanceProcessCase $Inbox @('-NoProfile','-NonInteractive','-InputFormat','Text','-OutputFormat','Text','-ExecutionPolicy','Bypass','-File',[IO.Path]::Combine($app,'Setup-WinBookSplit.ps1'),'-Check','-NonInteractive') $WorkRoot 300}
            else{$run=Invoke-NativeAcceptanceProcessCase $cmd @('/d','/v:off','/c',$setupWrapper) $WorkRoot 300}
            $records.setup.Add(@{role='application';environment=@{};stage=$id;process=$run});Assert-NativeAcceptanceProcess $run.probe
            if($run.failure -or $run.probe.ExitCode-ne 0 -or -not $run.environment_before_after_restored){throw 'Normal Store automatic setup/check/no-op failed.'}
            $frame=Get-NativeAcceptanceFrame $run.probe.Stdout '[SETUP] '
            if($frame.status-cne 'success' -or $frame.code-cne $(if($id-ceq 'install'){'setup_complete'}else{'setup_ready'}) -or $frame.changed-ne ($id-ceq 'install')){throw 'Automatic setup typed result differs.'}
            if($id-cne 'install' -and $snapshot-cne (Get-NativeAcceptanceEnvironmentSnapshot $venv)){throw 'Store setup check/no-op changed application environment.'}
        }
        $python=[IO.Path]::Combine($app,'.venv\Scripts\python.exe')
        $run=Invoke-NativeAcceptanceProcessCase $python @('-I','-B','-c','import json,sys; print(json.dumps({"base":sys._base_executable,"executable":sys.executable}))') $WorkRoot 10
        Assert-NativeAcceptanceProcess $run.probe;$identity=[WinBookSplitSupport.StrictJson]::Parse($run.probe.Stdout.Trim())
        if($run.failure -or $run.probe.ExitCode-ne 0 -or -not[StringComparer]::OrdinalIgnoreCase.Equals($identity.base,$base.Path)){throw 'Automatically created venv base differs from discovered existing Store runtime.'}
        $records.setup.Add(@{role='post_setup_verification';stage='venv_base';process=$run;identity=$identity});$report.setup=[ordered]@{entry='Setup-WinBookSplit.bat';runtime_selection='automatic';exit_code=0;base_runtime_ref=$tokens[$base.Path];venv_base_runtime_ref=$tokens[$base.Path]}
        $report.checks.normal_setup_bat_passed=$true;$report.checks.second_setup_noop=$true;$report.checks.readiness_readonly=$true
        $stage='pdf_prepare';$cases=New-NativeAcceptanceDirectory ([IO.Path]::Combine($WorkRoot,'store-cases'));$planPath=[IO.Path]::Combine($WorkRoot,'store-plan.json')
        $run=Invoke-NativeAcceptanceProcessCase $python @('-I','-B',$smoke,'prepare-store','--application-root',$app,'--fixture-root',[IO.Path]::Combine($KitRoot,'fixtures'),'--work-root',$cases,'--python-path',$python,'--documents-root',[Environment]::GetFolderPath('MyDocuments'),'--policy-sha256',$Bindings.production.policy_sha256,'--output',$planPath) $WorkRoot 120
        Assert-NativeAcceptanceProcess $run.probe;if($run.failure -or $run.probe.ExitCode-ne 0){throw 'Store PDF recipe preparation failed.'}
        $read=Read-NativeAcceptanceJson $planPath 8388608;$plan=$read.value;$rows=New-Object 'Collections.Generic.List[object]'
        $stage='pdf_cases'
        foreach($case in $plan.cases){
            if($case.environment.Count-ne 0 -or '-PythonPath'-cin @($case.arguments)){throw 'Store PDF recipe attempted runtime/environment injection.'}
            $exe=if($case.exe_selector-ceq 'cmd'){$cmd}elseif($case.exe_selector-ceq 'inbox_ps'){$Inbox}else{throw 'Unknown Store PDF host.'}
            $run=Invoke-NativeAcceptanceProcessCase $exe @($case.arguments) $case.cwd 180 @($case.answers)
            Assert-NativeAcceptanceProcess $run.probe
            if($run.failure -or $run.probe.ExitCode-ne 0){throw 'Store PDF application case failed.'}
            $rows.Add(@{id=$case.id;probe=$run.probe;command=$run.command;cwd=$run.cwd;answers=$run.answers;environment_before_after_restored=$run.environment_before_after_restored;failure=$null})
            if($case.id-ceq 'bat'){$records.bat.Add(@{role='application';environment=@{};process=$run})}
        }
        $processPath=[IO.Path]::Combine($WorkRoot,'store-process.json');Write-NativeAcceptanceJson $processPath @{protocol='winbooksplit.native-store-process';schema_version=2;plan_sha256=$read.sha256;cases=$rows.ToArray()}
        $oraclePath=[IO.Path]::Combine($WorkRoot,'store-oracle.json');$stage='pdf_validate'
        $run=Invoke-NativeAcceptanceProcessCase $python @('-I','-B',$smoke,'validate-store','--plan',$planPath,'--process-report',$processPath,'--output',$oraclePath) $WorkRoot 120
        Assert-NativeAcceptanceProcess $run.probe;if($run.failure -or $run.probe.ExitCode-ne 0){throw 'Store physical-page/console oracle failed.'}
        $oracle=(Read-NativeAcceptanceJson $oraclePath 8388608).value;if(-not $oracle.success){throw 'Store PDF oracle did not pass.'}
        $applicationRecords=@($rows.ToArray()|ForEach-Object{@{role='application';environment=@{};process=$_}})
        $records.pdf_validation.Add(@{role='post_setup_verification';process=$run;oracle=$oracle;cases=$applicationRecords})
        foreach($name in @('ordinary_application_bat_passed','manual_pdf_valid','auto1_pdf_valid','auto2_pdf_valid','full_page_coverage','source_and_neighbors_unchanged','unique_output_no_overwrite','unrelated_cwd_passed')){$report.checks[$name]=$true}
        $stage='policy_component';$policy=Invoke-NativeStoreExistingPolicyComponent $app $WorkRoot $Inbox
        $report.policy_component=$policy.component
        foreach($record in $policy.records){$records.preservation.Add($record)}
        $stage='preservation';$after=Get-NativeStoreExistingInventory $KitRoot
        if(($after.registration.Package|ConvertTo-Json -Depth 10 -Compress)-cne ($inventoryBefore.registration.Package|ConvertTo-Json -Depth 10 -Compress)){throw 'Store manager registration changed.'}
        $managerAfter=Get-NativeStoreExistingManagerState $after.registration
        if(($managerBefore|ConvertTo-Json -Depth 20 -Compress)-cne ($managerAfter|ConvertTo-Json -Depth 20 -Compress)){throw 'Store manager/PATH forwarding alias observation changed.'}
        $report.after=@{application_venv_exists=[IO.Directory]::Exists([IO.Path]::Combine($app,'.venv'));runtimes=@()}
        foreach($runtime in $after.runtimes){if(-not $tokens.Contains($runtime.path)){throw 'Store runtime inventory changed.'};$report.after.runtimes+=@(@{runtime_ref=$tokens[$runtime.path];version=$runtime.version;sha256=$runtime.sha256;managed_by_store=$true})}
        if(($report.before.runtimes|ConvertTo-Json -Compress)-cne ($report.after.runtimes|ConvertTo-Json -Compress)){throw 'Store installed-runtime bytes or inventory changed.'}
        $environmentAfter=[Environment]::GetEnvironmentVariables('Process')
        if($environmentBefore.Count-ne $environmentAfter.Count){throw 'Controller environment changed.'}
        foreach($key in $environmentBefore.Keys){if($environmentBefore[$key]-cne $environmentAfter[$key]){throw 'Controller environment changed.'}}
        $packageAfter=@{};foreach($runtime in $after.runtimes){$packageAfter[$runtime.path]=Get-NativeStoreExistingPackages $runtime.path $WorkRoot;if($packageBefore[$runtime.path].probe.Stdout-cne $packageAfter[$runtime.path].probe.Stdout){throw 'Existing-runtime global package inventory changed.'}}
        $records.preservation.Add(@{role='independent_inventory';manager_before=$managerBefore;manager_after=$managerAfter;runtime_hashes_before=$report.before.runtimes;runtime_hashes_after=$report.after.runtimes;packages_before=$packageBefore;packages_after=$packageAfter;parent_environment_keys_unchanged=$true})
        $report.checks.manager_state_unchanged=$true;$report.checks.parent_environment_unchanged=$true
        # Existing-base distribution names, versions and locations are observed
        # before and after; this does not claim all global files were hashed.
        $report.checks.global_packages_unchanged=$true
        foreach($relative in $verified.files.Keys){
            $actual=Get-NativeAcceptanceFileHash ([IO.Path]::Combine($app,$relative.Replace('/','\'))) 8388608
            if($actual.sha256-cne $verified.files[$relative].sha256 -or $actual.bytes-ne $verified.files[$relative].bytes){throw 'Exact extracted application source changed during Store acceptance.'}
        }
        foreach($relative in $Bindings.fixtures.Keys){$actual=Get-NativeAcceptanceFileHash ([IO.Path]::Combine($KitRoot,'fixtures',$relative)) 8388608;if($actual.sha256-cne $Bindings.fixtures[$relative].sha256 -or $actual.bytes-ne $Bindings.fixtures[$relative].bytes){throw 'Authored kit fixture source changed.'}}
        foreach($relative in $Bindings.helpers.Keys){$actual=Get-NativeAcceptanceFileHash ([IO.Path]::Combine($KitRoot,$relative.Replace('/','\'))) 8388608;if($actual.sha256-cne $Bindings.helpers[$relative].sha256 -or $actual.bytes-ne $Bindings.helpers[$relative].bytes){throw 'Authenticated kit helper source changed.'}}
        foreach($lease in $leases){$lease.AssertUnchanged()}
        if($report.host.owner_desktop -and $report.host.os-ceq 'Windows 10' -and $Bindings.production.version-ceq '1.0.3'){$report.status='passed';$report.notes='Actual native commands and physical-page oracle passed; receipts require independent review.'}
        else{$report.notes='Controlled/Windows11 rehearsal only; owner Windows10 exact-v1.0.3 gate remains NOT_RUN.'}
    }catch{
        $failure=$_.Exception.Message;$report.notes='Blocked at '+$stage+'. Native acceptance has not passed.'
        $report.failure=@{stage=$stage;message=$failure}
        $kind=if($stage-eq 'independent_existing_inventory'){'registration'}elseif($stage-like '*setup*'){'setup'}elseif($stage-like 'pdf*'){'pdf_validation'}elseif($stage-in @('preservation','policy_component')){'preservation'}else{'discovery'}
        $failedRecord=@{role='failed_stage';stage=$stage;message=$failure}
        foreach($key in @('StoreInventoryReceipt','StoreChildReceipt','StorePolicyReceipt','Probe')){if($_.Exception.Data.Contains($key)){$failedRecord[$key]=$_.Exception.Data[$key]}}
        if($_.Exception.Data.Contains('StorePolicyReceipt')){$report.policy_component=$_.Exception.Data['StorePolicyReceipt'].component}
        $records[$kind].Add($failedRecord)
    }
    finally{foreach($lease in $leases){$lease.Dispose()}}
    if($failure){Write-NativeAcceptanceJson ([IO.Path]::Combine($WorkRoot,'store-local-failure.json')) @{stage=$stage;message=$failure}}
    $returnRoot=New-NativeAcceptanceDirectory ([IO.Path]::Combine($WorkRoot,'store-return'))
    $safe=ConvertTo-NativeStoreSafeValue $report $tokens
    & {
        $summaries=@{registration=$safe.registration;discovery=$safe.discovery;setup=$safe.setup;
            bat=@{ordinary_application_bat_passed=$safe.checks.ordinary_application_bat_passed;unrelated_cwd_passed=$safe.checks.unrelated_cwd_passed};
            pdf_validation=@{manual_pdf_valid=$safe.checks.manual_pdf_valid;auto1_pdf_valid=$safe.checks.auto1_pdf_valid;auto2_pdf_valid=$safe.checks.auto2_pdf_valid;full_page_coverage=$safe.checks.full_page_coverage;source_and_neighbors_unchanged=$safe.checks.source_and_neighbors_unchanged;unique_output_no_overwrite=$safe.checks.unique_output_no_overwrite};
            preservation=@{before=$safe.before;after=$safe.after;forbidden_substitutions=$safe.forbidden_substitutions;fallback_acquisition=$safe.fallback_acquisition;inventory=$safe.inventory;policy_component=$safe.policy_component}}
        foreach($kind in $records.Keys){
            $relative=$kind+'.json';$path=[IO.Path]::Combine($returnRoot,$relative)
            if($records[$kind].Count-eq 0){$records[$kind].Add(@{role='not_run';status='NOT_RUN';reason='prerequisite_failed';failed_stage=$stage})}
            Write-NativeAcceptanceJson $path @{schema_version=2;kind=$kind;run_id=$safe.run_id;source_commit=$safe.source_commit;zip_sha256=$safe.zip_sha256;summary=$summaries[$kind];records=(ConvertTo-NativeStoreSafeValue $records[$kind].ToArray() $tokens)}
            $safe.artifacts+=@(@{kind=$kind;path=$relative;sha256=(Get-NativeAcceptanceFileHash $path 4194304).sha256})
        }
    }
    Write-NativeAcceptanceJson ([IO.Path]::Combine($returnRoot,'native-store-report.json')) $safe
    $returnZip=[IO.Path]::Combine($WorkRoot,'native-store-result.zip');$returnIdentity=Write-NativeStoreExistingReturnZip $returnRoot $returnZip
    Write-Host ('Return only: '+$returnZip);Write-Host ('Raw local artifacts retained: '+$WorkRoot)
    return [pscustomobject]@{success=($safe.status-ceq 'passed');scenario='store-existing';status=$safe.status;report=$safe;report_directory=$returnRoot;return_zip=$returnZip;return_identity=$returnIdentity}
}
