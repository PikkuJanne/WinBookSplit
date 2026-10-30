# Checkpoint-4 focused exact-package workflow. No runtime acquisition or injection.
[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$KitRoot,
      [Parameter(Mandatory=$true)][ValidatePattern('^[0-9a-f]{64}$')][string]$BindingsSha256)
$ErrorActionPreference='Stop'
$ProgressPreference='SilentlyContinue'
[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)
foreach($module in @('Microsoft.PowerShell.Utility','Microsoft.PowerShell.Management')){
    Import-Module -Name ([IO.Path]::Combine($PSHOME,'Modules',$module,$module+'.psd1')) -Scope Local -ErrorAction Stop
}
$root=[IO.Path]::GetFullPath($KitRoot)
$bindingPath=[IO.Path]::Combine($root,'bindings.json')
if((Get-FileHash -LiteralPath $bindingPath -Algorithm SHA256).Hash.ToLowerInvariant()-cne $BindingsSha256){throw 'Kit bindings hash differs from launcher.'}
$bindings=Get-Content -LiteralPath $bindingPath -Raw -Encoding UTF8|ConvertFrom-Json
foreach($property in $bindings.files.PSObject.Properties){
    $name=$property.Name
    if($name -match '(^/|\\|(^|/)\.\.?(/|$)|:)' -or [string]::IsNullOrWhiteSpace($name)){throw 'Unsafe kit member.'}
    $item=Get-Item -LiteralPath ([IO.Path]::Combine($root,$name.Replace('/','\')))
    if($item.PSIsContainer -or $item.Length-ne $property.Value.bytes -or
        (Get-FileHash -LiteralPath $item.FullName -Algorithm SHA256).Hash.ToLowerInvariant()-cne $property.Value.sha256){throw 'Kit member bytes differ from bindings.'}
}
foreach($name in @('Logging','Support','Runtime')){. ([IO.Path]::Combine($root,'support','engine','WinBookSplit.'+$name+'.ps1'))}
Initialize-WinBookSplitLogTypes
Initialize-WinBookSplitSupport
foreach($name in @('Invoke-NativeAcceptance','Invoke-NativeDiscoveryCollector','Invoke-StoreExistingAcceptance')){. ([IO.Path]::Combine($root,'tools','compatibility',$name+'.ps1'))}
$root=Get-NativeAcceptanceOrdinaryPath $root -Directory
$work=New-NativeAcceptanceDirectory ([IO.Path]::Combine($root,'RUN-'+[Guid]::NewGuid().ToString('N')))
$report=[ordered]@{schema_version=1;scope='FUNCTIONAL_NATIVE_EXACT_PACKAGE_NO_RUNTIME_SUBSTITUTION';status='NOT_RUN';
    source_commit=$bindings.source_commit;source_tree=$bindings.source_tree;package_sha256=$bindings.files.'package.zip'.sha256;
    host=@{powershell=$PSVersionTable.PSVersion.ToString();os=[Environment]::OSVersion.Version.ToString()};
    stage='start';failure=$null;records=@();checks=[ordered]@{};
    second_run_network_activity=@{status='NOT_RUN';reason='No packet/process network trace is collected; no setup frames plus unchanged venv do not prove zero network activity.'};
    windows10=@{status='NOT_RUN';reason='Only the actual host is tested.'};
    private_book=@{status='NOT_RUN';reason='The private book is not in this public synthetic kit.'}}
$environmentBefore=[Environment]::GetEnvironmentVariables('Process')
$source=$app=$null
function Add-FunctionalRecord {
    param([string]$Kind,$Value)
    $report.records+=@(@{kind=$Kind;value=$Value})
}
function Invoke-FunctionalCase {
    param([string]$Name,[string]$Exe,[string[]]$Arguments,[string[]]$Answers=@())
    $run=Invoke-NativeAcceptanceProcessCase -Path $Exe -Arguments $Arguments -WorkingDirectory $cwd -TimeoutSeconds 300 -Answers $Answers
    if($run.failure -is [Exception]){$run.failure=$run.failure.Message}
    Add-FunctionalRecord $Name $run
    Assert-NativeAcceptanceProcess $run.probe
    if($run.failure -or -not $run.environment_before_after_restored){throw ('Native process failed: '+$Name)}
    return $run
}
function Get-FunctionalOutcome {
    param($Run)
    $frame=Get-NativeAcceptanceFrame $Run.probe.Stdout '[OUTCOME] '
    if($frame.exit_code-ne $Run.probe.ExitCode){throw 'BAT exit differs from its outcome.'}
    return $frame
}
function Assert-FunctionalNoInstall {
    param($Run,[switch]$CheckFrame)
    if(((-not $CheckFrame) -and $Run.probe.Stdout -match '(?m)^\[SETUP\] ') -or [IO.Directory]::Exists([IO.Path]::Combine($app,'.venv'))){throw 'Read-only/declined case performed setup.'}
    if($CheckFrame){$frame=Get-NativeAcceptanceFrame $Run.probe.Stdout '[SETUP] ';if($frame.changed -or @($frame.operations).Count-ne 0 -or $frame.code-cne 'setup_required'){throw 'Readiness Check performed setup or misclassified missing package.'}}
    $documentsAfter=@([IO.Directory]::GetFileSystemEntries($documents))
    foreach($path in $documentsBefore){if($path -cnotin $documentsAfter){throw 'Read-only/declined case removed a default-output entry.'}}
    foreach($path in $documentsAfter){
        if($path -cin $documentsBefore){continue}
        # Existing failure diagnostics are expected; only the finite owned
        # console receipt directory is allowed, never chapter/staging output.
        if([IO.Path]::GetFileName($path)-cnotmatch '^\.WinBookSplit-console-[0-9a-f]{32}$' -or -not [IO.Directory]::Exists($path)){throw 'Read-only/declined case created a chapter/output entry.'}
        foreach($entry in [IO.Directory]::GetFileSystemEntries($path)){
            if([IO.Directory]::Exists($entry) -or [IO.Path]::GetFileName($entry)-cnotin @('.WinBookSplit-console-owner.json','console.log','WinBookSplit_Run.json')){throw 'Read-only/declined case created non-diagnostic output.'}
        }
    }
    if(@([IO.Directory]::GetDirectories($books,'*',[IO.SearchOption]::AllDirectories)).Count-ne 0 -or @([IO.Directory]::GetFiles($books,'*.pdf',[IO.SearchOption]::AllDirectories)).Count-ne 1){throw 'Read-only/declined case created chapter files or output directories.'}
}
try{
    $report.stage='input_guards'
    Assert-NativeStoreExistingInputs
    $os=Get-NativeAcceptanceOs -AllowWindows11DryRun
    $report.host.os_facts=ConvertTo-NativeAcceptanceSafeOs $os
    $report.checks.no_runtime_environment_injection=$true
    $report.stage='exact_extract'
    $manifest=(Read-NativeAcceptanceJson ([IO.Path]::Combine($root,'release-manifest.json'))).value
    if($manifest.source_commit-cne $bindings.source_commit -or $manifest.source_tree-cne $bindings.source_tree -or
        $manifest.archive.sha256-cne $report.package_sha256){throw 'Package source binding differs.'}
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $app=New-NativeAcceptanceDirectory ([IO.Path]::Combine($work,'application'))
    $zip=[IO.Compression.ZipFile]::OpenRead([IO.Path]::Combine($root,'package.zip'))
    try{
        if($zip.Entries.Count-ne @($manifest.files).Count){throw 'Package member count differs.'}
        foreach($row in $manifest.files){
            if($row.path -match '(^/|\\|(^|/)\.\.?(/|$)|:)' -or [string]::IsNullOrWhiteSpace($row.path)){throw 'Unsafe package member.'}
            $entries=@($zip.Entries|Where-Object{$_.FullName-ceq $row.path})
            if($entries.Count-ne 1 -or $entries[0].Length-ne $row.bytes){throw 'Package member differs.'}
            $target=[IO.Path]::Combine($app,$row.path.Replace('/','\'))
            $null=[IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($target))
            [IO.Compression.ZipFileExtensions]::ExtractToFile($entries[0],$target,$false)
            if((Get-NativeAcceptanceFileHash $target 8388608).sha256-cne $row.sha256){throw 'Extracted package hash differs.'}
        }
    }finally{$zip.Dispose()}
    if([IO.Directory]::Exists([IO.Path]::Combine($app,'.venv'))){throw 'Fresh package contains a prebuilt environment.'}
    $report.checks.exact_source_package_extraction=$true
    $books=New-NativeAcceptanceDirectory ([IO.Path]::Combine($work,'books'))
    $cwd=New-NativeAcceptanceDirectory ([IO.Path]::Combine($work,'unrelated-cwd'))
    $documents=[Environment]::GetFolderPath('MyDocuments');$documentsBefore=@([IO.Directory]::GetFileSystemEntries($documents)|Sort-Object)
    $source=[IO.Path]::Combine($books,([char]0x00c5).ToString()+([char]0x65e5)+([char]0x672c)+' [1] & combined.pdf')
    [IO.File]::Copy([IO.Path]::Combine($root,'combined.pdf'),$source,$false)
    $prior=[IO.Path]::Combine($books,'prior-output.txt');[IO.File]::WriteAllText($prior,'Authored prior output stays unchanged.');$priorHash=(Get-NativeAcceptanceFileHash $prior 8388608).sha256
    $inbox=[IO.Path]::Combine([Environment]::SystemDirectory,'WindowsPowerShell','v1.0','powershell.exe')
    $cmd=[IO.Path]::Combine([Environment]::SystemDirectory,'cmd.exe')
    $bat=[IO.Path]::Combine($app,'WinBookSplit.bat')
    if(($bat+$source)-match '[%"\r\n]'){throw 'Literal BAT wrapper path cannot contain percent, quotes or newlines.'}
    $wrapper=[IO.Path]::Combine($cwd,'invoke.cmd')
    [IO.File]::WriteAllText($wrapper,('@echo off'+"`r`n"+'setlocal DisableDelayedExpansion'+"`r`n"+'chcp 65001 >nul'+"`r`n"+'"'+$bat+'" "'+$source+'"'+"`r`n"),[Text.UTF8Encoding]::new($false))
    $report.stage='natural_discovery'
    $base=$processing=$null;$baseFailure=$processingFailure=$null
    try{$base=Resolve-WinBookSplitRuntime -ApplicationRoot $app -DocumentPath $source -InterpreterOnly}catch{$baseFailure=@{code=$_.Exception.Data['Code'];message=$_.Exception.Message;attempts=$_.Exception.Data['Attempts']}}
    if($base){try{$processing=Resolve-WinBookSplitRuntime -ApplicationRoot $app -DocumentPath $source}catch{$processingFailure=@{code=$_.Exception.Data['Code'];message=$_.Exception.Message;attempts=$_.Exception.Data['Attempts']}}}
    Add-FunctionalRecord 'natural_discovery' @{base=$base;processing=$processing;base_failure=$baseFailure;processing_failure=$processingFailure;verification_only_not_application_configuration=$true}
    # Always exercise actual packaged BAT safely, even when Store prerequisites
    # are unavailable. A healthy pre-existing base is cancelled at mode choice.
    $report.stage='ordinary_packaged_bat'
    $answers=if($processing){@('C','X')}else{@('N','X')}
    $decline=Invoke-FunctionalCase 'ordinary_bat_decline_or_prerequisite' $cmd @('/d','/v:off','/c',$wrapper) $answers
    $declineOutcome=Get-FunctionalOutcome $decline
    Assert-FunctionalNoInstall $decline
    if($declineOutcome.written_count-ne 0 -or $null-ne $declineOutcome.final_directory -or $declineOutcome.exit_code-eq 0){throw 'Declined/prerequisite BAT falsely claimed output.'}
    if(@([IO.Directory]::GetFiles($books,'*.pdf')).Count-ne 1){throw 'Declined/prerequisite BAT created chapter files.'}
    if($baseFailure){
        $known=@('no_candidates','missing_python','manager_not_registered','manager_no_runtimes','python_security_floor','python_compatibility_floor','unsupported_python','unsupported_architecture','unsupported_flavor','unstable_python')
        $incidental=@('rejected_entry_point','manager_entry_rejected','manager_registered','installed_runtimes_listed')
        $reasons=@($baseFailure.attempts|ForEach-Object{$_.Reason})
        if($baseFailure.code-cne 'runtime_not_found' -or @($reasons|Where-Object{$_ -cin $known}).Count-eq 0 -or @($reasons|Where-Object{$_ -cnotin ($known+$incidental)}).Count-gt 0){throw 'Unexpected natural discovery failure cannot qualify as missing prerequisite.'}
    }
    $report.checks.ordinary_packaged_bat_observed=$true
    $report.stage='existing_store_prerequisite'
    try{$inventory=Get-NativeStoreExistingInventory $app}catch{
        Add-FunctionalRecord 'store_inventory_unavailable' @{message=$_.Exception.Message;receipt=$_.Exception.Data['StoreInventoryReceipt']}
        if($_.Exception.Message -match 'Existing Store registration is absent or ambiguous|No compatible pre-existing Store interpreter|Existing Store runtime inventory is empty'){
            $report.failure=@{kind='missing_native_prerequisite';message=$_.Exception.Message};$report.status='NOT_RUN'
            throw 'FUNCTIONAL_PREREQUISITE_NOT_RUN'
        }
        throw
    }
    Add-FunctionalRecord 'store_inventory' $inventory
    $matching=@($inventory.runtimes|Where-Object{$base -and [StringComparer]::OrdinalIgnoreCase.Equals($_.path,$base.Path) -and $_.version-ceq $base.Version})
    if($matching.Count-ne 1 -or $base.Source-cne 'manager'){throw 'Natural application discovery did not select the verified Store interpreter.'}
    $storeAttempts=@($base.Attempts|Where-Object{$_.Accepted -and $_.Source-ceq 'manager_listing' -and [StringComparer]::OrdinalIgnoreCase.Equals($_.Path,$inventory.registration.Path) -and $_.Probe.PackageIdentityVerified})
    if($storeAttempts.Count-ne 1){throw 'Natural discovery did not itself complete the registered Store listing.'}
    if($processing -or $processingFailure.code-cne 'setup_required' -or @($processingFailure.attempts|Where-Object{$_.IdentityVerified -and $_.Reason-ceq 'missing_pypdf' -and [StringComparer]::OrdinalIgnoreCase.Equals($_.Path,$base.Path)}).Count-lt 1){
        $report.failure=@{kind='missing_native_prerequisite';message='Naturally selected Store base must lack required pypdf; no environment is modified to manufacture this condition.'};throw 'FUNCTIONAL_PREREQUISITE_NOT_RUN'
    }
    $managerBefore=Get-NativeStoreExistingManagerState $inventory.registration
    $packagesBefore=Get-NativeStoreExistingPackages $base.Path $cwd
    $distributionRows=[WinBookSplitSupport.StrictJson]::Parse($packagesBefore.probe.Stdout.Trim())
    foreach($distribution in $distributionRows){if([StringComparer]::OrdinalIgnoreCase.Equals($distribution[0],'pypdf')){throw 'Selected Store base contains pypdf; cold-start missing-package prerequisite is false.'}}
    if($declineOutcome.code-cne 'cancelled' -or $declineOutcome.exit_code-ne 130){throw 'Setup decline outcome differs.'}
    $report.checks.decline_preserved=$true
    $report.stage='read_only_cases'
    foreach($variant in @('noninteractive','preview','check')){
        if($variant-ceq 'check'){$argv=@('-NoProfile','-ExecutionPolicy','Bypass','-File',[IO.Path]::Combine($app,'Setup-WinBookSplit.ps1'),'-Check','-NonInteractive')}
        else{$argv=@('-NoProfile','-ExecutionPolicy','Bypass','-File',[IO.Path]::Combine($app,'WinBookSplit.ps1'),'-InputFile',$source,'-Mode','Manual','-StartPages','1,3,5','-NonInteractive','-NoPause');if($variant-ceq 'preview'){$argv+=@('-Preview')}}
        $run=Invoke-FunctionalCase $variant $inbox $argv
        Assert-FunctionalNoInstall $run -CheckFrame:($variant-ceq 'check')
        if($run.probe.ExitCode-ne 3){throw 'Read-only missing-package case must exit 3.'}
        if($variant-cne 'check'){$outcome=Get-FunctionalOutcome $run;if($outcome.code-cne 'setup_required'){throw 'Missing package incorrectly classified.'}}
    }
    $report.checks.read_only_no_setup=$true
    $report.stage='consent_setup_resume'
    $cold=Invoke-FunctionalCase 'cold_normal_bat' $cmd @('/d','/v:off','/c',$wrapper) @('Y','M','1,3,5','Y','N','X')
    $outcome=Get-FunctionalOutcome $cold
    if([regex]::Matches($cold.probe.Stdout,[regex]::Escape('First-time setup is required.')).Count-ne 1){throw 'Cold launch did not offer setup exactly once.'}
    $setup=Get-NativeAcceptanceFrame $cold.probe.Stdout '[SETUP] '
    if($setup.status-cne 'success' -or $setup.code-cne 'setup_complete' -or -not $setup.changed -or $setup.exit_code-ne 0){throw 'Consent setup did not complete exactly once.'}
    $report.checks.consent_setup_resume=$true
    $python=[IO.Path]::Combine($app,'.venv','Scripts','python.exe')
    $resolved=Resolve-WinBookSplitRuntime -ApplicationRoot $app -DocumentPath $source
    if($resolved.Source-cne 'application_venv' -or -not [StringComparer]::OrdinalIgnoreCase.Equals($resolved.Path,$python)){throw 'App runtime was not re-resolved to this application venv.'}
    Add-FunctionalRecord 'healthy_app_runtime' $resolved
    $identity=Invoke-FunctionalCase 'venv_base' $python @('-I','-B','-c','import json,sys;print(json.dumps({"base":sys._base_executable,"executable":sys.executable}))')
    $identityJson=[WinBookSplitSupport.StrictJson]::Parse($identity.probe.Stdout.Trim())
    if($identity.probe.ExitCode-ne 0 -or -not [StringComparer]::OrdinalIgnoreCase.Equals($identityJson.base,$base.Path)){throw 'Setup venv base differs from naturally discovered existing Store Python.'}
    $report.checks.venv_base_and_locked_package_origin=$true
    foreach($name in @('first','second')){
        if($name-ceq 'second'){
            $snapshot=Get-NativeAcceptanceEnvironmentSnapshot ([IO.Path]::Combine($app,'.venv'))
            $firstDirectory=$outcome.final_directory;$firstFiles=@{}
            foreach($item in [IO.Directory]::GetFiles($firstDirectory,'*.pdf')){$firstFiles[$item]=(Get-NativeAcceptanceFileHash $item 8388608).sha256}
            $repeat=Invoke-FunctionalCase 'repeat_normal_bat' $cmd @('/d','/v:off','/c',$wrapper) @('M','1,3,5','Y','N','X')
            $outcome=Get-FunctionalOutcome $repeat
            if($repeat.probe.Stdout -match '(?m)^\[SETUP\] |First-time setup is required|Downloading|Installing collected packages' -or
                $snapshot-cne (Get-NativeAcceptanceEnvironmentSnapshot ([IO.Path]::Combine($app,'.venv')))){throw 'Repeat changed the environment or performed setup/install.'}
            if([StringComparer]::OrdinalIgnoreCase.Equals($firstDirectory,$outcome.final_directory)){throw 'Repeat overwrote the first output directory.'}
            foreach($path in $firstFiles.Keys){if((Get-NativeAcceptanceFileHash $path 8388608).sha256-cne $firstFiles[$path]){throw 'Repeat changed prior output.'}}
            $report.checks.repeat_no_setup_or_reinstall_and_no_overwrite=$true
        }
        $outcomePath=[IO.Path]::Combine($work,$name+'-outcome.json');Write-NativeAcceptanceJson $outcomePath $outcome
        $oracle=Invoke-FunctionalCase ($name+'_physical_page_oracle') $python @('-I','-B',[IO.Path]::Combine($root,'tools','compatibility','functional_oracle.py'),'--expected',[IO.Path]::Combine($root,'expected.json'),'--source',$source,'--outcome',$outcomePath)
        if($oracle.probe.ExitCode-ne 0){throw 'Independent page/order/source/output oracle failed.'}
        $report.checks[$name+'_six_pages_in_order']=$true
    }
    $report.stage='preservation'
    $managerAfter=Get-NativeStoreExistingManagerState $inventory.registration
    $packagesAfter=Get-NativeStoreExistingPackages $base.Path $cwd
    if(($managerBefore|ConvertTo-Json -Depth 32 -Compress)-cne ($managerAfter|ConvertTo-Json -Depth 32 -Compress) -or $packagesBefore.probe.Stdout-cne $packagesAfter.probe.Stdout){throw 'Manager state or global packages changed.'}
    foreach($runtime in $inventory.runtimes){if((Get-NativeAcceptanceFileHash $runtime.path 8388608).sha256-cne $runtime.sha256){throw 'Existing runtime bytes changed.'}}
    $report.checks.manager_runtime_and_global_packages_preserved=$true
    $report.status='PASS_WITH_NETWORK_NOT_RUN'
    if($os.platform_scope-ceq 'NATIVE_WINDOWS10'){$report.windows10=@{status='PASS_WITH_NETWORK_NOT_RUN';reason='Actual Windows10 workflow completed; network observation remains NOT_RUN.'}}
}catch{
    Write-Host ('Stopped at '+$report.stage+': '+$_.Exception.Message)
    if($_.Exception.Message-cne 'FUNCTIONAL_PREREQUISITE_NOT_RUN'){$report.status='FAIL';$report.failure=@{stage=$report.stage;message=$_.Exception.Message}}
}finally{
    try{
        $environmentAfter=[Environment]::GetEnvironmentVariables('Process')
        if($environmentBefore.Count-ne $environmentAfter.Count){throw 'Controller environment changed.'}
        foreach($key in $environmentBefore.Keys){if($environmentBefore[$key]-cne $environmentAfter[$key]){throw 'Controller environment changed.'}}
        if($source -and (Get-NativeAcceptanceFileHash $source 8388608).sha256-cne $bindings.files.'combined.pdf'.sha256){throw 'Synthetic input changed.'}
        if($prior -and (Get-NativeAcceptanceFileHash $prior 8388608).sha256-cne $priorHash){throw 'Prior output marker changed.'}
        if($app){foreach($row in $manifest.files){if((Get-NativeAcceptanceFileHash ([IO.Path]::Combine($app,$row.path.Replace('/','\'))) 8388608).sha256-cne $row.sha256){throw 'Application bytes changed.'}}}
        $report.checks.source_application_prior_and_parent_environment_preserved=$true
    }catch{$report.status='FAIL';$report.failure=@{stage='final_preservation';message=$_.Exception.Message}}
    $raw=[IO.Path]::Combine($work,'raw-local-report.json');Write-NativeAcceptanceJson $raw $report
    $tokens=[ordered]@{};$tokens[$work]='<WORK>';$tokens[$root]='<KIT>';$tokens[[Environment]::GetFolderPath('UserProfile')]='<USER>'
    $safe=ConvertTo-NativeStoreSafeValue $report $tokens
    $returnPath=[IO.Path]::Combine($work,'functional-report.json');Write-NativeAcceptanceJson $returnPath $safe
    $returnZip=[IO.Path]::Combine($work,'SEND-THIS-REPORT.zip')
    $stream=[IO.File]::Open($returnZip,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::Read)
    $archive=New-Object IO.Compression.ZipArchive($stream,[IO.Compression.ZipArchiveMode]::Create,$false)
    try{$entry=$archive.CreateEntry('functional-report.json');$member=$entry.Open();try{$bytes=[IO.File]::ReadAllBytes($returnPath);$member.Write($bytes,0,$bytes.Length)}finally{$member.Dispose()}}finally{$archive.Dispose();$stream.Dispose()}
    Write-Host ('RESULT: '+$report.status)
    Write-Host ('Return only: '+$returnZip)
    Write-Host ('Raw local artifacts retained: '+$work)
}
if($report.status-ceq 'FAIL'){exit 1}
if($report.status-ceq 'NOT_RUN'){exit 3}
exit 0
