<#
.SYNOPSIS
Execute the exact setup Action from four authored native dependency failures.
.DESCRIPTION
Uses the existing native process/job/stream transport and reopened-PDF oracle.
Only fresh owned copies receive test ADS or environments. Results are private
local receipts, not automatically shareable reports or final-ZIP acceptance.
Requires explicit trusted runtimes and an authenticated source binding. This
test permits the emitted setup command's explicit hash-locked online install.
#>
param(
    [Parameter(Mandatory=$true)][string]$RepositoryRoot,
    [Parameter(Mandatory=$true)][string]$WorkParent,
    [Parameter(Mandatory=$true)][string]$SourceBinding,
    [Parameter(Mandatory=$true)][string]$SourceBindingSha256,
    [Parameter(Mandatory=$true)][string]$Python315,
    [Parameter(Mandatory=$true)][string]$StalePython,
    [Parameter(Mandatory=$true)][string]$PowerShell7,
    [Parameter(Mandatory=$true)][string]$AuthoredPdf,
    [Parameter(Mandatory=$true)][string]$AuthoredPdfSha256
)
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)
[Console]::InputEncoding=New-Object Text.UTF8Encoding($false)
Import-Module ([IO.Path]::Combine($PSHOME,'Modules','Microsoft.PowerShell.Security','Microsoft.PowerShell.Security.psd1')) -ErrorAction Stop
Import-Module ([IO.Path]::Combine($PSHOME,'Modules','Microsoft.PowerShell.Utility','Microsoft.PowerShell.Utility.psd1')) -ErrorAction Stop
. (Join-Path $RepositoryRoot 'engine\WinBookSplit.Logging.ps1')
. (Join-Path $RepositoryRoot 'engine\WinBookSplit.Diagnostics.ps1')
. (Join-Path $RepositoryRoot 'engine\WinBookSplit.Process.ps1')
. (Join-Path $RepositoryRoot 'engine\WinBookSplit.Support.ps1')
Initialize-WinBookSplitLogTypes;Initialize-WinBookSplitSupport;Initialize-WinBookSplitProcess
. (Join-Path $RepositoryRoot 'tools\compatibility\Invoke-NativeAcceptance.ps1')
foreach($path in @($RepositoryRoot,$WorkParent)){ $null=Get-NativeAcceptanceOrdinaryPath $path -Directory }
foreach($path in @($Python315,$StalePython,$PowerShell7,$AuthoredPdf)){ $null=Get-NativeAcceptanceOrdinaryPath $path }
if((Get-NativeAcceptanceFileHash $SourceBinding 8388608).sha256-cne $SourceBindingSha256){throw 'Source binding hash differs.'}
$binding=(Read-NativeAcceptanceJson $SourceBinding 8388608).value
function Assert-RemedySourceBinding {
    foreach($name in $binding.source_manifest.Keys){
        if((Get-NativeAcceptanceFileHash (Join-Path $RepositoryRoot $name) 8388608).sha256-cne $binding.source_manifest[$name]){throw 'Reviewed source bytes changed.'}
    }
}
Assert-RemedySourceBinding
if((Get-NativeAcceptanceFileHash $AuthoredPdf 8388608).sha256-cne $AuthoredPdfSha256){throw 'Authored PDF pin differs.'}
function Write-RemedyFixtureText {
    param([string]$Path,[string]$Text)
    $null=Get-NativeAcceptanceOrdinaryPath ([IO.Path]::GetDirectoryName($Path)) -Directory
    $raw=([Text.UTF8Encoding]::new($false,$true)).GetBytes($Text)
    if([IO.Path]::GetExtension($Path)-ieq '.ps1'){$raw=[byte[]]([Text.Encoding]::UTF8.GetPreamble()+$raw)}
    $stream=[IO.File]::Open($Path,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::None)
    try{$stream.Write($raw,0,$raw.Length)}finally{$stream.Dispose()}
}
function Get-RemedyEnvironmentHashes {
    param([string]$Root)
    $null=Get-NativeAcceptanceOrdinaryPath $Root -Directory
    $rows=@{};$files=@([IO.Directory]::EnumerateFiles($Root,'*',[IO.SearchOption]::AllDirectories))
    if($files.Count-gt 4096){throw 'Authored environment exceeds the bounded fixture.'}
    foreach($path in $files){$rows[$path.Substring($Root.Length+1)]=(Get-NativeAcceptanceFileHash $path 33554432 -AllowEmpty).sha256}
    return $rows
}
function Assert-RemedyEnvironmentHashes {
    param([string]$Root,[Collections.IDictionary]$Before)
    $actual=Get-RemedyEnvironmentHashes $Root
    if($actual.Count-ne $Before.Count){throw 'Preserved environment member count changed.'}
    foreach($name in $Before.Keys){if(-not $actual.ContainsKey($name) -or $actual[$name]-cne $Before[$name]){throw 'Preserved environment bytes changed.'}}
}
function Set-RemedyInternetMarks {
    param([string]$Root)
    if((New-Object IO.DriveInfo([IO.Path]::GetPathRoot($Root))).DriveFormat-cne 'NTFS'){throw 'NTFS ADS prerequisite unavailable.'}
    $rows=New-Object 'Collections.Generic.List[object]'
    foreach($path in @([IO.Directory]::EnumerateFiles($Root,'*.ps1',[IO.SearchOption]::AllDirectories))){
        $before=Get-NativeAcceptanceFileHash $path 8388608
        $text="[ZoneTransfer]`r`nZoneId=3`r`n"
        if($null-ne (Get-Item -LiteralPath $path -Stream Zone.Identifier -ErrorAction SilentlyContinue)){throw 'Fresh owned copy unexpectedly has Internet metadata.'}
        Set-Content -LiteralPath $path -Stream Zone.Identifier -Value $text -Encoding ASCII -NoNewline -ErrorAction Stop
        $raw=[Text.Encoding]::ASCII.GetBytes((Get-Content -LiteralPath $path -Stream Zone.Identifier -Raw -Encoding ASCII -ErrorAction Stop))
        if([Text.Encoding]::ASCII.GetString($raw)-cne $text){throw 'Internet metadata write/read differs.'}
        $rows.Add([ordered]@{path=$path;member_sha256=$before.sha256;ads_base64=[Convert]::ToBase64String($raw);zone_id=3})
    }
    if($rows.Count-lt 3){throw 'Actual shipped scripts/helpers were not marked.'}
    return $rows.ToArray()
}
function Assert-RemedyInternetMarks {
    param([object[]]$Before)
    foreach($row in $Before){
        if((Get-NativeAcceptanceFileHash $row.path 8388608).sha256-cne $row.member_sha256 -or
            [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes((Get-Content -LiteralPath $row.path -Stream Zone.Identifier -Raw -Encoding ASCII -ErrorAction Stop)))-cne $row.ads_base64){throw 'Shipped bytes/Internet metadata changed.'}
    }
}

# One fixed parent adapter. Authored paths are JSON data. Only the exact captured
# Action may become code, after a strict one-command literal AST check against
# the finite expected host/script/base/options. No arbitrary user logs are read.
$parent=@'
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
Import-Module ([IO.Path]::Combine($PSHOME,'Modules','Microsoft.PowerShell.Security','Microsoft.PowerShell.Security.psd1')) -ErrorAction Stop
Import-Module ([IO.Path]::Combine($PSHOME,'Modules','Microsoft.PowerShell.Utility','Microsoft.PowerShell.Utility.psd1')) -ErrorAction Stop
[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)
[Console]::InputEncoding=New-Object Text.UTF8Encoding($false)
$p=[IO.File]::ReadAllText($env:WBS_KIT_CONTROL_ROOT)|ConvertFrom-Json
$savedPath=$env:PATH
$before=@(Get-ExecutionPolicy -List|ForEach-Object{@{scope=$_.Scope.ToString();policy=$_.ExecutionPolicy.ToString()}})
if(@($before|Where-Object{$_.scope-ceq 'Process' -and $_.policy-ceq 'RemoteSigned'}).Count-ne 1 -or
   @($before|Where-Object{$_.scope-cin @('MachinePolicy','UserPolicy') -and $_.policy-cne 'Undefined'}).Count-ne 0){throw 'Native policy prerequisite unavailable; respect organization controls.'}
Write-Host ('[PARENT_POLICY_BEFORE] '+($before|ConvertTo-Json -Compress))
Write-Host ('[PARENT_HOST] '+(@{path=[Diagnostics.Process]::GetCurrentProcess().MainModule.FileName;version=$PSVersionTable.PSVersion.ToString()}|ConvertTo-Json -Compress))
$childExit=1
try{
    $env:PATH=[string]$p.discovery_path
    if($p.mode-ceq 'direct'){
        try{& ([string]$p.setup_path) -NonInteractive -PythonPath ([string]$p.base_python);$childExit=$LASTEXITCODE}
        catch{
            if($_.Exception -isnot [System.Management.Automation.PSSecurityException]){throw}
            Write-Host ('[DIRECT_POLICY_DENIAL] '+(@{exception=$_.Exception.GetType().FullName;id=$_.FullyQualifiedErrorId;setup_entered=$false}|ConvertTo-Json -Compress))
            $childExit=1
        }
    }elseif($p.mode-ceq 'action'){
        $tokens=$errors=$null
        $ast=[Management.Automation.Language.Parser]::ParseInput([string]$p.action,[ref]$tokens,[ref]$errors)
        if($errors.Count-ne 0 -or $null-ne $ast.BeginBlock -or $null-ne $ast.ProcessBlock -or $null-ne $ast.ParamBlock -or
           $ast.EndBlock.Statements.Count-ne 1){throw 'Captured action is not one literal command.'}
        $pipeline=$ast.EndBlock.Statements[0]
        if($pipeline -isnot [Management.Automation.Language.PipelineAst] -or $pipeline.PipelineElements.Count-ne 1){throw 'Captured action pipeline rejected.'}
        $command=$pipeline.PipelineElements[0]
        if($command -isnot [Management.Automation.Language.CommandAst] -or $command.InvocationOperator-ne 'Ampersand' -or $command.Redirections.Count-ne 0){throw 'Captured action invocation rejected.'}
        $values=New-Object 'Collections.Generic.List[string]'
        foreach($element in $command.CommandElements){
            if($element -is [Management.Automation.Language.StringConstantExpressionAst]){$values.Add([string]$element.Value)}
            elseif($element -is [Management.Automation.Language.CommandParameterAst] -and $null-eq $element.Argument){$values.Add('-'+$element.ParameterName)}
            else{throw 'Captured action contains nonliteral code.'}
        }
        $expected=@([string]$p.host_path,'-NoProfile','-ExecutionPolicy','Bypass','-File',[string]$p.setup_path,'-NonInteractive')
        if($p.repair){$expected+='-Repair'}
        $expected+=@('-PythonPath',[string]$p.base_python)
        if($values.Count-ne $expected.Count){throw 'Captured action option count differs.'}
        for($index=0;$index-lt $expected.Count;$index++){if($values[$index]-cne $expected[$index]){throw 'Captured action literal differs.'}}
        Write-Host ('[EXACT_ACTION] '+(@{text=[string]$p.action;ast_literal_verified=$true}|ConvertTo-Json -Compress))
        & ([ScriptBlock]::Create([string]$p.action))
        $childExit=$LASTEXITCODE
    }else{
        if($p.mode-cne 'operation'){throw 'Unknown finite recipe.'}
        & ([string]$p.host_path) @([string[]]$p.operation_arguments)
        $childExit=$LASTEXITCODE
    }
}finally{
    $env:PATH=$savedPath
    $after=@(Get-ExecutionPolicy -List|ForEach-Object{@{scope=$_.Scope.ToString();policy=$_.ExecutionPolicy.ToString()}})
    Write-Host ('[PARENT_POLICY_AFTER] '+($after|ConvertTo-Json -Compress))
    Write-Host ('[PARENT_PATH_RESTORED] '+(@{restored=($env:PATH-ceq $savedPath)}|ConvertTo-Json -Compress))
}
exit $childExit
'@
$system=[Environment]::GetFolderPath([Environment+SpecialFolder]::System)
$inbox=[IO.Path]::Combine($system,'WindowsPowerShell','v1.0','powershell.exe')
$work=New-NativeAcceptanceDirectory ([IO.Path]::Combine($WorkParent,'T6-'+[guid]::NewGuid().ToString('N')))
$report=[ordered]@{protocol='winbooksplit.setup-remedy-roundtrip';schema_version=1;success=$false;
    scope='Actual owned-copy Windows native tests; source-bound, not final-ZIP or Explorer acceptance';
    started_utc=[datetime]::UtcNow.ToString('o');source_binding=$binding;cases=@();source_files_unchanged=$false;
    os=Get-NativeAcceptanceOs -AllowWindows11DryRun}
$members=@('VERSION','WinBookSplit.ps1','WinBookSplit.bat','Setup-WinBookSplit.ps1','Setup-WinBookSplit.bat','requirements.txt','Export-WinBookSplitDiagnostics.ps1')
$members+=@(Get-ChildItem -LiteralPath (Join-Path $RepositoryRoot 'engine') -File|ForEach-Object{'engine\'+$_.Name})
try{
foreach($hostKind in @('PS51','PS766')){foreach($state in @('fresh','stale')){
    Assert-RemedySourceBinding
    $id=$hostKind+'-'+$state;$case=New-NativeAcceptanceDirectory (Join-Path $work $id)
    $appName='-'+[char]0xc5+[char]0x65e5+[char]0x672c+" [1] O'Brien & (app)! %WBS_KIT_X%"
    $app=New-NativeAcceptanceDirectory (Join-Path $case $appName)
    $null=New-NativeAcceptanceDirectory (Join-Path $app 'engine')
    $hashes=@{}
    foreach($name in $members){
        [IO.File]::Copy((Join-Path $RepositoryRoot $name),(Join-Path $app $name),$false)
        $hashes[$name]=(Get-NativeAcceptanceFileHash (Join-Path $app $name) 8388608).sha256
    }
    $marks=@(Set-RemedyInternetMarks $app)
    $inputPdf=Join-Path $case '-Authored input [1].pdf';[IO.File]::Copy($AuthoredPdf,$inputPdf,$false)
    $output=New-NativeAcceptanceDirectory (Join-Path $case 'output')
    $cwd=if($state-ceq 'fresh'){$system}else{New-NativeAcceptanceDirectory (Join-Path $case 'unrelated-cwd')}
    $adapter=Join-Path $case 'parent.ps1';Write-RemedyFixtureText $adapter $parent
    $old=@{}
    if($state-ceq 'stale'){
        $created=Invoke-NativeAcceptanceProcessCase $StalePython @('-I','-B','-m','venv',(Join-Path $app '.venv')) $case 120
        if($null-eq $created.probe){throw 'Stale fixture failed before start.'}
        Write-NativeAcceptanceJson (Join-Path $case 'stale-create-process.json') $created.probe
        Assert-NativeAcceptanceProcess $created.probe
        if($created.failure -or $created.probe.ExitCode-ne 0){throw 'Actual stale venv creation failed.'}
        Write-RemedyFixtureText (Join-Path $app '.venv\owned-preserve.txt') 'Retain this entire authored old environment.'
        $old=Get-RemedyEnvironmentHashes (Join-Path $app '.venv')
    }
    $operation=@('-NoProfile','-ExecutionPolicy','Bypass','-File',(Join-Path $app 'WinBookSplit.ps1'),'-InputFile',$inputPdf,
        '-OutputDirectory',$output,'-Mode','Manual','-StartPages','1,4,7','-NoPause')
    $payload=[ordered]@{mode='direct';host_path=$inbox;setup_path=(Join-Path $app 'Setup-WinBookSplit.ps1');
        base_python=$Python315;repair=($state-ceq 'stale');action='';operation_arguments=$operation;
        discovery_path=([IO.Path]::GetDirectoryName($Python315)+';'+$system+';'+[IO.Path]::GetDirectoryName($inbox))}
    $route=[ordered]@{id=$id;status='NOT_RUN';app=$app;state=$state;source_members=$hashes;internet_marks=$marks;
        old_environment=$old;steps=@();operation_arguments=$operation;original_source_sha256=$AuthoredPdfSha256}
    $report.cases+=@($route)
    foreach($stage in @('direct','before','action','retry')){
        Assert-RemedySourceBinding
        $payload.mode=if($stage-ceq 'direct'){'direct'}elseif($stage-ceq 'action'){'action'}else{'operation'}
        $payloadPath=Join-Path $case ($stage+'-payload.json');Write-NativeAcceptanceJson $payloadPath $payload
        $answers=@();if($stage-ceq 'retry'){$answers=@('Y','N')}
        $hostPath=if($hostKind-ceq 'PS51'){$inbox}else{$PowerShell7}
        $run=Invoke-NativeAcceptanceProcessCase $hostPath @('-NoProfile','-NonInteractive','-ExecutionPolicy','RemoteSigned','-File',$adapter) $cwd 180 -Answers $answers -Environment @{WBS_KIT_CONTROL_ROOT=$payloadPath;WBS_KIT_X='MUST_NOT_EXPAND'}
        if($null-eq $run.probe){throw 'Native parent failed before start.'}
        $probePath=Join-Path $case ($stage+'-process.json');Write-NativeAcceptanceJson $probePath $run.probe
        $route.steps+=@([ordered]@{stage=$stage;command=$run.command;cwd=$run.cwd;answers=$run.answers;probe=$run.probe;environment_restored=$run.environment_before_after_restored})
        Assert-NativeAcceptanceProcess $run.probe
        if($run.failure -or -not $run.environment_before_after_restored){throw 'Native remedy transport failed.'}
        $before=Get-NativeAcceptanceFrame $run.probe.Stdout '[PARENT_POLICY_BEFORE] '
        $after=Get-NativeAcceptanceFrame $run.probe.Stdout '[PARENT_POLICY_AFTER] '
        if(($before|ConvertTo-Json -Depth 5 -Compress)-cne ($after|ConvertTo-Json -Depth 5 -Compress) -or
            -not (Get-NativeAcceptanceFrame $run.probe.Stdout '[PARENT_PATH_RESTORED] ').restored){throw 'Parent settings changed.'}
        Assert-RemedyInternetMarks $marks
        foreach($name in $hashes.Keys){if((Get-NativeAcceptanceFileHash (Join-Path $app $name) 8388608).sha256-cne $hashes[$name]){throw 'Application bytes changed.'}}
        if((Get-NativeAcceptanceFileHash $inputPdf 8388608).sha256-cne $AuthoredPdfSha256){throw 'Original PDF changed.'}
        if($stage-ceq 'direct'){
            $null=Get-NativeAcceptanceFrame $run.probe.Stdout '[DIRECT_POLICY_DENIAL] '
            if($run.probe.ExitCode-eq 0 -or $run.probe.Stdout.Contains('[SETUP] ')){throw 'Direct marked setup was not denied before entry.'}
            if($state-ceq 'stale'){Assert-RemedyEnvironmentHashes (Join-Path $app '.venv') $old}
            elseif([IO.Directory]::Exists((Join-Path $app '.venv'))){throw 'Blocked setup created an environment.'}
        }elseif($stage-ceq 'before'){
            $outcome=Get-NativeAcceptanceFrame $run.probe.Stdout '[OUTCOME] '
            if($run.probe.ExitCode-ne 3 -or $outcome.exit_code-ne 3 -or $outcome.written_count-ne 0){throw 'Original dependency failure differs.'}
            $lines=@($run.probe.Stdout.Split("`n")|Where-Object{$_.StartsWith('Action (PowerShell): ')})
            if($lines.Count-ne 1){throw 'Exact emitted Action was not present once.'}
            $payload.action=$lines[0].Substring('Action (PowerShell): '.Length).TrimEnd([char]13)
            $route.emitted_action=$payload.action
            if($state-ceq 'stale'){
                if(-not $run.probe.Stdout.Contains('3.14.7; application_venv')){throw 'Actual stale selected identity was not reported.'}
                Assert-RemedyEnvironmentHashes (Join-Path $app '.venv') $old
            }elseif(-not $run.probe.Stdout.Contains('3.15.0; ')-or [IO.Directory]::Exists((Join-Path $app '.venv'))){throw 'Actual bare identity was absent or processing installed a venv.'}
        }elseif($stage-ceq 'action'){
            $exact=Get-NativeAcceptanceFrame $run.probe.Stdout '[EXACT_ACTION] '
            $setup=Get-NativeAcceptanceFrame $run.probe.Stdout '[SETUP] '
            if($exact.text-cne $route.emitted_action -or -not $exact.ast_literal_verified -or $run.probe.ExitCode-ne 0 -or
                $setup.exit_code-ne 0 -or $setup.code-cne 'setup_complete' -or $setup.selected_version-cne '3.15.0' -or $setup.pypdf_version-cne '6.20.0'){throw 'Exact emitted remedy did not create selected315 locked runtime.'}
            $route.setup=$setup
            if($state-ceq 'stale'){Assert-RemedyEnvironmentHashes $setup.backup_path $old}
        }else{
            $outcome=Get-NativeAcceptanceFrame $run.probe.Stdout '[OUTCOME] '
            if($run.probe.ExitCode-ne 0 -or $outcome.exit_code-ne 0 -or $outcome.written_count-ne 3){throw 'Original operation retry failed.'}
            $oraclePath=Join-Path $case 'pdf-oracle.json'
            $oracle=Invoke-NativeAcceptanceProcessCase (Join-Path $app '.venv\Scripts\python.exe') @('-I','-B',(Join-Path $RepositoryRoot 'tests\compatibility\validate_setup_remedy.py'),
                '--application-root',$app,'--input-file',$inputPdf,'--input-sha256',$AuthoredPdfSha256,'--output-base',$output,'--process-report',$probePath,'--output',$oraclePath) $case 60
            if($null-eq $oracle.probe){throw 'Existing PDF oracle failed before start.'}
            Write-NativeAcceptanceJson (Join-Path $case 'pdf-oracle-process.json') $oracle.probe
            Assert-NativeAcceptanceProcess $oracle.probe
            if($oracle.failure -or $oracle.probe.ExitCode-ne 0){throw 'Reopened physical PDF/publication oracle failed.'}
            $route.pdf_oracle=(Read-NativeAcceptanceJson $oraclePath 8388608).value
            if(-not $route.pdf_oracle.success){throw 'Reopened PDF oracle did not pass.'}
        }
    }
    $route.status='PASS'
}}
Assert-RemedySourceBinding;$report.source_files_unchanged=$true;$report.success=$true
}catch{$report.exception=$_.Exception.ToString();$report.script_stack=$_.ScriptStackTrace;Write-Error $_ -ErrorAction Continue}
finally{$report.finished_utc=[datetime]::UtcNow.ToString('o');$path=Join-Path $work 'native-result.json';Write-NativeAcceptanceJson $path $report;Write-Output ('RECEIPT '+$path)}
if($report.success){exit 0};exit 1
