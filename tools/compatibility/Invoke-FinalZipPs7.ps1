# Supplemental adapter only. Invoke after the existing authenticated native kit
# has completed successfully and its SDK functions are loaded in this session.
function Invoke-FinalZipPs7Acceptance {
    [CmdletBinding()]
    param([Parameter(Mandatory=$true)][string]$KitRoot,
        [Parameter(Mandatory=$true)][string]$BindingsSha256,
        [Parameter(Mandatory=$true)][string]$PriorReport,
        [Parameter(Mandatory=$true)][string]$PriorReportSha256,
        [Parameter(Mandatory=$true)][string]$NativePlan,
        [Parameter(Mandatory=$true)][string]$ExtractionReceipt,
        [Parameter(Mandatory=$true)][string]$PowerShell7,
        [Parameter(Mandatory=$true)][string]$AdapterPython,
        [Parameter(Mandatory=$true)][string]$AdapterPythonSha256,
        [Parameter(Mandatory=$true)][string]$WorkRoot,
        [string]$ExpectedPowerShell7Version='7.6.6')
    $ErrorActionPreference='Stop'
    $leases=New-Object 'Collections.Generic.List[IDisposable]'
    $heldDirectories=New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    try {
        $root=Get-NativeAcceptanceOrdinaryPath $KitRoot -Directory
        $read=Read-NativeAcceptanceJson ([IO.Path]::Combine($root,'kit-bindings.json'))
        if($read.sha256 -cne $BindingsSha256){throw 'Supplemental bindings pin differs.'}
        $bindings=$read.value;$sdk=[IO.Path]::Combine($root,'tools\compatibility\Invoke-NativeAcceptance.ps1')
        $sdkHash=Get-NativeAcceptanceFileHash $sdk 8388608
        if($sdkHash.sha256 -cne $bindings.helpers['tools/compatibility/Invoke-NativeAcceptance.ps1'].sha256){throw 'Supplemental SDK source differs.'}
        foreach($name in @('Invoke-NativeAcceptanceProcessCase','Assert-NativeAcceptanceProcess','Get-NativeAcceptanceEnvironmentSnapshot','Read-NativeAcceptanceJson')){
            $command=Get-Command $name -CommandType Function -ErrorAction Stop
            if(-not [StringComparer]::OrdinalIgnoreCase.Equals($command.ScriptBlock.File,$sdk)){throw 'Supplemental adapter needs the authenticated existing SDK definitions.'}
        }
        if($ExpectedPowerShell7Version -cnotmatch '^7\.[0-9]+\.[0-9]+$'){throw 'Supplemental PS7 expectation must be an actual stable version.'}
        $PowerShell7=Get-NativeAcceptanceOrdinaryPath $PowerShell7
        $AdapterPython=Get-NativeAcceptanceOrdinaryPath $AdapterPython
        if((Get-NativeAcceptanceFileHash $AdapterPython 8388608).sha256 -cne $AdapterPythonSha256){throw 'Supplemental Python adapter differs from its reviewed source.'}
        if((Get-NativeAcceptanceFileHash $PriorReport 8388608).sha256 -cne $PriorReportSha256){throw 'Supplemental prior report differs from its external pin.'}
        $prior=(Read-NativeAcceptanceJson $PriorReport 8388608).value
        $nativeRead=Read-NativeAcceptanceJson $NativePlan 8388608;$native=$nativeRead.value
        if($prior.success -isnot [bool] -or -not $prior.success -or $prior.cleanup_safe -isnot [bool] -or -not $prior.cleanup_safe -or
            $prior.identities.bindings_sha256 -cne $BindingsSha256 -or $nativeRead.sha256 -cne $prior.oracle.plan_sha256){throw 'Supplemental original accepted plan differs.'}
        $app=Get-NativeAcceptanceOrdinaryPath $native.application_root -Directory
        $python=Get-NativeAcceptanceOrdinaryPath $native.python_path
        foreach($directory in @($root,$app,[IO.Path]::GetDirectoryName($AdapterPython),[IO.Path]::GetDirectoryName($PowerShell7),
            [IO.Path]::GetDirectoryName($PriorReport),[IO.Path]::GetDirectoryName($NativePlan),[IO.Path]::GetDirectoryName($ExtractionReceipt),[IO.Path]::GetDirectoryName($WorkRoot))){
            $path=Get-NativeAcceptanceOrdinaryPath $directory -Directory
            $ancestors=New-Object 'Collections.Generic.List[string]'
            while($path){$ancestors.Add($path);$path=[IO.Path]::GetDirectoryName($path)}
            for($i=$ancestors.Count-1;$i -ge 0;$i--){if($heldDirectories.Add($ancestors[$i])){$leases.Add((New-Object WinBookSplitLogging.PathLease($ancestors[$i],$true)))}}
        }
        $work=New-NativeAcceptanceDirectory $WorkRoot
        $leases.Add((New-Object WinBookSplitLogging.PathLease($work,$true)))
        $caseWork=New-NativeAcceptanceDirectory ([IO.Path]::Combine($work,'cases'))
        $leases.Add((New-Object WinBookSplitLogging.PathLease($caseWork,$true)))
        foreach($path in @($sdk,$AdapterPython,$PowerShell7,$PriorReport,$NativePlan,$ExtractionReceipt,[IO.Path]::Combine($root,'kit-bindings.json'))){
            $leases.Add((New-Object WinBookSplitLogging.PathLease($path,$false,$false,$true)))
        }
        foreach($name in @('tools/compatibility/native_smoke.py','tools/release/frozen_assets.py','engine/winbooksplit_windows.py')){
            $path=[IO.Path]::Combine($root,$name.Replace('/','\'))
            if((Get-NativeAcceptanceFileHash $path 8388608).sha256 -cne $bindings.helpers[$name].sha256){throw 'Supplemental reused helper differs.'}
            $leases.Add((New-Object WinBookSplitLogging.PathLease($path,$false,$false,$true)))
        }
        foreach($name in $bindings.production.assets.Keys){
            $path=[IO.Path]::Combine($root,'assets',$name)
            if((Get-NativeAcceptanceFileHash $path 33554432).sha256 -cne $bindings.production.assets[$name].sha256){throw 'Supplemental frozen asset differs.'}
            $leases.Add((New-Object WinBookSplitLogging.PathLease($path,$false,$false,$true)))
        }
        foreach($item in $native.application_files){$leases.Add((New-Object WinBookSplitLogging.PathLease($item.path,$false,$false,$true)))}
        foreach($name in @('simple-10.pdf','nested-12.pdf')){
            $path=[IO.Path]::Combine($root,'fixtures',$name)
            $leases.Add((New-Object WinBookSplitLogging.PathLease($path,$false,$false,$true)))
            $fixtureHash=Get-NativeAcceptanceFileHash $path 8388608
            if($fixtureHash.sha256 -cne $bindings.fixtures[$name].sha256 -or $fixtureHash.bytes -ne $bindings.fixtures[$name].bytes){throw 'Supplemental pinned fixture differs.'}
        }
        $fullProbe=@('-I','-B',[IO.Path]::Combine($app,'engine\winbooksplit_compatibility.py'),'--probe','--policy',[IO.Path]::Combine($app,'engine\WinBookSplit.Compatibility.json'))
        $run=Invoke-NativeAcceptanceProcessCase $python $fullProbe $work 120
        $runtimeBefore=[IO.Path]::Combine($work,'runtime-before.json');Write-NativeAcceptanceJson $runtimeBefore $run.probe
        Assert-NativeAcceptanceProcess $run.probe
        if($run.failure -or $run.probe.ExitCode -ne 0){throw 'Supplemental application-local full probe failed.'}
        $common=@('--kit-root',$root,'--bindings-sha256',$BindingsSha256,'--prior-report',$PriorReport,'--prior-report-sha256',$PriorReportSha256,
            '--native-plan',$NativePlan,'--extraction-receipt',$ExtractionReceipt,'--expected-pwsh-version',$ExpectedPowerShell7Version)
        $planPath=[IO.Path]::Combine($work,'ps7-plan.json')
        $argv=@('-I','-B',$AdapterPython,'prepare')+$common+@('--pwsh-path',$PowerShell7,'--work-root',$caseWork,'--runtime-probe',$runtimeBefore,'--output',$planPath)
        $run=Invoke-NativeAcceptanceProcessCase $python $argv $work 120
        Write-NativeAcceptanceJson ([IO.Path]::Combine($work,'prepare-process.json')) $run.probe
        Assert-NativeAcceptanceProcess $run.probe
        if($run.failure -or $run.probe.ExitCode -ne 0){throw 'Supplemental same-ZIP preparation failed.'}
        $read=Read-NativeAcceptanceJson $planPath 8388608;$plan=$read.value;$planHash=$read.sha256
        $leases.Add((New-Object WinBookSplitLogging.PathLease($planPath,$false,$false,$true)))
        $ids=@('version','setup-check','manual','auto2')
        if($plan.protocol -cne 'winbooksplit.final-zip-ps7-plan' -or $plan.cases.Count -ne 4 -or
            -not [StringComparer]::OrdinalIgnoreCase.Equals($plan.pwsh_path,$PowerShell7)){throw 'Supplemental finite plan differs.'}
        foreach($item in $plan.immutable_files){$leases.Add((New-Object WinBookSplitLogging.PathLease($item.path,$false,$false,$true)))}
        foreach($case in $plan.cases){
            $leases.Add((New-Object WinBookSplitLogging.PathLease($case.cwd,$true)))
            if($case.ContainsKey('output_base')){$leases.Add((New-Object WinBookSplitLogging.PathLease($case.output_base,$true)))}
        }
        $before=Get-NativeAcceptanceEnvironmentSnapshot ([IO.Path]::Combine($app,'.venv'))
        $factsCode=@'
$ErrorActionPreference='Stop'
Import-Module ([IO.Path]::Combine($PSHOME,'Modules','Microsoft.PowerShell.Utility','Microsoft.PowerShell.Utility.psd1')) -ErrorAction Stop
Import-Module ([IO.Path]::Combine($PSHOME,'Modules','CimCmdlets','CimCmdlets.psd1')) -ErrorAction Stop
$rows=@(Get-CimInstance Win32_OperatingSystem -ErrorAction Stop);$cpus=@(Get-CimInstance Win32_Processor -ErrorAction Stop)
if($rows.Count -ne 1 -or $cpus.Count -lt 1 -or $cpus.Count -gt 64){throw 'Native CIM evidence is unbounded.'};$os=$rows[0]
[ordered]@{Caption=[string]$os.Caption;Version=[string]$os.Version;BuildNumber=[string]$os.BuildNumber;ProductType=[int]$os.ProductType;
processor_architectures=@($cpus|ForEach-Object{[int]$_.Architecture});powershell_version=$PSVersionTable.PSVersion.ToString();
shell_executable=[Diagnostics.Process]::GetCurrentProcess().MainModule.FileName;process_bits=([IntPtr]::Size*8);
os_bits=$(if([Environment]::Is64BitOperatingSystem){64}else{32});clr=[Environment]::Version.ToString()}|ConvertTo-Json -Compress
'@
        $encoded=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($factsCode))
        $facts=Invoke-NativeAcceptanceProcessCase $PowerShell7 @('-NoProfile','-NonInteractive','-EncodedCommand',$encoded) $work 120
        Write-NativeAcceptanceJson ([IO.Path]::Combine($work,'facts-process.json')) $facts
        Assert-NativeAcceptanceProcess $facts.probe
        if($facts.failure -or $facts.probe.ExitCode -ne 0){throw 'Supplemental actual PS7/CIM facts failed.'}
        $rows=New-Object 'Collections.Generic.List[object]'
        for($i=0;$i -lt 4;$i++){
            $case=$plan.cases[$i]
            if($case.id -cne $ids[$i] -or $case.expected_exit -ne 0 -or $case.timeout_seconds -ne 120){throw 'Supplemental finite recipe differs.'}
            $run=Invoke-NativeAcceptanceProcessCase $PowerShell7 @($case.arguments) $case.cwd 120
            $rows.Add([ordered]@{id=$case.id;probe=$run.probe;command=$run.command;cwd=$run.cwd;answers=$run.answers})
            Write-NativeAcceptanceJson ([IO.Path]::Combine($work,$case.id+'-process.json')) $rows[$i]
            Assert-NativeAcceptanceProcess $run.probe
            if($run.failure -or -not $run.environment_before_after_restored -or $run.probe.ExitCode -ne 0){throw 'Supplemental native case failed; artifacts preserved.'}
        }
        $run=Invoke-NativeAcceptanceProcessCase $python $fullProbe $work 120
        Write-NativeAcceptanceJson ([IO.Path]::Combine($work,'runtime-after.json')) $run.probe
        Assert-NativeAcceptanceProcess $run.probe
        if($run.failure -or $run.probe.ExitCode -ne 0){throw 'Supplemental final full probe failed.'}
        $sameEnvironment=$before -ceq (Get-NativeAcceptanceEnvironmentSnapshot ([IO.Path]::Combine($app,'.venv')))
        $processPath=[IO.Path]::Combine($work,'ps7-process.json')
        Write-NativeAcceptanceJson $processPath ([ordered]@{protocol='winbooksplit.final-zip-ps7-process';schema_version=1;plan_sha256=$planHash;
            facts=$facts.probe;runtime_after=$run.probe;environment_unchanged=$sameEnvironment;cases=$rows.ToArray()})
        $leases.Add((New-Object WinBookSplitLogging.PathLease($processPath,$false,$false,$true)))
        $report=[IO.Path]::Combine($work,'ps7-result.json')
        $run=Invoke-NativeAcceptanceProcessCase $python (@('-I','-B',$AdapterPython,'validate')+$common+@('--plan',$planPath,'--process-report',$processPath,'--output',$report)) $work 120
        Write-NativeAcceptanceJson ([IO.Path]::Combine($work,'validate-process.json')) $run.probe
        Assert-NativeAcceptanceProcess $run.probe
        if($run.failure -or $run.probe.ExitCode -ne 0){throw 'Supplemental physical/console/source oracle failed.'}
        foreach($lease in $leases){$lease.AssertUnchanged()}
        return [pscustomobject]@{ReportPath=$report;ReportSha256=(Get-NativeAcceptanceFileHash $report 8388608).sha256;Success=$true}
    } finally {for($i=$leases.Count-1;$i -ge 0;$i--){$leases[$i].Dispose()}}
}
