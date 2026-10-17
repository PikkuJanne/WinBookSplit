[CmdletBinding()]
param([string]$RepositoryRoot,[string]$WorkRoot,[string]$ReportPath)
$ErrorActionPreference='Stop'
foreach($module in @('Microsoft.PowerShell.Utility','Microsoft.PowerShell.Management','Microsoft.PowerShell.Security')){
    Import-Module -Name ([IO.Path]::Combine($PSHOME,'Modules',$module,$module+'.psd1')) -Scope Local -ErrorAction Stop
}
. ([IO.Path]::Combine($RepositoryRoot,'engine','WinBookSplit.Logging.ps1'))
. ([IO.Path]::Combine($RepositoryRoot,'engine','WinBookSplit.Support.ps1'))
. ([IO.Path]::Combine($RepositoryRoot,'engine','WinBookSplit.Runtime.ps1'))
. ([IO.Path]::Combine($RepositoryRoot,'tools','compatibility','Invoke-NativeAcceptance.ps1'))
Initialize-WinBookSplitLogTypes;Initialize-WinBookSplitSupport
$environmentBefore=[Environment]::GetEnvironmentVariables('Process')
$attempts=New-Object 'Collections.Generic.List[object]'
$registrations=@(Get-WinBookSplitRegisteredManagerPaths -ApplicationRoot $RepositoryRoot -Attempts $attempts -TimeoutSeconds 5 -WithRegistration)
if($registrations.Count-lt 1 -or $registrations.Count-gt 2){throw 'A bounded verified official manager reference is required for the observed welcome marker.'}
$probes=New-Object 'Collections.Generic.List[object]'
$index=0
foreach($registration in $registrations){
    $index++;$cache=New-NativeAcceptanceDirectory ([IO.Path]::Combine($WorkRoot,'cache-'+$index))
    $null=[IO.Directory]::CreateDirectory([IO.Path]::Combine($cache,'Microsoft','Windows','Caches'))
    $manager=[IO.Path]::Combine($registration.Package.InstallLocation,'pymanager.exe')
    $before=(Get-NativeAcceptanceFileHash $manager 8388608).sha256
    $prior=[Environment]::GetEnvironmentVariable('LOCALAPPDATA','Process')
    $listing=$null
    try{
        [Environment]::SetEnvironmentVariable('LOCALAPPDATA',$cache,'Process')
        $listing=Invoke-WinBookSplitManagerListing -Registration $registration -ApplicationRoot $RepositoryRoot -TimeoutSeconds 5
    }finally{[Environment]::SetEnvironmentVariable('LOCALAPPDATA',$prior,'Process')}
    Assert-NativeAcceptanceProcess $listing
    if($listing.ExitCode-ne 0 -or -not $listing.PackageIdentityVerified){throw 'Verified installed-only manager reference listing failed.'}
    $after=(Get-NativeAcceptanceFileHash $manager 8388608).sha256
    if($before-cne $after){throw 'Manager reference image changed.'}
    $probes.Add(@{registration=$registration;listing_command=@('list','--format=exe');listing=$listing;
        manager_sha256_before=$before;manager_sha256_after=$after;cache_root=$cache;marker_absent_before=$true;
        child_localappdata_only=$true;parent_localappdata_restored=([Environment]::GetEnvironmentVariable('LOCALAPPDATA','Process')-ceq $prior)})
}
$environmentAfter=[Environment]::GetEnvironmentVariables('Process')
$unchanged=$environmentBefore.Count-eq $environmentAfter.Count
foreach($key in $environmentBefore.Keys){if($environmentBefore[$key]-cne $environmentAfter[$key]){$unchanged=$false}}
if(-not $unchanged){throw 'Manager reference controller environment changed.'}
Write-NativeAcceptanceJson $ReportPath @{kind='VERIFIED_INSTALLED_ONLY_MANAGER_CACHE_REFERENCE';probes=$probes.ToArray();query_attempts=$attempts.ToArray();
    parent_environment_unchanged=$true;runtime_acquisition_attempted=$false;supplied_to_application=$false;
    note='Separate owned developer side-effect observation only. No runtime is selected, installed, copied or supplied to application/setup.'}
