param([Parameter(Mandatory=$true)][string]$ApplicationRoot,[Parameter(Mandatory=$true)][string]$OutputPath)
$ErrorActionPreference='Stop'
Set-Location -LiteralPath $ApplicationRoot
. (Join-Path $ApplicationRoot 'engine/WinBookSplit.Runtime.ps1')
function HashText([string]$Text){$sha=[Security.Cryptography.SHA256]::Create();try{return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Text)))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}}
function SourceHashes { [ordered]@{Runtime=(Get-FileHash -LiteralPath (Join-Path $ApplicationRoot 'engine/WinBookSplit.Runtime.ps1') -Algorithm SHA256).Hash.ToLowerInvariant();Process=(Get-FileHash -LiteralPath (Join-Path $ApplicationRoot 'engine/WinBookSplit.Process.ps1') -Algorithm SHA256).Hash.ToLowerInvariant();Policy=(Get-FileHash -LiteralPath (Join-Path $ApplicationRoot 'engine/WinBookSplit.Compatibility.json') -Algorithm SHA256).Hash.ToLowerInvariant();Helper=(Get-FileHash -LiteralPath (Join-Path $ApplicationRoot 'engine/winbooksplit_compatibility.py') -Algorithm SHA256).Hash.ToLowerInvariant()} }
$sourceBefore=SourceHashes
$inventoryAttempts=New-Object 'Collections.Generic.List[object]'
$registrations=@(Get-WinBookSplitRegisteredManagerPaths -ApplicationRoot $ApplicationRoot -Attempts $inventoryAttempts -TimeoutSeconds 5 -WithRegistration)
if($registrations.Count -ne 1){throw 'This host currently expects one existing registered python.org manager for verification-only inventory.'}
$registration=$registrations[0]
$listing=Invoke-WinBookSplitManagerListing -Registration $registration -ApplicationRoot $ApplicationRoot -TimeoutSeconds 5
$listedPaths=@(ConvertFrom-WinBookSplitPythonListing -Probe $listing -Kind Manager)
function Snapshot {
    $runtimeHashes=@(foreach($verificationPath in $listedPaths) {[ordered]@{Path=$verificationPath;Sha256=(Get-FileHash -LiteralPath $verificationPath -Algorithm SHA256).Hash.ToLowerInvariant();Canonical=[WinBookSplit.Preflight.Probe]::Canonical($verificationPath)}})
    return [ordered]@{PathEnvironmentSha256=(HashText $env:PATH);ManagerPackage=$registration.Package.PackageFullName;ManagerExeSha256=(Get-FileHash -LiteralPath (Join-Path $registration.Package.InstallLocation 'pymanager.exe') -Algorithm SHA256).Hash.ToLowerInvariant();RuntimeHashesVerificationOnly=$runtimeHashes;AliasAttributes=[IO.File]::GetAttributes($registration.Path).ToString();AppVenvExists=[IO.Directory]::Exists((Join-Path $ApplicationRoot '.venv'))}
}
function InvokeAutomatic([bool]$InterpreterOnly) {
    $result=$null;$failure=$null
    try {if($InterpreterOnly){$result=Resolve-WinBookSplitRuntime -ApplicationRoot $ApplicationRoot -InterpreterOnly -TimeoutSeconds 3}else{$result=Resolve-WinBookSplitRuntime -ApplicationRoot $ApplicationRoot -TimeoutSeconds 3}}
    catch{$failure=$_.Exception}
    $record=[ordered]@{InterpreterOnly=$InterpreterOnly;Success=($null -ne $result);Result=$result;Code=$null;Message=$null;Attempts=$null}
    if($failure){$record.Code=$failure.Data['Code'];$record.Message=$failure.Message;$record.Attempts=$failure.Data['Attempts']}
    return $record
}
$stateBefore=Snapshot
$interpreter=InvokeAutomatic $true
$processing=InvokeAutomatic $false
$stateAfter=Snapshot
$sourceAfter=SourceHashes
$checks=[ordered]@{RealRegistration=($registration.Package.PackageFamilyName -ceq 'PythonSoftwareFoundation.PythonManager_3847v3x7pw1km');VerificationListingBound=($listing.PackageIdentityVerified -and $listing.ExitCode -eq 0 -and $listing.StreamsComplete);NoAppVenv=(-not $stateBefore.AppVenvExists -and -not $stateAfter.AppVenvExists);InstallationAndPathUnchanged=(($stateBefore|ConvertTo-Json -Depth 8 -Compress) -ceq ($stateAfter|ConvertTo-Json -Depth 8 -Compress));SourceUnchanged=(($sourceBefore|ConvertTo-Json -Compress) -ceq ($sourceAfter|ConvertTo-Json -Compress))}
foreach($mode in @([pscustomobject]@{Name='InterpreterOnly';Record=$interpreter},[pscustomobject]@{Name='Processing';Record=$processing})) {
    $listingAttempts=@($mode.Record.Attempts|Where-Object{$_.Source -eq 'manager_listing' -and $_.Accepted -and $null -ne $_.Probe -and $_.Probe.PackageIdentityVerified})
    $runtimeAttempts=@($mode.Record.Attempts|Where-Object{$_.Source -eq 'manager' -and $_.Reason -eq 'python_security_floor' -and $_.IdentityVerified -and $_.ObservedVersion -eq '3.14.7'})
    $checks[$mode.Name+'RejectedUnsupportedRuntime']=(-not $mode.Record.Success -and $mode.Record.Code -eq 'runtime_not_found' -and $runtimeAttempts.Count -eq 1)
    $checks[$mode.Name+'BoundManagerListing']=($listingAttempts.Count -eq 1)
    $checks[$mode.Name+'ConcreteIdentity']=($runtimeAttempts.Count -eq 1 -and $runtimeAttempts[0].Path.Equals($listedPaths[0],[StringComparison]::OrdinalIgnoreCase) -and $runtimeAttempts[0].Probe.StreamsComplete -and $runtimeAttempts[0].Probe.ParentStopped -and $runtimeAttempts[0].Probe.DescendantsStopped)
}
$receipt=[ordered]@{schema_version=1;kind='ACTUAL_WINDOWS11_PYTHONORG_AUTOMATIC_T03';policy=[ordered]@{Process=(Get-ExecutionPolicy -Scope Process).ToString();Effective=(Get-ExecutionPolicy).ToString();Purpose='RemoteSigned automatic discovery component; not MOTW acceptance'};host=[ordered]@{PowerShell=$PSVersionTable.PSVersion.ToString();Edition=$PSVersionTable.PSEdition;OS=[Environment]::OSVersion.VersionString;Process64Bit=[Environment]::Is64BitProcess};source=[ordered]@{Head=(& git rev-parse HEAD);Tree=(& git rev-parse 'HEAD^{tree}');WorktreeSourceHashesBefore=$sourceBefore;WorktreeSourceHashesAfter=$sourceAfter};commands=@('Get-WinBookSplitRegisteredManagerPaths -ApplicationRoot <WORKSPACE> -Attempts <LIST> -TimeoutSeconds 5 -WithRegistration [verification-only inventory]','Invoke-WinBookSplitManagerListing -Registration <REAL_REGISTRATION> -ApplicationRoot <WORKSPACE> -TimeoutSeconds 5 [verification-only inventory]','Resolve-WinBookSplitRuntime -ApplicationRoot <WORKSPACE> -InterpreterOnly -TimeoutSeconds 3','Resolve-WinBookSplitRuntime -ApplicationRoot <WORKSPACE> -TimeoutSeconds 3');verificationInventory=[ordered]@{Registration=$registration;Attempts=$inventoryAttempts;Listing=$listing};stateBefore=$stateBefore;interpreterOnly=$interpreter;processing=$processing;stateAfter=$stateAfter;checks=$checks;limitations=@('Actual automatic discovery reaches and independently verifies existing Python 3.14.7, then preserves reviewed security-floor rejection; no compatible runtime success claimed.','Windows11/python.org component only; native Store and owner Windows10 exact-ZIP gate remain NOT_RUN.','No explicit PythonPath, hidden selected-runtime input, injected environment, setup, installation/update, PATH/alias setting change, venv creation or fallback acquisition.','Independently enumerated runtime paths are hashed only for unchanged installation verification and are never passed into automatic application discovery.','Get-WinBookSplitDependencySummary UX is T04 and was not repaired or accepted by this test.')}
$json=$receipt|ConvertTo-Json -Depth 18
$json=$json.Replace($env:USERPROFILE.Replace('\','\\'),'<USERPROFILE>').Replace($env:USERPROFILE,'<USERPROFILE>').Replace($ApplicationRoot.Replace('\','\\'),'<WORKSPACE>').Replace($ApplicationRoot,'<WORKSPACE>')
[IO.File]::WriteAllText($OutputPath,$json+"`n",(New-Object Text.UTF8Encoding($false)))
[ordered]@{Output=$OutputPath;Checks=$checks;AllPassed=(-not ($checks.Values -contains $false))}|ConvertTo-Json -Depth 4 -Compress
if($checks.Values -contains $false){exit 1}