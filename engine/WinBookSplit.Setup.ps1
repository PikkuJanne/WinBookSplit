# Explicit application-local setup only. Dot-sourcing performs no setup or discovery.
. (Join-Path $PSScriptRoot 'WinBookSplit.Diagnostics.ps1')
. (Join-Path $PSScriptRoot 'WinBookSplit.Runtime.ps1')
. (Join-Path $PSScriptRoot 'WinBookSplit.Logging.ps1')

function Initialize-WinBookSplitSetupTypes {
    if ('WinBookSplitSetup.DirectoryLease' -as [type]) { return }
    Add-Type -TypeDefinition @'
using System;
using System.IO;
using System.Text;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
namespace WinBookSplitSetup {
    // Hold an ordinary directory's identity and deny other rename/delete opens.
    // Failed environments are preserved by sibling rename, never recursively deleted.
    public sealed class DirectoryLease : IDisposable {
        [StructLayout(LayoutKind.Sequential)] struct Information {
            public uint Attributes, CreationLow, CreationHigh, AccessLow, AccessHigh, WriteLow, WriteHigh;
            public uint Volume, SizeHigh, SizeLow, Links, IndexHigh, IndexLow;
        }
        [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)]
        struct RenameInformation { public uint Replace; public IntPtr Root; public uint Length; public char Name; }
        [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        static extern SafeFileHandle CreateFileW(string path,uint access,uint sharing,IntPtr security,uint creation,uint flags,IntPtr template);
        [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        static extern bool CreateDirectoryW(string path,IntPtr security);
        [DllImport("kernel32.dll", SetLastError=true)]
        static extern bool GetFileInformationByHandle(SafeFileHandle handle,out Information information);
        [DllImport("kernel32.dll", SetLastError=true)]
        static extern bool SetFileInformationByHandle(SafeFileHandle handle,int kind,IntPtr information,uint size);
        SafeFileHandle handle;
        readonly Information identity;
        public string Path { get; private set; }
        public string Identity { get { return FormatIdentity(identity); } }
        public static void CreateExclusive(string path) {
            if (!CreateDirectoryW(path,IntPtr.Zero)) throw new Win32Exception(Marshal.GetLastWin32Error());
        }
        static SafeFileHandle Open(string path,uint access,uint sharing) {
            var result=CreateFileW(path,access,sharing,IntPtr.Zero,3,0x02200000,IntPtr.Zero);
            if (result.IsInvalid) { result.Dispose(); throw new Win32Exception(Marshal.GetLastWin32Error()); }
            return result;
        }
        static Information Read(SafeFileHandle value) {
            Information result;
            if (!GetFileInformationByHandle(value,out result)) throw new Win32Exception(Marshal.GetLastWin32Error());
            if ((result.Attributes&0x400)!=0 || (result.Attributes&0x10)==0) throw new IOException("Setup requires an ordinary directory.");
            return result;
        }
        static string FormatIdentity(Information value) {
            return value.Volume.ToString("x8")+":"+value.IndexHigh.ToString("x8")+value.IndexLow.ToString("x8");
        }
        public static string ReadIdentity(string path) {
            // A read-attributes snapshot does not reserve, rename, or mutate Check inputs.
            using (var value=Open(System.IO.Path.GetFullPath(path),0x80,7)) return FormatIdentity(Read(value));
        }
        public DirectoryLease(string path) {
            Path=System.IO.Path.GetFullPath(path);
            handle=Open(Path,0x80010000,3);
            try { identity=Read(handle); AssertUnchanged(); }
            catch { handle.Dispose(); throw; }
        }
        public void AssertUnchanged() {
            if (handle==null || handle.IsClosed) throw new IOException("Setup directory lease is closed.");
            var held=Read(handle);
            using (var named=Open(Path,0x80,7)) {
                var observed=Read(named);
                if (held.Volume!=identity.Volume || held.IndexHigh!=identity.IndexHigh || held.IndexLow!=identity.IndexLow ||
                    observed.Volume!=identity.Volume || observed.IndexHigh!=identity.IndexHigh || observed.IndexLow!=identity.IndexLow)
                    throw new IOException("Setup directory identity changed.");
            }
        }
        public void MoveTo(string target) {
            target=System.IO.Path.GetFullPath(target);
            if (!String.Equals(System.IO.Path.GetDirectoryName(Path),System.IO.Path.GetDirectoryName(target),StringComparison.OrdinalIgnoreCase) ||
                String.Equals(Path,target,StringComparison.OrdinalIgnoreCase)) throw new IOException("Setup preservation requires a new sibling path.");
            AssertUnchanged();
            byte[] name=Encoding.Unicode.GetBytes(target);
            int offset=Marshal.OffsetOf(typeof(RenameInformation),"Name").ToInt32();
            int size=checked(offset+name.Length+2);
            IntPtr buffer=Marshal.AllocHGlobal(size);
            try {
                Marshal.Copy(new byte[size],0,buffer,size);
                var info=new RenameInformation { Replace=0,Root=IntPtr.Zero,Length=(uint)name.Length,Name='\0' };
                Marshal.StructureToPtr(info,buffer,false); Marshal.Copy(name,0,IntPtr.Add(buffer,offset),name.Length);
                if (!SetFileInformationByHandle(handle,3,buffer,(uint)size)) throw new Win32Exception(Marshal.GetLastWin32Error());
                Path=target; AssertUnchanged();
            } finally { Marshal.FreeHGlobal(buffer); }
        }
        public void Dispose() { if (handle!=null) { handle.Dispose(); handle=null; } }
    }
}
'@ -ErrorAction Stop
}

function New-WinBookSplitSetupError {
    param([string]$Code, [string]$Message, $Probe = $null)
    $failure = New-Object InvalidOperationException($Message)
    $failure.Data['SetupCode']=$Code
    if ($null -ne $Probe) { $failure.Data['Probe']=$Probe }
    return $failure
}

function Test-WinBookSplitSetupEntryExists {
    param([string]$Path)
    try { $null=[IO.File]::GetAttributes($Path); return $true }
    catch {
        $failure=$_.Exception
        while ($failure.InnerException) { $failure=$failure.InnerException }
        if ($failure -is [IO.FileNotFoundException] -or $failure -is [IO.DirectoryNotFoundException]) { return $false }
        throw
    }
}

function Assert-WinBookSplitSetupNotCancelled {
    param([Threading.CancellationToken]$CancellationToken)
    if ($CancellationToken.IsCancellationRequested) { throw (New-WinBookSplitSetupError 'setup_cancelled' 'Setup was cancelled before further changes.') }
}

function Invoke-WinBookSplitSetupProcess {
    param([string]$Path, [string[]]$Arguments, [string]$ApplicationRoot, [string]$Phase, $Operations,
        [double]$TimeoutSeconds=300, [Threading.CancellationToken]$CancellationToken=[Threading.CancellationToken]::None,
        [switch]$IsolatePipConfiguration)
    $probe=Invoke-WinBookSplitProcess -Path $Path -Arguments $Arguments -WorkingDirectory $ApplicationRoot -TimeoutSeconds $TimeoutSeconds -CancellationToken $CancellationToken -IsolatePipConfiguration:$IsolatePipConfiguration
    $Operations.Add([pscustomobject]@{phase=$Phase;path=$Path;arguments=$Arguments;probe=$probe})
    if (-not $probe.ParentStopped -or -not $probe.DescendantsStopped -or $probe.StopError) {
        $failure=New-WinBookSplitSetupError 'setup_failed' ('Setup process shutdown could not be proved complete: '+$probe.StopError) $probe
        $failure.Data['UnsafeCleanup']=$true; throw $failure
    }
    if ($probe.Cancelled) { throw (New-WinBookSplitSetupError 'setup_cancelled' 'Setup was cancelled; its owned process tree was proved stopped.' $probe) }
    if ($probe.TimedOut) { throw (New-WinBookSplitSetupError 'setup_timeout' 'Setup timed out; its owned process tree was proved stopped.' $probe) }
    if ($probe.ExitCode -ne 0 -or -not $probe.JobAssigned -or -not $probe.StreamsComplete -or $probe.StartError -or $probe.StreamError -or $probe.StdoutTruncated -or $probe.StderrTruncated) {
        throw (New-WinBookSplitSetupError 'setup_failed' ('Setup '+$Phase+' failed. '+$probe.Stderr+' '+$probe.StartError) $probe)
    }
    return $probe
}

function Get-WinBookSplitSetupPackageSource {
    param([string]$ApplicationRoot, $Policy, [string]$Wheelhouse, $FileLeases)
    try {
    $requirements=Get-WinBookSplitTrustedPath ([IO.Path]::Combine($ApplicationRoot,'requirements.txt'))
    $FileLeases.Add([WinBookSplitLogging.PathLease]::new($requirements,$false,$false,$true))
    $stream=[IO.File]::OpenRead($requirements)
    try {
        if ($stream.Length -lt 1 -or $stream.Length -gt 4096) { throw (New-WinBookSplitSetupError 'setup_source_invalid' 'The runtime requirements lock is missing or oversized.') }
        $bytes=New-Object byte[] ([int]$stream.Length);$offset=0
        while ($offset -lt $bytes.Length) {
            $count=$stream.Read($bytes,$offset,$bytes.Length-$offset)
            if ($count -eq 0) { throw 'Incomplete runtime requirements lock read.' }
            $offset+=$count
        }
        if ($stream.ReadByte() -ne -1) { throw 'Runtime requirements lock changed during its bounded read.' }
    } finally { $stream.Dispose() }
    $text=(New-Object Text.UTF8Encoding($false,$true)).GetString($bytes)
    $lines=@($text -split '\r?\n' | Where-Object {-not [string]::IsNullOrWhiteSpace($_) -and -not $_.TrimStart().StartsWith('#')})
    $expected='pypdf=='+$Policy.pypdf.version+' --hash=sha256:'+$Policy.pypdf.wheel_sha256
    if ($lines.Count -ne 1 -or $lines[0].Trim() -cne $expected) { throw (New-WinBookSplitSetupError 'setup_source_invalid' 'The runtime requirements lock does not match the reviewed pypdf policy.') }
    $arguments=@('--index-url','https://pypi.org/simple')
    $description='reviewed PyPI wheel from https://pypi.org/simple'
    if (-not [string]::IsNullOrWhiteSpace($Wheelhouse)) {
        $directory=Get-WinBookSplitTrustedPath -Path $Wheelhouse -Directory
        $FileLeases.Add([WinBookSplitLogging.PathLease]::new($directory,$true))
        $name='pypdf-'+$Policy.pypdf.version+'-py3-none-any.whl'
        $wheel=Get-WinBookSplitTrustedPath ([IO.Path]::Combine($directory,$name))
        $FileLeases.Add([WinBookSplitLogging.PathLease]::new($wheel,$false,$false,$true))
        $stream=[IO.File]::OpenRead($wheel); $hashAlgorithm=[Security.Cryptography.SHA256]::Create()
        try {
            if ($stream.Length -lt 1 -or $stream.Length -gt 16777216) { throw 'The approved wheel must be a nonempty file no larger than 16 MiB.' }
            $hash=[BitConverter]::ToString($hashAlgorithm.ComputeHash($stream)).Replace('-','').ToLowerInvariant()
        } finally { $stream.Dispose(); $hashAlgorithm.Dispose() }
        if ($hash -cne $Policy.pypdf.wheel_sha256) { throw (New-WinBookSplitSetupError 'setup_source_invalid' 'The offline wheel SHA256 does not match the reviewed policy.') }
        # A directory find-links source can follow HTML links despite --no-index.
        # Pass only the exact reviewed wheel file, held against changes throughout pip.
        $arguments=@('--no-index','--find-links',$wheel)
        $description='verified offline wheel '+$wheel
    }
    return [pscustomobject]@{Requirements=$requirements;Arguments=$arguments;Description=$description}
    } catch {
        if ($_.Exception.Data.Contains('SetupCode')) { throw }
        throw (New-WinBookSplitSetupError 'setup_source_invalid' ('Cannot validate the reviewed package source: '+$_.Exception.Message))
    }
}

function Invoke-WinBookSplitSetup {
    param([Parameter(Mandatory=$true)][string]$ApplicationRoot, [string]$PythonPath,
        [switch]$Check, [switch]$NonInteractive, [string]$Wheelhouse, [switch]$Repair,
        [ValidateRange(0.05,3600)][double]$TimeoutSeconds=300,
        [Threading.CancellationToken]$CancellationToken=[Threading.CancellationToken]::None)
    $operations=New-Object 'Collections.Generic.List[object]'
    $attempts=New-Object 'Collections.Generic.List[object]'
    $fileLeases=New-Object 'Collections.Generic.List[object]'
    $parents=$null; $lock=$null; $marker=$null; $newEnvironment=$null; $oldEnvironment=$null
    $root=$null; $venv=$null; $backup=$null; $incomplete=$null
    $result=[pscustomobject][ordered]@{protocol='winbooksplit.setup';schema_version=1;status='failed';code='setup_failed';exit_code=6;
        message='Setup did not complete.';check=[bool]$Check;changed=$false;environment_path=$null;backup_path=$null;incomplete_path=$null;
        rollback_status='not_needed';selected_python=$null;selected_version=$null;pypdf_version=$null;operations=@();attempts=@();
        observed_python=$null;observed_version=$null;observed_source=$null;action=$null;action_shell=$null}
    try {
        Assert-WinBookSplitSetupNotCancelled $CancellationToken
        $root=Get-WinBookSplitTrustedPath -Path $ApplicationRoot -Directory
        $venv=[IO.Path]::Combine($root,'.venv'); $result.environment_path=$venv
        $policyPath=Get-WinBookSplitTrustedPath ([IO.Path]::Combine($root,'engine','WinBookSplit.Compatibility.json'))
        $policyIdentity=Get-WinBookSplitCompatibilityPolicy -PolicyPath $policyPath -WithIdentity
        $policy=$policyIdentity.Policy
        $guidance='Use '+(Get-WinBookSplitPythonPolicyDescription $policy)+'. PowerShell setup: '+(Get-WinBookSplitSetupCommand -ApplicationRoot $root)
        $lockPath=[IO.Path]::Combine($root,'.WinBookSplit-setup.lock')
        try { $busy=Test-WinBookSplitSetupEntryExists $lockPath }
        catch { throw (New-WinBookSplitSetupError 'setup_busy' 'The setup lock cannot be inspected safely; preserve it and retry after the other setup exits.') }
        if ($busy) { throw (New-WinBookSplitSetupError 'setup_busy' 'Another setup or a preserved setup lock exists. Wait for it to finish; no environment was changed.') }
        $exists=Test-WinBookSplitSetupEntryExists $venv
        $existingIdentity=$null
        if ($exists) {
            $null=Get-WinBookSplitTrustedPath -Path $venv -Directory
            Initialize-WinBookSplitSetupTypes
            $existingIdentity=[WinBookSplitSetup.DirectoryLease]::ReadIdentity($venv)
        }
        $preservedPrefixes=@([IO.Path]::Combine($root,'.WinBookSplit-venv-backup-'),[IO.Path]::Combine($root,'.WinBookSplit-incomplete-'))
        $probeTimeout=[Math]::Min($TimeoutSeconds,10.0)
        if ($exists) {
            $healthy=Test-WinBookSplitPythonCandidate ([IO.Path]::Combine($venv,'Scripts','python.exe')) 'application_venv' $root $attempts $probeTimeout
            if ($healthy) {
                $result.status='success';$result.code='setup_ready';$result.exit_code=0;$result.message='Application .venv is healthy; no setup changes or package downloads were needed.'
                $result.selected_python=$healthy.Path;$result.selected_version=$healthy.Version;$result.pypdf_version=$healthy.PypdfVersion
                return $result
            }
            if ($Check -or -not $Repair) {
                $rejected=$attempts[$attempts.Count-1]
                $result.observed_python=$rejected.Path;$result.observed_source='application_venv'
                if ($rejected.IdentityVerified) {
                    $result.observed_version=$rejected.ObservedVersion
                }
                $hintBase=Find-WinBookSplitSetupBase -ApplicationRoot $root -Policy $policy -PythonPath $PythonPath
                $hintPath=$null;if($hintBase){$hintPath=$hintBase.Path}
                $result.action=Get-WinBookSplitSetupCommand -ApplicationRoot $root -Repair -PythonPath $hintPath -Wheelhouse $Wheelhouse
                $result.action_shell='PowerShell'
                throw (New-WinBookSplitSetupError 'setup_repair_required' ('Existing application .venv was preserved. '+$rejected.Message+' PowerShell repair: '+$result.action))
            }
        }
        if (-not [string]::IsNullOrWhiteSpace($PythonPath)) {
            try { $explicit=Get-WinBookSplitTrustedPath $PythonPath }
            catch { throw (New-WinBookSplitSetupError 'runtime_invalid' ('The explicit base interpreter path is not trusted: '+$_.Exception.Message+' '+$guidance)) }
            if ($explicit.StartsWith($venv+'\',[StringComparison]::OrdinalIgnoreCase)) { throw (New-WinBookSplitSetupError 'runtime_invalid' 'Select a base interpreter outside the application .venv; repair must not move its selected interpreter.') }
            foreach ($prefix in $preservedPrefixes) {
                if ($explicit.StartsWith($prefix,[StringComparison]::OrdinalIgnoreCase)) { throw (New-WinBookSplitSetupError 'runtime_invalid' 'Preserved setup artifacts are not working environments. Select a trusted base interpreter outside the application preservation directories.') }
            }
        }
        if (-not $exists -and $Check) {
            try {
                $ready=Resolve-WinBookSplitRuntime -ApplicationRoot $root -PythonPath $PythonPath -TimeoutSeconds $probeTimeout
                $result.status='success';$result.code='setup_ready';$result.exit_code=0;$result.message='Processing dependencies are ready; Check made no setup changes.'
                $result.selected_python=$ready.Path;$result.selected_version=$ready.Version;$result.pypdf_version=$ready.PypdfVersion
                return $result
            } catch { if (Test-WinBookSplitInterruption $_.Exception) { throw } }
        }
        if (-not [string]::IsNullOrWhiteSpace($PythonPath)) {
            $base=Test-WinBookSplitPythonCandidate $explicit 'explicit' $root $attempts $probeTimeout -InterpreterOnly
            if (-not $base) { throw (New-WinBookSplitSetupError 'runtime_invalid' ('The selected base interpreter failed preflight. '+$attempts[$attempts.Count-1].Message+' '+$guidance)) }
        } else {
            $base=Resolve-WinBookSplitAutomaticPython -ApplicationRoot $root -Attempts $attempts -Policy $policy -TimeoutSeconds $probeTimeout -InterpreterOnly -ExcludedDirectories @($venv) -ExcludedDirectoryPrefixes $preservedPrefixes
            if (-not $base) { throw (New-WinBookSplitSetupError 'runtime_not_found' ('No compatible trusted base Python was found. '+$guidance)) }
        }
        $result.selected_python=$base.Path;$result.selected_version=$base.Version
        if ($Check) {
            $result.action=Get-WinBookSplitSetupCommand -ApplicationRoot $root -PythonPath $base.Path -Wheelhouse $Wheelhouse
            $result.action_shell='PowerShell'
            throw (New-WinBookSplitSetupError 'setup_required' ('Compatible Python is available; create the application environment and install locked pypdf. PowerShell setup: '+$result.action))
        }
        Initialize-WinBookSplitLogTypes
        $fileLeases.Add([WinBookSplitLogging.PathLease]::new($policyPath,$false,$false,$true))
        try { $heldPolicyIdentity=Get-WinBookSplitCompatibilityPolicy -PolicyPath $policyPath -WithIdentity }
        catch { throw (New-WinBookSplitSetupError 'setup_source_invalid' ('Compatibility policy changed or cannot be validated under its held lease: '+$_.Exception.Message)) }
        if ($heldPolicyIdentity.Sha256 -cne $policyIdentity.Sha256) { throw (New-WinBookSplitSetupError 'setup_source_invalid' 'Compatibility policy changed during setup preparation. No environment was changed; retry with the reviewed application files.') }
        $source=Get-WinBookSplitSetupPackageSource -ApplicationRoot $root -Policy $policy -Wheelhouse $Wheelhouse -FileLeases $fileLeases
        $id=[guid]::NewGuid().ToString('N')
        $backup=[IO.Path]::Combine($root,('.WinBookSplit-venv-backup-'+$id))
        $incomplete=[IO.Path]::Combine($root,('.WinBookSplit-incomplete-'+$id))
        if ($backup.Length -gt 247 -or [IO.Path]::Combine($venv,'.WinBookSplit-setup-owner.json').Length -gt 259) { throw (New-WinBookSplitSetupError 'setup_failed' 'The application path is too long for safe setup records. Extract to a shorter per-user directory.') }
        Assert-WinBookSplitSetupNotCancelled $CancellationToken
        if (-not $NonInteractive) {
            Write-Host ('Python '+$base.Version+': '+$base.Path)
            Write-Host ('Environment: '+$venv+'. Package source: '+$source.Description+'. Existing repair backups will be preserved.')
            $answer=Read-Host 'Proceed with local setup? [y/N]'
            if ($answer.Trim() -notin @('y','yes')) { throw (New-WinBookSplitSetupError 'setup_cancelled' 'Setup was cancelled before changes.') }
        }
        Assert-WinBookSplitSetupNotCancelled $CancellationToken
        $parents=New-WinBookSplitRecordLeases -Directory $root -IncludeMarker $false
        try { $lock=New-Object IO.FileStream($lockPath,[IO.FileMode]::CreateNew,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None,4096,[IO.FileOptions]::DeleteOnClose) }
        catch { throw (New-WinBookSplitSetupError 'setup_busy' 'Could not reserve the exclusive setup lock. Another setup may be active, or the app root is unwritable. Extract to a writable per-user directory.') }
        $lockBytes=[Text.Encoding]::UTF8.GetBytes('{"protocol":"winbooksplit.setup-lock","schema_version":1,"run_id":"'+$id+'"}')
        $lock.Write($lockBytes,0,$lockBytes.Length);$lock.Flush($true)
        foreach ($lease in $parents) { $lease.AssertUnchanged() }
        foreach ($lease in $fileLeases) { $lease.AssertUnchanged() }
        Initialize-WinBookSplitSetupTypes
        # Recheck existence under the exclusive lock; never adopt a concurrent entry.
        if ((Test-WinBookSplitSetupEntryExists $venv) -ne $exists) { throw (New-WinBookSplitSetupError 'setup_busy' 'Application .venv changed while setup was being prepared. Preserve it and retry.') }
        if ($exists) {
            $null=Get-WinBookSplitTrustedPath -Path $venv -Directory
            $oldEnvironment=[WinBookSplitSetup.DirectoryLease]::new($venv)
            if ($oldEnvironment.Identity -cne $existingIdentity) {
                $oldEnvironment.Dispose();$oldEnvironment=$null
                throw (New-WinBookSplitSetupError 'setup_busy' 'Application .venv identity changed while setup was being prepared. The current environment was preserved; retry after reviewing it.')
            }
            $oldEnvironment.MoveTo($backup);$result.backup_path=$backup;$result.changed=$true;$result.rollback_status='pending'
        }
        Assert-WinBookSplitSetupNotCancelled $CancellationToken
        [WinBookSplitSetup.DirectoryLease]::CreateExclusive($venv)
        $result.changed=$true
        $newEnvironment=[WinBookSplitSetup.DirectoryLease]::new($venv)
        $entries=[IO.Directory]::EnumerateFileSystemEntries($venv).GetEnumerator()
        try { if ($entries.MoveNext()) { throw (New-WinBookSplitSetupError 'setup_failed' 'New environment was not empty; its contents will be preserved.') } }
        finally { $entries.Dispose() }
        $markerPath=[IO.Path]::Combine($venv,'.WinBookSplit-setup-owner.json')
        $marker=New-Object IO.FileStream($markerPath,[IO.FileMode]::CreateNew,[IO.FileAccess]::ReadWrite,([IO.FileShare]::Read -bor [IO.FileShare]::Delete),4096,[IO.FileOptions]::DeleteOnClose)
        $markerBytes=[Text.Encoding]::UTF8.GetBytes('{"protocol":"winbooksplit.setup-owner","schema_version":1,"run_id":"'+$id+'"}')
        $marker.Write($markerBytes,0,$markerBytes.Length);$marker.Flush($true)
        $newEnvironment.AssertUnchanged()
        $null=Invoke-WinBookSplitSetupProcess -Path $base.Path -Arguments @('-I','-B','-X','utf8','-m','venv','--copies',$venv) -ApplicationRoot $root -Phase 'venv_create' -Operations $operations -TimeoutSeconds $TimeoutSeconds -CancellationToken $CancellationToken
        $newEnvironment.AssertUnchanged()
        $python=[IO.Path]::Combine($venv,'Scripts','python.exe')
        $interpreter=Test-WinBookSplitPythonCandidate $python 'application_venv' $root $attempts $probeTimeout -InterpreterOnly
        if (-not $interpreter) { throw (New-WinBookSplitSetupError 'setup_failed' ('New environment interpreter failed preflight: '+$attempts[$attempts.Count-1].Message)) }
        foreach ($lease in $fileLeases) { $lease.AssertUnchanged() }
        $arguments=@('-I','-B','-X','utf8','-m','pip','--isolated','--disable-pip-version-check','--no-input','install','--require-hashes','--only-binary=:all:','--no-deps','--no-cache-dir','--retries','2','--timeout','15')+$source.Arguments+@('-r',$source.Requirements)
        $null=Invoke-WinBookSplitSetupProcess -Path $python -Arguments $arguments -ApplicationRoot $root -Phase 'pip_install' -Operations $operations -TimeoutSeconds $TimeoutSeconds -CancellationToken $CancellationToken -IsolatePipConfiguration
        $newEnvironment.AssertUnchanged()
        $null=Invoke-WinBookSplitSetupProcess -Path $python -Arguments @('-I','-B','-X','utf8','-m','pip','--isolated','--disable-pip-version-check','--no-input','check') -ApplicationRoot $root -Phase 'pip_check' -Operations $operations -TimeoutSeconds $TimeoutSeconds -CancellationToken $CancellationToken -IsolatePipConfiguration
        Assert-WinBookSplitSetupNotCancelled $CancellationToken
        $healthy=Test-WinBookSplitPythonCandidate $python 'application_venv' $root $attempts $probeTimeout
        if (-not $healthy) { throw (New-WinBookSplitSetupError 'setup_failed' ('Installed environment failed processing preflight: '+$attempts[$attempts.Count-1].Message)) }
        $newEnvironment.AssertUnchanged();foreach ($lease in $parents) { $lease.AssertUnchanged() }
        foreach ($lease in $fileLeases) { $lease.AssertUnchanged() }
        $result.status='success';$result.code='setup_complete';$result.exit_code=0;$result.message='Application environment is ready. Drag a PDF onto WinBookSplit.bat to begin.'
        $result.selected_python=$healthy.Path;$result.selected_version=$healthy.Version;$result.pypdf_version=$healthy.PypdfVersion
        if ($oldEnvironment) { $result.rollback_status='backup_preserved';$result.message+=' Previous environment backup: '+$backup }
    } catch {
        $failure=$_.Exception
        $code='setup_failed'
        if ($failure.Data.Contains('SetupCode')) { $code=[string]$failure.Data['SetupCode'] }
        elseif ($failure.Data.Contains('Code')) { $code=[string]$failure.Data['Code'] }
        $result.code=$code;$result.message=$failure.Message
        if ($code -in @('setup_cancelled','processor_cancelled')) { $result.status='cancelled';$result.exit_code=130 }
        elseif ($code -in @('setup_timeout','processor_timeout')) { $result.status='timeout';$result.exit_code=130 }
        elseif ($code -in @('runtime_invalid','runtime_not_found','setup_required','setup_repair_required','setup_source_invalid')) { $result.exit_code=3 }
        $unsafe=[bool]$failure.Data['UnsafeCleanup'] -or [bool]$failure.Data['FatalDiscovery']
        $probe=$failure.Data['Probe']
        if ($null -ne $probe -and (-not $probe.ParentStopped -or -not $probe.DescendantsStopped -or $probe.StopError)) { $unsafe=$true }
        if ($newEnvironment -or $oldEnvironment) {
            try {
                if ($unsafe) { throw 'Owned process shutdown or path safety is uncertain; no environment was moved for rollback.' }
                foreach ($lease in $parents) { $lease.AssertUnchanged() }
                if ($marker) { $marker.Dispose();$marker=$null }
                if ($newEnvironment) {
                    $newEnvironment.MoveTo($incomplete);$result.incomplete_path=$incomplete
                    $result.message+=' Incomplete environment preserved at '+$incomplete+'. It is a retained artifact, not a working environment.'
                }
                if ($oldEnvironment) { $oldEnvironment.MoveTo($venv);$result.backup_path=$null;$result.rollback_status='restored';$result.message+=' Previous environment restored to its original .venv path.' }
            } catch {
                $result.status='failed';$result.code='setup_rollback_failed';$result.exit_code=6;$result.rollback_status='preserved_unrestored'
                $result.message+=' Rollback was refused or failed; all remaining environment contents were preserved. '+$_.Exception.Message
            }
        }
    } finally {
        if ($marker) { $marker.Dispose() }
        if ($newEnvironment) { $newEnvironment.Dispose() }
        if ($oldEnvironment) { $oldEnvironment.Dispose() }
        if ($lock) { $lock.Dispose() }
        if ($parents) { foreach ($lease in $parents) { $lease.Dispose() } }
        foreach ($lease in $fileLeases) { $lease.Dispose() }
        $result.operations=$operations.ToArray();$result.attempts=$attempts.ToArray()
    }
    return $result
}
