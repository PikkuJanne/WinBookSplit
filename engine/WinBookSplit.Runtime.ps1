# Dependency preflight only. Dot-sourcing performs no discovery or installation.
# Probe transport shares the owned engine job; discovery never owns file cleanup.
. (Join-Path $PSScriptRoot 'WinBookSplit.Process.ps1')
function Assert-WinBookSplitProbeNotInterrupted {
    param($Probe)
    if ($null -ne $Probe -and ($Probe.Cancelled -or $Probe.TimedOut)) {
        $code = 'processor_cancelled'
        $message = 'Dependency preflight was cancelled.'
        if ($Probe.TimedOut) { $code = 'processor_timeout'; $message = 'Dependency preflight timed out.' }
        if ($Probe.ParentStopped -and $Probe.DescendantsStopped -and -not $Probe.StopError) {
            $message += ' Its owned process tree was proved stopped.'
        }
        else { $message += ' Shutdown could not be proved complete: ' + $Probe.StopError }
        $failure = New-Object InvalidOperationException($message)
        $failure.Data['Code'] = $code
        $failure.Data['Probe'] = $Probe
        throw $failure
    }
}

function Test-WinBookSplitInterruption {
    param($Failure)
    return ($Failure.Data.Contains('Code') -and $Failure.Data['Code'] -in @('processor_cancelled', 'processor_timeout'))
}

function Initialize-WinBookSplitRuntimeProbe {
    if ('WinBookSplit.Preflight.Probe' -as [type]) { return }
    Add-Type -TypeDefinition @'
using System;
using System.Collections;
using System.Diagnostics;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
namespace WinBookSplit.Preflight {
    public sealed class Result {
        public int? ExitCode; public int? Pid; public bool TimedOut;
        public bool ParentStopped; public bool? DescendantsStopped = null;
        public bool StreamsComplete; public string StartError;
        public string Stdout; public string Stderr;
        public long StdoutTotalBytes; public long StderrTotalBytes;
        public bool StdoutTruncated; public bool StderrTruncated;
        public string StreamError; public double ElapsedSeconds;
    }
    sealed class Tail {
        internal readonly object Gate = new object();
        internal byte[] Bytes = new byte[65536]; internal int Count;
        internal long Total; internal bool Done; internal string Error;
        internal void Read(Stream stream) {
            try {
                byte[] block = new byte[8192]; int count;
                while ((count = stream.Read(block, 0, block.Length)) != 0) {
                    lock (Gate) {
                        Total += count;
                        int discard = Math.Max(0, Count + count - Bytes.Length);
                        if (discard != 0) { Buffer.BlockCopy(Bytes, discard, Bytes, 0, Count - discard); Count -= discard; }
                        Buffer.BlockCopy(block, 0, Bytes, Count, count); Count += count;
                    }
                }
            } catch (Exception error) { lock (Gate) { Error = error.Message; } }
            finally {
                try { stream.Dispose(); } catch (Exception error) { lock (Gate) { Error = error.Message; } }
                lock (Gate) { Done = true; }
            }
        }
        internal string Text() { lock (Gate) { return Encoding.UTF8.GetString(Bytes, 0, Count); } }
        internal bool Complete() { lock (Gate) { return Done; } }
    }
    sealed class State {
        internal Process Child; internal Tail Out = new Tail(); internal Tail Err = new Tail();
        readonly object Gate = new object(); bool Returning; int Completed; int Started;
        internal void Start(Thread thread) {
            lock (Gate) { Started++; }
            try { thread.Start(); } catch { lock (Gate) { Started--; } throw; }
        }
        internal void Finished() {
            lock (Gate) { Completed++; if (Returning && Completed == Started) { try { Child.Dispose(); } catch { } } }
        }
        internal void Return() { lock (Gate) { Returning = true; if (Completed == Started) Child.Dispose(); } }
    }
    public sealed class PathDetails { public string Path; public uint Links; }
    public static class Probe {
        [StructLayout(LayoutKind.Sequential)] struct FileInfo {
            public uint Attributes, CreationLow, CreationHigh, AccessLow, AccessHigh, WriteLow, WriteHigh;
            public uint Volume, SizeHigh, SizeLow, Links, IndexHigh, IndexLow;
        }
        [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        static extern IntPtr CreateFileW(string name, uint access, uint share, IntPtr security, uint creation, uint flags, IntPtr template);
        [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        static extern uint GetFinalPathNameByHandleW(IntPtr handle, StringBuilder path, uint size, uint flags);
        [DllImport("kernel32.dll", SetLastError=true)] static extern bool GetFileInformationByHandle(IntPtr handle, out FileInfo info);
        [DllImport("kernel32.dll", SetLastError=true)] static extern bool CloseHandle(IntPtr handle);
        public static PathDetails Canonical(string path) {
            IntPtr handle=CreateFileW(path,0x80,7,IntPtr.Zero,3,0x02200000,IntPtr.Zero);
            if (handle==new IntPtr(-1)) throw new Win32Exception(Marshal.GetLastWin32Error());
            try {
                StringBuilder text=new StringBuilder(32768); FileInfo info;
                uint length=GetFinalPathNameByHandleW(handle,text,(uint)text.Capacity,0);
                if (length==0 || length>=text.Capacity || !GetFileInformationByHandle(handle,out info)) throw new Win32Exception(Marshal.GetLastWin32Error());
                string final=text.ToString();
                if (final.StartsWith(@"\\?\UNC\",StringComparison.OrdinalIgnoreCase)) final=@"\\"+final.Substring(8);
                else if (final.StartsWith(@"\\?\",StringComparison.OrdinalIgnoreCase)) final=final.Substring(4);
                return new PathDetails { Path=final, Links=info.Links };
            } finally { if (!CloseHandle(handle)) throw new Win32Exception(Marshal.GetLastWin32Error()); }
        }
        public static Result Run(string path, string arguments, string cwd, int timeoutMilliseconds) {
            Result result = new Result(); Stopwatch clock = Stopwatch.StartNew();
            State state = new State(); state.Child = new Process(); bool started = false;
            try {
                ProcessStartInfo info = new ProcessStartInfo(path, arguments);
                info.WorkingDirectory = cwd; info.UseShellExecute = false; info.CreateNoWindow = true;
                info.RedirectStandardInput = true; info.RedirectStandardOutput = true; info.RedirectStandardError = true;
                foreach (string name in new ArrayList(info.EnvironmentVariables.Keys)) {
                    if (name.StartsWith("PYTHON", StringComparison.OrdinalIgnoreCase) ||
                        name.StartsWith("PYLAUNCHER", StringComparison.OrdinalIgnoreCase)) info.EnvironmentVariables.Remove(name);
                }
                info.EnvironmentVariables["PYTHONUTF8"] = "1";
                info.EnvironmentVariables["PYTHONIOENCODING"] = "utf-8:backslashreplace";
                info.EnvironmentVariables["PYTHON_MANAGER_AUTOMATIC_INSTALL"] = "false";
                info.EnvironmentVariables["PYLAUNCHER_NO_SEARCH_PATH"] = "1";
                state.Child.StartInfo = info; started = state.Child.Start();
                result.Pid = state.Child.Id; state.Child.StandardInput.Close();
                Thread output = new Thread(delegate() { state.Out.Read(state.Child.StandardOutput.BaseStream); state.Finished(); });
                Thread error = new Thread(delegate() { state.Err.Read(state.Child.StandardError.BaseStream); state.Finished(); });
                output.IsBackground = true; error.IsBackground = true; state.Start(output); state.Start(error);
                while (!(state.Child.HasExited && state.Out.Complete() && state.Err.Complete())) {
                    if (clock.ElapsedMilliseconds >= timeoutMilliseconds) { result.TimedOut = true; break; }
                    Thread.Sleep(5);
                }
                if (result.TimedOut && !state.Child.HasExited) {
                    try { state.Child.Kill(); } catch (Exception failure) { result.StartError = "Parent termination: " + failure.Message; }
                }
                // One bounded shutdown grace; inherited open pipes never turn
                // into an unbounded GetResult, stream Dispose or WaitForExit.
                if (!state.Child.HasExited) state.Child.WaitForExit(250);
                result.ParentStopped = state.Child.HasExited;
                if (result.ParentStopped) result.ExitCode = state.Child.ExitCode;
            } catch (Exception error) {
                result.StartError = error.Message;
                if (started) {
                    try { if (!state.Child.HasExited) state.Child.Kill(); state.Child.WaitForExit(250); result.ParentStopped = state.Child.HasExited; }
                    catch (Exception shutdown) { result.StartError += "; parent shutdown: " + shutdown.Message; }
                } else { result.ParentStopped = true; }
            } finally {
                result.StreamsComplete = state.Out.Complete() && state.Err.Complete();
                result.Stdout = state.Out.Text(); result.Stderr = state.Err.Text();
                lock (state.Out.Gate) { result.StdoutTotalBytes = state.Out.Total; result.StdoutTruncated = state.Out.Total > 65536; }
                lock (state.Err.Gate) { result.StderrTotalBytes = state.Err.Total; result.StderrTruncated = state.Err.Total > 65536; }
                result.StreamError = state.Out.Error ?? state.Err.Error;
                result.ElapsedSeconds = clock.Elapsed.TotalSeconds;
                try { if (started) state.Return(); else state.Child.Dispose(); }
                catch (Exception close) { result.StreamError = "Probe handle close: " + close.Message; }
            }
            return result;
        }
    }
}
'@ -ErrorAction Stop
}

function Invoke-WinBookSplitRuntimeProbe {
    param([string]$Path, [string[]]$Arguments, [string]$ApplicationRoot,
          [ValidateRange(0.05, 60)][double]$TimeoutSeconds = 10)
    return Invoke-WinBookSplitProcess -Path $Path -Arguments $Arguments -WorkingDirectory $ApplicationRoot -TimeoutSeconds $TimeoutSeconds
}

function Get-WinBookSplitTrustedPath {
    param([string]$Path, [switch]$Directory, [switch]$Automatic)
    if ([string]::IsNullOrWhiteSpace($Path) -or $Path -notmatch '^(?:[A-Za-z]:[\\/]|\\\\[^\\]+\\[^\\]+\\)') {
        throw 'Select an absolute literal FileSystem path.'
    }
    $full = [IO.Path]::GetFullPath($Path)
    $attributes = [IO.File]::GetAttributes($full)
    if ([bool]($attributes -band [IO.FileAttributes]::Directory) -ne [bool]$Directory) {
        throw 'The selected path has the wrong file/directory type.'
    }
    $current = $full
    while ($current) {
        if ([IO.File]::GetAttributes($current) -band [IO.FileAttributes]::ReparsePoint) {
            throw 'Dependency paths and their ancestors must not be reparse points.'
        }
        $current = [IO.Path]::GetDirectoryName($current)
    }
    Initialize-WinBookSplitRuntimeProbe
    $canonical = [WinBookSplit.Preflight.Probe]::Canonical($full)
    if (-not $canonical.Path.TrimEnd('\').Equals($full.TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Dependency aliases such as short 8.3 names are ambiguous; select the ordinary long path explicitly.'
    }
    if ($Automatic -and -not $Directory -and $canonical.Links -ne 1) {
        throw 'Automatic dependency discovery refuses hard-linked executable aliases; select a trusted ordinary executable explicitly.'
    }
    return $full
}

function Test-WinBookSplitAutomaticPath {
    param([string]$Path, [string]$ApplicationRoot, [string]$DocumentPath)
    $full = [IO.Path]::GetFullPath($Path)
    $cwd = [IO.Path]::GetFullPath((Get-Location).ProviderPath).TrimEnd('\')
    Initialize-WinBookSplitRuntimeProbe
    $cwd = ([WinBookSplit.Preflight.Probe]::Canonical($cwd)).Path.TrimEnd('\')
    if (-not $cwd.Equals($ApplicationRoot.TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase) -and
        ($full.StartsWith($cwd + '\', [StringComparison]::OrdinalIgnoreCase))) { return $false }
    if (-not [string]::IsNullOrWhiteSpace($DocumentPath)) {
        $bookDirectory = [IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($DocumentPath)).TrimEnd('\')
        if ([IO.Directory]::Exists($bookDirectory)) { $bookDirectory = ([WinBookSplit.Preflight.Probe]::Canonical($bookDirectory)).Path.TrimEnd('\') }
        if ($full.StartsWith($bookDirectory + '\', [StringComparison]::OrdinalIgnoreCase)) { return $false }
    }
    return $true
}

function Get-WinBookSplitPathCandidates {
    param([string]$Name, [string]$ApplicationRoot, [string]$DocumentPath, $Attempts)
    $seen = @{}; $directories = @($env:PATH -split ';')
    if ($directories.Count -gt 256) { throw 'PATH exceeds the bounded 256-directory discovery limit.' }
    foreach ($directory in $directories) {
        $directory = $directory.Trim().Trim('"')
        if ([string]::IsNullOrWhiteSpace($directory) -or $directory -notmatch '^(?:[A-Za-z]:[\\/]|\\\\[^\\]+\\[^\\]+\\)') { continue }
        try {
            $candidate = [IO.Path]::Combine([IO.Path]::GetFullPath($directory), $Name)
            if ($seen.ContainsKey($candidate)) { continue }; $seen[$candidate] = $true
            if (-not (Test-WinBookSplitAutomaticPath $candidate $ApplicationRoot $DocumentPath)) {
                $Attempts.Add([pscustomobject]@{ Path = $candidate; Source = 'PATH'; Accepted = $false; Message = 'Document-directory or unrelated-CWD candidate excluded.'; Probe = $null })
                continue
            }
            if ([IO.File]::Exists($candidate)) { $candidate }
        } catch {
            $Attempts.Add([pscustomobject]@{ Path = $directory; Source = 'PATH'; Accepted = $false; Message = $_.Exception.Message; Probe = $null })
        }
    }
}

function New-WinBookSplitDependencyError {
    param([string]$Code, [string]$Message, $Attempts)
    $error = New-Object InvalidOperationException($Message)
    $error.Data['Code'] = $Code; $error.Data['Attempts'] = @($Attempts.ToArray())
    return $error
}

function Test-WinBookSplitPythonCandidate {
    param([string]$Path, [string]$Source, [string]$ApplicationRoot, $Attempts, [double]$TimeoutSeconds)
    $probe = $null
    try {
        $path = Get-WinBookSplitTrustedPath -Path $Path -Automatic:($Source -in @('PATH', 'py_launcher'))
        $code = @'
import importlib.metadata, json, os, pathlib, platform, struct, sys, sysconfig
r = dict(protocol='winbooksplit.runtime', schema_version=1, ok=False, executable=sys.executable,
         version=platform.python_version(), implementation=sys.implementation.name, platform=sys.platform,
         machine=platform.machine(), bits=struct.calcsize('P')*8,
         gil_disabled=bool(sysconfig.get_config_var('Py_GIL_DISABLED')),
         isolated=bool(sys.flags.isolated), dont_write_bytecode=bool(sys.flags.dont_write_bytecode))
try:
    assert r['implementation']=='cpython' and sys.version_info[:3]==(3,14,8), 'Expected CPython 3.14.8 exactly.'
    assert r['platform']=='win32' and r['machine']=='AMD64' and r['bits']==64 and not r['gil_disabled'], 'Expected regular Windows x64 CPython.'
    assert r['isolated'] and r['dont_write_bytecode'], 'Isolated flags are missing.'
    import pypdf
    r.update(pypdf_version=pypdf.__version__, pypdf_path=str(pathlib.Path(pypdf.__file__).resolve()))
    assert r['pypdf_version']=='6.20.0' and importlib.metadata.version('pypdf')=='6.20.0', 'Expected pypdf 6.20.0 exactly.'
    roots=[pathlib.Path(sysconfig.get_paths()[key]).resolve() for key in ('purelib','platlib')]
    origin=pathlib.Path(r['pypdf_path'])
    assert origin.is_file() and any(origin.is_relative_to(root) for root in roots), 'Unexpected pypdf import origin.'
    package=origin.parent
    for name,module in tuple(sys.modules.items()):
        if name=='pypdf' or name.startswith('pypdf.'):
            file=getattr(module,'__file__',None)
            assert file and pathlib.Path(file).resolve().is_relative_to(package), 'Unexpected pypdf module origin.'
    r.update(ok=True, message='Supported interpreter and pypdf imported in isolation.')
except Exception as error:
    r['message']=str(error)
print(json.dumps(r, ensure_ascii=True))
raise SystemExit(0 if r['ok'] else 1)
'@
        $probe = Invoke-WinBookSplitRuntimeProbe -Path $path -Arguments @('-I', '-B', '-X', 'utf8', '-c', $code) -ApplicationRoot $ApplicationRoot -TimeoutSeconds $TimeoutSeconds
        Assert-WinBookSplitProbeNotInterrupted $probe
        if ($probe.TimedOut -or $probe.Cancelled -or -not $probe.ParentStopped -or -not $probe.DescendantsStopped -or -not $probe.JobAssigned -or -not $probe.StreamsComplete -or $probe.StartError -or $probe.StopError -or $probe.StreamError -or $probe.StdoutTruncated -or $probe.StderrTruncated) {
            throw 'The bounded interpreter probe did not complete safely.'
        }
        $lines = @($probe.Stdout -split '\r?\n' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        if ($lines.Count -ne 1) { throw 'Expected one isolated interpreter JSON record.' }
        if (-not $lines[0].TrimStart().StartsWith('{')) { throw 'The interpreter JSON root must be an object.' }
        $record = $lines[0] | ConvertFrom-Json -ErrorAction Stop
        if ($record -isnot [pscustomobject]) { throw 'The interpreter JSON record must be one object.' }
        foreach ($field in @('protocol', 'version', 'implementation', 'platform', 'machine', 'executable', 'message')) {
            if ($record.$field -isnot [string] -or [string]::IsNullOrWhiteSpace($record.$field)) {
                throw ('Interpreter JSON requires a nonempty scalar string: ' + $field)
            }
        }
        if ($record.ok -is [bool] -and -not $record.ok) {
            throw ('Incompatible interpreter/dependency: ' + $record.message + ' Observed Python ' + $record.version + '.')
        }
        foreach ($field in @('pypdf_version', 'pypdf_path')) {
            if ($record.$field -isnot [string] -or [string]::IsNullOrWhiteSpace($record.$field)) { throw ('Interpreter JSON requires a nonempty scalar string: ' + $field) }
        }
        if (-not (Test-SplitInteger $record.schema_version) -or -not (Test-SplitInteger $record.bits) -or
            $record.protocol -cne 'winbooksplit.runtime' -or $record.schema_version -ne 1 -or
            $record.ok -isnot [bool] -or -not $record.ok -or $probe.ExitCode -ne 0 -or
            $record.version -cne '3.14.8' -or $record.implementation -cne 'cpython' -or
            $record.platform -cne 'win32' -or $record.machine -cne 'AMD64' -or $record.bits -ne 64 -or
            $record.gil_disabled -isnot [bool] -or $record.gil_disabled -or
            $record.isolated -isnot [bool] -or -not $record.isolated -or
            $record.dont_write_bytecode -isnot [bool] -or -not $record.dont_write_bytecode -or
            $record.pypdf_version -cne '6.20.0' -or
            -not ([IO.Path]::GetFullPath($record.executable).Equals($path, [StringComparison]::OrdinalIgnoreCase))) {
            throw ('Incompatible interpreter/dependency: ' + $record.message)
        }
        $null = Get-WinBookSplitTrustedPath $record.pypdf_path
        $Attempts.Add([pscustomobject]@{ Path = $path; Source = $Source; Accepted = $true; Message = $record.message; Probe = $probe })
        return [pscustomobject]@{ Path = $path; Version = $record.version; Source = $Source;
            PypdfVersion = $record.pypdf_version; PypdfPath = $record.pypdf_path;
            Details = $record; Probe = $probe; Arguments = @('-I', '-B'); Attempts = @($Attempts.ToArray()) }
    } catch {
        if (Test-WinBookSplitInterruption $_.Exception) { throw }
        $Attempts.Add([pscustomobject]@{ Path = $Path; Source = $Source; Accepted = $false; Message = $_.Exception.Message; Probe = $probe })
        return $null
    }
}

function Resolve-WinBookSplitRuntime {
    param([Parameter(Mandatory = $true)][string]$ApplicationRoot, [string]$PythonPath,
          [string]$DocumentPath, [ValidateRange(0.05, 60)][double]$TimeoutSeconds = 10)
    $root = Get-WinBookSplitTrustedPath -Path $ApplicationRoot -Directory
    $attempts = New-Object 'System.Collections.Generic.List[object]'
    $setup = "Use regular Windows x64 CPython 3.14.8 and pypdf 6.20.0. Set -PythonPath or create a fresh application .venv; installation is an explicit setup action."
    if (-not [string]::IsNullOrWhiteSpace($PythonPath)) {
        $selected = Test-WinBookSplitPythonCandidate $PythonPath 'explicit' $root $attempts $TimeoutSeconds
        if ($selected) { return $selected }
        throw (New-WinBookSplitDependencyError 'runtime_invalid' ("Python preflight failed for '$PythonPath': " + $attempts[-1].Message + ' ' + $setup + " For that selected compatible interpreter use & '$($PythonPath.Replace("'", "''"))' -I -m pip install --require-hashes --only-binary=:all: -r '$([IO.Path]::Combine($root, 'requirements.txt').Replace("'", "''"))'.") $attempts)
    }
    $venv = [IO.Path]::Combine($root, '.venv')
    if ((Get-Item -LiteralPath $venv -Force -ErrorAction SilentlyContinue) -or (Test-Path -LiteralPath $venv)) {
        $candidate = [IO.Path]::Combine($venv, 'Scripts\python.exe')
        $selected = Test-WinBookSplitPythonCandidate $candidate 'application_venv' $root $attempts $TimeoutSeconds
        if ($selected) { return $selected }
        throw (New-WinBookSplitDependencyError 'runtime_invalid' ("Application .venv preflight failed for '$candidate': " + $attempts[-1].Message + ' ' + $setup + " After explicitly creating a compatible venv use & '$($candidate.Replace("'", "''"))' -I -m pip install --require-hashes --only-binary=:all: -r '$([IO.Path]::Combine($root, 'requirements.txt').Replace("'", "''"))'.") $attempts)
    }
    $seen = @{}
    foreach ($launcher in @(Get-WinBookSplitPathCandidates 'py.exe' $root $DocumentPath $attempts)) {
        $probe = $null
        try {
            $launcher = Get-WinBookSplitTrustedPath $launcher -Automatic
            $probe = Invoke-WinBookSplitRuntimeProbe $launcher @('-0p') $root -TimeoutSeconds $TimeoutSeconds
            Assert-WinBookSplitProbeNotInterrupted $probe
            if ($probe.ExitCode -ne 0 -or $probe.TimedOut -or $probe.Cancelled -or -not $probe.ParentStopped -or -not $probe.DescendantsStopped -or -not $probe.JobAssigned -or -not $probe.StreamsComplete -or
                $probe.StartError -or $probe.StopError -or $probe.StreamError -or $probe.StdoutTruncated -or $probe.StderrTruncated) { throw 'Read-only Python launcher listing failed.' }
            $attempts.Add([pscustomobject]@{ Path = $launcher; Source = 'py_launcher_listing'; Accepted = $true; Message = 'Read-only -0p listing; no runtime launch/install requested.'; Probe = $probe })
            foreach ($line in (($probe.Stdout + "`n" + $probe.Stderr) -split '\r?\n')) {
                if ($line -match '([A-Za-z]:[\\/].*\.exe)\s*$') {
                    $candidate = $Matches[1].Trim()
                    if ($seen.ContainsKey($candidate)) { continue }; $seen[$candidate] = $true
                    if (-not (Test-WinBookSplitAutomaticPath $candidate $root $DocumentPath)) { continue }
                    $selected = Test-WinBookSplitPythonCandidate $candidate 'py_launcher' $root $attempts $TimeoutSeconds
                    if ($selected) { return $selected }
                }
            }
        } catch {
            if (Test-WinBookSplitInterruption $_.Exception) { throw }
            $attempts.Add([pscustomobject]@{ Path = $launcher; Source = 'py_launcher_listing'; Accepted = $false; Message = $_.Exception.Message; Probe = $probe })
        }
    }
    foreach ($candidate in @(Get-WinBookSplitPathCandidates 'python.exe' $root $DocumentPath $attempts)) {
        if ($seen.ContainsKey($candidate)) { continue }; $seen[$candidate] = $true
        $selected = Test-WinBookSplitPythonCandidate $candidate 'PATH' $root $attempts $TimeoutSeconds
        if ($selected) { return $selected }
    }
    throw (New-WinBookSplitDependencyError 'runtime_not_found' ('No compatible trusted Python was found. ' + $setup) $attempts)
}

function Test-WinBookSplitConverterCandidate {
    param([string]$Path, [string]$Source, [string]$ApplicationRoot, $Attempts, [double]$TimeoutSeconds)
    $probe = $null
    try {
        $path = Get-WinBookSplitTrustedPath $Path -Automatic:($Source -in @('PATH', 'known_location'))
        $probe = Invoke-WinBookSplitRuntimeProbe $path @('--version') $ApplicationRoot -TimeoutSeconds $TimeoutSeconds
        Assert-WinBookSplitProbeNotInterrupted $probe
        if ($probe.ExitCode -ne 0 -or $probe.TimedOut -or $probe.Cancelled -or -not $probe.ParentStopped -or -not $probe.DescendantsStopped -or -not $probe.JobAssigned -or -not $probe.StreamsComplete -or
            $probe.StartError -or $probe.StopError -or $probe.StreamError -or $probe.StdoutTruncated -or $probe.StderrTruncated -or
            $probe.Stdout -cnotmatch '\Aebook-convert(?:\.exe)? \(calibre ([0-9]+\.[0-9]+\.[0-9]+)\)(?:\r?\nCreated by: Kovid Goyal <kovid@kovidgoyal\.net>)?(?:\r?\n)?\z') {
            throw 'The converter did not return a complete ebook-convert version probe.'
        }
        $version = $Matches[1]
        if ($version -cne '9.15.0') { throw ('Expected Calibre 9.15.0 exactly; observed ' + $version + '.') }
        $Attempts.Add([pscustomobject]@{ Path = $path; Source = $Source; Accepted = $true; Message = 'Supported converter version.'; Probe = $probe })
        return [pscustomobject]@{ Path = $path; Version = $version; Source = $Source; Probe = $probe; Attempts = @($Attempts.ToArray()) }
    } catch {
        if (Test-WinBookSplitInterruption $_.Exception) { throw }
        $Attempts.Add([pscustomobject]@{ Path = $Path; Source = $Source; Accepted = $false; Message = $_.Exception.Message; Probe = $probe })
        return $null
    }
}

function Resolve-WinBookSplitConverter {
    param([Parameter(Mandatory = $true)][string]$ApplicationRoot, [string]$CalibrePath,
          [string]$DocumentPath, [ValidateRange(0.05, 60)][double]$TimeoutSeconds = 10)
    $root = Get-WinBookSplitTrustedPath -Path $ApplicationRoot -Directory
    $attempts = New-Object 'System.Collections.Generic.List[object]'
    $guidance = 'Select trusted Calibre 9.15.0 ebook-convert.exe with -CalibrePath, a trusted PATH entry, or a standard install; PDF-only processing does not require Calibre.'
    if (-not [string]::IsNullOrWhiteSpace($CalibrePath)) {
        $selected = Test-WinBookSplitConverterCandidate $CalibrePath 'explicit' $root $attempts $TimeoutSeconds
        if ($selected) { return $selected }
        throw (New-WinBookSplitDependencyError 'converter_invalid' ("Converter preflight failed for '$CalibrePath': " + $attempts[-1].Message + ' ' + $guidance) $attempts)
    }
    foreach ($candidate in @(Get-WinBookSplitPathCandidates 'ebook-convert.exe' $root $DocumentPath $attempts)) {
        $selected = Test-WinBookSplitConverterCandidate $candidate 'PATH' $root $attempts $TimeoutSeconds
        if ($selected) { return $selected }
    }
    foreach ($candidate in @('C:\Program Files\Calibre2\ebook-convert.exe', 'C:\Program Files (x86)\Calibre2\ebook-convert.exe',
                            [IO.Path]::Combine($env:LOCALAPPDATA, 'Programs\Calibre\ebook-convert.exe'))) {
        if ([IO.File]::Exists($candidate) -and (Test-WinBookSplitAutomaticPath $candidate $root $DocumentPath)) {
            $selected = Test-WinBookSplitConverterCandidate $candidate 'known_location' $root $attempts $TimeoutSeconds
            if ($selected) { return $selected }
        }
    }
    throw (New-WinBookSplitDependencyError 'converter_not_found' ('No compatible trusted converter was found. ' + $guidance) $attempts)
}
