# External task08 draft primitives only. No application is launched by loading
# this file. Entry-point wiring follows the frozen native_smoke/bindings schema.

function Initialize-NativeAcceptanceDirectory {
    if ('WinBookSplitNative.Directory' -as [type]) { return }
    Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
namespace WinBookSplitNative {
    public static class Directory {
        [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        static extern bool CreateDirectoryW(string path, IntPtr security);
        [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        static extern bool SetEnvironmentVariableW(string name, string value);
        public static void CreateExclusive(string path) {
            if (!CreateDirectoryW(path, IntPtr.Zero))
                throw new Win32Exception(Marshal.GetLastWin32Error());
        }
        public static void SetConsoleFixtureEnvironment(string name, string value) {
            // Exactly the task's fixed console fixture variables; no general
            // environment-override capability or production transport change.
            string[] allowed={"WBS_KIT_BAT","WBS_KIT_INPUT1","WBS_KIT_INPUT2",
                              "WBS_KIT_EXPAND","WBS_KIT_X","WBS_KIT_CONTROL_ROOT","WBS_KIT_CONTROL_APP","WBS_KIT_CONTROL_PYTHON",
                              "WBS_CLI_HELP_APP","CALIBRE_CONFIG_DIRECTORY",
                              "CALIBRE_CACHE_DIRECTORY","TEMP","TMP"};
            if (!Array.Exists(allowed, item=>String.Equals(item,name,StringComparison.OrdinalIgnoreCase)))
                throw new ArgumentException("Unknown native console fixture environment name.");
            // Win32 distinguishes an empty value from removal; preserve both.
            if (!SetEnvironmentVariableW(name,value))
                throw new Win32Exception(Marshal.GetLastWin32Error());
        }
        public static void RestoreConsoleFixtureEnvironment(string name, string value, bool existed) {
            // PowerShell may convert $null to String.Empty for a CLR string
            // parameter. Select removal inside the CLR with an explicit bool.
            SetConsoleFixtureEnvironment(name, existed ? value : null);
        }
    }
}
'@ -ErrorAction Stop
}

function Get-NativeAcceptanceOrdinaryPath {
    param([Parameter(Mandatory=$true)][string]$Path,[switch]$Directory)
    if (-not [IO.Path]::IsPathRooted($Path) -or $Path.StartsWith('\\') -or $Path.IndexOf([char]0) -ge 0) {
        throw 'Native kit paths require an absolute local literal path.'
    }
    $full=[IO.Path]::GetFullPath($Path)
    if (-not [StringComparer]::OrdinalIgnoreCase.Equals($full,$Path.Replace('/','\'))) {
        throw 'Native kit paths require their canonical literal spelling.'
    }
    $current=$full
    while ($null -ne $current) {
        $attributes=[IO.File]::GetAttributes($current)
        if (($attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw 'Native kit paths cannot contain reparse entries.'
        }
        if ($current -ne $full -or $Directory) {
            if (($attributes -band [IO.FileAttributes]::Directory) -eq 0) { throw 'Native kit directory is not ordinary.' }
        } elseif (($attributes -band [IO.FileAttributes]::Directory) -ne 0) { throw 'Native kit file is not ordinary.' }
        $current=[IO.Path]::GetDirectoryName($current)
        if ([string]::IsNullOrEmpty($current)) { $current=$null }
    }
    return $full
}

function Get-NativeAcceptanceFileHash {
    param([Parameter(Mandatory=$true)][string]$Path,[ValidateRange(1,536870912)][long]$MaximumBytes,[switch]$AllowEmpty)
    $full=Get-NativeAcceptanceOrdinaryPath $Path
    $lease=New-Object WinBookSplitLogging.PathLease($full,$false,$false,$true)
    $stream=$algorithm=$null
    try {
        $stream=[IO.File]::Open($full,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read)
        $lease.AssertSameHandle($stream.SafeFileHandle)
        if ($stream.Length -lt 0 -or (-not $AllowEmpty -and $stream.Length -eq 0) -or $stream.Length -gt $MaximumBytes) { throw 'Native kit file exceeds its bounded size.' }
        $size=$stream.Length; $algorithm=[Security.Cryptography.SHA256]::Create()
        $hash=([BitConverter]::ToString($algorithm.ComputeHash($stream))).Replace('-','').ToLowerInvariant()
        if ($stream.Length -ne $size) { throw 'Native kit file length changed.' }
        $lease.AssertUnchanged()
        return [pscustomobject]@{bytes=$size;sha256=$hash}
    } finally {
        if ($algorithm) { $algorithm.Dispose() }
        if ($stream) { $stream.Dispose() }
        $lease.Dispose()
    }
}

function Read-NativeAcceptanceJson {
    param([Parameter(Mandatory=$true)][string]$Path,[ValidateRange(1,8388608)][int]$MaximumBytes=131072)
    $full=Get-NativeAcceptanceOrdinaryPath $Path
    $lease=New-Object WinBookSplitLogging.PathLease($full,$false,$false,$true)
    $stream=$null
    try {
        $stream=[IO.File]::Open($full,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read)
        $lease.AssertSameHandle($stream.SafeFileHandle)
        if ($stream.Length -lt 1 -or $stream.Length -gt $MaximumBytes) { throw 'Native kit JSON exceeds its bounded size.' }
        $bytes=New-Object byte[] ([int]$stream.Length); $count=0
        while ($count -lt $bytes.Length) {
            $read=$stream.Read($bytes,$count,$bytes.Length-$count)
            if ($read -eq 0) { throw 'Native kit JSON ended early.' }
            $count+=$read
        }
        if ($stream.ReadByte() -ne -1) { throw 'Native kit JSON length changed.' }
        $text=([Text.UTF8Encoding]::new($false,$true)).GetString($bytes)
        $value=[WinBookSplitSupport.StrictJson]::Parse($text)
        $algorithm=[Security.Cryptography.SHA256]::Create()
        try { $hash=([BitConverter]::ToString($algorithm.ComputeHash($bytes))).Replace('-','').ToLowerInvariant() }
        finally { $algorithm.Dispose() }
        $lease.AssertUnchanged()
        return [pscustomobject]@{value=$value;sha256=$hash;bytes=$bytes.Length}
    } finally { if ($stream) { $stream.Dispose() }; $lease.Dispose() }
}

function Write-NativeAcceptanceJson {
    param([Parameter(Mandatory=$true)][string]$Path,[Parameter(Mandatory=$true)]$Record,
          [ValidateRange(1024,8388608)][int]$MaximumBytes=8388608)
    if (-not [IO.Path]::IsPathRooted($Path)) { throw 'Native receipt requires an absolute literal sibling.' }
    $full=[IO.Path]::GetFullPath($Path)
    if (-not [StringComparer]::OrdinalIgnoreCase.Equals($full,$Path.Replace('/','\'))) { throw 'Native receipt requires its literal canonical path.' }
    $parent=Get-NativeAcceptanceOrdinaryPath ([IO.Path]::GetDirectoryName($full)) -Directory
    $name=[IO.Path]::GetFileName($full)
    if ($name -match '[<>:"/\\|?*\x00-\x1f\x7f]' -or $name -ne $name.TrimEnd(' ','.') -or
        $name -match '^(?i:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)' -or $name -in @('','.','..')) {
        throw 'Native receipt requires one ordinary Windows filename.'
    }
    $lease=New-Object WinBookSplitLogging.PathLease($parent,$true)
    $stream=$null
    try {
        $bytes=([Text.UTF8Encoding]::new($false,$true)).GetBytes(($Record | ConvertTo-Json -Depth 100 -Compress))
        if ($bytes.Length -gt $MaximumBytes) { throw 'Native receipt exceeds its finite bound.' }
        $lease.AssertUnchanged()
        $stream=[IO.File]::Open($full,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::Read)
        $stream.Write($bytes,0,$bytes.Length); $stream.Flush(); $lease.AssertUnchanged()
    } finally { if ($stream) { $stream.Dispose() }; $lease.Dispose() }
}

function Assert-NativeAcceptanceProcess {
    param([Parameter(Mandatory=$true)]$Probe,[ValidateSet('normal','timeout','cancel')][string]$ExpectedStop='normal')
    foreach ($key in @('Pid','ExitCode','JobAssigned','ParentStopped','DescendantsStopped','StreamsComplete','InputWriterStopped',
        'StartError','StopError','StreamError','ResultError','InteractionError','InputError','StdoutTruncated','StderrTruncated',
        'TimedOut','Cancelled','Stdout','Stderr','StdoutTotalBytes','StderrTotalBytes')) {
        $present=if ($Probe -is [Collections.IDictionary]) { $key -cin @($Probe.Keys) } else { $null -ne $Probe.PSObject.Properties[$key] }
        if (-not $present) { throw ('Missing native process evidence field: '+$key) }
    }
    if (-not (Test-SplitInteger $Probe.Pid) -or $Probe.Pid -lt 1) { throw 'Native process PID is malformed.' }
    if ($ExpectedStop -eq 'normal' -and -not (Test-SplitInteger $Probe.ExitCode)) { throw 'Native process exit is malformed.' }
    foreach ($key in @('JobAssigned','ParentStopped','DescendantsStopped','StreamsComplete','InputWriterStopped')) {
        if ($Probe.$key -isnot [bool] -or $Probe.$key -ne $true) { throw ('Unproved native stop/drain/job: '+$key) }
    }
    foreach ($key in @('StartError','StopError','StreamError','ResultError','InteractionError','InputError')) {
        if ($null -ne $Probe.$key) { throw ('Native process supervision failed: '+$key) }
    }
    foreach ($key in @('StdoutTruncated','StderrTruncated','TimedOut','Cancelled')) {
        if ($Probe.$key -isnot [bool]) { throw ('Malformed native process flag: '+$key) }
    }
    if ($Probe.StdoutTruncated -or $Probe.StderrTruncated -or
        $Probe.TimedOut -ne ($ExpectedStop -eq 'timeout') -or $Probe.Cancelled -ne ($ExpectedStop -eq 'cancel')) {
        throw 'Native process interruption/truncation differs from its authored contract.'
    }
    foreach ($key in @('Stdout','Stderr')) {
        $total=$key+'TotalBytes'
        if ($Probe.$key -isnot [string] -or -not (Test-SplitInteger $Probe.$total) -or
            $Probe.$total -ne ([Text.UTF8Encoding]::new($false,$true)).GetByteCount($Probe.$key)) {
            throw 'Native process UTF8 stream/count differs.'
        }
    }
}

function Invoke-NativeAcceptanceProcessCase {
    param([Parameter(Mandatory=$true)][string]$Path,[Parameter(Mandatory=$true)][AllowEmptyCollection()][string[]]$Arguments,
          [Parameter(Mandatory=$true)][string]$WorkingDirectory,[ValidateRange(1,600)][double]$TimeoutSeconds=120,
          [AllowEmptyCollection()][string[]]$Answers=@(),[Collections.IDictionary]$Environment=@{})
    $null=Get-NativeAcceptanceOrdinaryPath $Path
    $null=Get-NativeAcceptanceOrdinaryPath $WorkingDirectory -Directory
    if ($Answers.Count -gt 8) { throw 'Native console recipe exceeds the existing reply queue.' }
    Initialize-NativeAcceptanceDirectory
    $previous=@{}; $restored=$false; $session=$null; $probe=$null; $failure=$null
    try {
        $snapshot=[Environment]::GetEnvironmentVariables('Process')
        foreach ($key in $Environment.Keys) {
            if ($key -isnot [string] -or $Environment[$key] -isnot [string] -or
                $Environment[$key].IndexOf([char]0) -ge 0) { throw 'Malformed fixed console environment value.' }
            $previous[$key]=[pscustomobject]@{exists=$snapshot.Contains($key);value=$snapshot[$key]}
            [WinBookSplitNative.Directory]::SetConsoleFixtureEnvironment($key,$Environment[$key])
        }
        $parameters=@{Path=$Path;Arguments=$Arguments;WorkingDirectory=$WorkingDirectory;
            TimeoutSeconds=$TimeoutSeconds;CaptureLimitBytes=1048576;ResultLimitBytes=8388608}
        if ($Answers.Count -eq 0) {
            $probe=Invoke-WinBookSplitProcess @parameters
        } else {
            $session=Start-WinBookSplitProcessSession @parameters
            foreach ($answer in $Answers) { $session.SendReply($answer) }
            $clock=[Diagnostics.Stopwatch]::StartNew()
            while (-not $session.Completed -and $clock.Elapsed.TotalSeconds -lt ($TimeoutSeconds+8)) {
                Start-Sleep -Milliseconds 20
            }
            if (-not $session.Completed) { $session.Cancel(); throw 'Native session did not complete inside its finite owned supervision.' }
            $probe=$session.Result
        }
    } catch { $failure=$_.Exception }
    finally {
        if ($session) {
            try { $session.Dispose(); if ($session.Completed -and $null -eq $probe) { $probe=$session.Result } }
            catch { if ($null -eq $failure) { $failure=$_.Exception } }
        }
        foreach ($key in $previous.Keys) {
            try {
                $value=$null; if ($previous[$key].exists) { $value=[string]$previous[$key].value }
                [WinBookSplitNative.Directory]::RestoreConsoleFixtureEnvironment($key,$value,[bool]$previous[$key].exists)
            } catch { if ($null -eq $failure) { $failure=$_.Exception } }
        }
        $snapshot=[Environment]::GetEnvironmentVariables('Process'); $restored=$true
        foreach ($key in $previous.Keys) {
            if ($snapshot.Contains($key) -ne $previous[$key].exists -or $snapshot[$key] -cne $previous[$key].value) { $restored=$false }
        }
    }
    return [pscustomobject]@{probe=$probe;command=(@($Path)+@($Arguments));cwd=$WorkingDirectory;answers=@($Answers);
        environment_before_after_restored=$restored;failure=$failure}
}

function Get-NativeAcceptanceOs {
    param([switch]$AllowWindows11DryRun)
    Import-Module ([IO.Path]::Combine($PSHOME,'Modules\CimCmdlets\CimCmdlets.psd1')) -ErrorAction Stop
    $rows=@(Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop)
    $cpus=@(Get-CimInstance -ClassName Win32_Processor -ErrorAction Stop)
    if ($rows.Count -ne 1 -or $cpus.Count -lt 1 -or $cpus.Count -gt 64) { throw 'Native OS observation is not bounded and unambiguous.' }
    $observed=$rows[0]; $architectures=@($cpus|ForEach-Object {[int]$_.Architecture})
    $key=[Microsoft.Win32.Registry]::LocalMachine.OpenSubKey('SOFTWARE\Microsoft\Windows NT\CurrentVersion',$false)
    try { $display=[string]$key.GetValue('DisplayVersion',$null);$ubr=$key.GetValue('UBR',$null) }
    finally { if ($key) { $key.Dispose() } }
    $scope='BLOCKED'
    $native=[Environment]::Is64BitOperatingSystem -and [IntPtr]::Size -eq 8 -and [int]$observed.ProductType -eq 1 -and
        @($architectures|Where-Object {$_ -ne 9}).Count -eq 0 -and $PSVersionTable.PSVersion.Major -eq 5 -and $PSVersionTable.PSVersion.Minor -eq 1
    if ($native -and $observed.BuildNumber -ceq '19045' -and $observed.Version -ceq '10.0.19045' -and
        $display -ceq '22H2' -and $observed.Caption -match '^Microsoft Windows 10 (Home(?: Single Language| Country Specific| China)?|Pro(?: for Workstations| N for Workstations| Education)?|Enterprise(?: G)?|Education|IoT Enterprise|SE)(?: N)?$') { $scope='NATIVE_WINDOWS10' }
    elseif ($native -and $AllowWindows11DryRun -and [int]$observed.BuildNumber -ge 22000 -and
        $observed.Version -ceq ('10.0.'+$observed.BuildNumber) -and $observed.Caption -match '^Microsoft Windows 11 (Home(?: Single Language| Country Specific| China)?|Pro(?: for Workstations| N for Workstations| Education)?|Enterprise(?: G)?|Education|IoT Enterprise|SE)(?: N)?$') { $scope='NATIVE_WINDOWS11_DRY_RUN' }
    return [ordered]@{Caption=[string]$observed.Caption;Version=[string]$observed.Version;BuildNumber=[string]$observed.BuildNumber;
        OSArchitecture=[string]$observed.OSArchitecture;ProductType=[int]$observed.ProductType;UBR=$ubr;
        powershell_version=$PSVersionTable.PSVersion.ToString();process_bits=([IntPtr]::Size*8);DisplayVersion=$display;
        processor_architectures=$architectures;os_bits=$(if ([Environment]::Is64BitOperatingSystem) {64} else {32});platform_scope=$scope}
}

function ConvertTo-NativeAcceptanceSafeOs {
    param($Record)
    if ($null -eq $Record) {return $null}
    $safe=[ordered]@{Caption=$null;Version=$null;BuildNumber=$null;OSArchitecture=$null;ProductType=$null;UBR=$null;
        powershell_version=$null;process_bits=$null;DisplayVersion=$null;processor_architectures=@();os_bits=$null;platform_scope='BLOCKED'}
    $patterns=@{Caption='^Microsoft Windows (10|11) (Home(?: Single Language| Country Specific| China)?|Pro(?: for Workstations| N for Workstations| Education)?|Enterprise(?: G)?|Education|IoT Enterprise|SE)(?: N)?$';Version='^10\.0\.[0-9]{5}$';
        BuildNumber='^[0-9]{5}$';OSArchitecture='^(64-bit|64ビット|64 bits|64-Bit|64-bittinen)$';
        powershell_version='^5\.1\.[0-9]+\.[0-9]+$';DisplayVersion='^[0-9]{2}H[12]$'}
    foreach($key in $patterns.Keys){if($Record[$key] -is [string] -and $Record[$key] -cmatch $patterns[$key]){$safe[$key]=$Record[$key]}}
    foreach($key in @('ProductType','UBR','process_bits','os_bits')){if((Test-SplitInteger $Record[$key]) -and $Record[$key] -ge 0 -and $Record[$key] -le 1000000){$safe[$key]=$Record[$key]}}
    if($Record.processor_architectures -is [array] -and $Record.processor_architectures.Count -le 64 -and
        @($Record.processor_architectures|Where-Object{-not(Test-SplitInteger $_) -or $_ -lt 0 -or $_ -gt 12}).Count -eq 0){$safe.processor_architectures=@($Record.processor_architectures)}
    if($Record.platform_scope -cin @('NATIVE_WINDOWS10','NATIVE_WINDOWS11_DRY_RUN','BLOCKED')){$safe.platform_scope=$Record.platform_scope}
    return $safe
}

function New-NativeAcceptanceDirectory {
    param([Parameter(Mandatory=$true)][string]$Path)
    $parent=Get-NativeAcceptanceOrdinaryPath ([IO.Path]::GetDirectoryName($Path)) -Directory
    $lease=New-Object WinBookSplitLogging.PathLease($parent,$true)
    try { $lease.AssertUnchanged(); Initialize-NativeAcceptanceDirectory; [WinBookSplitNative.Directory]::CreateExclusive($Path); $lease.AssertUnchanged() }
    finally { $lease.Dispose() }
    return (Get-NativeAcceptanceOrdinaryPath $Path -Directory)
}

function Test-NativeAcceptancePortableTargetLength {
    param([Parameter(Mandatory=$true)][string]$Target)
    # The vendor's v9.15.0 installer compares the appended path with wcslen.
    # String.Length uses the same UTF-16 units, including surrogate pairs.
    return [IO.Path]::Combine([IO.Path]::GetFullPath($Target),'Calibre Portable').Length -le 58
}

function New-NativeAcceptancePortableDirectory {
    param([Parameter(Mandatory=$true)][string]$Profile,
          [Parameter(Mandatory=$true)]$Leases)
    $null=Get-NativeAcceptanceOrdinaryPath $Profile -Directory
    # Retain the short target's own root-to-profile ancestry in the caller's
    # existing lease stack, including across native execution and final writes.
    $parents=New-Object 'Collections.Generic.List[string]';$current=$Profile
    while($current){$parents.Add($current);$current=[IO.Path]::GetDirectoryName($current)}
    for($i=$parents.Count-1;$i -ge 0;$i--){$Leases.Add((New-Object WinBookSplitLogging.PathLease($parents[$i],$true)))}
    # A separate short, exclusively created sibling remains private per user.
    for($attempt=0;$attempt -lt 4;$attempt++){
        $target=[IO.Path]::Combine($Profile,'WBS-C-'+[Guid]::NewGuid().ToString('N').Substring(0,12))
        if(-not (Test-NativeAcceptancePortableTargetLength $target)){
            $error=New-Object IO.IOException('Private portable provisioning is blocked by the vendor path-length bound.')
            $error.Data['NativeBlocked']='portable_target_too_long';throw $error
        }
        try{
            $created=New-NativeAcceptanceDirectory $target
            $lease=New-Object WinBookSplitLogging.PathLease($created,$true);$Leases.Add($lease)
            $lease.AssertUnchanged()
            $entries=[IO.Directory]::EnumerateFileSystemEntries($created).GetEnumerator()
            try{if($entries.MoveNext()){throw 'Fresh portable target unexpectedly contains an entry.'}}finally{$entries.Dispose()}
            return $created
        }catch{
            $collision=$_.Exception
            while($collision.InnerException){$collision=$collision.InnerException}
            if($collision -isnot [ComponentModel.Win32Exception] -or $collision.NativeErrorCode -notin @(80,183)){throw}
            if($attempt -eq 3){throw 'No exclusive private portable target could be reserved within the bounded attempts.'}
        }
    }
}

function Get-NativeAcceptanceDownload {
    param([Parameter(Mandatory=$true)]$Expected,[Parameter(Mandatory=$true)][string]$Destination)
    if ($Expected.url -isnot [string] -or $Expected.sha256 -isnot [string] -or $Expected.sha256 -cnotmatch '^[0-9a-f]{64}$' -or
        -not (Test-SplitInteger $Expected.bytes) -or $Expected.bytes -lt 1 -or $Expected.bytes -gt 268435456) { throw 'Invalid fixed acquisition identity.' }
    $uri=[Uri]$Expected.url
    if ($uri.Scheme -cne 'https' -or $uri.UserInfo -or $uri.Host -notin @('www.python.org','files.pythonhosted.org','download.calibre-ebook.com')) { throw 'Acquisition is outside its fixed primary-source hosts.' }
    $null=Get-NativeAcceptanceOrdinaryPath ([IO.Path]::GetDirectoryName($Destination)) -Directory
    $parentLease=New-Object WinBookSplitLogging.PathLease([IO.Path]::GetDirectoryName($Destination),$true)
    $stream=$input=$response=$null
    $timer=[Diagnostics.Stopwatch]::StartNew()
    $request=[Net.HttpWebRequest]::Create($uri);$request.Method='GET';$request.Timeout=30000;$request.ReadWriteTimeout=30000
    $request.AllowAutoRedirect=$true;$request.MaximumAutomaticRedirections=3
    try {
        $response=$request.GetResponse()
        if ($response.ResponseUri.Scheme -cne 'https' -or $response.ResponseUri.Host -cne $uri.Host -or
            ($response.ContentLength -ne -1 -and $response.ContentLength -ne $Expected.bytes)) { throw 'Acquisition response identity differs from the fixed source.' }
        $parentLease.AssertUnchanged();$input=$response.GetResponseStream();$stream=[IO.File]::Open($Destination,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::Read)
        $buffer=New-Object byte[] 65536;$total=0L
        while (($count=$input.Read($buffer,0,$buffer.Length)) -gt 0) {
            $total+=$count
            if ($total -gt $Expected.bytes -or $timer.Elapsed.TotalSeconds -gt 300) { throw 'Acquisition exceeded its fixed bytes or deadline.' }
            $stream.Write($buffer,0,$count)
        }
        if ($total -ne $Expected.bytes) { throw 'Acquisition is incomplete.' }
        $stream.Flush()
        $parentLease.AssertUnchanged()
    } finally { if ($stream) {$stream.Dispose()};if ($input) {$input.Dispose()};if ($response) {$response.Dispose()};$parentLease.Dispose() }
    $actual=Get-NativeAcceptanceFileHash $Destination 268435456
    if ($actual.sha256 -cne $Expected.sha256 -or $actual.bytes -ne $Expected.bytes) { throw 'Acquisition hash differs from the reviewed primary-source pin.' }
    return $Destination
}

function Expand-NativeAcceptancePython {
    param([Parameter(Mandatory=$true)][string]$Archive,[Parameter(Mandatory=$true)][string]$Destination,[Parameter(Mandatory=$true)]$Expected)
    Add-Type -AssemblyName System.IO.Compression -ErrorAction Stop
    $null=New-NativeAcceptanceDirectory $Destination
    $leases=New-Object 'Collections.Generic.List[IDisposable]';$dirs=@{};$stream=$zip=$null
    try {
        $dirs[$Destination]=New-Object WinBookSplitLogging.PathLease($Destination,$true);$leases.Add($dirs[$Destination])
        $lease=New-Object WinBookSplitLogging.PathLease($Archive,$false,$false,$true);$leases.Add($lease)
        $stream=[IO.File]::Open($Archive,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read);$lease.AssertSameHandle($stream.SafeFileHandle)
        $algorithm=[Security.Cryptography.SHA256]::Create()
        try{$digest=([BitConverter]::ToString($algorithm.ComputeHash($stream))).Replace('-','').ToLowerInvariant()}
        finally{$algorithm.Dispose()}
        if($stream.Length -ne $Expected.bytes -or $digest -cne $Expected.sha256){throw 'Held Python archive no longer matches its primary-source pin.'};$stream.Position=0
        $zip=New-Object IO.Compression.ZipArchive($stream,[IO.Compression.ZipArchiveMode]::Read,$false)
        if ($zip.Entries.Count -lt 1 -or $zip.Entries.Count -gt 4096) { throw 'Python archive member count exceeds its bound.' }
        $seen=@{};$total=0L
        foreach ($entry in $zip.Entries) {
            $name=$entry.FullName.TrimEnd('/')
            if (-not $name -or $name.StartsWith('/') -or $name.Contains('\') -or $name -match '[<>:"|?*\x00-\x1f\x7f]' -or
                ($entry.ExternalAttributes -band 0xF0000000L) -eq 0xA0000000L -or $entry.Length -lt 0 -or $entry.Length -gt 33554432) { throw 'Unsafe Python archive member.' }
            foreach ($part in $name.Split('/')) {
                if ($part -in @('','.','..') -or $part -ne $part.TrimEnd(' ','.') -or $part -match '^(?i:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)') { throw 'Unsafe Python archive path.' }
            }
            if ($seen.ContainsKey($name)) { throw 'Colliding Python archive paths.' };$seen[$name]=$entry.FullName.EndsWith('/')
            $total+=$entry.Length;if ($total -gt 268435456) { throw 'Python archive expansion exceeds its bound.' }
        }
        $ancestorSpellings=@{}
        foreach($name in $seen.Keys){
            $parts=$name.Split('/')
            for($i=1;$i -lt $parts.Count;$i++){
                $ancestor=[string]::Join('/',$parts[0..($i-1)])
                if($seen.ContainsKey($ancestor) -and -not $seen[$ancestor]){throw 'Python archive file/ancestor collision.'}
                if($ancestorSpellings.ContainsKey($ancestor) -and $ancestorSpellings[$ancestor] -cne $ancestor){throw 'Python archive ancestor spelling collision.'}
                $ancestorSpellings[$ancestor]=$ancestor
            }
        }
        foreach ($entry in $zip.Entries) {
            $relative=$entry.FullName.TrimEnd('/');$directory=$entry.FullName.EndsWith('/')
            $parts=$relative.Split('/');$parent=$Destination
            $directoryParts=if ($directory) {$parts.Count} else {$parts.Count-1}
            for ($i=0;$i -lt $directoryParts;$i++) {
                $child=[IO.Path]::Combine($parent,$parts[$i])
                if (-not $dirs.ContainsKey($child)) { $dirs[$parent].AssertUnchanged();$null=New-NativeAcceptanceDirectory $child;$dirs[$child]=New-Object WinBookSplitLogging.PathLease($child,$true);$leases.Add($dirs[$child]) }
                $parent=$child
            }
            if ($directory) { continue }
            $dirs[$parent].AssertUnchanged();$target=[IO.Path]::Combine($parent,$parts[$parts.Count-1]);$input=$output=$null
            try {
                $input=$entry.Open();$output=[IO.File]::Open($target,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::Read)
                $buffer=New-Object byte[] 65536;$copied=0L
                while (($count=$input.Read($buffer,0,$buffer.Length)) -gt 0) { $copied+=$count;if ($copied -gt $entry.Length) {throw 'Python archive member exceeded its stated length.'};$output.Write($buffer,0,$count) }
                if ($copied -ne $entry.Length) {throw 'Python archive member is incomplete.'};$output.Flush()
            } finally {if($input){$input.Dispose()};if($output){$output.Dispose()}}
            $null=Get-NativeAcceptanceOrdinaryPath $target
        }
        foreach ($held in $leases) {$held.AssertUnchanged()}
    } finally { if($zip){$zip.Dispose()};if($stream){$stream.Dispose()};for($i=$leases.Count-1;$i -ge 0;$i--){$leases[$i].Dispose()} }
    return (Get-NativeAcceptanceOrdinaryPath ([IO.Path]::Combine($Destination,'python.exe')))
}

function Get-NativeAcceptanceFrame {
    param([string]$Text,[string]$Prefix)
    $lines=@($Text.Split([char]10)|Where-Object {$_.StartsWith($Prefix,[StringComparison]::Ordinal)})
    if ($lines.Count -ne 1) {throw 'Expected exactly one fixed native frame.'}
    return [WinBookSplitSupport.StrictJson]::Parse($lines[0].Substring($Prefix.Length).TrimEnd([char]13))
}

function Test-NativeAcceptanceProductionIdentity {
    param($Left,$Right)
    foreach($record in @($Left,$Right)){
        if($record -isnot [Collections.IDictionary] -or $record.Count -ne 6 -or $record.assets -isnot [Collections.IDictionary] -or $record.assets.Count -ne 3){return $false}
        foreach($key in @('version','source_commit','source_tree','payload_sha256','policy_sha256')){
            $pattern=if($key -ceq 'version'){'^[0-9]+\.[0-9]+\.[0-9]+$'}elseif($key -cin @('source_commit','source_tree')){'^[0-9a-f]{40}$'}else{'^[0-9a-f]{64}$'}
            if($record[$key] -isnot [string] -or $record[$key] -cnotmatch $pattern){return $false}
        }
    }
    foreach($key in @('version','source_commit','source_tree','payload_sha256','policy_sha256')){
        if($key -cnotin @($Left.Keys) -or $key -cnotin @($Right.Keys) -or $Left[$key] -isnot [string] -or $Right[$key] -isnot [string] -or $Left[$key] -cne $Right[$key]){return $false}
    }
    $archiveName='WinBookSplit-v'+$Right.version+'.zip'
    foreach($name in @($archiveName,'release-manifest.json','SHA256SUMS.txt')){
        if($name -cnotin @($Left.assets.Keys) -or $name -cnotin @($Right.assets.Keys)){return $false}
        $first=$Left.assets[$name];$second=$Right.assets[$name]
        foreach($row in @($first,$second)){if($row -isnot [Collections.IDictionary] -or $row.Count -ne 2 -or
            'bytes' -cnotin @($row.Keys) -or 'sha256' -cnotin @($row.Keys) -or -not(Test-SplitInteger $row.bytes) -or $row.bytes -lt 1 -or
            $row.bytes -gt 33554432 -or $row.sha256 -isnot [string] -or $row.sha256 -cnotmatch '^[0-9a-f]{64}$'){return $false}}
        if($first.bytes -ne $second.bytes -or $first.sha256 -cne $second.sha256){return $false}
    }
    return $true
}

function Get-NativeAcceptanceEnvironmentSnapshot {
    param([string]$Root)
    $records=New-Object 'Collections.Generic.List[object]'
    $items=[IO.Directory]::EnumerateFileSystemEntries($Root,'*',[IO.SearchOption]::AllDirectories).GetEnumerator()
    try {
        $count=0
        while ($items.MoveNext()) {
            $count++;if ($count -gt 8192) {throw 'Environment inventory exceeded its bound.'}
            $path=[string]$items.Current;$attributes=[IO.File]::GetAttributes($path)
            $directory=($attributes -band [IO.FileAttributes]::Directory) -ne 0
            $null=Get-NativeAcceptanceOrdinaryPath $path -Directory:$directory
            $row=[ordered]@{relative=$path.Substring($Root.Length);last_write_utc=[IO.File]::GetLastWriteTimeUtc($path).Ticks;directory=$directory}
            if (-not $directory) {$row.file=Get-NativeAcceptanceFileHash $path 67108864 -AllowEmpty}
            $records.Add($row)
        }
    } finally { $items.Dispose() }
    return ($records.ToArray()|ConvertTo-Json -Depth 10 -Compress)
}

function Assert-NativeAcceptanceCaseEnvironment {
    param($Case,[string]$WorkRoot,[string]$ApplicationRoot)
    if ($Case.environment -isnot [Collections.IDictionary]) {throw 'Malformed finite environment plan.'}
    if ($Case.kind -ceq 'reference') {
        if (@($Case.environment.Keys).Count -ne 4 -or $Case.id -cnotin @('ref-epub','ref-azw3')) {throw 'Reference environment differs from fixed recipe.'}
        foreach ($key in @('CALIBRE_CONFIG_DIRECTORY','CALIBRE_CACHE_DIRECTORY','TEMP','TMP')) {
            $leaf=if($key -ceq 'CALIBRE_CONFIG_DIRECTORY'){'cfg'}elseif($key -ceq 'CALIBRE_CACHE_DIRECTORY'){'cache'}else{'tmp'}
            $expected=[IO.Path]::Combine($WorkRoot,$Case.id,$leaf)
            if ($Case.environment[$key] -isnot [string] -or -not [StringComparer]::OrdinalIgnoreCase.Equals($Case.environment[$key],$expected)) {throw 'Reference settings are outside owned finite directories.'}
            $null=Get-NativeAcceptanceOrdinaryPath $expected -Directory
        }
    } elseif ($Case.kind -ceq 'bat') {
        if (@($Case.environment.Keys).Count -ne 4 -or $Case.environment.WBS_KIT_X -cne 'MUST_NOT_EXPAND' -or
            -not [StringComparer]::OrdinalIgnoreCase.Equals($Case.environment.WBS_KIT_BAT,[IO.Path]::Combine($ApplicationRoot,'WinBookSplit.bat'))) {throw 'BAT environment differs from its finite literal recipe.'}
        foreach($key in @('WBS_KIT_INPUT1','WBS_KIT_INPUT2')) {
            $path=Get-NativeAcceptanceOrdinaryPath $Case.environment[$key]
            if(-not $path.StartsWith($WorkRoot+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'BAT input is outside owned authored fixture work.'}
        }
    } elseif ($Case.environment.Count -ne 0) {throw 'Application cases cannot override their environment.'}
}

function Invoke-NativeAcceptanceTransportControl {
    param([string]$Python,[string]$Helper,[string]$WorkRoot,[ValidateSet('timeout','cancel')][string]$Kind)
    $session=$probe=$failure=$null;$ready=$false;$result=$null
    try{
        $timeout=if($Kind -ceq 'timeout'){3.0}else{12.0}
        $session=Start-WinBookSplitProcessSession -Path $Python -Arguments @('-I','-B','-X','utf8',$Helper,'wait-with-descendant') -WorkingDirectory $WorkRoot -TimeoutSeconds $timeout -CaptureLimitBytes 1048576 -ResultLimitBytes 8388608
        $clock=[Diagnostics.Stopwatch]::StartNew()
        while(-not $session.Completed -and $clock.Elapsed.TotalSeconds -lt 5){
            $request=$session.TryReadRequest()
            if($request){$event=[WinBookSplitSupport.StrictJson]::Parse($request);if((Test-SplitInteger $event.owned_child_pid) -and $event.owned_child_pid -gt 0){$ready=$true;break}}
            Start-Sleep -Milliseconds 10
        }
        if(-not $ready){$session.Cancel();throw 'Controlled native descendant was not observed.'}
        if($Kind -ceq 'cancel'){$session.Cancel()}
        while(-not $session.Completed -and $clock.Elapsed.TotalSeconds -lt 20){Start-Sleep -Milliseconds 10}
        if(-not $session.Completed){$session.Cancel();throw 'Controlled native shutdown exceeded its bound.'}
        $probe=$session.Result;Assert-NativeAcceptanceProcess $probe -ExpectedStop $Kind
        Write-NativeAcceptanceJson ([IO.Path]::Combine($WorkRoot,'control-'+$Kind+'.json')) $probe
        $result=[ordered]@{id=('owned_'+$Kind);scope='CONTROLLED_NATIVE';status='PASS';code='expected_stop_proved';
            probe_sha256=(Get-NativeAcceptanceFileHash ([IO.Path]::Combine($WorkRoot,'control-'+$Kind+'.json')) 8388608).sha256;
            helper_sha256=(Get-NativeAcceptanceFileHash $Helper 1048576).sha256}
    }catch{$failure=$_.Exception}
    finally{
        if($session){try{$session.Dispose();if($session.Completed -and $null -eq $probe){$probe=$session.Result}}catch{if($null -eq $failure){$failure=$_.Exception}}}
    }
    if($failure){
        $safe=$false
        if($probe){try{$stop=if($probe.Cancelled){'cancel'}elseif($probe.TimedOut){'timeout'}else{'normal'};Assert-NativeAcceptanceProcess $probe -ExpectedStop $stop;$safe=$true}catch{$safe=$false}}
        $error=New-Object IO.IOException('Controlled native transport failed; artifacts retained.',$failure)
        $error.Data['ControlUnsafe']=(-not $safe);if($probe){$error.Data['Probe']=$probe}
        throw $error
    }
    return $result
}

function Invoke-NativeAcceptanceDiscoveryControl {
    param([string]$Inbox,[string]$ApplicationRoot,[string]$SelectedPython,[string]$WorkRoot)
    $control=New-NativeAcceptanceDirectory ([IO.Path]::Combine($WorkRoot,'discovery-control'))
    # Authored static child source, not path interpolation or shell evaluation.
    # Its PATH override exists only inside this owned child process, and is
    # restored before its report. The actual candidate remains byte-identical.
    $source=@'
$ErrorActionPreference='Stop';$PSModuleAutoloadingPreference='None'
Import-Module ([IO.Path]::Combine($PSHOME,'Modules\Microsoft.PowerShell.Utility\Microsoft.PowerShell.Utility.psd1')) -ErrorAction Stop
Import-Module ([IO.Path]::Combine($PSHOME,'Modules\Microsoft.PowerShell.Management\Microsoft.PowerShell.Management.psd1')) -ErrorAction Stop
$root=$env:WBS_KIT_CONTROL_ROOT;$app=$env:WBS_KIT_CONTROL_APP;$python=$env:WBS_KIT_CONTROL_PYTHON
$old=$env:PATH;$restored=$false
try{
  $bin=[IO.Path]::Combine($root,'bin');[IO.Directory]::CreateDirectory($bin)|Out-Null
  $cwd=[IO.Path]::Combine($root,'cwd');[IO.Directory]::CreateDirectory($cwd)|Out-Null
  $exe=[IO.Path]::Combine($bin,'py.exe')
  $cs='using System; public static class AuthoredLegacyFailure { public static int Main(string[] args) { Console.Error.WriteLine("Unknown option: -0"); return 2; } }'
  Add-Type -TypeDefinition $cs -OutputAssembly $exe -OutputType ConsoleApplication -ErrorAction Stop
  . ([IO.Path]::Combine($app,'engine\WinBookSplit.Diagnostics.ps1'))
  . ([IO.Path]::Combine($app,'engine\WinBookSplit.Runtime.ps1'))
  $env:PATH=$bin+';'+[IO.Path]::GetDirectoryName($python)
  Set-Location -LiteralPath $cwd
  $runtime=Resolve-WinBookSplitRuntime -ApplicationRoot $app
  $algorithm=[Security.Cryptography.SHA256]::Create()
  try{$stream=[IO.File]::OpenRead($exe);try{$hash=([BitConverter]::ToString($algorithm.ComputeHash($stream))).Replace('-','').ToLowerInvariant()}finally{$stream.Dispose()}}
  finally{$algorithm.Dispose()}
  $env:PATH=$old;$restored=$env:PATH -ceq $old
  $record=@{protocol='winbooksplit.discovery-control';schema_version=1;scope='CONTROLLED_NATIVE';runtime=$runtime;path_restored=$restored;fixture_exe_sha256=$hash;fixture_source=$cs;compiler_host=$PSVersionTable.PSVersion.ToString()}
  Write-Output ('[DISCOVERY-CONTROL] '+($record|ConvertTo-Json -Depth 30 -Compress))
}finally{$env:PATH=$old}
'@
    $encoded=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($source))
    $run=Invoke-NativeAcceptanceProcessCase $Inbox @('-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-EncodedCommand',$encoded) $WorkRoot 120 -Environment @{
        WBS_KIT_CONTROL_ROOT=$control;WBS_KIT_CONTROL_APP=$ApplicationRoot;WBS_KIT_CONTROL_PYTHON=$SelectedPython}
    try {
    Write-NativeAcceptanceJson ([IO.Path]::Combine($WorkRoot,'control-discovery-process.json')) $run.probe
    Assert-NativeAcceptanceProcess $run.probe
    if($run.failure -or -not $run.environment_before_after_restored -or $run.probe.ExitCode -ne 0){throw 'Controlled legacy listing fallthrough failed.'}
    $record=Get-NativeAcceptanceFrame $run.probe.Stdout '[DISCOVERY-CONTROL] '
    $python=$record.runtime
    if($record.path_restored -isnot [bool] -or -not $record.path_restored -or $record.scope -cne 'CONTROLLED_NATIVE' -or
        -not [StringComparer]::OrdinalIgnoreCase.Equals($python.Path,$SelectedPython)){throw 'Controlled discovery did not bind the actual PATH candidate.'}
    $attempts=@($python.Attempts|Where-Object{$_.Source -ceq 'py_launcher_listing' -and $_.Path -ceq [IO.Path]::Combine($control,'bin\py.exe')})
    if($attempts.Count -ne 1 -or $attempts[0].Accepted -isnot [bool] -or $attempts[0].Accepted -or $null -eq $attempts[0].Probe -or
        $attempts[0].Probe.ExitCode -ne 2 -or $attempts[0].Probe.Stderr -cne "Unknown option: -0`r`n"){throw 'Controlled native launcher failure was not retained.'}
    Assert-NativeAcceptanceProcess $attempts[0].Probe
    $accepted=@($python.Attempts|Where-Object{$_.Accepted -eq $true -and $_.Source -ceq 'PATH' -and [StringComparer]::OrdinalIgnoreCase.Equals($_.Path,$SelectedPython)})
    if($accepted.Count -ne 1){throw 'Controlled native PATH candidate was not independently accepted.'}
    Assert-NativeAcceptanceProcess $accepted[0].Probe
    Write-NativeAcceptanceJson ([IO.Path]::Combine($WorkRoot,'control-discovery-local.json')) $record
    $algorithm=[Security.Cryptography.SHA256]::Create()
    try{$sourceHash=([BitConverter]::ToString($algorithm.ComputeHash([Text.Encoding]::UTF8.GetBytes($source)))).Replace('-','').ToLowerInvariant()}
    finally{$algorithm.Dispose()}
    return [ordered]@{id='failed_legacy_path_fallback';scope='CONTROLLED_NATIVE';status='PASS';code='failed_listing_then_path_accepted';
        probe_sha256=(Get-NativeAcceptanceFileHash ([IO.Path]::Combine($WorkRoot,'control-discovery-process.json')) 8388608).sha256;
        helper_sha256=$sourceHash;fixture_exe_sha256=$record.fixture_exe_sha256}
    } catch {
        $safe=$false;if($run.probe){try{Assert-NativeAcceptanceProcess $run.probe;$safe=$true}catch{$safe=$false}}
        $error=New-Object IO.IOException('Controlled native discovery failed; artifacts retained.',$_.Exception)
        $error.Data['ControlUnsafe']=(-not $safe);if($run.probe){$error.Data['Probe']=$run.probe};throw $error
    }
}

function Invoke-NativeAcceptance {
    [CmdletBinding()]
    param([Parameter(Mandatory=$true)][string]$KitRoot,[Parameter(Mandatory=$true)][string]$BindingsSha256,
        [string]$BarePythonPath,[string]$BaselinePythonPath,[string]$CalibrePath,[string]$Wheelhouse,[string]$PowerShell7,[switch]$ProvisionPrerequisites,
        [switch]$AllowWindows11DryRun,[ValidateSet('AGENT_LAUNCHED','USER_LAUNCHED_AUTOMATED')][string]$Execution='USER_LAUNCHED_AUTOMATED')
    $ErrorActionPreference='Stop'
    Initialize-WinBookSplitLogTypes;Initialize-WinBookSplitSupport;Initialize-WinBookSplitProcess
    $started=[DateTime]::UtcNow.ToString('o');$os=$bindings=$safeNative=$oracle=$runtimeSelection=$markedSummary=$run=$null;$failure=$failureStage=$null;$work=$null;$unsafe=$false
    $leases=New-Object 'Collections.Generic.List[IDisposable]';$stages=New-Object 'Collections.Generic.List[object]';$controls=New-Object 'Collections.Generic.List[object]'
    $stageIds=@('os','kit_identity','prerequisites','verify_extract','version','help','setup_install','setup_check','setup_noop','marked_setup_remedies','controls','prepare','native_cases','validate','redact','source_identity')
    $stage='os';$productionVerified=$false;$kitVerified=$false
    try {
        $root=Get-NativeAcceptanceOrdinaryPath $KitRoot -Directory
        $system=[Environment]::SystemDirectory;$inbox=Get-NativeAcceptanceOrdinaryPath ([IO.Path]::Combine($system,'WindowsPowerShell\v1.0\powershell.exe'))
        $cache=[IO.Path]::Combine([Environment]::GetFolderPath('UserProfile'),'.cache')
        if(-not [IO.Directory]::Exists($cache)){$null=New-NativeAcceptanceDirectory $cache}
        $null=Get-NativeAcceptanceOrdinaryPath $cache -Directory
        $ancestors=New-Object 'Collections.Generic.List[string]';$current=$cache
        while($current){$ancestors.Add($current);$current=[IO.Path]::GetDirectoryName($current)}
        for($i=$ancestors.Count-1;$i -ge 0;$i--){$leases.Add((New-Object WinBookSplitLogging.PathLease($ancestors[$i],$true)))}
        $work=New-NativeAcceptanceDirectory ([IO.Path]::Combine($cache,'WBS-Native-'+[Guid]::NewGuid().ToString('N')))
        $leases.Add((New-Object WinBookSplitLogging.PathLease($work,$true)))
        if (-not [StringComparer]::OrdinalIgnoreCase.Equals([IO.Path]::GetFullPath($PSHOME),[IO.Path]::GetDirectoryName($inbox))) {throw 'Central native kit requires the genuine inbox PowerShell5.1.'}
        $os=Get-NativeAcceptanceOs -AllowWindows11DryRun:$AllowWindows11DryRun
        if ($os.platform_scope -ceq 'BLOCKED') {throw 'The observed native OS is outside the requested acceptance scope.'}
        $stages.Add([ordered]@{id=$stage;status='PASS';code='completed';exit_code=0;probe_sha256=$null})
        $stage='kit_identity'
        $read=Read-NativeAcceptanceJson ([IO.Path]::Combine($root,'kit-bindings.json'))
        if ($BindingsSha256 -cnotmatch '^[0-9a-f]{64}$' -or $read.sha256 -cne $BindingsSha256) {throw 'Kit bindings differ from generated wrapper identity.'}
        $bindings=$read.value
        if ($bindings -isnot [Collections.IDictionary] -or $bindings.Count -ne 7 -or $bindings.protocol -cne 'winbooksplit.native-kit' -or
            -not(Test-SplitInteger $bindings.schema_version) -or $bindings.schema_version -ne 1 -or
            $bindings.helpers -isnot [Collections.IDictionary] -or $bindings.helpers.Count -ne 19 -or
            $bindings.fixtures -isnot [Collections.IDictionary] -or $bindings.fixtures.Count -ne 6) {throw 'Native kit bindings schema differs from its finite closure.'}
        if(-not(Test-NativeAcceptanceProductionIdentity $bindings.production $bindings.production)){throw 'Expected production pins are not a finite safe identity.'}
        $helperNames=@('engine/WinBookSplit.Logging.ps1','engine/WinBookSplit.Diagnostics.ps1','engine/WinBookSplit.Process.ps1','engine/WinBookSplit.Support.ps1',
            'engine/WinBookSplit.Outcomes.json','engine/winbooksplit_windows.py','tools/compatibility/Invoke-NativeAcceptance.ps1','tools/compatibility/kit_builder.py',
            'tools/compatibility/native_smoke.py','tools/release/frozen_assets.py','tests/cli/Probe-Help.ps1','tests/process/session_child.py',
            'engine/WinBookSplit.Runtime.ps1','engine/WinBookSplit.Compatibility.ps1','engine/WinBookSplit.Paths.ps1',
            'engine/WinBookSplit.Compatibility.json','engine/winbooksplit_compatibility.py',
            'tools/compatibility/Invoke-SetupRemedyRoundTrip.ps1','tests/compatibility/validate_setup_remedy.py')
        $fixtureNames=@('flat-6.pdf','nested-12.pdf','original-three-chapters.azw3','original-three-chapters.epub','simple-10.pdf','native-fixtures.json')
        $kit=$bindings.kit
        if($kit -isnot [Collections.IDictionary] -or $kit.Count -ne 6 -or $kit.helpers_payload_sha256 -isnot [string] -or
            $kit.helpers_payload_sha256 -cnotmatch '^[0-9a-f]{64}$' -or $kit.fixtures_payload_sha256 -isnot [string] -or
            $kit.fixtures_payload_sha256 -cnotmatch '^[0-9a-f]{64}$'){throw 'Invalid fixed kit provenance.'}
        if($kit.mode -ceq 'COMMITTED'){
            if($kit.source_commit -isnot [string] -or $kit.source_commit -cnotmatch '^[0-9a-f]{40}$' -or
                $kit.source_tree -isnot [string] -or $kit.source_tree -cnotmatch '^[0-9a-f]{40}$' -or $null -ne $kit.repository_observation){throw 'Invalid committed kit source identity.'}
        }elseif($kit.mode -ceq 'EXTERNAL_DRAFT_REHEARSAL'){
            if($null -ne $kit.source_commit -or $null -ne $kit.source_tree -or $kit.repository_observation -isnot [Collections.IDictionary] -or
                $kit.repository_observation.Count -ne 2 -or $kit.repository_observation.commit -cnotmatch '^[0-9a-f]{40}$' -or
                $kit.repository_observation.tree -cnotmatch '^[0-9a-f]{40}$'){throw 'External draft identity is not explicitly disclosed.'}
        }else{throw 'Unknown kit provenance mode.'}
        foreach($kind in @('helpers','fixtures')) {
            $names=if($kind -ceq 'helpers'){$helperNames}else{$fixtureNames}
            foreach($name in $names){
                if($name -cnotin @($bindings[$kind].Keys)){throw 'Missing finite kit member.'}
                $relative=if($kind -ceq 'helpers'){$name}else{'fixtures/'+$name};$path=[IO.Path]::Combine($root,$relative.Replace('/','\'))
                $expected=$bindings[$kind][$name]
                if($expected -isnot [Collections.IDictionary] -or $expected.Count -ne 2 -or -not(Test-SplitInteger $expected.bytes) -or
                    $expected.bytes -lt 1 -or $expected.bytes -gt 8388608 -or $expected.sha256 -isnot [string] -or $expected.sha256 -cnotmatch '^[0-9a-f]{64}$'){throw 'Invalid kit member identity.'}
                $actual=Get-NativeAcceptanceFileHash $path 8388608
                if($actual.bytes -ne $expected.bytes -or $actual.sha256 -cne $expected.sha256){throw 'Kit member identity mismatch.'}
                $leases.Add((New-Object WinBookSplitLogging.PathLease($path,$false,$false,$true)))
            }
        }
        foreach($name in $bindings.production.assets.Keys){
            $path=[IO.Path]::Combine($root,'assets',$name);$actual=Get-NativeAcceptanceFileHash $path 33554432;$expected=$bindings.production.assets[$name]
            if($actual.sha256 -cne $expected.sha256 -or $actual.bytes -ne $expected.bytes){throw 'An external production asset pin differs.'}
            $leases.Add((New-Object WinBookSplitLogging.PathLease($path,$false,$false,$true)))
        }
        $kitVerified=$true
        if($bindings.helpers['engine/WinBookSplit.Compatibility.json'].sha256-cne $bindings.production.policy_sha256){throw 'Authenticated discovery policy differs from exact production.'}
        $stages.Add([ordered]@{id=$stage;status='PASS';code='completed';exit_code=0;probe_sha256=$null})
        $stage='prerequisites';$downloads=New-NativeAcceptanceDirectory ([IO.Path]::Combine($work,'downloads'))
        $discoveryPolicy=Get-WinBookSplitCompatibilityPolicy -PolicyPath ([IO.Path]::Combine($root,'engine','WinBookSplit.Compatibility.json'))
        $owner=Find-WinBookSplitSetupBase -ApplicationRoot $root -Policy $discoveryPolicy -PythonPath $BarePythonPath
        $ownerUsable=$null-ne $owner -and $owner.Version -cmatch '\A3\.15\.[0-9]+\z'
        $runtimeSelection=[ordered]@{owner_version=$(if($owner){$owner.Version}else{$null});owner_regular315_usable=$ownerUsable;
            owner_selected=$ownerUsable;selected_version=$null;required_baseline_version='3.15.0';baseline_version=$null;baseline_kind='explicit3150_test_baseline';
            reason=$(if($ownerUsable){'usable_regular315_discovered'}elseif($owner){'supported_legacy_branch_is_not_315_kit_target'}else{'no_usable_regular315_discovered'})}
        if($ownerUsable){$BarePythonPath=$owner.Path}else{$BarePythonPath=$null}
        if([string]::IsNullOrEmpty($BaselinePythonPath) -and $ownerUsable -and $owner.Version-ceq '3.15.0'){
            # Read-only package probe: an already-installed pypdf cannot serve
            # the required bare missing-pypdf regression. Keep owner selection
            # primary, and provision a clean baseline only for that fixture.
            $packageAttempts=New-Object 'Collections.Generic.List[object]'
            $null=Test-WinBookSplitPythonCandidate $owner.Path 'explicit' $root $packageAttempts 10
            if($packageAttempts.Count-eq 1 -and $packageAttempts[0].Reason-ceq 'missing_pypdf'){
                $BaselinePythonPath=$owner.Path
                $runtimeSelection.baseline_kind='existing3150_without_pypdf'
            }else{
                $missingRoot=New-NativeAcceptanceDirectory ([IO.Path]::Combine($work,'owned-missing-package-baseline'))
                $leases.Add((New-Object WinBookSplitLogging.PathLease($missingRoot,$true)))
                $run=Invoke-NativeAcceptanceProcessCase $owner.Path @('-I','-B','-m','venv','--without-pip',$missingRoot) $work 120
                Write-NativeAcceptanceJson ([IO.Path]::Combine($work,'baseline-venv-process.json')) $run.probe
                Assert-NativeAcceptanceProcess $run.probe
                if($run.failure -or $run.probe.ExitCode-ne 0){throw 'Owned clean315.0 missing-package baseline creation failed.'}
                $BaselinePythonPath=[IO.Path]::Combine($missingRoot,'Scripts','python.exe')
                $runtimeSelection.baseline_kind='owned3150_venv_without_pip_from_existing3150'
            }
        }
        if([string]::IsNullOrEmpty($BaselinePythonPath)){
            if(-not $ProvisionPrerequisites){throw 'The exact315.0 private test baseline or authorized private provisioning is required.'}
            $zip=Get-NativeAcceptanceDownload $bindings.prerequisites.python ([IO.Path]::Combine($downloads,$bindings.prerequisites.python.name))
            $BaselinePythonPath=Expand-NativeAcceptancePython $zip ([IO.Path]::Combine($work,'python315-baseline')) $bindings.prerequisites.python
            $runtimeSelection.baseline_kind='reviewed_private3150_archive'
        }else{$BaselinePythonPath=Get-NativeAcceptanceOrdinaryPath $BaselinePythonPath}
        $baselineAttempts=New-Object 'Collections.Generic.List[object]'
        $baseline=Test-WinBookSplitPythonCandidate $BaselinePythonPath 'explicit' $root $baselineAttempts 10 -InterpreterOnly
        if($null-eq $baseline -or $baseline.Version-cne '3.15.0'){throw 'Required actual regular315.0 baseline differs; no version substitution permitted.'}
        $runtimeSelection.baseline_version=$baseline.Version
        if([string]::IsNullOrEmpty($BarePythonPath)){$BarePythonPath=$BaselinePythonPath}
        $chosenAttempts=New-Object 'Collections.Generic.List[object]'
        $chosen=Test-WinBookSplitPythonCandidate $BarePythonPath 'explicit' $root $chosenAttempts 10 -InterpreterOnly
        if($null-eq $chosen -or $chosen.Version-cnotmatch '\A3\.15\.[0-9]+\z'){throw 'Native kit selected runtime is not a usable regular315 branch.'}
        $runtimeSelection.selected_version=$chosen.Version
        foreach($kind in @('selected_identity','baseline_identity')){
            $candidate=if($kind-ceq 'selected_identity'){$chosen}else{$baseline}
            $details=$candidate.Details
            $identity=[ordered]@{source=$(if($kind-ceq 'selected_identity' -and $ownerUsable){$owner.Source}else{$candidate.Source});
                executable_sha256=(Get-NativeAcceptanceFileHash $candidate.Path 8388608).sha256;version=$candidate.Version}
            foreach($field in @('implementation','platform','machine','bits','gil_disabled','releaselevel','serial','isolated','dont_write_bytecode','policy_sha256')){$identity[$field]=$details.$field}
            $runtimeSelection[$kind]=$identity
        }
        Write-NativeAcceptanceJson ([IO.Path]::Combine($work,'runtime-selection-local.json')) @{selection=$runtimeSelection;selected=$chosen;baseline=$baseline}
        if([string]::IsNullOrEmpty($Wheelhouse)){
            if(-not $ProvisionPrerequisites){throw 'An exact wheelhouse or authorized private provisioning is required.'}
            $Wheelhouse=New-NativeAcceptanceDirectory ([IO.Path]::Combine($work,'wheelhouse'))
            $null=Get-NativeAcceptanceDownload $bindings.prerequisites.pypdf ([IO.Path]::Combine($Wheelhouse,$bindings.prerequisites.pypdf.name))
        }else{$Wheelhouse=Get-NativeAcceptanceOrdinaryPath $Wheelhouse -Directory}
        $expectedWheel=$bindings.prerequisites.pypdf;$actualWheel=Get-NativeAcceptanceFileHash ([IO.Path]::Combine($Wheelhouse,$expectedWheel.name)) 1048576
        if($actualWheel.bytes -ne $expectedWheel.bytes -or $actualWheel.sha256 -cne $expectedWheel.sha256){throw 'Offline wheel differs from reviewed source.'}
        if([string]::IsNullOrEmpty($CalibrePath)){
            if(-not $ProvisionPrerequisites){throw 'An exact converter or authorized private provisioning is required.'}
            $installer=Get-NativeAcceptanceDownload $bindings.prerequisites.calibre ([IO.Path]::Combine($downloads,$bindings.prerequisites.calibre.name))
            $leases.Add((New-Object WinBookSplitLogging.PathLease($installer,$false,$false,$true)))
            $actualInstaller=Get-NativeAcceptanceFileHash $installer 268435456
            if($actualInstaller.sha256 -cne $bindings.prerequisites.calibre.sha256 -or $actualInstaller.bytes -ne $bindings.prerequisites.calibre.bytes){throw 'Held converter installer differs from its primary-source pin.'}
            $destination=New-NativeAcceptancePortableDirectory ([Environment]::GetFolderPath('UserProfile')) $leases
            Write-NativeAcceptanceJson ([IO.Path]::Combine($work,'calibre-target-local.json')) ([ordered]@{
                protocol='winbooksplit.native-portable-target';schema_version=1;destination=$destination;
                final_portable_root=[IO.Path]::Combine($destination,'Calibre Portable');
                final_path_utf16_units=[IO.Path]::Combine($destination,'Calibre Portable').Length;
                maximum_utf16_units=58;installer_sha256=$actualInstaller.sha256})
            $run=Invoke-NativeAcceptanceProcessCase $installer @($destination) $work 300
            Write-NativeAcceptanceJson ([IO.Path]::Combine($work,'calibre-provision.json')) $run.probe
            Assert-NativeAcceptanceProcess $run.probe
            if($run.failure -or -not $run.environment_before_after_restored -or $run.probe.ExitCode -ne 0){throw 'Private converter provisioning failed.'}
            $candidates=@([IO.Path]::Combine($destination,'Calibre Portable\Calibre\ebook-convert.exe'),[IO.Path]::Combine($destination,'Calibre\ebook-convert.exe'))
            $found=@($candidates|Where-Object{[IO.File]::Exists($_)})
            if($found.Count -ne 1){throw 'Private converter destination is ambiguous.'};$CalibrePath=$found[0]
        }
        $CalibrePath=Get-NativeAcceptanceOrdinaryPath $CalibrePath
        $fixtureCatalog=(Read-NativeAcceptanceJson ([IO.Path]::Combine($root,'fixtures\native-fixtures.json'))).value
        if((Get-NativeAcceptanceFileHash $CalibrePath 8388608).sha256 -cne $fixtureCatalog.calibre.sha256){throw 'Converter differs from the reviewed genuine fixture tool.'}
        $stages.Add([ordered]@{id=$stage;status='PASS';code='completed';exit_code=0;probe_sha256=$null})
        $stage='verify_extract';$app=[IO.Path]::Combine($work,'application');$extraction=[IO.Path]::Combine($work,'extraction.json')
        $kitBuilder=[IO.Path]::Combine($root,'tools\compatibility\kit_builder.py');$smoke=[IO.Path]::Combine($root,'tools\compatibility\native_smoke.py')
        $argv=@('-I','-B',$kitBuilder,'verify-extract','--kit-root',$root,'--bindings-sha256',$BindingsSha256,'--destination',$app,'--output',$extraction)
        $run=Invoke-NativeAcceptanceProcessCase $BarePythonPath $argv $work 120
        Write-NativeAcceptanceJson ([IO.Path]::Combine($work,'verify-extract-process.json')) $run.probe
        Assert-NativeAcceptanceProcess $run.probe
        if($run.failure -or $run.probe.ExitCode -ne 0){throw 'Exact production asset verification/extraction failed.'}
        $verified=(Read-NativeAcceptanceJson $extraction 131072).value
        if($verified.protocol -cne 'winbooksplit.native-extraction' -or $verified.success -isnot [bool] -or -not $verified.success -or
            -not(Test-NativeAcceptanceProductionIdentity $verified.production $bindings.production)){throw 'Extraction identity differs from exact bindings.'}
        $productionVerified=$true
        $stages.Add([ordered]@{id=$stage;status='PASS';code='completed';exit_code=0;probe_sha256=(Get-NativeAcceptanceFileHash ([IO.Path]::Combine($work,'verify-extract-process.json')) 8388608).sha256})
        $main=[IO.Path]::Combine($app,'WinBookSplit.ps1');$setup=[IO.Path]::Combine($app,'Setup-WinBookSplit.ps1')
        $baseArgs=@('-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File')
        foreach($stage in @('version','help','setup_install','setup_check','setup_noop')){
            if($stage -ceq 'version'){$argv=$baseArgs+@($main,'-Version');$env=@{}}
            elseif($stage -ceq 'help'){$argv=$baseArgs+@([IO.Path]::Combine($root,'tests\cli\Probe-Help.ps1'));$env=@{WBS_CLI_HELP_APP=$main}}
            else{$argv=$baseArgs+@($setup,'-PythonPath',$BarePythonPath,'-Wheelhouse',$Wheelhouse,'-NonInteractive');$env=@{};if($stage -ceq 'setup_check'){$argv+= '-Check'}}
            if($stage -ceq 'setup_noop'){$before=Get-NativeAcceptanceEnvironmentSnapshot ([IO.Path]::Combine($app,'.venv'))}
            $run=Invoke-NativeAcceptanceProcessCase $inbox $argv $work 300 -Environment $env
            $rawPath=[IO.Path]::Combine($work,$stage+'-process.json');Write-NativeAcceptanceJson $rawPath $run.probe
            Assert-NativeAcceptanceProcess $run.probe
            if($run.failure -or -not $run.environment_before_after_restored -or $run.probe.ExitCode -ne 0){throw 'A mandatory native package stage failed.'}
            if($stage -ceq 'version') {if($run.probe.Stdout.Trim() -cne ('WinBookSplit '+$bindings.production.version)){throw 'Package Version differs from exact release identity.'}}
            elseif($stage -ceq 'help'){$help=[WinBookSplitSupport.StrictJson]::Parse($run.probe.Stdout);foreach($parameter in @('InputFile','PythonPath','Preview','Mode','NoPause')){if($parameter -cnotin $help.parameters){throw 'Actual native Get-Help omitted shipped parameters.'}}}
            else{
                $frame=Get-NativeAcceptanceFrame $run.probe.Stdout '[SETUP] '
                $expected=if($stage -ceq 'setup_install'){'setup_complete'}else{'setup_ready'}
                if($frame.code -cne $expected -or $frame.status -cne 'success' -or $frame.exit_code -ne 0 -or
                    $frame.changed -isnot [bool] -or $frame.changed -ne ($stage -ceq 'setup_install') -or
                    -not [StringComparer]::OrdinalIgnoreCase.Equals($frame.selected_python,[IO.Path]::Combine($app,'.venv\Scripts\python.exe'))){throw 'Shipped setup result differs from final-path install/check/no-op contract.'}
                if($stage -ceq 'setup_noop' -and $before -cne (Get-NativeAcceptanceEnvironmentSnapshot ([IO.Path]::Combine($app,'.venv')))){throw 'Healthy setup no-op changed environment bytes or timestamps.'}
            }
            $stages.Add([ordered]@{id=$stage;status='PASS';code='completed';exit_code=0;probe_sha256=(Get-NativeAcceptanceFileHash $rawPath 8388608).sha256})
        }
        $python=[IO.Path]::Combine($app,'.venv\Scripts\python.exe');$stage='marked_setup_remedies'
        $sourceMap=@{};foreach($name in $helperNames){$sourceMap[$name]=$bindings.helpers[$name].sha256}
        $applicationMap=@{};foreach($name in $verified.files.Keys){$applicationMap[$name]=$verified.files[$name].sha256}
        $remedyBinding=[IO.Path]::Combine($work,'remedy-source-binding.json')
        Write-NativeAcceptanceJson $remedyBinding @{protocol='winbooksplit.native-kit-remedy-binding';schema_version=1;
            source_manifest=$sourceMap;application_manifest=$applicationMap;production=$bindings.production;kit=$bindings.kit}
        $bindingHash=(Get-NativeAcceptanceFileHash $remedyBinding 8388608).sha256
        $argv=@('-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',[IO.Path]::Combine($root,'tools','compatibility','Invoke-SetupRemedyRoundTrip.ps1'),
            '-RepositoryRoot',$root,'-ApplicationSourceRoot',$app,'-WorkParent',$work,'-SourceBinding',$remedyBinding,
            '-SourceBindingSha256',$bindingHash,'-Python315',$BaselinePythonPath,'-AuthoredPdf',[IO.Path]::Combine($root,'fixtures','simple-10.pdf'),
            '-AuthoredPdfSha256',$bindings.fixtures['simple-10.pdf'].sha256,'-IncludeSecurityControls','-Wheelhouse',$Wheelhouse)
        if(-not[string]::IsNullOrEmpty($PowerShell7)){$argv+=@('-PowerShell7',$PowerShell7)}
        $run=Invoke-NativeAcceptanceProcessCase $inbox $argv $work 600
        $markedProcess=[IO.Path]::Combine($work,'marked-setup-remedies-process.json');Write-NativeAcceptanceJson $markedProcess $run.probe
        Assert-NativeAcceptanceProcess $run.probe
        $paths=@($run.probe.Stdout.Split("`n")|Where-Object{$_.StartsWith('RECEIPT ')})
        if($paths.Count-ne 1){throw 'Exact extracted-package remedy receipt was not emitted once.'}
        $markedPath=$paths[0].Substring(8).TrimEnd([char]13)
        if(-not [IO.Path]::GetFullPath($markedPath).StartsWith($work+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Remedy receipt escaped owned kit work.'}
        $marked=(Read-NativeAcceptanceJson $markedPath 8388608).value
        if($run.failure -or $run.probe.ExitCode-ne 0 -or $marked.success-isnot [bool] -or -not $marked.success -or
            $marked.source_files_unchanged-isnot [bool] -or -not $marked.source_files_unchanged){throw 'Exact extracted-package marked BAT/Action/repair/PDF regressions failed.'}
        $markedSummary=[ordered]@{status='PASS';required_runtime_version='3.15.0';exact_asset_bound=$true;
            Action_roundtrips=$marked.cases.Count;marked_BAT_routes=@($marked.security_controls|Where-Object{$_.scope-ceq 'ACTUAL_MARKED_APPLICATION_BAT'}).Count;
            mocked_policy_refusals=@($marked.security_controls|Where-Object{$_.scope-like 'MOCKED_POLICY_QUERY*'}).Count;
            raw_receipt_sha256=(Get-NativeAcceptanceFileHash $markedPath 8388608).sha256;
            probe_sha256=(Get-NativeAcceptanceFileHash $markedProcess 8388608).sha256;actual_inbox_policy_snapshots=$marked.driver_policies_before_after;
            fixture='Authored broken venv, not an owner legacy runtime; marks assigned by automation; all source/ADS/PDF checks verified'}
        $safeRoutes=New-Object 'Collections.Generic.List[object]'
        foreach($route in @($marked.cases)+@($marked.security_controls)){
            if($route.status-cne 'PASS'){throw 'Marked route did not actually pass.'}
            $safeRoute=[ordered]@{id=$route.id;status=$route.status;kind=$(if('simulation'-cin @($route.Keys)){'MOCKED_POLICY_QUERY'}elseif('environment_fixture'-cin @($route.Keys)){'ACTION'}else{'BAT'});
                fixture=$(if('environment_fixture'-cin @($route.Keys)){$route.environment_fixture}elseif('simulation'-cin @($route.Keys)){'MOCKED_POLICY_QUERY'}else{'AUTHORED_BROKEN_VENV'});
                exits=@();child_hosts=@();pdf=$null}
            if('simulation'-cin @($route.Keys)){
                $safeRoute.exits=@(@{stage='simulation';exit_code=$route.probe.ExitCode})
                $safeRoute.simulated_positive_launch_refused=$route.simulation.positive_launch_refused
            }else{
                foreach($step in $route.steps){
                    $safeRoute.exits+=@(@{stage=$step.stage;exit_code=$step.probe.ExitCode})
                    $frames=@($step.probe.Stdout.Split("`n")|Where-Object{$_.StartsWith('[SETUP] ')})
                    if($frames.Count-eq 1){
                        $setupRecord=Get-NativeAcceptanceFrame $step.probe.Stdout '[SETUP] '
                        $safeRoute.child_hosts+=@(@{stage=$step.stage;executable_status=$setupRecord.host.executable_status;
                            version=$setupRecord.host.version;edition=$setupRecord.host.edition;execution_policy=$setupRecord.host.execution_policy;
                            selected_version=$setupRecord.selected_version;pypdf_version=$setupRecord.pypdf_version})
                    }
                }
                if($route.pdf_oracle.success-isnot [bool] -or -not $route.pdf_oracle.success){throw 'Marked PDF oracle did not prove success.'}
                $safeRoute.pdf=@{success=$true;runtime_version=$route.pdf_oracle.console.run_manifest.runtime_versions.python;
                    written_count=$route.pdf_oracle.written_count;physical_pages=$route.pdf_oracle.physical_pages;
                    source_sha256=$route.original_source_sha256}
            }
            $safeRoutes.Add($safeRoute)
        }
        $markedSummary.routes=$safeRoutes.ToArray()
        $stages.Add([ordered]@{id=$stage;status='PASS';code='completed';exit_code=0;probe_sha256=$markedSummary.probe_sha256})
        $stage='controls'
        foreach($kind in @('timeout','cancel')){$controls.Add((Invoke-NativeAcceptanceTransportControl $python ([IO.Path]::Combine($root,'tests\process\session_child.py')) $work $kind))}
        # The second exact extraction has no .venv, preserving authoritative
        # application semantics while the actual full runtime is on PATH.
        $controlApp=[IO.Path]::Combine($work,'control-application');$controlExtraction=[IO.Path]::Combine($work,'control-extraction.json')
        $run=Invoke-NativeAcceptanceProcessCase $BarePythonPath @('-I','-B',$kitBuilder,'verify-extract','--kit-root',$root,'--bindings-sha256',$BindingsSha256,'--destination',$controlApp,'--output',$controlExtraction) $work 120
        Assert-NativeAcceptanceProcess $run.probe;if($run.failure -or $run.probe.ExitCode -ne 0){throw 'Controlled exact package extraction failed.'}
        $controls.Add((Invoke-NativeAcceptanceDiscoveryControl $inbox $controlApp $python $work))
        $stages.Add([ordered]@{id=$stage;status='PASS';code='completed';exit_code=0;probe_sha256=$null})
        $stage='prepare';$fixtureWork=New-NativeAcceptanceDirectory ([IO.Path]::Combine($work,'cases'))
        $documents=Get-NativeAcceptanceOrdinaryPath ([Environment]::GetFolderPath('MyDocuments')) -Directory
        $planPath=[IO.Path]::Combine($work,'plan.json')
        $argv=@('-I','-B',$smoke,'prepare','--application-root',$app,'--fixture-root',[IO.Path]::Combine($root,'fixtures'),
            '--work-root',$fixtureWork,'--python-path',$python,'--bare-python-path',$BaselinePythonPath,'--calibre-path',$CalibrePath,
            '--documents-root',$documents,'--policy-sha256',$bindings.production.policy_sha256,'--output',$planPath)
        $run=Invoke-NativeAcceptanceProcessCase $python $argv $work 120
        Write-NativeAcceptanceJson ([IO.Path]::Combine($work,'prepare-process.json')) $run.probe;Assert-NativeAcceptanceProcess $run.probe
        if($run.failure -or $run.probe.ExitCode -ne 0){throw 'Native authored fixture preparation failed.'}
        $read=Read-NativeAcceptanceJson $planPath 8388608;$plan=$read.value;$planHash=$read.sha256
        $ids=@('ref-epub','ref-azw3','manual','manual-first','auto1','auto2','repeat','preview','invalid-manual','no-bookmarks','cancel-plan','missing-python','missing-pypdf','epub-discard','epub-keep','azw3-discard','azw3-keep','bat-one','bat-two')
        if($plan.protocol -cne 'winbooksplit.native-plan' -or $plan.cases.Count -ne $ids.Count -or -not [StringComparer]::OrdinalIgnoreCase.Equals($plan.application_root,$app) -or
            -not [StringComparer]::OrdinalIgnoreCase.Equals($plan.work_root,$fixtureWork) -or $plan.policy_sha256 -cne $bindings.production.policy_sha256){throw 'Finite native case plan differs from concrete bindings.'}
        $stages.Add([ordered]@{id=$stage;status='PASS';code='completed';exit_code=0;probe_sha256=(Get-NativeAcceptanceFileHash ([IO.Path]::Combine($work,'prepare-process.json')) 8388608).sha256})
        $stage='native_cases';$rows=New-Object 'Collections.Generic.List[object]'
        for($i=0;$i -lt $ids.Count;$i++){
            $case=$plan.cases[$i]
            if($case.id -cne $ids[$i] -or -not(Test-SplitInteger $case.timeout_seconds) -or $case.timeout_seconds -lt 1 -or $case.timeout_seconds -gt 300){throw 'Finite native recipe is malformed.'}
            if($unsafe){$rows.Add([ordered]@{id=$case.id;status='NOT_RUN';probe=$null;command=@();cwd=$case.cwd;answers=@();environment_before_after_restored=$false});continue}
            Assert-NativeAcceptanceCaseEnvironment $case $fixtureWork $app
            $exe=if($case.exe_selector -ceq 'calibre'){$CalibrePath}elseif($case.exe_selector -ceq 'inbox_ps'){$inbox}elseif($case.exe_selector -ceq 'cmd'){[IO.Path]::Combine($system,'cmd.exe')}else{throw 'Unknown finite native executable selector.'}
            $run=Invoke-NativeAcceptanceProcessCase $exe @($case.arguments) $case.cwd $case.timeout_seconds @($case.answers) $case.environment
            $rows.Add([ordered]@{id=$case.id;status='COMPLETED';probe=$run.probe;command=$run.command;cwd=$run.cwd;answers=$run.answers;environment_before_after_restored=$run.environment_before_after_restored})
            try{Assert-NativeAcceptanceProcess $run.probe;if($run.failure -or -not $run.environment_before_after_restored){throw 'Native recipe environment/restoration failed.'}}
            catch{$unsafe=$true;$failure=$_.Exception.Message}
        }
        $processPath=[IO.Path]::Combine($work,'process-report.json')
        Write-NativeAcceptanceJson $processPath ([ordered]@{protocol='winbooksplit.native-process';schema_version=1;plan_sha256=$planHash;cases=$rows.ToArray()})
        if($unsafe){throw 'Native process safety proof failed; all local artifacts were preserved.'}
        $stages.Add([ordered]@{id=$stage;status='PASS';code='completed';exit_code=0;probe_sha256=(Get-NativeAcceptanceFileHash $processPath 8388608).sha256})
        $stage='validate';$oraclePath=[IO.Path]::Combine($work,'oracle.json')
        $run=Invoke-NativeAcceptanceProcessCase $python @('-I','-B',$smoke,'validate','--plan',$planPath,'--process-report',$processPath,'--output',$oraclePath) $work 120
        Write-NativeAcceptanceJson ([IO.Path]::Combine($work,'validate-process.json')) $run.probe;Assert-NativeAcceptanceProcess $run.probe
        if($run.failure -or -not [IO.File]::Exists($oraclePath)){throw 'Native physical/console oracle did not produce a bounded result.'}
        $oracle=(Read-NativeAcceptanceJson $oraclePath 8388608).value
        if($oracle.success -isnot [bool] -or $run.probe.ExitCode -ne $(if($oracle.success){0}else{1})){throw 'Native oracle exit and typed result disagree.'}
        if(-not $oracle.success){$failure='Native application acceptance oracle failed.';$failureStage=$stage}
        $stages.Add([ordered]@{id=$stage;status=$(if($oracle.success){'PASS'}else{'FAIL'});code=$(if($oracle.success){'completed'}else{'oracle_failed'});
            exit_code=$run.probe.ExitCode;probe_sha256=(Get-NativeAcceptanceFileHash ([IO.Path]::Combine($work,'validate-process.json')) 8388608).sha256})
        $stage='redact';$raw=[IO.Path]::Combine($work,'oracle-os.json');$safe=[IO.Path]::Combine($work,'sanitized-native.json')
        Write-NativeAcceptanceJson $raw ([ordered]@{oracle=$oracle;os=$os;execution=$Execution;runtime_selection=$runtimeSelection;marked_setup_regressions=$markedSummary})
        $run=Invoke-NativeAcceptanceProcessCase $python @('-I','-B',$smoke,'export','--raw-report',$raw,'--output',$safe) $work 60
        Write-NativeAcceptanceJson ([IO.Path]::Combine($work,'export-process.json')) $run.probe;Assert-NativeAcceptanceProcess $run.probe
        if($run.failure -or $run.probe.ExitCode -ne $(if($oracle.success){0}else{1})){throw 'Fixed native redaction failed.'}
        $safeNative=(Read-NativeAcceptanceJson $safe 8388608).value
        $runtimeSelection=$safeNative.runtime_selection;$markedSummary=$safeNative.marked_setup_regressions
        $stages.Add([ordered]@{id=$stage;status='PASS';code='completed';exit_code=$run.probe.ExitCode;probe_sha256=(Get-NativeAcceptanceFileHash ([IO.Path]::Combine($work,'export-process.json')) 8388608).sha256})
        $stage='source_identity'
        foreach($name in $helperNames){if((Get-NativeAcceptanceFileHash ([IO.Path]::Combine($root,$name.Replace('/','\'))) 8388608).sha256 -cne $bindings.helpers[$name].sha256){throw 'Frozen native helper source changed.'}}
        foreach($held in $leases){$held.AssertUnchanged()}
        $stages.Add([ordered]@{id=$stage;status='PASS';code='completed';exit_code=0;probe_sha256=$null})
    }catch{
        $primary=$_.Exception;$failure=$primary.Message;$failureStage=$stage
        if($primary.Data.Contains('FatalDiscovery') -and $primary.Data['FatalDiscovery']){$unsafe=$true}
        if($primary.Data.Contains('Probe') -and $null-ne $primary.Data['Probe']){try{Assert-NativeAcceptanceProcess $primary.Data['Probe']}catch{$unsafe=$true}}
        if($primary.Data.Contains('ControlUnsafe')){$unsafe=$unsafe -or [bool]$primary.Data['ControlUnsafe']}
        elseif($run -and $run.probe){try{Assert-NativeAcceptanceProcess $run.probe}catch{$unsafe=$true}}
        $blocked=$primary.Data.Contains('NativeBlocked') -and $primary.Data['NativeBlocked'] -ceq 'portable_target_too_long'
        $stages.Add([ordered]@{id=$stage;status=$(if($blocked){'BLOCKED'}else{'FAIL'});
            code=$(if($blocked){'portable_target_too_long'}else{'stage_failed'});exit_code=$null;probe_sha256=$null})
    }
    foreach($id in $stageIds){if(@($stages|Where-Object{$_.id -ceq $id}).Count -eq 0){$stages.Add([ordered]@{id=$id;status='NOT_RUN';code='prerequisite_failed';exit_code=$null;probe_sha256=$null})}}
    $success=$null -eq $failure -and -not $unsafe -and @($stages|Where-Object{$_.status -cne 'PASS'}).Count -eq 0 -and $null -ne $safeNative
    $safeOs=if($safeNative){$safeNative.os}else{ConvertTo-NativeAcceptanceSafeOs $os}
    $safeOracle=if($safeNative){$safeNative.oracle}else{$null}
    $safeCases=if($safeOracle){$safeOracle.cases}else{@('ref-epub','ref-azw3','manual','manual-first','auto1','auto2','repeat','preview','invalid-manual','no-bookmarks','cancel-plan','missing-python','missing-pypdf','epub-discard','epub-keep','azw3-discard','azw3-keep','bat-one','bat-two')|ForEach-Object{[ordered]@{id=$_;status='NOT_RUN';exit_code=$null;error_code='prerequisite_failed'}}}
    $identities=if($kitVerified){[ordered]@{production=$(if($productionVerified){$bindings.production}else{$null});kit=$bindings.kit;
        bindings_sha256=$BindingsSha256;helpers=$bindings.helpers;fixtures=$bindings.fixtures}}else{$null}
    $result=[ordered]@{protocol='winbooksplit.native-acceptance';schema_version=1;evidence_kind='NATIVE_AUTOMATED';execution=$Execution;
        started_utc=$started;finished_utc=[DateTime]::UtcNow.ToString('o');success=$success;os=$safeOs;oracle=$safeOracle;cases=@($safeCases);identities=$identities;
        expected_production=$(if($kitVerified){$bindings.production}else{$null});runtime_selection=$(if($kitVerified -and $runtimeSelection){$runtimeSelection}else{$null});marked_setup_regressions=$markedSummary;stages=$stages.ToArray();controls=$controls.ToArray();cleanup_safe=(-not $unsafe)}
    try{if($work){
        if($failure){Write-NativeAcceptanceJson ([IO.Path]::Combine($work,'local-failure.json')) ([ordered]@{stage=$failureStage;message=$failure;unsafe=$unsafe;os=$os})}
        $resultPath=[IO.Path]::Combine($work,'native-result.json');Write-NativeAcceptanceJson $resultPath $result
        Add-Type -AssemblyName System.IO.Compression -ErrorAction Stop
        $file=[IO.File]::Open([IO.Path]::Combine($work,'native-result.zip'),[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::Read);$zip=$null
        try{$zip=New-Object IO.Compression.ZipArchive($file,[IO.Compression.ZipArchiveMode]::Create,$false);$entry=$zip.CreateEntry('native-result.json',[IO.Compression.CompressionLevel]::NoCompression);$entryStream=$entry.Open();try{$bytes=[IO.File]::ReadAllBytes($resultPath);$entryStream.Write($bytes,0,$bytes.Length)}finally{$entryStream.Dispose()}}finally{if($zip){$zip.Dispose()};$file.Dispose()}
        Write-Host ('Return only: '+[IO.Path]::Combine($work,'native-result.zip'))
        Write-Host ('Raw local artifacts retained: '+$work)
    }else{Write-Host 'Native acceptance was blocked before a fresh work directory could be reserved; no pass recorded.'}}
    finally{for($i=$leases.Count-1;$i -ge 0;$i--){$leases[$i].Dispose()}}
    return [pscustomobject]$result
}
