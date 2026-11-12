# Explicit/consented portable preparation. Dot-sourcing has no side effects.
function New-WinBookSplitCalibreFailure {
    param([string]$Code, [string]$Message)
    $failure = [InvalidOperationException]::new($Message)
    $failure.Data['Code'] = $Code
    return $failure
}

function Get-WinBookSplitCalibreRoot {
    $local = $env:LOCALAPPDATA
    if ([string]::IsNullOrWhiteSpace($local)) {
        $local = [Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)
    }
    $local = Get-WinBookSplitTrustedPath -Path $local -Directory
    if ($local -notmatch '^[A-Za-z]:\\') { throw 'Managed Calibre needs a local per-user application-data folder.' }
    return [IO.Path]::Combine($local, 'WinBookSplit', 'Calibre')
}

function Invoke-WinBookSplitCalibreWorker {
    param([string]$ApplicationRoot, [string]$PythonPath, [ValidateSet('find','prepare')][string]$Action,
          [string]$InstallerPath)
    $helper = Get-WinBookSplitTrustedPath -Path ([IO.Path]::Combine($ApplicationRoot, 'engine', 'winbooksplit_calibre.py'))
    $root = Get-WinBookSplitCalibreRoot
    $arguments = @('-I','-B','-X','utf8',$helper,$Action,$root)
    if (-not [string]::IsNullOrWhiteSpace($InstallerPath)) {
        $arguments += @('--offline', (Get-WinBookSplitTrustedPath -Path $InstallerPath))
    }
    $deadline = 60
    if ($Action -ceq 'prepare') { $deadline = 480 }
    $probe = Invoke-WinBookSplitProcess -Path $PythonPath -Arguments $arguments -WorkingDirectory $ApplicationRoot -TimeoutSeconds $deadline
    Assert-WinBookSplitDiscoveryStopped $probe
    Assert-WinBookSplitProbeNotInterrupted $probe
    $badResult = 'The Calibre preparation helper did not return a valid complete result.'
    if ($probe.StartError -or $probe.StreamError -or $probe.StdoutTruncated -or $probe.StderrTruncated -or
        -not $probe.JobAssigned -or -not $probe.StreamsComplete -or $probe.ExitCode -cnotin @(0,3)) {
        throw (New-WinBookSplitCalibreFailure 'converter_setup_failed' $badResult)
    }
    try { $record = ConvertFrom-Json -InputObject $probe.Stdout -ErrorAction Stop }
    catch { throw (New-WinBookSplitCalibreFailure 'converter_setup_failed' $badResult) }
    if ($null -eq $record -or $record.PSObject.Properties.Name -cnotcontains 'code' -or $record.code -isnot [string]) {
        throw (New-WinBookSplitCalibreFailure 'converter_setup_failed' $badResult)
    }
    if ($probe.ExitCode -eq 3) {
        $code = $record.code
        if ($code -cnotin @('converter_download_failed','converter_integrity_failed','converter_managed_invalid','converter_setup_failed') -or
            $record.PSObject.Properties.Name -cnotcontains 'message' -or $record.message -isnot [string]) {
            throw (New-WinBookSplitCalibreFailure 'converter_setup_failed' $badResult)
        }
        throw (New-WinBookSplitCalibreFailure $code $record.message)
    }
    if ($record.code -cnotin @('ready','absent') -or
        @('path','version','changed' | Where-Object { $record.PSObject.Properties.Name -cnotcontains $_ }).Count -ne 0 -or
        $record.changed -isnot [bool] -or ($Action -ceq 'find' -and $record.changed)) {
        throw (New-WinBookSplitCalibreFailure 'converter_setup_failed' $badResult)
    }
    if ($record.code -ceq 'ready') {
        if ($record.path -isnot [string] -or $record.version -isnot [string] -or
            -not $record.path.StartsWith($root + '\', [StringComparison]::OrdinalIgnoreCase)) {
            throw (New-WinBookSplitCalibreFailure 'converter_setup_failed' $badResult)
        }
        $record.path = Get-WinBookSplitTrustedPath -Path $record.path -Automatic
    }
    elseif ($Action -ceq 'prepare' -or $null -ne $record.path -or $null -ne $record.version) {
        throw (New-WinBookSplitCalibreFailure 'converter_setup_failed' $badResult)
    }
    return $record
}

function Get-WinBookSplitManagedConverterPath {
    param([string]$ApplicationRoot, [string]$PythonPath)
    return (Invoke-WinBookSplitCalibreWorker -ApplicationRoot $ApplicationRoot -PythonPath $PythonPath -Action find).path
}

function Invoke-WinBookSplitCalibreSetup {
    param([string]$ApplicationRoot, [string]$PythonPath, [string]$InstallerPath, [hashtable]$Discovery)
    $record = Invoke-WinBookSplitCalibreWorker -ApplicationRoot $ApplicationRoot -PythonPath $PythonPath -Action prepare -InstallerPath $InstallerPath
    $selected = Resolve-WinBookSplitConverter @Discovery
    if ($selected.Source -ceq 'managed' -and
        (-not $record.path.Equals($selected.Path, [StringComparison]::OrdinalIgnoreCase) -or $record.version -cne $selected.Version)) {
        throw (New-WinBookSplitCalibreFailure 'converter_setup_failed' 'Calibre was prepared but its verified path or version changed.')
    }
    return [pscustomobject]@{ converter = $selected; changed = $record.changed }
}

function Remove-WinBookSplitCalibreProbeState {
    param($Stage, [hashtable]$Created)
    $Stage.AssertUnchanged()
    $children = [Collections.Generic.List[string]]::new()
    foreach ($child in [IO.Directory]::EnumerateFileSystemEntries($Stage.Path)) {
        $name = [IO.Path]::GetFileName($child)
        if (-not $Created.ContainsKey($name) -or
            [WinBookSplitSetup.DirectoryLease]::ReadIdentity($child) -cne $Created[$name]) {
            throw 'The temporary Calibre version probe state contains an unowned member.'
        }
        $children.Add($child)
    }
    if ($children.Count -ne $Created.Count) { throw 'A temporary Calibre probe directory changed.' }
    $count = 0
    foreach ($child in $children) {
        $pending = [Collections.Generic.Stack[string]]::new()
        $pending.Push($child)
        while ($pending.Count -gt 0) {
            $current = $pending.Pop()
            $attributes = [IO.File]::GetAttributes($current)
            $count++
            if ($count -gt 10000 -or ($attributes -band [IO.FileAttributes]::ReparsePoint) -or
                -not ([IO.Path]::GetFullPath($current)).StartsWith($Stage.Path + '\', [StringComparison]::OrdinalIgnoreCase)) {
                throw 'The temporary Calibre version probe state is excessive or unsafe to clean.'
            }
            if ($attributes -band [IO.FileAttributes]::Directory) {
                foreach ($member in [IO.Directory]::EnumerateFileSystemEntries($current)) {
                    if ($pending.Count -ge 10000) { throw 'The temporary Calibre version probe state exceeds the inspection limit.' }
                    $pending.Push($member)
                }
            }
        }
    }
    foreach ($child in $children) {
        $Stage.AssertUnchanged()
        [IO.Directory]::Delete($child, $true)
    }
    $Stage.DeleteEmpty()
}

function Invoke-WinBookSplitCalibreVersionProbe {
    param([string]$Path, [string]$ApplicationRoot, [double]$TimeoutSeconds)
    Initialize-WinBookSplitSetupTypes
    Initialize-WinBookSplitLogTypes
    $base = Get-WinBookSplitTrustedPath -Path ([IO.Path]::GetTempPath()) -Directory
    if ($base -notmatch '^[A-Za-z]:\\') { throw 'Calibre version checks need an ordinary local temporary directory.' }
    if ($base.Length -gt 3) { $base = $base.TrimEnd('\') }
    $ancestors = [Collections.Generic.List[string]]::new()
    for ($current = $base; $current; $current = [IO.Path]::GetDirectoryName($current)) { $ancestors.Insert(0,$current) }
    $leases = [Collections.Generic.List[IDisposable]]::new()
    $stateLeases = [Collections.Generic.List[IDisposable]]::new()
    $created = @{}; $original = @{}
    $variables = [ordered]@{ CALIBRE_CONFIG_DIRECTORY = 'Config'; CALIBRE_CACHE_DIRECTORY = 'Cache'; CALIBRE_TEMP_DIR = 'Temp' }
    $stage = $null; $probe = $null; $failure = $null; $cleanupFailure = $null
    $environmentChanged = $false; $invoked = $false
    try {
        foreach ($directory in $ancestors) { $leases.Add([WinBookSplitLogging.PathLease]::new($directory,$true)) }
        foreach ($lease in $leases) { $lease.AssertUnchanged() }
        $folder = [IO.Path]::Combine($base, '.WinBookSplit-Calibre-Probe-' + [Guid]::NewGuid().ToString('N'))
        [WinBookSplitSetup.DirectoryLease]::CreateExclusive($folder)
        $stage = [WinBookSplitSetup.DirectoryLease]::new($folder)
        foreach ($name in $variables.Values) {
            $stage.AssertUnchanged()
            $statePath = [IO.Path]::Combine($folder,$name)
            [WinBookSplitSetup.DirectoryLease]::CreateExclusive($statePath)
            $lease = [WinBookSplitLogging.PathLease]::new($statePath,$true)
            $stateLeases.Add($lease)
            $created[$name] = [WinBookSplitSetup.DirectoryLease]::ReadIdentity($statePath)
        }
        $environment = [Environment]::GetEnvironmentVariables('Process')
        foreach ($name in $environment.Keys) { if ($name -ilike 'CALIBRE_*') { $original[$name] = [string]$environment[$name] } }
        $environmentChanged = $true
        foreach ($name in $original.Keys) { [Environment]::SetEnvironmentVariable($name,$null,'Process') }
        foreach ($name in $variables.Keys) {
            [Environment]::SetEnvironmentVariable($name,[IO.Path]::Combine($folder,$variables[$name]),'Process')
        }
        $invoked = $true
        $probe = Invoke-WinBookSplitRuntimeProbe $Path @('--version') $ApplicationRoot -TimeoutSeconds $TimeoutSeconds
    } catch { $failure = $_.Exception }
    finally {
        if ($environmentChanged) {
            foreach ($name in $variables.Keys) { [Environment]::SetEnvironmentVariable($name,$null,'Process') }
            foreach ($name in $original.Keys) { [Environment]::SetEnvironmentVariable($name,$original[$name],'Process') }
        }
        foreach ($lease in $stateLeases) { $lease.Dispose() }
        try {
            if ($null -ne $stage -and (-not $invoked -or
                ($null -ne $probe -and $probe.ParentStopped -and $probe.DescendantsStopped -and -not $probe.StopError))) {
                Remove-WinBookSplitCalibreProbeState $stage $created
            }
        } catch { $cleanupFailure = $_.Exception }
        finally {
            if ($null -ne $stage) { $stage.Dispose() }
            foreach ($lease in $leases) { $lease.Dispose() }
        }
    }
    if ($null -ne $failure) { throw $failure }
    if ($null -ne $cleanupFailure -and -not ($probe -and ($probe.Cancelled -or $probe.TimedOut))) {
        $errorRecord = [InvalidOperationException]::new('Cannot clean temporary Calibre version probe state: ' + $cleanupFailure.Message)
        $errorRecord.Data['Probe'] = $probe
        throw $errorRecord
    }
    return $probe
}
