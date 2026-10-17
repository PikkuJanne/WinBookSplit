# Convenience delivery only. The frozen kit, native receipts and runtime selection
# are unchanged. Never interpret a returned ZIP as a passing acceptance result.
[CmdletBinding()]
param([switch]$NoOpenResults)

$ErrorActionPreference = 'Stop'
$PSModuleAutoloadingPreference = 'None'
Import-Module ([IO.Path]::Combine($PSHOME, 'Modules\Microsoft.PowerShell.Utility\Microsoft.PowerShell.Utility.psd1')) -ErrorAction Stop
Import-Module ([IO.Path]::Combine($PSHOME, 'Modules\Microsoft.PowerShell.Management\Microsoft.PowerShell.Management.psd1')) -ErrorAction Stop

function Get-OwnerDeliveryStreamHash {
    param([IO.Stream]$Stream)
    $algorithm = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($algorithm.ComputeHash($Stream))).Replace('-', '').ToLowerInvariant() }
    finally { $algorithm.Dispose() }
}

function Add-OwnerDeliveryAncestors {
    param([string]$Path, $Leases)
    $paths = New-Object 'Collections.Generic.List[string]'
    $current = [IO.Path]::GetFullPath($Path)
    while ($current) { $paths.Add($current); $current = [IO.Path]::GetDirectoryName($current) }
    for ($i = $paths.Count - 1; $i -ge 0; $i--) {
        $Leases.Add((New-Object WinBookSplitLogging.PathLease($paths[$i], $true)))
    }
}

function Get-OwnerDeliveryPrintedReport {
    param([string]$StdoutPath)
    $stream = [IO.File]::Open($StdoutPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
    try {
        if ($stream.Length -gt 1048576) { throw 'The test console output exceeded its delivery bound.' }
        $bytes = New-Object byte[] ([int]$stream.Length); $offset = 0
        while ($offset -lt $bytes.Length) {
            $count = $stream.Read($bytes, $offset, $bytes.Length - $offset)
            if ($count -eq 0) { throw 'The test console output ended early.' }
            $offset += $count
        }
        if ($stream.ReadByte() -ne -1) { throw 'The test console output changed while reading.' }
    } finally { $stream.Dispose() }
    # Start-Process console logs are convenience text, not native raw receipts.
    # Accommodate BOMs and the inbox host's local console encoding without
    # modifying the environment or the frozen controller's argument vector.
    $encodings = @([Text.UTF8Encoding]::new($false, $true), [Text.Encoding]::Unicode,
        [Console]::OutputEncoding, [Text.Encoding]::Default,
        [Text.Encoding]::GetEncoding([Globalization.CultureInfo]::CurrentCulture.TextInfo.OEMCodePage))
    $candidates = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
    foreach ($encoding in $encodings) {
        try { $text = $encoding.GetString($bytes).TrimStart([char]0xfeff) } catch { continue }
        $lines = @([regex]::Matches($text, '(?m)^Return only: ([^\r\n]+)\r?$'))
        if ($lines.Count -gt 1) { throw 'The test printed more than one return file. No report was selected.' }
        if ($lines.Count -eq 1) {
            $candidate = $lines[0].Groups[1].Value
            if ([IO.File]::Exists($candidate)) { $null = $candidates.Add($candidate) }
        }
    }
    if ($candidates.Count -ne 1) { throw 'The test did not identify one readable return ZIP. No cached or older report was selected.' }
    foreach ($candidate in $candidates) { return $candidate }
}

function Copy-OwnerDeliveryReport {
    param([string]$Source, [string]$Destination, [datetime]$StartedUtc, [datetime]$FinishedUtc, $Leases)
    $profile = [Environment]::GetFolderPath('UserProfile')
    $cachePrefix = [IO.Path]::Combine($profile, '.cache') + '\'
    $bootstrapPrefix = $profile + '\'
    $cacheName = if ($Source.StartsWith($cachePrefix, [StringComparison]::OrdinalIgnoreCase)) { $Source.Substring($cachePrefix.Length) } else { '' }
    $bootstrapName = if ($Source.StartsWith($bootstrapPrefix, [StringComparison]::OrdinalIgnoreCase)) { $Source.Substring($bootstrapPrefix.Length) } else { '' }
    if ($cacheName -cnotmatch '^WBS-Native-[0-9a-f]{32}\\native-(?:store-result|result)\.zip$' -and
        $bootstrapName -cnotmatch '^WBS-native-bootstrap-[0-9a-f]{32}\.zip$') {
        throw 'The printed return ZIP is outside the frozen kit return locations.'
    }
    $null = Get-NativeAcceptanceOrdinaryPath $Source
    Add-OwnerDeliveryAncestors ([IO.Path]::GetDirectoryName($Source)) $Leases
    $sourceLease = New-Object WinBookSplitLogging.PathLease($Source, $false, $false, $true)
    $Leases.Add($sourceLease)
    $inputStream = [IO.File]::Open($Source, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
    $outputStream = $null
    try {
        $sourceLease.AssertSameHandle($inputStream.SafeFileHandle)
        if ($inputStream.Length -lt 1 -or $inputStream.Length -gt 33554432) { throw 'The return ZIP is outside its finite size bound.' }
        if ([IO.File]::GetCreationTimeUtc($Source) -lt $StartedUtc.AddSeconds(-2) -or
            [IO.File]::GetCreationTimeUtc($Source) -gt $FinishedUtc.AddSeconds(2) -or
            [IO.File]::GetLastWriteTimeUtc($Source) -lt $StartedUtc.AddSeconds(-2) -or
            [IO.File]::GetLastWriteTimeUtc($Source) -gt $FinishedUtc.AddSeconds(2)) {
            throw 'The return ZIP was not created during this launch. No older result was selected.'
        }
        $sourceHash = Get-OwnerDeliveryStreamHash $inputStream
        $size = $inputStream.Length; $inputStream.Position = 0
        $outputStream = [IO.File]::Open($Destination, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::Read)
        $inputStream.CopyTo($outputStream); $outputStream.Flush(); $outputStream.Dispose(); $outputStream = $null
        $copy = Get-NativeAcceptanceFileHash $Destination 33554432
        if ($copy.bytes -ne $size -or $copy.sha256 -cne $sourceHash) { throw 'The copied return ZIP did not match the original bytes.' }
        $sourceLease.AssertUnchanged()
        return [pscustomobject]@{bytes = $size; sha256 = $sourceHash}
    } finally { if ($outputStream) { $outputStream.Dispose() }; $inputStream.Dispose() }
}

function Invoke-OwnerDelivery {
    param([string]$Root, [switch]$NoOpenResults)
    $leases = New-Object 'Collections.Generic.List[IDisposable]'
    $archiveStream = $archive = $child = $null
    $results = $stderr = $null; $exitCode = 2
    try {
        $Root = [IO.Path]::GetFullPath($Root)
        $originalZip = [IO.Path]::Combine($Root, 'WinBookSplit-v103-native-kit-243d31e.zip')
        $kit = [IO.Path]::Combine($Root, 'kit')
        $archiveStream = [IO.File]::Open($originalZip, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
        if ($archiveStream.Length -ne 1559722 -or
            (Get-OwnerDeliveryStreamHash $archiveStream) -cne 'e2f2a36a727edec378b891f5b125a18bea89dc8e41946719ac8b8a6d2e1837fc') {
            throw 'The supplied test kit is incomplete or changed. Use the original transferred bundle.'
        }
        $archiveStream.Position = 0
        Add-Type -AssemblyName System.IO.Compression
        $archive = New-Object IO.Compression.ZipArchive($archiveStream, [IO.Compression.ZipArchiveMode]::Read, $true)
        if ($archive.Entries.Count -ne 35) { throw 'The original kit member count differs.' }
        # Bootstrap only the byte-pinned original logging definitions from the
        # authenticated archive; no unchecked on-disk helper is evaluated.
        $loggingEntry = $archive.GetEntry('engine/WinBookSplit.Logging.ps1')
        $reader = New-Object IO.StreamReader($loggingEntry.Open(), [Text.UTF8Encoding]::new($false, $true))
        try { $loggingText = $reader.ReadToEnd() } finally { $reader.Dispose() }
        . ([scriptblock]::Create($loggingText))
        Initialize-WinBookSplitLogTypes
        Add-OwnerDeliveryAncestors $Root $leases
        $zipLease = New-Object WinBookSplitLogging.PathLease($originalZip, $false, $false, $true)
        $zipLease.AssertSameHandle($archiveStream.SafeFileHandle); $leases.Add($zipLease)
        $names = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        foreach ($entry in $archive.Entries) {
            if (-not $names.Add($entry.FullName) -or $entry.FullName -match '(^/|\\|(^|/)\.\.(/|$)|:)' -or
                $entry.FullName.EndsWith('/') -or $entry.Length -gt 8388608) { throw 'The original kit member name or bound differs.' }
            $path = [IO.Path]::Combine($kit, $entry.FullName.Replace('/', '\'))
            Add-OwnerDeliveryAncestors ([IO.Path]::GetDirectoryName($path)) $leases
            $lease = New-Object WinBookSplitLogging.PathLease($path, $false, $false, $true); $leases.Add($lease)
            $file = [IO.File]::Open($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
            $member = $entry.Open()
            try {
                $lease.AssertSameHandle($file.SafeFileHandle)
                if ($file.Length -ne $entry.Length -or (Get-OwnerDeliveryStreamHash $file) -cne (Get-OwnerDeliveryStreamHash $member)) {
                    throw 'An extracted test helper differs from the original kit. Extract a fresh transfer bundle.'
                }
            } finally { $member.Dispose(); $file.Dispose() }
        }
        . ([IO.Path]::Combine($kit, 'engine', 'WinBookSplit.Diagnostics.ps1'))
        . ([IO.Path]::Combine($kit, 'tools', 'compatibility', 'Invoke-NativeAcceptance.ps1'))
        Initialize-NativeAcceptanceDirectory
        $results = [IO.Path]::Combine($Root, 'RESULTS-' + [guid]::NewGuid().ToString('N'))
        [WinBookSplitNative.Directory]::CreateExclusive($results)
        $leases.Add((New-Object WinBookSplitLogging.PathLease($results, $true)))
        $stdout = [IO.Path]::Combine($results, 'PRIVATE-console-output.txt')
        $stderr = [IO.Path]::Combine($results, 'PRIVATE-console-errors.txt')
        foreach ($log in @($stdout, $stderr)) {
            $file = [IO.File]::Open($log, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::ReadWrite)
            try { $leases.Add((New-Object WinBookSplitLogging.PathLease($log, $false))) } finally { $file.Dispose() }
        }
        $system = [Environment]::GetFolderPath('System')
        $inbox = [IO.Path]::Combine($system, 'WindowsPowerShell', 'v1.0', 'powershell.exe')
        $null = Get-NativeAcceptanceOrdinaryPath $inbox
        $arguments = @('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File',
            [IO.Path]::Combine($kit, 'Run-Windows10.ps1'), '-Scenario', 'store-existing')
        $command = ($arguments | ForEach-Object { ConvertTo-NativeArgument $_ }) -join ' '
        Write-Host 'Running the automatic Windows 10 test. Please wait; no typing is needed.'
        Write-Host 'The checks can take several minutes. Keep this window open.'
        # Same executable/argv as the unchanged Run-Windows10.cmd. Start-Process
        # inherits the existing environment: no runtime inputs or new environment.
        # The frozen controller keeps its original bounded child-process checks.
        $started = [datetime]::UtcNow
        $child = Start-Process -FilePath $inbox -ArgumentList $command -WorkingDirectory $Root -WindowStyle Hidden -PassThru -Wait -RedirectStandardOutput $stdout -RedirectStandardError $stderr
        $finished = [datetime]::UtcNow; $child.Refresh()
        if (-not $child.HasExited -or $null -eq $child.ExitCode -or $child.ExitCode -isnot [int]) {
            throw 'The test process did not provide a completed native exit code.'
        }
        $exitCode = $child.ExitCode
        foreach ($lease in $leases) { $lease.AssertUnchanged() }
        if ((Get-Item -LiteralPath $stderr).Length -gt 1048576) { throw 'The test error output exceeded its delivery bound.' }
        $source = Get-OwnerDeliveryPrintedReport $stdout
        $sendBack = [IO.Path]::Combine($results, 'SEND-THIS-REPORT.zip')
        $identity = Copy-OwnerDeliveryReport $source $sendBack $started $finished $leases
        Write-NativeAcceptanceJson ([IO.Path]::Combine($results, 'PRIVATE-delivery.json')) ([ordered]@{
            protocol = 'winbooksplit.owner-delivery'; schema_version = 1
            native_exit_code = $exitCode; source_report = $source; copy = $identity
            native_acceptance = 'NOT_EVALUATED_BY_DELIVERY'; report_bytes_changed = $false
        })
        Write-Host ''
        Write-Host 'REPORT READY - send only this file back in the chat:'
        Write-Host $sendBack
        if ($exitCode -ne 0) { Write-Host 'The test reported a failure or blocker. Send the report anyway.' }
        Write-Host 'Keep the other diagnostic files on this computer. The report still needs review.'
        if (-not $NoOpenResults) {
            try { Start-Process -FilePath ([IO.Path]::Combine([Environment]::GetFolderPath('Windows'), 'explorer.exe')) -ArgumentList ('/select,' + (ConvertTo-NativeArgument $sendBack)) -WindowStyle Normal | Out-Null }
            catch { Write-Host ('Open the results folder shown above and send SEND-THIS-REPORT.zip. The folder could not open automatically: ' + $_.Exception.Message) }
        }
        return [int]$exitCode
    } catch {
        Write-Host ''
        Write-Host 'TEST STOPPED - no passing result has been recorded.'
        $message = $_.Exception.Message.Replace([Environment]::GetFolderPath('UserProfile'), '%USERPROFILE%')
        Write-Host ('Please send this short message in the chat: ' + $message)
        if ($stderr -and [IO.File]::Exists($stderr) -and (Get-Item -LiteralPath $stderr).Length -le 1048576) {
            $reader = New-Object IO.StreamReader($stderr, $true)
            try {
                $buffer = New-Object char[] 768; $count = $reader.ReadBlock($buffer, 0, $buffer.Length)
                if ($count -gt 0) {
                    $detail = [string]::new($buffer, 0, $count)
                    $detail = $detail.Replace([Environment]::GetFolderPath('UserProfile'), '%USERPROFILE%').Replace($Root, '<TEST_FOLDER>')
                    Write-Host ('Windows said: ' + ([regex]::Replace($detail, '[\x00-\x1f\x7f-\x9f]', ' ')).Trim())
                }
            } finally { $reader.Dispose() }
        }
        Write-Host 'Do not change Python or security settings, and do not retry with a different launcher.'
        if ($results) { Write-Host ('Private diagnostics remain here: ' + $results) }
        return $(if ($exitCode -ne 0) { [int]$exitCode } else { 2 })
    } finally {
        if ($child) { $child.Dispose() }
        if ($archive) { $archive.Dispose() }; if ($archiveStream) { $archiveStream.Dispose() }
        for ($i = $leases.Count - 1; $i -ge 0; $i--) { $leases[$i].Dispose() }
    }
}

if ($MyInvocation.InvocationName -ne '.') {
    exit (Invoke-OwnerDelivery -Root $PSScriptRoot -NoOpenResults:$NoOpenResults)
}
