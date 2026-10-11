param([string]$RepositoryRoot, [string]$WorkRoot)

BeforeAll {
    . (Join-Path $RepositoryRoot 'engine\WinBookSplit.Logging.ps1')
    . (Join-Path $RepositoryRoot 'engine\WinBookSplit.Paths.ps1')
    . (Join-Path $RepositoryRoot 'engine\WinBookSplit.Diagnostics.ps1')
    Initialize-WinBookSplitLogTypes
    $tokens=$null; $errors=$null
    $ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $RepositoryRoot 'WinBookSplit.ps1'),[ref]$tokens,[ref]$errors)
    if (@($errors).Count) {throw 'The trusted Main contains syntax errors.'}
    foreach ($name in @('Open-WinBookSplitConsoleRecords','New-WinBookSplitRunRecord','Complete-WinBookSplitConsoleLog','New-WinBookSplitOutcome')) {
        $functions=@($ast.FindAll({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -ceq $name},$true))
        if ($functions.Count -ne 1) {throw ('Expected one trusted Main function: '+$name)}
        . ([scriptblock]::Create($functions[0].Extent.Text))
    }
    function New-AuthoredRecordDirectory {
        $directory = Join-Path $WorkRoot ('records-' + [guid]::NewGuid().ToString('N'))
        $null = New-Item -ItemType Directory -Path $directory
        [IO.File]::WriteAllText((Join-Path $directory '.WinBookSplit-console-owner.json'), '{"kind":"console"}', [Text.UTF8Encoding]::new($false))
        return $directory
    }
    function Set-AuthoredDependencyRecordState {
        param($Records, [bool]$IsPreview=$false)
        $script:consoleDir=$Records.Directory; $script:consoleRunId=$Records.RunId
        $script:consoleLogWriter=$Records.Writer; $script:closedConsoleLogWriter=$Records.Writer
        $script:consoleRecordLeases=$Records.Leases; $script:runtime=$null; $script:converter=$null
        $script:confirmedPlan=$null; $script:sourceObservation=@{path='C:\PRIVATE\dependency.pdf';binding='metadata_only';size_bytes=128;sha256=$null}
        $script:runStartedUtc=[DateTime]::UtcNow.ToString('o'); $script:ver='1.0.0'; $script:inputExt='.pdf'
        $script:Preview=$IsPreview; $script:NonInteractive=$false; $script:NoPause=$true; $script:KeepConvertedPdf=$false
        $script:ConversionTimeout=120; $script:ProcessTimeout=300
    }
    function Close-AuthoredConsoleRecords {
        param($Records)
        if ($null -ne $Records.Writer) {$Records.Writer.Dispose();$Records.Writer.CloseLease()}
        foreach ($lease in $Records.Leases) {$lease.Dispose()}
    }
}

Describe 'V101 existing console records for preflight refusals' -Tag 'V101-AC-037' {
    It 'saves the full synthetic diagnosis under real leases with null accepted runtimes and unchanged root schema' {
        $base=Join-Path $WorkRoot ('d'+[guid]::NewGuid().ToString('N').Substring(0,12))
        $null=New-Item -ItemType Directory -Path $base
        $records=Open-WinBookSplitConsoleRecords -Base $base -Preview $false
        try {
            Set-AuthoredDependencyRecordState $records
            $diagnosis='[DEPENDENCY-ERROR] '+(@{Code='runtime_invalid';Attempts=@(@{Reason='python_security_floor';Message='PRIVATE raw diagnosis';Observed='3.14.2'});Probe=$null} | ConvertTo-Json -Depth 8 -Compress)
            $records.Writer.WriteLine($diagnosis)
            $outcome=New-WinBookSplitOutcome 'failed' 'runtime_invalid' 'Concise authored security-floor summary'
            Complete-WinBookSplitConsoleLog $outcome
            $manifest=Get-Content -LiteralPath (Join-Path $records.Directory 'WinBookSplit_Run.json') -Raw | ConvertFrom-Json
            @($manifest.PSObject.Properties.Name) -join ',' | Should-Be -Expected 'protocol,version,run_id,application_version,started_utc,finished_utc,diagnostics_finalized,runtime_versions,settings,source_identity,plan,engine_result,outcome,warnings,log'
            $manifest.diagnostics_finalized | Should-BeTrue
            $manifest.runtime_versions.python | Should -BeNullOrEmpty
            $manifest.runtime_versions.pypdf | Should -BeNullOrEmpty
            $manifest.runtime_versions.calibre | Should -BeNullOrEmpty
            $manifest.plan | Should -BeNullOrEmpty
            $manifest.engine_result | Should -BeNullOrEmpty
            $manifest.outcome.exit_code | Should-Be -Expected 3
            $manifest.outcome.written_count | Should-Be -Expected 0
            [IO.File]::ReadAllText($records.LogPath).Contains($diagnosis) | Should-BeTrue
            { [IO.Directory]::Move($records.Directory,$records.Directory+'-replaced') } | Should-Throw
            { [IO.File]::WriteAllText((Join-Path $records.Directory '.WinBookSplit-console-owner.json'),'foreign') } | Should-Throw
        } finally {Close-AuthoredConsoleRecords $records}
    }
    It 'keeps the reused Preview record path entirely in memory' {
        $base=Join-Path $WorkRoot ('d'+[guid]::NewGuid().ToString('N').Substring(0,12))
        $null=New-Item -ItemType Directory -Path $base
        $records=Open-WinBookSplitConsoleRecords -Base $base -Preview $true
        try {
            Set-AuthoredDependencyRecordState $records $true
            $records.Directory | Should -BeNullOrEmpty
            $records.LogPath | Should -BeNullOrEmpty
            @($records.Leases).Count | Should-Be -Expected 0
            $outcome=New-WinBookSplitOutcome 'failed' 'runtime_invalid' 'SYNTHETIC Preview refusal'
            Complete-WinBookSplitConsoleLog $outcome
            @(Get-ChildItem -LiteralPath $base -Force).Count | Should-Be -Expected 0
        } finally {Close-AuthoredConsoleRecords $records}
    }
    It 'retains an existing marker and its partial reservation when preparation fails without overwriting data' {
        $base=Join-Path $WorkRoot ('d'+[guid]::NewGuid().ToString('N').Substring(0,12))
        $null=New-Item -ItemType Directory -Path $base
        $script:partialReservation=New-WinBookSplitConsoleDirectory -Base $base
        $marker=Join-Path $script:partialReservation.path '.WinBookSplit-console-owner.json'
        [IO.File]::WriteAllText($marker,'Authored foreign marker must remain byte-for-byte.')
        Mock New-WinBookSplitConsoleDirectory {$script:partialReservation}
        $failure=$null
        try {$null=Open-WinBookSplitConsoleRecords -Base $base -Preview $false} catch {$failure=$_.Exception}
        $failure | Should -Not -BeNullOrEmpty
        $failure.Data['DiagnosticDirectory'] | Should-Be -Expected $script:partialReservation.path
        [IO.File]::ReadAllText($marker) | Should-Be -Expected 'Authored foreign marker must remain byte-for-byte.'
        [IO.Directory]::Exists($script:partialReservation.path) | Should-BeTrue
        @(Get-ChildItem -LiteralPath $script:partialReservation.path -Force).Count | Should-Be -Expected 1
    }
    It 'leaves primary dependency cancellation and timeout outcomes intact under an authored manifest finalizer fault' {
        foreach ($row in @(@('failed','runtime_invalid',3),@('cancelled','processor_cancelled',130),@('timeout','processor_timeout',130))) {
            $base=Join-Path $WorkRoot ('d'+[guid]::NewGuid().ToString('N').Substring(0,12))
            $null=New-Item -ItemType Directory -Path $base
            $records=Open-WinBookSplitConsoleRecords -Base $base -Preview $false
            try {
                Set-AuthoredDependencyRecordState $records
                $outcome=New-WinBookSplitOutcome $row[0] $row[1] 'SYNTHETIC primary failure with a controlled secondary fault'
                $before=$outcome | ConvertTo-Json -Depth 8 -Compress
                Mock Write-WinBookSplitRunManifest {throw 'Authored secondary manifest finalizer fault.'}
                {Complete-WinBookSplitConsoleLog $outcome} | Should-Throw
                ($outcome | ConvertTo-Json -Depth 8 -Compress) | Should-Be -Expected $before
                $outcome.exit_code | Should-Be -Expected $row[2]
                [IO.Directory]::Exists($records.Directory) | Should-BeTrue
                [IO.File]::ReadAllText($records.LogPath) | Should-MatchString -Expected '[OPERATION-OUTCOME]'
                [IO.File]::Exists((Join-Path $records.Directory 'WinBookSplit_Run.json')) | Should-BeFalse
            } finally {Close-AuthoredConsoleRecords $records}
        }
    }
}

Describe 'Bounded UTF-8 local run records' -Tag 'AC-063' {
    It 'holds the base before reservation and the fresh console before its first marker write' {
        $base = Join-Path $WorkRoot ('base-' + [guid]::NewGuid().ToString('N'))
        $null = New-Item -ItemType Directory -Path $base
        $leases = New-WinBookSplitRecordLeases -Directory $base -IncludeMarker $false
        try {
            { [IO.Directory]::Move($base, $base + '-replaced') } | Should-Throw
            $reservation = New-WinBookSplitConsoleDirectory -Base $base
            $leases.Add([WinBookSplitLogging.PathLease]::new($reservation.path, $true))
            { [IO.Directory]::Move($reservation.path, $reservation.path + '-replaced') } | Should-Throw
            $marker = Join-Path $reservation.path '.WinBookSplit-console-owner.json'
            [IO.File]::WriteAllText($marker, '{"kind":"authored"}', [Text.UTF8Encoding]::new($false))
            (Get-Item -LiteralPath $marker).DirectoryName | Should-Be -Expected $reservation.path
        }
        finally { foreach ($lease in $leases) { $lease.Dispose() } }
    }
    It 'counts UTF-8 bytes and refuses body overflow while reserving final status' {
        $stream = [IO.MemoryStream]::new()
        $writer = New-WinBookSplitLogWriter -Stream $stream -BodyLimitBytes 8 -FinalReserveBytes 64
        try {
            $writer.Write([string][char]0x00e4 + [char]0x00e4)
            $writer.BytesWritten | Should-Be -Expected 4
            { $writer.Write('overflow') } | Should-Throw
            $writer.LimitReached | Should-BeTrue
            $writer.BytesWritten | Should-Be -Expected 4
            $writer.WriteFinal('[OPERATION-OUTCOME] failed: bounded log')
            ($writer.BytesWritten -le $writer.MaxBytes) | Should-BeTrue
            [Text.UTF8Encoding]::new($false, $true).GetString($stream.ToArray()).EndsWith("failed: bounded log`r`n") | Should-BeTrue
        }
        finally { $writer.Dispose(); $writer.CloseLease() }
    }
    It 'never partly writes an oversized terminal record' {
        $stream = [IO.MemoryStream]::new()
        $writer = New-WinBookSplitLogWriter -Stream $stream -BodyLimitBytes 4 -FinalReserveBytes 8
        try {
            $writer.Write('body')
            { $writer.WriteFinal('terminal too large') } | Should-Throw
            [Text.Encoding]::UTF8.GetString($stream.ToArray()) | Should-Be -Expected 'body'
        }
        finally { $writer.Dispose(); $writer.CloseLease() }
    }
    It 'publishes readable UTF-8 only after closing the owned pending file and refuses overwrite' {
        $directory = New-AuthoredRecordDirectory
        $leases = New-WinBookSplitRecordLeases $directory
        $record = @{ protocol='winbooksplit.run'; version=1; title=([string][char]0x65e5 + [char]0x00e4); outcome=@{ exit_code=0 } }
        try {
            Write-WinBookSplitRunManifest $directory $record $leases
            $path = Join-Path $directory 'WinBookSplit_Run.json'
            $before = [IO.File]::ReadAllBytes($path)
            ([Text.UTF8Encoding]::new($false, $true).GetString($before) | ConvertFrom-Json).title | Should-Be -Expected $record.title
            (Test-Path -LiteralPath (Join-Path $directory 'WinBookSplit_Run.pending.json')) | Should-BeFalse
            { Write-WinBookSplitRunManifest $directory @{ outcome=@{ exit_code=6 } } $leases } | Should-Throw
            [Convert]::ToBase64String([IO.File]::ReadAllBytes($path)) | Should-Be -Expected ([Convert]::ToBase64String($before))
            (Test-Path -LiteralPath (Join-Path $directory 'WinBookSplit_Run.pending.json')) | Should-BeTrue
            (Get-Content -LiteralPath (Join-Path $directory 'WinBookSplit_Run.pending.json') -Raw | ConvertFrom-Json).diagnostics_finalized | Should-BeFalse
        }
        finally { foreach ($lease in $leases) { $lease.Dispose() } }
    }
    It 'holds its marker and directory against writes or replacement during finalization' {
        $directory = New-AuthoredRecordDirectory
        $leases = New-WinBookSplitRecordLeases $directory
        try {
            { [IO.File]::WriteAllText((Join-Path $directory '.WinBookSplit-console-owner.json'), 'tampered') } | Should-Throw
            { [IO.Directory]::Move($directory, $directory + '-replacement') } | Should-Throw
            Write-WinBookSplitRunManifest $directory @{ outcome=@{ exit_code=130 } } $leases
        }
        finally { foreach ($lease in $leases) { $lease.Dispose() } }
    }
    It 'can append a corrective failure after log close without replacing another file' {
        $directory = New-AuthoredRecordDirectory
        $path = Join-Path $directory 'console.log'
        $stream = [IO.File]::Open($path, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::Read)
        $writer = New-WinBookSplitLogWriter -Stream $stream -Path $path -BodyLimitBytes 128 -FinalReserveBytes 128
        try {
            $writer.WriteFinal('[OPERATION-OUTCOME] success')
            $writer.Dispose()
            { [IO.File]::Move($path, $path + '-replaced') } | Should-Throw
            $writer.AppendCorrection('[OPERATION-OUTCOME] incomplete')
            [IO.File]::ReadAllText($path).EndsWith("[OPERATION-OUTCOME] incomplete`r`n") | Should-BeTrue
            $writer.BytesWritten | Should-Be -Expected ((Get-Item -LiteralPath $path).Length)
            [IO.File]::AppendAllText($path, 'foreign change')
            { $writer.AppendCorrection('[OPERATION-OUTCOME] refused') } | Should-Throw
        }
        finally { $writer.Dispose(); $writer.CloseLease() }
    }
}
