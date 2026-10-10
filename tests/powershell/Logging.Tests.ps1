param([string]$RepositoryRoot, [string]$WorkRoot)

BeforeAll {
    . (Join-Path $RepositoryRoot 'engine\WinBookSplit.Logging.ps1')
    . (Join-Path $RepositoryRoot 'engine\WinBookSplit.Paths.ps1')
    Initialize-WinBookSplitLogTypes
    function New-AuthoredRecordDirectory {
        $directory = Join-Path $WorkRoot ('records-' + [guid]::NewGuid().ToString('N'))
        $null = New-Item -ItemType Directory -Path $directory
        [IO.File]::WriteAllText((Join-Path $directory '.WinBookSplit-console-owner.json'), '{"kind":"console"}', [Text.UTF8Encoding]::new($false))
        return $directory
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
