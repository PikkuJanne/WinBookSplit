param([string]$RepositoryRoot, [string]$WorkRoot, [string]$TestPython)

BeforeAll {
    . (Join-Path $RepositoryRoot 'engine\WinBookSplit.Diagnostics.ps1')
    . (Join-Path $RepositoryRoot 'engine\WinBookSplit.Process.ps1')
    if (-not [IO.File]::Exists($TestPython)) { throw 'Process tests require the explicit pinned TestPython.' }
    $script:child = Join-Path $RepositoryRoot 'tests\process\native_child.py'
    function Invoke-Control {
        param([string]$Kind, [string[]]$Data = @(), [int]$ResultLimit = 8388608)
        Invoke-WinBookSplitProcess -Path $TestPython -WorkingDirectory $RepositoryRoot -TimeoutSeconds 10 `
            -Arguments (@('-I','-B','-X','utf8',$script:child,$Kind)+$Data) -ResultLimitBytes $ResultLimit
    }
}

Describe 'M2 bounded owned executable supervision' -Tag 'AC-046', 'AC-047' {
    It 'drains dual two-MiB streams and retains a separate valid final frame' {
        $result = Invoke-Control 'flood'
        $result.ExitCode | Should-Be -Expected 1
        $result.JobAssigned | Should-BeTrue
        $result.ParentStopped | Should-BeTrue
        $result.DescendantsStopped | Should-BeTrue
        $result.StreamsComplete | Should-BeTrue
        $result.StdoutTruncated | Should-BeTrue
        $result.StderrTruncated | Should-BeTrue
        ($result.StdoutTotalBytes -gt (2*1024*1024)) | Should-BeTrue
        ($result.StderrTotalBytes -gt (2*1024*1024)) | Should-BeTrue
        @($result.ResultRecords).Count | Should-Be -Expected 1
        $frame = ConvertFrom-SplitResult -Stdout ($result.ResultRecords -join "`n") -ExitCode $result.ExitCode -Mode 'manual'
        $frame.code | Should-Be -Expected 'invalid_start_pages'
    }
    It 'retains immediate blank lines and both no-newline tails' {
        $result = Invoke-Control 'fast-tail'
        $result.ExitCode | Should-Be -Expected 23
        $result.Stdout | Should-Be -Expected "`n`nstdout blank`n`nno-newline-output"
        $result.Stderr | Should-Be -Expected "`n`nFINAL_STDERR_NO_NEWLINE"
        $result.StreamsComplete | Should-BeTrue
    }
    It 'captures a valid frame larger than the bounded human stream tail' {
        $result = Invoke-Control 'large-frame'
        $result.StdoutTruncated | Should-BeTrue
        ($null -eq $result.ResultError) | Should-BeTrue
        @($result.ResultRecords).Count | Should-Be -Expected 1
        $frame = ConvertFrom-SplitResult -Stdout ($result.ResultRecords -join "`n") -ExitCode 1 -Mode 'manual'
        ($frame.message.Length -gt 100000) | Should-BeTrue
    }
    It 'reports a result frame that exceeds its separate bounded budget' {
        $result = Invoke-Control 'oversized-frame' -ResultLimit 1024
        ([string]::IsNullOrEmpty($result.ResultError)) | Should-BeFalse
        $result.StreamsComplete | Should-BeTrue
        $result.ExitCode | Should-Be -Expected 1
    }
    It 'keeps human text after a partial large frame without displaying machine fragments' {
        $result = Invoke-Control 'large-frame-human-tail'
        $result.HumanStdout | Should-Be -Expected "`nHuman last no-newline"
        @($result.ResultRecords).Count | Should-Be -Expected 1
        $result.StdoutTruncated | Should-BeTrue
    }
    It 'rejects duplicate structured frames rather than selecting a final success' {
        $result = Invoke-Control 'duplicate-frames'
        { ConvertFrom-SplitResult -Stdout ($result.ResultRecords -join "`n") -ExitCode 1 -Mode 'manual' } | Should-Throw
    }
    It 'rejects absent and malformed frames after complete native failure' {
        foreach ($kind in @('no-frame','malformed-frame')) {
            $result = Invoke-Control $kind
            $result.StreamsComplete | Should-BeTrue
            { ConvertFrom-SplitResult -Stdout ($result.ResultRecords -join "`n") -ExitCode 1 -Mode 'manual' } | Should-Throw
        }
    }
    It 'bounds an exact owned parent timeout and waits for confirmed stop' {
        $result = Invoke-WinBookSplitProcess -Path $TestPython -WorkingDirectory $RepositoryRoot -TimeoutSeconds .2 `
            -Arguments @('-I','-B','-c','import time;time.sleep(8)')
        $result.TimedOut | Should-BeTrue
        $result.Cancelled | Should-BeFalse
        $result.ParentStopped | Should-BeTrue
        $result.DescendantsStopped | Should-BeTrue
        $result.StreamsComplete | Should-BeTrue
        ($result.ElapsedSeconds -lt 3) | Should-BeTrue
    }
    It 'distinguishes injected cancellation from timeout' {
        $cancel = New-Object Threading.CancellationTokenSource
        try {
            $cancel.CancelAfter(150)
            $result = Invoke-WinBookSplitProcess -Path $TestPython -WorkingDirectory $RepositoryRoot -TimeoutSeconds 8 `
                -Arguments @('-I','-B','-c','import time;time.sleep(8)') -CancellationToken $cancel.Token
            $result.Cancelled | Should-BeTrue
            $result.TimedOut | Should-BeFalse
            $result.ParentStopped | Should-BeTrue
            $result.DescendantsStopped | Should-BeTrue
        }
        finally { $cancel.Dispose() }
    }
    It 'checks an already-cancelled token before launching executable side effects' {
        $marker = Join-Path $TestDrive 'pre-cancel-side-effect.txt'
        $savedMarker = $env:WBS_PROCESS_PRE_CANCEL_MARKER
        $cancel = New-Object Threading.CancellationTokenSource
        try {
            $env:WBS_PROCESS_PRE_CANCEL_MARKER = $marker
            $cancel.Cancel()
            $result = Invoke-WinBookSplitProcess -Path $TestPython -WorkingDirectory $RepositoryRoot -TimeoutSeconds 8 `
                -Arguments @('-I','-B','-c','import os,pathlib;pathlib.Path(os.environ["WBS_PROCESS_PRE_CANCEL_MARKER"]).write_text("unexpected")') `
                -CancellationToken $cancel.Token
            $result.Cancelled | Should-BeTrue
            ($null -eq $result.Pid) | Should-BeTrue
            Test-Path -LiteralPath $marker | Should-BeFalse
        }
        finally { $cancel.Dispose(); $env:WBS_PROCESS_PRE_CANCEL_MARKER = $savedMarker }
    }
}

Describe 'M2 UTF-8 and argument data boundary' -Tag 'AC-048', 'AC-049' {
    It 'preserves UTF-8 final text when a byte budget splits a multibyte character' {
        $result = Invoke-Control 'unicode-boundary'
        $result.ExitCode | Should-Be -Expected 0
        foreach ($text in @($result.Stdout,$result.Stderr)) {
            $text.Contains([char]0xfffd) | Should-BeFalse
            ([Text.Encoding]::UTF8.GetByteCount($text) -le 65536) | Should-BeTrue
        }
        $result.StdoutTruncated | Should-BeTrue
        $result.StderrTruncated | Should-BeTrue
    }
    It 'marshals empty, quoted, trailing-backslash and executable-looking text literally' {
        $unicode=([char]0x65E5).ToString()+[char]0x672C
        $arguments = @('', 'space value', '[%!] & (literal)', "O'Brien", 'a"b', 'trailing\', $unicode,
            '__import__("os").system("echo injected")', '$(Set-Content injected.txt)')
        $result = Invoke-Control 'arguments' -Data $arguments
        $actual = $result.Stdout | ConvertFrom-Json
        $actual.argv -join '|' | Should-Be -Expected ($arguments -join '|')
        $actual.utf8_env | Should-Be -Expected '1'
        $actual.io_env | Should-Be -Expected 'utf-8:backslashreplace'
    }
}
