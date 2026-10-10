param([string]$RepositoryRoot, [string]$WorkRoot, [string]$TestPython)

BeforeAll {
    . (Join-Path $RepositoryRoot 'engine\WinBookSplit.Diagnostics.ps1')
    . (Join-Path $RepositoryRoot 'engine\WinBookSplit.Process.ps1')
    if (-not [IO.File]::Exists($TestPython)) { throw 'Session tests require pinned TestPython.' }
    Initialize-WinBookSplitProcess
    $script:sessionChild = Join-Path $RepositoryRoot 'tests\process\session_child.py'
    function Start-ControlSession {
        param([string]$Kind, [double]$Timeout = 5, [Threading.CancellationToken]$CancellationToken = [Threading.CancellationToken]::None)
        Start-WinBookSplitProcessSession -Path $TestPython -WorkingDirectory $RepositoryRoot -TimeoutSeconds $Timeout `
            -Arguments @('-I','-B','-X','utf8',$script:sessionChild,$Kind) -CancellationToken $CancellationToken
    }
    function Read-ControlRequest {
        param($Session)
        $deadline = [Diagnostics.Stopwatch]::StartNew()
        while ($deadline.Elapsed.TotalSeconds -lt 3) {
            $request = $Session.TryReadRequest()
            if ($null -ne $request) { return $request }
            if ($Session.Completed) { throw 'Session completed before its authored request.' }
            Start-Sleep -Milliseconds 5
        }
        throw 'Authored session request exceeded bounded wait.'
    }
    function Wait-ControlSession {
        param($Session)
        $deadline = [Diagnostics.Stopwatch]::StartNew()
        while (-not $Session.Completed -and $deadline.Elapsed.TotalSeconds -lt 5) { Start-Sleep -Milliseconds 5 }
        if (-not $Session.Completed) { throw 'Authored session completion exceeded bounded wait.' }
        return $Session.Result
    }
    function Assert-OwnedStop {
        param($Result)
        $Result.JobAssigned | Should-BeTrue
        $Result.ParentStopped | Should-BeTrue
        $Result.DescendantsStopped | Should-BeTrue
        $Result.StreamsComplete | Should-BeTrue
        $Result.InputWriterStopped | Should-BeTrue
        ([string]::IsNullOrEmpty($Result.StopError)) | Should-BeTrue
    }
    if (-not ('WinBookSplit.SessionTests.BlockingReader' -as [type])) {
        Add-Type -TypeDefinition @'
using System.IO;
using System.Threading;
namespace WinBookSplit.SessionTests {
    public sealed class BlockingReader : TextReader {
        readonly ManualResetEvent ready = new ManualResetEvent(false);
        public override string ReadLine() { ready.WaitOne(); return "authored released line"; }
        public void Release() { ready.Set(); }
    }
}
'@
    }
}

Describe 'M3 bounded supervised interaction channel' -Tag 'AC-060', 'AC-061', 'AC-062' {
    It 'queues two requests, preserves literal UTF-8 replies and leaves one ordinary terminal result' {
        $session = Start-ControlSession 'handshake'
        try {
            $first = Read-ControlRequest $session | ConvertFrom-Json
            $first.sequence | Should-Be -Expected 1
            $literal = ([char]0x00c5).ToString() + [char]0x65e5 + [char]0x672c + ' & % ! $(literal)'
            $first.label | Should-Be -Expected $literal
            $session.SendReply((@{value=$literal} | ConvertTo-Json -Compress))
            $second = Read-ControlRequest $session | ConvertFrom-Json
            $second.sequence | Should-Be -Expected 2
            $session.SendReply('{"action":"authored literal reply"}')
            $result = Wait-ControlSession $session
            Assert-OwnedStop $result
            $result.ExitCode | Should-Be -Expected 2
            @($result.ResultRecords).Count | Should-Be -Expected 1
            $result.InteractionCount | Should-Be -Expected 2
            $result.QueuedReplyCount | Should-Be -Expected 2
            $result.ReplyCount | Should-Be -Expected 2
            ([string]::IsNullOrEmpty($result.InteractionError)) | Should-BeTrue
            ([string]::IsNullOrEmpty($result.InputError)) | Should-BeTrue
            $result.Stdout.Contains('[WBS-INTERACTION] ') | Should-BeTrue
            $result.HumanStdout | Should-Be -Expected "Human before`r`n`r`nHuman final no-newline`r`n"
            $result.Stderr | Should-Be -Expected "AUTHORED_SESSION_STDERR`r`n"
            $frame = ConvertFrom-SplitResult -Stdout ($result.ResultRecords -join "`n") -ExitCode $result.ExitCode -Mode 'manual'
            $answers = $frame.message | ConvertFrom-Json
            $answers[0].value | Should-Be -Expected $literal
            $answers[1].action | Should-Be -Expected 'authored literal reply'
        }
        finally { $session.Dispose() }
    }
    It 'retains a large interaction independently of the bounded human tail' {
        $session = Start-ControlSession 'large-event'
        try {
            $request = Read-ControlRequest $session | ConvertFrom-Json
            ($request.pad.Length -gt 100000) | Should-BeTrue
            $session.SendReply('{"action":"reply"}')
            $result = Wait-ControlSession $session
            Assert-OwnedStop $result
            $result.StdoutTruncated | Should-BeTrue
            $result.HumanStdout | Should-Be -Expected "Human after event`r`n"
            @($result.ResultRecords).Count | Should-Be -Expected 1
            ([string]::IsNullOrEmpty($result.InteractionError)) | Should-BeTrue
        }
        finally { $session.Dispose() }
    }
    It 'bounds event queue overflow and proves owned shutdown' {
        $session = Start-ControlSession 'queue-flood'
        try {
            $result = Wait-ControlSession $session
            Assert-OwnedStop $result
            $result.InteractionError.Contains('bounded queue') | Should-BeTrue
            $result.HumanStdout.Contains('[WBS-INTERACTION] ') | Should-BeFalse
        }
        finally { $session.Dispose() }
    }
    It 'accepts exactly eight MiB of event JSON with CRLF and rejects one extra payload byte' {
        $session = Start-ControlSession 'exact-event-cap'
        try {
            $request = Read-ControlRequest $session
            ([Text.Encoding]::UTF8.GetByteCount($request)) | Should-Be -Expected 8388608
            $session.SendReply('{}')
            $result = Wait-ControlSession $session
            Assert-OwnedStop $result
            $result.InteractionCount | Should-Be -Expected 1
            $result.ReplyCount | Should-Be -Expected 1
            ([string]::IsNullOrEmpty($result.InteractionError)) | Should-BeTrue
            @($result.ResultRecords).Count | Should-Be -Expected 1
        }
        finally { $session.Dispose() }
        $session = Start-ControlSession 'one-over-event-cap'
        try {
            $result = Wait-ControlSession $session
            Assert-OwnedStop $result
            $result.InteractionError.Contains('bounded event limit') | Should-BeTrue
            $result.HumanStdout.Contains('[WBS-INTERACTION] ') | Should-BeFalse
        }
        finally { $session.Dispose() }
    }
    It 'rejects an event larger than eight MiB without an unbounded retained record' {
        $session = Start-ControlSession 'oversized-event'
        try {
            $result = Wait-ControlSession $session
            Assert-OwnedStop $result
            $result.InteractionError.Contains('bounded event limit') | Should-BeTrue
            $result.StdoutTruncated | Should-BeTrue
            $result.HumanStdout.Contains('[WBS-INTERACTION] ') | Should-BeFalse
        }
        finally { $session.Dispose() }
    }
    It 'keeps a deadline live while the caller withholds a reply and a descendant owns streams' {
        $session = Start-ControlSession 'wait-with-descendant' -Timeout .8
        try {
            $request = Read-ControlRequest $session | ConvertFrom-Json
            ($request.owned_child_pid -gt 0) | Should-BeTrue
            $result = Wait-ControlSession $session
            Assert-OwnedStop $result
            $result.TimedOut | Should-BeTrue
            $result.Cancelled | Should-BeFalse
            ($result.ElapsedSeconds -lt 3) | Should-BeTrue
        }
        finally { $session.Dispose() }
    }
    It 'distinguishes explicit session cancellation from its deadline' {
        $session = Start-ControlSession 'wait-with-descendant'
        try {
            $null = Read-ControlRequest $session
            $session.Cancel()
            $result = Wait-ControlSession $session
            Assert-OwnedStop $result
            $result.Cancelled | Should-BeTrue
            $result.TimedOut | Should-BeFalse
        }
        finally { $session.Dispose() }
    }
    It 'observes cancellation tokens independently of request polling' {
        $cancel = New-Object Threading.CancellationTokenSource
        $session = $null
        try {
            $session = Start-ControlSession 'ignore-stdin' -CancellationToken $cancel.Token
            $null = Read-ControlRequest $session
            $cancel.CancelAfter(150)
            $result = Wait-ControlSession $session
            Assert-OwnedStop $result
            $result.Cancelled | Should-BeTrue
            $result.TimedOut | Should-BeFalse
        }
        finally { if ($null -ne $session) { $session.Dispose() }; $cancel.Dispose() }
    }
    It 'keeps a blocked stdin writer off the caller and stops it at the owned deadline' {
        $session = Start-ControlSession 'ignore-stdin' -Timeout 1
        try {
            $null = Read-ControlRequest $session
            $clock = [Diagnostics.Stopwatch]::StartNew()
            $session.SendReply(('{"value":"' + ('Q' * 65000) + '"}'))
            ($clock.Elapsed.TotalSeconds -lt .5) | Should-BeTrue
            $result = Wait-ControlSession $session
            Assert-OwnedStop $result
            $result.TimedOut | Should-BeTrue
            $result.QueuedReplyCount | Should-Be -Expected 1
            $result.ReplyCount | Should-Be -Expected 0
            ([string]::IsNullOrEmpty($result.InputError)) | Should-BeTrue
        }
        finally { $session.Dispose() }
    }
    It 'rejects oversized, multi-line, NUL and invalid UTF-16 reply data' {
        $session = Start-ControlSession 'ignore-stdin'
        try {
            $null = Read-ControlRequest $session
            { $session.SendReply(('X' * 65537)) } | Should-Throw
            { $session.SendReply("line`nline") } | Should-Throw
            { $session.SendReply("line`rline") } | Should-Throw
            { $session.SendReply(('line' + [char]0 + 'line')) } | Should-Throw
            { $session.SendReply(([char]0xd800).ToString()) } | Should-Throw
            $session.Cancel()
            Assert-OwnedStop (Wait-ControlSession $session)
        }
        finally { $session.Dispose() }
    }
    It 'preserves NUL stdin for the ordinary noninteractive transport' {
        $result = Invoke-WinBookSplitProcess -Path $TestPython -WorkingDirectory $RepositoryRoot -TimeoutSeconds 5 `
            -Arguments @('-I','-B','-X','utf8',$script:sessionChild,'nul')
        Assert-OwnedStop $result
        $result.ExitCode | Should-Be -Expected 2
        $result.InteractionCount | Should-Be -Expected 0
        $result.ReplyCount | Should-Be -Expected 0
        $result.QueuedReplyCount | Should-Be -Expected 0
        ($null -eq $result.InteractionError) | Should-BeTrue
        ($null -eq $result.InputError) | Should-BeTrue
        $frame = ConvertFrom-SplitResult -Stdout ($result.ResultRecords -join "`n") -ExitCode 2 -Mode 'manual'
        $frame.message | Should-Be -Expected 'stdin EOF'
    }
}

Describe 'M3 background console line ownership' -Tag 'AC-060', 'AC-061' {
    It 'distinguishes an authored blank line from EOF' {
        foreach ($text in @("`n",'')) {
            $reader = New-Object IO.StringReader($text)
            $line = Start-WinBookSplitConsoleLine -Reader $reader
            try {
                $clock = [Diagnostics.Stopwatch]::StartNew()
                while (-not $line.Completed -and $clock.Elapsed.TotalSeconds -lt 2) { Start-Sleep -Milliseconds 5 }
                $line.Completed | Should-BeTrue
                if ($text.Length -eq 0) { ($null -eq $line.Value) | Should-BeTrue }
                else { ($null -ne $line.Value -and $line.Value.Length -eq 0) | Should-BeTrue }
                ([string]::IsNullOrEmpty($line.Error)) | Should-BeTrue
            }
            finally { $reader.Dispose() }
        }
    }
    It 'does not block supervision on a synchronized pending console reader or start a competing read' {
        $sourceReader = New-Object WinBookSplit.SessionTests.BlockingReader
        $reader = [IO.TextReader]::Synchronized($sourceReader)
        $session = $null
        $line = $null
        try {
            $clock = [Diagnostics.Stopwatch]::StartNew()
            $line = Start-WinBookSplitConsoleLine -Reader $reader
            ($clock.Elapsed.TotalSeconds -lt .5) | Should-BeTrue
            $line.Completed | Should-BeFalse
            { Start-WinBookSplitConsoleLine -Reader $reader } | Should-Throw
            $session = Start-ControlSession 'ignore-stdin' -Timeout .3
            $result = Wait-ControlSession $session
            Assert-OwnedStop $result
            $result.TimedOut | Should-BeTrue
            $line.Completed | Should-BeFalse
        }
        finally {
            $sourceReader.Release()
            $clock = [Diagnostics.Stopwatch]::StartNew()
            while ($null -ne $line -and -not $line.Completed -and $clock.Elapsed.TotalSeconds -lt 2) { Start-Sleep -Milliseconds 5 }
            if ($null -ne $line) {
                $line.Completed | Should-BeTrue
                $line.Value | Should-Be -Expected 'authored released line'
            }
            if ($null -ne $session) { $session.Dispose() }
            $reader.Dispose()
        }
    }
}
