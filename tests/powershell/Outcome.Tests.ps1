param([string]$RepositoryRoot, [string]$WorkRoot, [string]$TestPython)

BeforeAll {
    . (Join-Path $RepositoryRoot 'engine\WinBookSplit.Diagnostics.ps1')
    . (Join-Path $RepositoryRoot 'engine\WinBookSplit.Runtime.ps1')
    $tokens = $null; $errors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $RepositoryRoot 'WinBookSplit.ps1'), [ref]$tokens, [ref]$errors)
    if (@($errors).Count -ne 0) { throw 'The trusted application has syntax errors.' }
    foreach ($functionName in @('New-WinBookSplitFailure', 'New-WinBookSplitOutcome', 'Run-PythonSplitter')) {
        $functions = @($ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and
            $node.Name -ceq $functionName }, $true))
        if ($functions.Count -ne 1) { throw ('Expected one trusted application function: ' + $functionName) }
        . ([scriptblock]::Create($functions[0].Extent.Text))
    }
    function New-AuthoredResult {
        param([string]$Status, [string]$Code, [int]$ExitCode, $Execution = $null, $Diagnostic = $null)
        $count = 0
        if ($null -ne $Execution) { $count = $Execution.written_count }
        return [pscustomobject]@{ protocol='winbooksplit.result'; version=1; mode='manual'; status=$Status;
            code=$Code; exit_code=$ExitCode; message='Authored structural outcome fixture'; warnings=@();
            fallback_modes=@(); written_count=$count; execution=$Execution; diagnostic=$Diagnostic }
    }
    function New-AuthoredExecution {
        $runId = 'a' * 32
        $directory = Join-Path $WorkRoot 'authored-completed-folder'
        return [pscustomobject]@{ written_count=1; outputs=@([pscustomobject]@{ filename='01 - Synthetic.pdf' });
            coverage=[pscustomobject]@{ complete=$true }; run_id=$runId; final_directory=$directory;
            manifest_filename='WinBookSplit_Manifest.json'; manifest=[pscustomobject]@{ status='complete'; run_id=$runId;
                final_directory=$directory; written_count=1 } }
    }
    function New-AuthoredTransport {
        param($Result)
        $json = $Result | ConvertTo-Json -Depth 100 -Compress
        return [pscustomobject]@{ Cancelled=$false; TimedOut=$false; ParentStopped=$true; DescendantsStopped=$true;
            JobAssigned=$true; StreamsComplete=$true; ExitCode=$Result.exit_code; StartError=$null; StopError=$null;
            StreamError=$null; ResultError=$null; ResultRecords=@($json); Stdout=$json; Stderr=''; HumanStdout='';
            StdoutTotalBytes=[Text.Encoding]::UTF8.GetByteCount($json); StderrTotalBytes=0;
            StdoutTruncated=$false; StderrTruncated=$false; Pid=1234; ElapsedSeconds=0.1 }
    }
    function New-AuthoredFailingWriter {
        $script:authoredWriteCalls = 0
        $writer = New-Object PSObject
        $writer | Add-Member -MemberType ScriptMethod -Name WriteLine -Value {
            $script:authoredWriteCalls++
            if ($script:authoredWriteCalls -gt 1) { throw 'Authored console record failure after ENGINE intent.' }
        }
        return $writer
    }
}

Describe 'M3 independently pinned structured outcomes' -Tag 'AC-050', 'AC-052' {
    It 'maps every documented code with the actual host JSON integer representation' {
        $expected = @{
            split_complete=0;
            invalid_arguments=2; invalid_mode=2; invalid_start_pages=2; input_invalid=2; output_base_invalid=2;
            output_path_too_long=2; output_destination_changed=2; source_output_alias=2;
            dependency_missing=3; runtime_invalid=3; runtime_not_found=3; converter_invalid=3; converter_not_found=3; processor_start_failed=3;
            conversion_start_failed=4; conversion_failed=4; conversion_output_invalid=4; conversion_cleanup_failed=4;
            conversion_source_changed=4; conversion_ownership_failed=4;
            no_bookmarks=5; no_usable_bookmarks=5; no_bookmarks_at_level=5;
            invalid_document=6; invalid_outline=6; unreadable_document=6; invalid_plan=6; invalid_prepared_split=6;
            source_changed=6; output_exists=6; invalid_execution=6; output_write_failed=6; output_validation_failed=6;
            output_ownership_failed=6; output_manifest_failed=6; output_publish_failed=6; output_handle_close_failed=6;
            processor_protocol_failed=6; console_setup_failed=6; console_finalize_failed=6; processor_failed=6;
            unsupported_document=7;
            output_cancelled=130; conversion_cancelled=130; conversion_timeout=130; processing_cancelled=130;
            processor_cancelled=130; processor_timeout=130; cancelled=130
        }
        @($script:splitOutcomeContract.codes.PSObject.Properties).Count | Should-Be -Expected $expected.Count
        foreach ($entry in $expected.GetEnumerator()) { Get-SplitExitCode $entry.Key | Should-Be -Expected $entry.Value }
        Test-SplitInteger ([int]2) | Should-BeTrue
        Test-SplitInteger ([long]2) | Should-BeTrue
        Test-SplitInteger $true | Should-BeFalse
        Test-SplitInteger '2' | Should-BeFalse
        { Get-SplitExitCode 'unknown_outcome' } | Should-Throw
    }
    It 'requires explicit cancellation or timeout and rejects output claims on either' {
        foreach ($case in @(@('cancelled','output_cancelled'), @('timeout','conversion_timeout'))) {
            $result = New-AuthoredResult $case[0] $case[1] 130
            $json = $result | ConvertTo-Json -Depth 100 -Compress
            (ConvertFrom-SplitResult $json 130 'manual').status | Should-Be -Expected $case[0]
            { ConvertFrom-SplitResult $json 6 'manual' } | Should-Throw
            $result.status='error'
            { ConvertFrom-SplitResult ($result | ConvertTo-Json -Depth 100 -Compress) 130 'manual' } | Should-Throw
            $result.status=$case[0]; $result.execution=New-AuthoredExecution; $result.written_count=1
            { ConvertFrom-SplitResult ($result | ConvertTo-Json -Depth 100 -Compress) 130 'manual' } | Should-Throw
        }
    }
    It 'keeps primary conversion cancellation or timeout through failed cleanup' {
        foreach ($case in @(@('cancelled','conversion_cancelled'), @('timeout','conversion_timeout'))) {
            $result = New-AuthoredResult $case[0] 'conversion_cleanup_failed' 130 $null ([pscustomobject]@{ primary_code=$case[1] })
            (ConvertFrom-SplitResult ($result | ConvertTo-Json -Depth 100 -Compress) 130 'manual').status | Should-Be -Expected $case[0]
            $result.diagnostic.primary_code='conversion_failed'
            { ConvertFrom-SplitResult ($result | ConvertTo-Json -Depth 100 -Compress) 130 'manual' } | Should-Throw
        }
    }
    It 'validates retained incomplete execution with every completed-result guard' {
        $result = New-AuthoredResult 'incomplete' 'output_handle_close_failed' 6 (New-AuthoredExecution)
        (ConvertFrom-SplitResult ($result | ConvertTo-Json -Depth 100 -Compress) 6 'manual').written_count | Should-Be -Expected 1
        $result.execution.manifest.written_count=0
        { ConvertFrom-SplitResult ($result | ConvertTo-Json -Depth 100 -Compress) 6 'manual' } | Should-Throw
        $result = New-AuthoredResult 'incomplete' 'output_write_failed' 6 (New-AuthoredExecution)
        { ConvertFrom-SplitResult ($result | ConvertTo-Json -Depth 100 -Compress) 6 'manual' } | Should-Throw
    }
}

Describe 'M3 actual application functions under authored transport and log controls' -Tag 'AC-050', 'AC-051', 'AC-052' {
    BeforeEach {
        $script:enginePath=Join-Path $RepositoryRoot 'engine\winbooksplit_engine.py'
        $script:InputFile=Join-Path $WorkRoot 'authored-input.pdf'
        $script:outputDir=$WorkRoot; $script:pythonExe=$TestPython; $script:ProcessTimeout=30
        $script:consoleLogWriter=New-AuthoredFailingWriter
    }
    It 'retains completed engine evidence when the later console record write fails' {
        $script:authoredTransport=New-AuthoredTransport (New-AuthoredResult 'success' 'split_complete' 0 (New-AuthoredExecution))
        Mock Invoke-WinBookSplitProcess { $script:authoredTransport }
        $failure = $null
        try { $null=Run-PythonSplitter 'manual' '1' } catch { $failure=$_.Exception }
        $failure.Data['Code'] | Should-Be -Expected 'console_finalize_failed'
        $script:lastEngineResult.status | Should-Be -Expected 'success'
        $outcome=New-WinBookSplitOutcome 'incomplete' 'console_finalize_failed' $failure.Message 'manual' $script:lastEngineResult
        $outcome.exit_code | Should-Be -Expected 6
        $outcome.written_count | Should-Be -Expected 0
        $outcome.final_directory | Should-Be -Expected $script:lastEngineResult.execution.final_directory
        $outcome.engine_result.written_count | Should-Be -Expected 1
    }
    It 'preserves an engine failure through a secondary console record error' {
        $script:authoredTransport=New-AuthoredTransport (New-AuthoredResult 'write_error' 'output_write_failed' 6)
        Mock Invoke-WinBookSplitProcess { $script:authoredTransport }
        $result=Run-PythonSplitter 'manual' '1'
        $result.code | Should-Be -Expected 'output_write_failed'
        $result.exit_code | Should-Be -Expected 6
        $script:consoleRecordError | Should-MatchString -Expected 'Authored console record failure'
    }
    It 'retains completed evidence after supervisor finalization failure only when shutdown is proved' {
        $script:authoredTransport=New-AuthoredTransport (New-AuthoredResult 'success' 'split_complete' 0 (New-AuthoredExecution))
        $script:authoredTransport.StopError='Authored native handler or handle finalization failure.'
        Mock Invoke-WinBookSplitProcess { $script:authoredTransport }
        $failure=$null
        try { $null=Run-PythonSplitter 'manual' '1' } catch { $failure=$_.Exception }
        $failure.Data['Code'] | Should-Be -Expected 'processor_protocol_failed'
        $failure.Message | Should-MatchString -Expected 'Authored native handler or handle finalization failure'
        $script:lastEngineResult.status | Should-Be -Expected 'success'
        $script:lastEngineResult.execution.final_directory | Should-Be -Expected (Join-Path $WorkRoot 'authored-completed-folder')
        $outcome=New-WinBookSplitOutcome 'incomplete' $failure.Data['Code'] $failure.Message 'manual' $script:lastEngineResult
        $outcome.exit_code | Should-Be -Expected 6
        $outcome.written_count | Should-Be -Expected 0
        $outcome.final_directory | Should-Be -Expected $script:lastEngineResult.execution.final_directory
        $outcome.engine_result.written_count | Should-Be -Expected 1
        $outcome.message | Should-MatchString -Expected 'Authored native handler or handle finalization failure'

        foreach ($missingProof in @('ParentStopped','DescendantsStopped')) {
            $script:consoleLogWriter=New-AuthoredFailingWriter
            $script:authoredTransport=New-AuthoredTransport (New-AuthoredResult 'success' 'split_complete' 0 (New-AuthoredExecution))
            $script:authoredTransport.$missingProof=$false
            $script:authoredTransport.StopError='Authored uncertain owned shutdown.'
            $failure=$null
            try { $null=Run-PythonSplitter 'manual' '1' } catch { $failure=$_.Exception }
            $failure.Data['Code'] | Should-Be -Expected 'processor_protocol_failed'
            ($null -eq $script:lastEngineResult) | Should-BeTrue
        }
    }
    It 'preserves native cancellation or timeout through a secondary console record error' {
        foreach ($code in @('processor_cancelled','processor_timeout')) {
            $script:consoleLogWriter=New-AuthoredFailingWriter
            $script:authoredTransport=New-AuthoredTransport (New-AuthoredResult 'write_error' 'output_write_failed' 6)
            $script:authoredTransport.ResultRecords=@()
            $script:authoredTransport.Cancelled=($code -eq 'processor_cancelled')
            $script:authoredTransport.TimedOut=($code -eq 'processor_timeout')
            Mock Invoke-WinBookSplitProcess { $script:authoredTransport }
            $failure=$null
            try { $null=Run-PythonSplitter 'manual' '1' } catch { $failure=$_.Exception }
            $failure.Data['Code'] | Should-Be -Expected $code
            Get-SplitExitCode $failure.Data['Code'] | Should-Be -Expected 130
            ($null -eq $script:lastEngineResult) | Should-BeTrue
            $script:consoleRecordError | Should-MatchString -Expected 'Authored console record failure'
        }
    }
    It 'propagates preflight cancellation or timeout without trying another dependency candidate' {
        foreach ($code in @('processor_cancelled','processor_timeout')) {
            $script:authoredTransport=New-AuthoredTransport (New-AuthoredResult 'write_error' 'output_write_failed' 6)
            $script:authoredTransport.Cancelled=($code -eq 'processor_cancelled')
            $script:authoredTransport.TimedOut=($code -eq 'processor_timeout')
            Mock Invoke-WinBookSplitRuntimeProbe { $script:authoredTransport }
            Mock Get-WinBookSplitPathCandidates { throw 'A cancelled explicit probe cannot try another dependency.' }
            foreach ($converter in @($false,$true)) {
                $failure=$null
                try {
                    if ($converter) { $null=Resolve-WinBookSplitConverter -ApplicationRoot $RepositoryRoot -CalibrePath $TestPython }
                    else { $null=Resolve-WinBookSplitRuntime -ApplicationRoot $RepositoryRoot -PythonPath $TestPython }
                } catch { $failure=$_.Exception }
                $failure.Data['Code'] | Should-Be -Expected $code
                Get-SplitExitCode $failure.Data['Code'] | Should-Be -Expected 130
            }
            Should-Invoke -CommandName Get-WinBookSplitPathCandidates -Times 0 -Exactly
        }
    }
}
