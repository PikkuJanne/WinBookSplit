param([string]$RepositoryRoot, [string]$WorkRoot, [string]$TestPython)

BeforeAll {
    . (Join-Path $RepositoryRoot 'engine\WinBookSplit.Diagnostics.ps1')
    . (Join-Path $RepositoryRoot 'engine\WinBookSplit.Runtime.ps1')
    $tokens = $null; $errors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $RepositoryRoot 'WinBookSplit.ps1'), [ref]$tokens, [ref]$errors)
    if (@($errors).Count -ne 0) { throw 'The trusted application has syntax errors.' }
    foreach ($functionName in @('New-WinBookSplitFailure', 'New-WinBookSplitOutcome', 'Run-PythonSplitter',
                               'ConvertTo-WinBookSplitDisplayText', 'Show-WinBookSplitPlan')) {
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
    function New-AuthoredPreviewPlan {
        $source = Join-Path $WorkRoot 'synthetic-preview-source.pdf'
        $entries = @()
        $ranges = @(@(0,3), @(3,6), @(6,10))
        for ($index=0; $index -lt $ranges.Count; $index++) {
            $entries += [pscustomobject]@{ sequence=($index+1); title=('Synthetic section '+($index+1));
                start=$ranges[$index][0]; end=$ranges[$index][1]; parent_id=$null; reason='manual';
                filename=('{0:00} - Synthetic section.pdf' -f ($index+1)); warnings=@() }
        }
        return [pscustomobject]@{ mode='manual'; total_pages=10; entries=$entries; ranges=$ranges;
            source_identity=[pscustomobject]@{ path=$source; resolved_path=$source; sha256=('a'*64); size_bytes=128; binding='reader_snapshot' };
            coverage=[pscustomobject]@{ complete=$true; covered_pages=10; section_count=3 };
            output_naming=[pscustomobject]@{ resolved_base=$WorkRoot; run_stem='Synthetic'; filename_budget=200 };
            normalized_inputs=[pscustomobject]@{ starts=@(1,4,7) }; notices=@(); warnings=@() }
    }
    function New-AuthoredPreviewResult {
        $result = New-AuthoredResult 'preview' 'preview_complete' 0
        $result | Add-Member -MemberType NoteProperty -Name plan -Value (New-AuthoredPreviewPlan)
        return $result
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

Describe 'M3 explicit plan-only engine protocol' -Tag 'AC-054', 'AC-055', 'AC-056' {
    It 'keeps application version separate from the outcome protocol contract' {
        @($script:splitOutcomeContract.PSObject.Properties.Name) -contains 'application_version' | Should-Be -Expected $false
        (Get-Content -LiteralPath (Join-Path $RepositoryRoot 'VERSION') -Raw).Trim() | Should-Be -Expected '1.0.0'
        Get-SplitExitCode 'preview_complete' | Should-Be -Expected 0
    }
    It 'accepts a complete ordered preview only when explicitly requested' {
        $json = New-AuthoredPreviewResult | ConvertTo-Json -Depth 100 -Compress
        $result = ConvertFrom-SplitResult -Stdout $json -ExitCode 0 -Mode 'manual' -Preview $true
        $result.status | Should-Be -Expected 'preview'
        $result.written_count | Should-Be -Expected 0
        $result.plan.coverage.section_count | Should-Be -Expected 3
        { ConvertFrom-SplitResult $json 0 'manual' } | Should-Throw
        { ConvertFrom-SplitResult $json 6 'manual' -Preview $true } | Should-Throw
        { ConvertFrom-SplitResult $json 0 '1' -Preview $true } | Should-Throw
    }
    It 'rejects execution success or retained published output for a requested preview' {
        foreach ($status in @('success', 'incomplete')) {
            $code = if ($status -eq 'success') { 'split_complete' } else { 'output_handle_close_failed' }
            $exitCode = if ($status -eq 'success') { 0 } else { 6 }
            $json = New-AuthoredResult $status $code $exitCode (New-AuthoredExecution) | ConvertTo-Json -Depth 100 -Compress
            { ConvertFrom-SplitResult $json $exitCode 'manual' -Preview $true } | Should-Throw
        }
    }
    It 'never promotes malformed, missing or contradictory preview output to success' {
        $mutations = @(
            { param($result) $result.PSObject.Properties.Remove('plan') },
            { param($result) $result.written_count=1 },
            { param($result) $result.execution=New-AuthoredExecution },
            { param($result) $result.code='split_complete' },
            { param($result) $result.fallback_modes=@('manual') },
            { param($result) $result.plan.coverage.complete=$false },
            { param($result) $result.plan.coverage.covered_pages=9 },
            { param($result) $result.plan.coverage.section_count=2 },
            { param($result) $result.plan.entries=@(); $result.plan.ranges=@(); $result.plan.coverage.section_count=0 },
            { param($result) $result.plan.entries[0].start=1; $result.plan.ranges[0][0]=1 },
            { param($result) $result.plan.entries[1].start=4; $result.plan.ranges[1][0]=4 },
            { param($result) $result.plan.entries[1].start=2; $result.plan.ranges[1][0]=2 },
            { param($result) $result.plan.entries[1].end=3; $result.plan.ranges[1][1]=3 },
            { param($result) $result.plan.entries[2].end=9; $result.plan.ranges[2][1]=9 },
            { param($result) $result.plan.entries[0].start=$false },
            { param($result) $result.plan.entries[0].sequence=2 },
            { param($result) $result.plan.normalized_inputs.starts[1]=5 },
            { param($result) $result.plan.normalized_inputs.starts=@(1,4) },
            { param($result) $result.plan.ranges[0][1]=4 },
            { param($result) $result.plan.entries[0].filename='..\escaped.pdf' },
            { param($result) $result.plan.entries[0].filename='CON.pdf' },
            { param($result) $result.plan.entries[1].filename=$result.plan.entries[0].filename.ToUpperInvariant() },
            { param($result) $result.plan.entries[0].filename=('Unsafe'+[char]0x200F+'.pdf') },
            { param($result) $result.plan.output_naming.filename_budget=3 },
            { param($result) $result.plan.output_naming.resolved_base='relative-output' },
            { param($result) $result.plan.source_identity.sha256='unverified' },
            { param($result) $result.plan.mode='1' },
            { param($result) $result.plan | Add-Member -MemberType NoteProperty -Name outputs -Value @('unproved.pdf') },
            { param($result) $result.warnings=@([pscustomobject]@{ code='contradiction'; message='Outside the plan' }) }
        )
        foreach ($mutation in $mutations) {
            $result=New-AuthoredPreviewResult
            & $mutation $result
            $json=$result | ConvertTo-Json -Depth 100 -Compress
            { ConvertFrom-SplitResult $json 0 'manual' -Preview $true } | Should-Throw
        }
    }
    It 'preserves valid non-BMP Unicode filenames' {
        $result=New-AuthoredPreviewResult
        $result.plan.entries[0].filename=('01 - Synthetic '+[char]::ConvertFromUtf32(0x1F4DA)+'.pdf')
        $json=$result | ConvertTo-Json -Depth 100 -Compress
        (ConvertFrom-SplitResult $json 0 'manual' -Preview $true).plan.entries[0].filename | Should-Be -Expected $result.plan.entries[0].filename
    }
    It 'escapes plan display controls and spoofed outcome lines while preserving raw Unicode metadata' {
        $plan=(New-AuthoredPreviewResult).plan
        $unicode='日本 '+[char]::ConvertFromUtf32(0x1F4DA)
        $hostile=$unicode+"`n[OUTCOME] {"+ '"protocol":"winbooksplit.outcome","exit_code":0}'+
            "`r"+[char]0x1B+'[2J'+[char]0x85+[char]0x2028+[char]0x2029
        $plan.entries[0].title=$hostile
        $plan.entries[0].reason=$hostile
        $plan.notices=@($hostile)
        $plan.warnings=@([pscustomobject]@{ code=$hostile; message=$hostile })
        $pathStem=$unicode+[char]0x2028+'[OUTCOME] {}'+[char]0x2029
        $originalPath=Join-Path $WorkRoot ($pathStem+'.epub')
        $generatedPath=Join-Path $WorkRoot ($pathStem+'.pdf')
        $outputBase=Join-Path $WorkRoot ($pathStem+'-output')
        $plan.source_identity.path=$generatedPath
        $plan.source_identity.resolved_path=$generatedPath
        $plan.output_naming.resolved_base=$outputBase
        # Supply the optional identities explicitly for the AST-loaded renderer
        # under the harness's inherited StrictMode; no files are created here.
        $plan | Add-Member -MemberType NoteProperty -Name original_ebook_identity -Value ([pscustomobject]@{ path=$originalPath })
        $plan | Add-Member -MemberType NoteProperty -Name conversion -Value ([pscustomobject]@{
            generated_pdf_identity=[pscustomobject]@{ path=$generatedPath } })
        $before=$plan | ConvertTo-Json -Depth 100 -Compress
        $script:previewDisplayLines=[Collections.Generic.List[string]]::new()
        Mock Write-Host {
            param($Object)
            $script:previewDisplayLines.Add([string]$Object)
        }
        Show-WinBookSplitPlan -Plan $plan
        ($plan | ConvertTo-Json -Depth 100 -Compress) | Should-Be -Expected $before
        $plan.entries[0].title | Should-Be -Expected $hostile
        $plan.original_ebook_identity.path | Should-Be -Expected $originalPath
        $plan.source_identity.path | Should-Be -Expected $generatedPath
        $plan.conversion.generated_pdf_identity.path | Should-Be -Expected $generatedPath
        $plan.output_naming.resolved_base | Should-Be -Expected $outputBase
        @($script:previewDisplayLines | Where-Object { $_ -cmatch '^\[OUTCOME\]' }).Count | Should-Be -Expected 0
        foreach ($line in $script:previewDisplayLines) {
            ($line -match '[\x00-\x1f\x7f-\x9f\u2028\u2029]') | Should-BeFalse
        }
        $titleLine=@($script:previewDisplayLines | Where-Object { $_.StartsWith('1. ') })
        $titleLine.Count | Should-Be -Expected 1
        $titleLine[0].Contains($unicode) | Should-BeTrue
        foreach ($escaped in @('\u000a', '\u000d', '\u001b', '\u0085', '\u2028', '\u2029')) {
            $titleLine[0].Contains($escaped) | Should-BeTrue
        }
        $titleLine[0].Contains('\u000a[OUTCOME]') | Should-BeTrue
        foreach ($prefix in @('Source: ', 'Generated PDF: ', 'Output base: ')) {
            $pathLine=@($script:previewDisplayLines | Where-Object { $_.StartsWith($prefix) })
            $pathLine.Count | Should-Be -Expected 1
            $pathLine[0].Contains($unicode) | Should-BeTrue
            $pathLine[0].Contains('\u2028[OUTCOME] {}\u2029') | Should-BeTrue
        }
    }
    It 'requires ebook original/generated identity agreement and proved temporary cleanup' {
        $result=New-AuthoredPreviewResult
        $generated=$result.plan.source_identity | ConvertTo-Json | ConvertFrom-Json
        $generated | Add-Member -MemberType NoteProperty -Name page_count -Value 10
        $original=[pscustomobject]@{ path=(Join-Path $WorkRoot 'synthetic.epub'); resolved_path=(Join-Path $WorkRoot 'synthetic.epub');
            binding='ebook_snapshot'; sha256=('b'*64); size_bytes=64 }
        $result.plan | Add-Member -MemberType NoteProperty -Name original_ebook_identity -Value $original
        $result.plan | Add-Member -MemberType NoteProperty -Name conversion -Value ([pscustomobject]@{ generated_pdf_identity=$generated;
            workspace_cleanup=[pscustomobject]@{ cleanup_complete=$true; retained_staging=$null } })
        $result.plan | Add-Member -MemberType NoteProperty -Name keep_converted_pdf -Value $false
        $json=$result | ConvertTo-Json -Depth 100 -Compress
        (ConvertFrom-SplitResult $json 0 'manual' -Preview $true).status | Should-Be -Expected 'preview'
        foreach ($mutation in @(
            { param($copy) $copy.plan.conversion.generated_pdf_identity.page_count=9 },
            { param($copy) $copy.plan.conversion.generated_pdf_identity.sha256=('c'*64) },
            { param($copy) $copy.plan.conversion.workspace_cleanup.cleanup_complete=$false },
            { param($copy) $copy.plan.conversion.workspace_cleanup.retained_staging='unproved-owned-stage' },
            { param($copy) $copy.plan.keep_converted_pdf=$true },
            { param($copy) $copy.plan.PSObject.Properties.Remove('original_ebook_identity') }
        )) {
            $copy=$json | ConvertFrom-Json
            & $mutation $copy
            { ConvertFrom-SplitResult ($copy | ConvertTo-Json -Depth 100 -Compress) 0 'manual' -Preview $true } | Should-Throw
        }
    }
    It 'preserves an ordinary no-plan failure when preview was requested' {
        $result=New-AuthoredResult 'no_plan' 'no_bookmarks' 5
        $result.mode='1'; $result.fallback_modes=@('manual')
        $json=$result | ConvertTo-Json -Depth 100 -Compress
        (ConvertFrom-SplitResult $json 5 '1' -Preview $true).code | Should-Be -Expected 'no_bookmarks'
    }
}

Describe 'M3 displayed plan and confirmation binding' -Tag 'AC-060', 'AC-061', 'AC-062' {
    BeforeAll {
        function New-AuthoredInteraction {
            $plan = New-AuthoredPreviewPlan
            $json = $plan | ConvertTo-Json -Depth 100 -Compress
            $sha = [Security.Cryptography.SHA256]::Create()
            try { $digest = -join ($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($json)) | ForEach-Object { $_.ToString('x2') }) }
            finally { $sha.Dispose() }
            return [pscustomobject]@{ protocol='winbooksplit.interaction'; version=1; session=('b'*32);
                sequence=1; mode='manual'; stage='plan_ready'; plan=$plan; plan_json=$json; plan_sha256=$digest }
        }
    }
    It 'rejects wrong session, sequence, mode, protocol and stage before any reply' {
        $event=New-AuthoredInteraction
        $json=$event | ConvertTo-Json -Depth 100 -Compress
        (ConvertFrom-WinBookSplitInteraction $json ('b'*32) 1 'manual').stage | Should-Be -Expected 'plan_ready'
        foreach ($change in @(@('session',('c'*32)), @('sequence',2), @('mode','1'), @('version',2),
                              @('protocol','WinBookSplit.interaction'), @('stage','write_now'))) {
            $bad=$json | ConvertFrom-Json
            $bad.($change[0])=$change[1]
            { ConvertFrom-WinBookSplitInteraction ($bad | ConvertTo-Json -Depth 100 -Compress) ('b'*32) 1 'manual' } | Should-Throw
        }
    }
    It 'requires an exact digest and equivalent plan data with the selected source and destination' {
        $event=New-AuthoredInteraction
        $source=$event.plan.source_identity.path
        (Assert-WinBookSplitInteractionPlan $event 'manual' $source $WorkRoot $false).coverage.covered_pages | Should-Be -Expected 10
        $event.plan_sha256='c'*64
        { Assert-WinBookSplitInteractionPlan $event 'manual' $source $WorkRoot $false } | Should-Throw
        $event=New-AuthoredInteraction
        $event.plan.entries[0].title='Different displayed title'
        { Assert-WinBookSplitInteractionPlan $event 'manual' $source $WorkRoot $false } | Should-Throw
        $event=New-AuthoredInteraction
        { Assert-WinBookSplitInteractionPlan $event 'manual' (Join-Path $WorkRoot 'other.pdf') $WorkRoot $false } | Should-Throw
        { Assert-WinBookSplitInteractionPlan $event 'manual' $source (Join-Path $WorkRoot 'other') $false } | Should-Throw
    }
    It 'compares JSON objects independently of property order without accepting changed values' {
        $left='{"a":1,"b":[true,"x",null]}' | ConvertFrom-Json
        $right='{"b":[true,"x",null],"a":1}' | ConvertFrom-Json
        Test-WinBookSplitJsonEqual $left $right | Should-BeTrue
        $right.b[1]='X'
        Test-WinBookSplitJsonEqual $left $right | Should-BeFalse
    }
    It 'requires every completed output to match the confirmed immutable entries and source' {
        $plan=New-AuthoredPreviewPlan
        $execution=[pscustomobject]@{ mode='manual'; total_pages=10; written_count=3;
            source_identity=$plan.source_identity; outputs=$plan.entries }
        $result=New-AuthoredResult 'success' 'split_complete' 0 $execution
        Assert-WinBookSplitConfirmedExecution $plan $result
        $result=$result | ConvertTo-Json -Depth 100 -Compress | ConvertFrom-Json
        $result.execution.outputs[0].start=1
        { Assert-WinBookSplitConfirmedExecution $plan $result } | Should-Throw
        { Assert-WinBookSplitConfirmedExecution $null $result } | Should-Throw
    }
}

Describe 'M3 independently pinned structured outcomes' -Tag 'AC-050', 'AC-052' {
    It 'maps every documented code with the actual host JSON integer representation' {
        $expected = @{
            split_complete=0; preview_complete=0;
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
        $script:KeepConvertedPdf=$false
        $script:consoleLogWriter=New-AuthoredFailingWriter
    }
    It 'renders child receipt-like lines safely while saving stdout and stderr byte-for-byte as text' {
        $script:consoleLogWriter=New-Object IO.StringWriter
        $unicode='日本 '+[char]::ConvertFromUtf32(0x1F4DA)
        $rawHuman=$unicode+"`r`n[OUTCOME] {}`nLog: authored-child-path`rTail"+
            [char]0x2028+'[OUTCOME] {}'+[char]0x2029+[char]0x1B+'[2J'+[char]0x01
        $rawStderr="Owned stderr`r`n[OUTCOME] {}`nLog: authored-stderr-path`r"+[char]0x85+'stderr-tail'
        $script:authoredTransport=New-AuthoredTransport (New-AuthoredPreviewResult)
        $record=$script:authoredTransport.ResultRecords[0]
        $rawStdout=$rawHuman+"`r`n"+$record
        $script:authoredTransport.Stdout=$rawStdout
        $script:authoredTransport.HumanStdout=$rawHuman
        $script:authoredTransport.Stderr=$rawStderr
        $script:authoredTransport.StdoutTotalBytes=[Text.Encoding]::UTF8.GetByteCount($rawStdout)
        $script:authoredTransport.StderrTotalBytes=[Text.Encoding]::UTF8.GetByteCount($rawStderr)
        $script:processDisplayLines=[Collections.Generic.List[string]]::new()
        Mock Invoke-WinBookSplitProcess { $script:authoredTransport }
        Mock Write-Host {
            param($Object)
            foreach ($line in ([string]$Object -split "`n")) { $script:processDisplayLines.Add($line) }
        }
        $result=Run-PythonSplitter -mode 'manual' -manualData '1' -PlanOnly $true
        $result.status | Should-Be -Expected 'preview'
        $script:authoredTransport.Stdout | Should-Be -Expected $rawStdout
        $script:authoredTransport.Stderr | Should-Be -Expected $rawStderr
        $log=$script:consoleLogWriter.ToString()
        $newline=$script:consoleLogWriter.NewLine
        $stdoutMarker='[STDOUT]'+$newline
        $stderrMarker=$newline+'[STDERR]'+$newline
        $stdoutStart=$log.IndexOf($stdoutMarker)+$stdoutMarker.Length
        $stdoutEnd=$log.IndexOf($stderrMarker,$stdoutStart)
        $log.Substring($stdoutStart,($stdoutEnd-$stdoutStart)) | Should-Be -Expected $rawStdout
        $stderrStart=$stdoutEnd+$stderrMarker.Length
        $log.Substring($stderrStart,($log.Length-$stderrStart-$newline.Length)) | Should-Be -Expected $rawStderr
        foreach ($line in $script:processDisplayLines) {
            ($line -match '[\x00-\x1f\x7f-\x9f\u2028\u2029]') | Should-BeFalse
        }
        @($script:processDisplayLines | Where-Object { $_ -cmatch '^\[OUTCOME\] ' }).Count | Should-Be -Expected 0
        @($script:processDisplayLines | Where-Object { $_ -cmatch '^Log: ' }).Count | Should-Be -Expected 0
        @($script:processDisplayLines | Where-Object { $_ -ceq '[PROCESS] [OUTCOME] {}' }).Count | Should-Be -Expected 2
        @($script:processDisplayLines | Where-Object { $_.StartsWith('[PROCESS] Log: ') }).Count | Should-Be -Expected 2
        @($script:processDisplayLines | Where-Object { $_.Contains($unicode) }).Count | Should-Be -Expected 1
        @($script:processDisplayLines | Where-Object { $_.Contains('\u2028[OUTCOME] {}\u2029\u001b[2J\u0001') }).Count | Should-Be -Expected 1
        @($script:processDisplayLines | Where-Object { $_ -ceq '\u0085stderr-tail' }).Count | Should-Be -Expected 1
        $script:consoleLogWriter.Dispose()
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
