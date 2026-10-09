# Pure protocol and decision functions; dot-sourcing never starts processing.
function Test-SplitInteger {
    param($Value)
    return ($Value -is [int] -or $Value -is [long])
}

function ConvertTo-NativeArgument {
    param([AllowEmptyString()][string]$Value)
    # Windows CommandLineToArgvW/CRT quoting: double trailing slashes and
    # slashes preceding a quote, so user text remains exactly one argument.
    $escaped = [regex]::Replace($Value, '(\\*)"', '$1$1\"')
    $escaped = [regex]::Replace($escaped, '(\\+)$', '$1$1')
    return '"' + $escaped + '"'
}

function ConvertFrom-SplitResult {
    param([string]$Stdout, [int]$ExitCode, [string]$Mode)
    $records = @()
    foreach ($line in ($Stdout -split '\r?\n')) {
        if ($line.TrimStart().StartsWith('{')) {
            $candidate = $line | ConvertFrom-Json -ErrorAction Stop
            if ($candidate.protocol -is [string] -and $candidate.protocol -ceq 'winbooksplit.result') { $records += $candidate }
        }
    }
    if ($records.Count -ne 1) { throw 'Expected exactly one structured engine result.' }
    $result = $records[0]
    foreach ($field in @('protocol', 'version', 'mode', 'status', 'code', 'message',
                         'warnings', 'fallback_modes', 'exit_code', 'written_count', 'execution')) {
        if ($result.PSObject.Properties.Name -notcontains $field) { throw "Missing result field: $field" }
    }
    if ($result.protocol -isnot [string] -or $result.protocol -cne 'winbooksplit.result' -or
        -not (Test-SplitInteger $result.version) -or $result.version -ne 1 -or
        $result.mode -isnot [string] -or $result.mode -cne $Mode -or
        $Mode -cnotin @('1', '2', 'manual') -or $result.message -isnot [string] -or
        [string]::IsNullOrWhiteSpace($result.message) -or $result.code -isnot [string] -or
        [string]::IsNullOrWhiteSpace($result.code) -or
        -not (Test-SplitInteger $result.exit_code) -or $result.exit_code -ne $ExitCode -or
        $result.status -isnot [string] -or
        -not (Test-SplitInteger $result.written_count) -or $result.warnings -isnot [array] -or
        $result.fallback_modes -isnot [array]) { throw 'Invalid engine result metadata.' }
    $expectedFallback = @()
    switch -CaseSensitive ($result.status) {
        'success' {
            if ($ExitCode -ne 0 -or $result.code -cne 'split_complete' -or $result.written_count -lt 1 -or
                $null -eq $result.execution -or -not (Test-SplitInteger $result.execution.written_count) -or
                $result.execution.written_count -ne $result.written_count -or
                $result.execution.outputs -isnot [array] -or
                @($result.execution.outputs).Count -ne $result.written_count -or
                @($result.execution.outputs | Where-Object { $null -eq $_ }).Count -gt 0 -or
                $result.execution.coverage.complete -isnot [bool] -or
                -not $result.execution.coverage.complete -or
                $result.execution.run_id -isnot [string] -or $result.execution.run_id -cnotmatch '^[0-9a-f]{32}$' -or
                $result.execution.final_directory -isnot [string] -or
                -not [IO.Path]::IsPathRooted($result.execution.final_directory) -or
                $result.execution.manifest_filename -cne 'WinBookSplit_Manifest.json' -or
                $null -eq $result.execution.manifest -or $result.execution.manifest.status -cne 'complete' -or
                $result.execution.manifest.run_id -cne $result.execution.run_id -or
                $result.execution.manifest.final_directory -cne $result.execution.final_directory -or
                $result.execution.manifest.written_count -ne $result.written_count) { throw 'Invalid successful engine result.' }
        }
        'no_plan' {
            if ($ExitCode -ne 55 -or $result.code -cnotin @('no_bookmarks', 'no_usable_bookmarks', 'no_bookmarks_at_level') -or
                $Mode -eq 'manual') { throw 'Invalid no-plan engine result.' }
            $expectedFallback = @('manual')
            if ($result.code -ceq 'no_bookmarks_at_level') {
                if ($Mode -ne '2') { throw 'Level 2 fallback requires Level 2 mode.' }
                $expectedFallback = @('1', 'manual')
            }
        }
        'invalid_input' {
            if ($ExitCode -ne 1 -or $result.code -cnotin @('invalid_document', 'invalid_outline', 'invalid_start_pages', 'invalid_mode', 'invalid_arguments')) {
                throw 'Invalid input-failure result.'
            }
        }
        'read_error' {
            if ($ExitCode -ne 1 -or $result.code -cne 'unreadable_document') { throw 'Invalid read-failure result.' }
        }
        'write_error' {
            if ($ExitCode -ne 1 -or $result.code -cne 'output_write_failed') { throw 'Invalid write-failure result.' }
        }
        'error' {
            if ($ExitCode -ne 1 -or $result.code -cnotin @('invalid_plan', 'invalid_prepared_split',
                'source_changed', 'source_output_alias', 'output_exists', 'invalid_execution',
                'output_validation_failed', 'output_ownership_failed', 'output_base_invalid',
                'output_manifest_failed', 'output_publish_failed', 'output_cancelled',
                'output_handle_close_failed', 'output_path_too_long',
                'output_destination_changed')) { throw 'Invalid failure result.' }
        }
        default { throw 'Unknown engine result category.' }
    }
    if ($result.status -cne 'success' -and ($result.written_count -ne 0 -or $null -ne $result.execution)) {
        throw 'A failed result cannot claim successful output.'
    }
    if (@($result.fallback_modes).Count -ne $expectedFallback.Count) { throw 'Invalid fallback choices.' }
    for ($index = 0; $index -lt $expectedFallback.Count; $index++) {
        if ($result.fallback_modes[$index] -isnot [string] -or
            $result.fallback_modes[$index] -cne $expectedFallback[$index]) { throw 'Invalid fallback choice.' }
    }
    return $result
}

function Get-SplitDecision {
    param($Result, [AllowEmptyString()][string]$Choice = '')
    $message = $Result.message
    switch -CaseSensitive ($Result.code) {
        'no_bookmarks' { $message = 'The PDF has no outline bookmarks. Manual start pages are available.' }
        'no_usable_bookmarks' { $message = 'The outline exists, but has no usable Level 1 destinations. Manual start pages are available.' }
        'no_bookmarks_at_level' { $message = 'No usable direct Level 2 bookmarks were found. Choose Level 1 or manual start pages.' }
        'invalid_document' { $message = 'The PDF contains no physical pages and cannot be split.' }
    }
    $decision, $retryMode = 'cancel', $null
    $allowed = @($Result.fallback_modes)
    if ($allowed.Count -gt 0) {
        switch ($Choice.Trim().ToUpperInvariant()) {
            '' { $decision = 'pending' }
            { $_ -in @('M', 'Y') } {
                if ($allowed -contains 'manual') { $decision, $retryMode = 'retry', 'manual' }
                else { $decision = 'invalid' }
            }
            '1' {
                if ($allowed -contains '1') { $decision, $retryMode = 'retry', '1' }
                else { $decision = 'invalid' }
            }
            { $_ -in @('N', 'C') } { $decision = 'cancel' }
            default { $decision = 'invalid' }
        }
    }
    return [pscustomobject]@{ message = $message; fallback_modes = $allowed;
                             decision = $decision; retry_mode = $retryMode }
}
