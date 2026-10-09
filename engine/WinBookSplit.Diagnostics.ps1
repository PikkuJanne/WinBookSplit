# Pure protocol and decision functions; dot-sourcing never starts processing.
$script:splitOutcomeContract = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'WinBookSplit.Outcomes.json') -Raw -Encoding UTF8 | ConvertFrom-Json -ErrorAction Stop
if ($script:splitOutcomeContract.schema_version -ne 1 -or $null -eq $script:splitOutcomeContract.codes) {
    throw 'Invalid shipped outcome contract.'
}

function Get-SplitExitCode {
    param([string]$Code)
    $property = $script:splitOutcomeContract.codes.PSObject.Properties[$Code]
    if ($null -eq $property -or -not (Test-SplitInteger $property.Value) -or $property.Value -notin @(0, 2, 3, 4, 5, 6, 7, 130)) {
        throw ('Unknown or invalid outcome code: ' + $Code)
    }
    return [int]$property.Value
}

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

function Assert-SplitPreviewSourceIdentity {
    param($Identity, [string]$Binding)
    if ($null -eq $Identity) { throw 'Missing preview source identity.' }
    foreach ($field in @('path', 'resolved_path', 'sha256', 'size_bytes', 'binding')) {
        if ($Identity.PSObject.Properties.Name -notcontains $field) { throw ('Missing preview identity field: ' + $field) }
    }
    if ($Identity.path -isnot [string] -or -not [IO.Path]::IsPathRooted($Identity.path) -or
        $Identity.resolved_path -isnot [string] -or -not [IO.Path]::IsPathRooted($Identity.resolved_path) -or
        $Identity.sha256 -isnot [string] -or $Identity.sha256 -cnotmatch '^[0-9a-f]{64}$' -or
        -not (Test-SplitInteger $Identity.size_bytes) -or $Identity.size_bytes -lt 1 -or
        $Identity.binding -isnot [string] -or $Identity.binding -cne $Binding) { throw 'Invalid preview source identity.' }
}

function Assert-SplitPreviewPlan {
    param($Plan, [string]$Mode)
    if ($null -eq $Plan) { throw 'Missing preview plan.' }
    foreach ($field in @('mode', 'total_pages', 'source_identity', 'entries', 'ranges', 'normalized_inputs',
                         'notices', 'warnings', 'coverage', 'output_naming')) {
        if ($Plan.PSObject.Properties.Name -notcontains $field) { throw ('Missing preview plan field: ' + $field) }
    }
    if ($Plan.mode -isnot [string] -or $Plan.mode -cne $Mode -or -not (Test-SplitInteger $Plan.total_pages) -or $Plan.total_pages -lt 1 -or
        $Plan.entries -isnot [array] -or $Plan.entries.Count -lt 1 -or $Plan.ranges -isnot [array] -or
        $Plan.ranges.Count -ne $Plan.entries.Count -or $Plan.notices -isnot [array] -or $Plan.warnings -isnot [array] -or
        $null -eq $Plan.normalized_inputs -or $Plan.coverage.complete -isnot [bool] -or -not $Plan.coverage.complete -or
        -not (Test-SplitInteger $Plan.coverage.covered_pages) -or $Plan.coverage.covered_pages -ne $Plan.total_pages -or
        -not (Test-SplitInteger $Plan.coverage.section_count) -or $Plan.coverage.section_count -ne $Plan.entries.Count -or
        $Plan.output_naming.resolved_base -isnot [string] -or -not [IO.Path]::IsPathRooted($Plan.output_naming.resolved_base) -or
        $Plan.output_naming.run_stem -isnot [string] -or [string]::IsNullOrWhiteSpace($Plan.output_naming.run_stem) -or
        $Plan.output_naming.run_stem -match '[<>:"/\\|?*\x00-\x1f]' -or $Plan.output_naming.run_stem.Length -gt 64 -or
        $Plan.output_naming.run_stem -cne $Plan.output_naming.run_stem.Trim(' ', '.') -or
        -not (Test-SplitInteger $Plan.output_naming.filename_budget) -or
        $Plan.output_naming.filename_budget -lt 1 -or $Plan.output_naming.filename_budget -gt 255) { throw 'Invalid preview plan metadata.' }
    foreach ($notice in $Plan.notices) { if ($notice -isnot [string]) { throw 'Invalid preview notice.' } }
    foreach ($field in @('outputs', 'execution', 'written_count', 'final_directory')) {
        if ($Plan.PSObject.Properties.Name -contains $field) { throw 'A preview plan cannot claim emitted output.' }
    }
    Assert-SplitPreviewSourceIdentity $Plan.source_identity 'reader_snapshot'
    if ($Mode -ceq 'manual') {
        if ($Plan.normalized_inputs.starts -isnot [array] -or
            $Plan.normalized_inputs.starts.Count -ne $Plan.entries.Count) { throw 'Invalid preview manual starts.' }
    }
    elseif ($Plan.normalized_inputs.bookmarks -isnot [array]) { throw 'Invalid preview normalized bookmarks.' }
    $filenames = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $previousEnd = 0
    for ($index = 0; $index -lt $Plan.entries.Count; $index++) {
        $entry = $Plan.entries[$index]
        if ($null -eq $entry) { throw 'Missing preview section.' }
        foreach ($field in @('sequence', 'title', 'start', 'end', 'parent_id', 'reason', 'filename', 'warnings')) {
            if ($entry.PSObject.Properties.Name -notcontains $field) { throw ('Missing preview section field: ' + $field) }
        }
        $range = $Plan.ranges[$index]
        $filename = $entry.filename
        if ($Mode -ceq 'manual' -and (-not (Test-SplitInteger $Plan.normalized_inputs.starts[$index]) -or
            $Plan.normalized_inputs.starts[$index] -ne ($entry.start + 1))) { throw 'Preview manual starts disagree with its sections.' }
        if (-not (Test-SplitInteger $entry.sequence) -or $entry.sequence -ne ($index + 1) -or
            -not (Test-SplitInteger $entry.start) -or -not (Test-SplitInteger $entry.end) -or
            $entry.start -ne $previousEnd -or $entry.start -ge $entry.end -or $entry.end -gt $Plan.total_pages -or
            $range -isnot [array] -or $range.Count -ne 2 -or
            -not (Test-SplitInteger $range[0]) -or -not (Test-SplitInteger $range[1]) -or
            $range[0] -ne $entry.start -or $range[1] -ne $entry.end -or
            $entry.title -isnot [string] -or $entry.reason -isnot [string] -or
            ($null -ne $entry.parent_id -and $entry.parent_id -isnot [string]) -or $entry.warnings -isnot [array] -or
            $filename -isnot [string] -or -not $filename.EndsWith('.pdf', [StringComparison]::Ordinal) -or
            $filename -in @('.pdf', '..pdf') -or $filename -cne $filename.Trim() -or
            $filename -match '[<>:"/\\|?*\x00-\x1f]' -or $filename.Length -gt $Plan.output_naming.filename_budget -or
            $filename -match '^(CON|PRN|AUX|NUL|CONIN\$|CONOUT\$|COM[1-9¹²³]|LPT[1-9¹²³])\s*\.' -or
            -not $filenames.Add($filename)) { throw 'Invalid preview section partition or filename.' }
        for ($character = 0; $character -lt $filename.Length; $character++) {
            $category = [Globalization.CharUnicodeInfo]::GetUnicodeCategory($filename, $character)
            if ($category -in @([Globalization.UnicodeCategory]::Control, [Globalization.UnicodeCategory]::Format,
                               [Globalization.UnicodeCategory]::Surrogate)) { throw 'Invalid preview filename character.' }
            if ([char]::IsHighSurrogate($filename[$character])) { $character++ }
        }
        $previousEnd = $entry.end
    }
    if ($previousEnd -ne $Plan.total_pages) { throw 'Preview sections must cover the final physical page.' }
    if ($Plan.PSObject.Properties.Name -contains 'original_ebook_identity' -or
        $Plan.PSObject.Properties.Name -contains 'conversion' -or $Plan.PSObject.Properties.Name -contains 'keep_converted_pdf') {
        Assert-SplitPreviewSourceIdentity $Plan.original_ebook_identity 'ebook_snapshot'
        $generated = $Plan.conversion.generated_pdf_identity
        Assert-SplitPreviewSourceIdentity $generated 'reader_snapshot'
        if ($Plan.keep_converted_pdf -isnot [bool] -or $Plan.keep_converted_pdf -or
            -not (Test-SplitInteger $generated.page_count) -or $generated.page_count -ne $Plan.total_pages -or
            $generated.path -cne $Plan.source_identity.path -or $generated.resolved_path -cne $Plan.source_identity.resolved_path -or
            $generated.sha256 -cne $Plan.source_identity.sha256 -or $generated.size_bytes -ne $Plan.source_identity.size_bytes -or
            $Plan.conversion.workspace_cleanup.cleanup_complete -isnot [bool] -or
            -not $Plan.conversion.workspace_cleanup.cleanup_complete -or
            $null -ne $Plan.conversion.workspace_cleanup.retained_staging) { throw 'Invalid converted preview source or cleanup.' }
    }
}

function ConvertFrom-SplitResult {
    param([string]$Stdout, [int]$ExitCode, [string]$Mode, [bool]$Preview = $false)
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
    $mappedExit = Get-SplitExitCode -Code $result.code
    if ($result.code -ceq 'conversion_cleanup_failed' -and
        $result.diagnostic.primary_code -cin @('conversion_cancelled', 'conversion_timeout')) { $mappedExit = 130 }
    if ($ExitCode -ne $mappedExit) { throw 'Outcome code does not match the documented native exit.' }
    if ($ExitCode -eq 130 -and $result.status -cnotin @('cancelled', 'timeout')) { throw 'Cancellation requires its explicit result category.' }
    switch -CaseSensitive ($result.status) {
        'preview' {
            if (-not $Preview -or $ExitCode -ne 0 -or $result.code -cne 'preview_complete' -or
                $result.written_count -ne 0 -or $null -ne $result.execution -or
                $result.PSObject.Properties.Name -notcontains 'plan') { throw 'Invalid or unexpected preview result.' }
            Assert-SplitPreviewPlan -Plan $result.plan -Mode $Mode
            if (($result.warnings | ConvertTo-Json -Depth 100 -Compress) -cne
                ($result.plan.warnings | ConvertTo-Json -Depth 100 -Compress)) { throw 'Preview warnings disagree with its plan.' }
        }
        'success' {
            if ($Preview -or $ExitCode -ne 0 -or $result.code -cne 'split_complete' -or $result.written_count -lt 1 -or
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
        'incomplete' {
            if ($Preview -or $ExitCode -ne 6 -or $result.code -cne 'output_handle_close_failed' -or $null -eq $result.execution) {
                throw 'Invalid incomplete published result.'
            }
            # Apply every ordinary completed-output check to the retained folder.
            $complete = $result | ConvertTo-Json -Depth 100 -Compress | ConvertFrom-Json -ErrorAction Stop
            $complete.status = 'success'; $complete.code = 'split_complete'; $complete.exit_code = 0
            $null = ConvertFrom-SplitResult -Stdout ($complete | ConvertTo-Json -Depth 100 -Compress) -ExitCode 0 -Mode $Mode
        }
        'no_plan' {
            if ($ExitCode -ne 5 -or $result.code -cnotin @('no_bookmarks', 'no_usable_bookmarks', 'no_bookmarks_at_level') -or
                $Mode -eq 'manual') { throw 'Invalid no-plan engine result.' }
            $expectedFallback = @('manual')
            if ($result.code -ceq 'no_bookmarks_at_level') {
                if ($Mode -ne '2') { throw 'Level 2 fallback requires Level 2 mode.' }
                $expectedFallback = @('1', 'manual')
            }
        }
        'invalid_input' {
            if ($result.code -cnotin @('input_invalid', 'invalid_document', 'invalid_outline', 'invalid_start_pages', 'invalid_mode', 'invalid_arguments')) {
                throw 'Invalid input-failure result.'
            }
        }
        'read_error' {
            if ($ExitCode -ne 6 -or $result.code -cne 'unreadable_document') { throw 'Invalid read-failure result.' }
        }
        'write_error' {
            if ($ExitCode -ne 6 -or $result.code -cne 'output_write_failed') { throw 'Invalid write-failure result.' }
        }
        'error' {
            if ($result.code -cnotin @('invalid_plan', 'invalid_prepared_split', 'dependency_missing',
                'source_changed', 'source_output_alias', 'output_exists', 'invalid_execution',
                'output_validation_failed', 'output_ownership_failed', 'output_base_invalid',
                'output_manifest_failed', 'output_publish_failed', 'output_cancelled',
                'output_handle_close_failed', 'output_path_too_long',
                'output_destination_changed', 'converter_not_found',
                'conversion_start_failed', 'conversion_failed', 'conversion_output_invalid',
                'conversion_timeout', 'conversion_cancelled', 'conversion_cleanup_failed',
                'conversion_source_changed', 'conversion_ownership_failed')) { throw 'Invalid failure result.' }
        }
        'cancelled' {
            if ($ExitCode -ne 130 -or $result.code -cnotin @('output_cancelled', 'conversion_cancelled',
                'processing_cancelled', 'conversion_cleanup_failed') -or
                ($result.code -ceq 'conversion_cleanup_failed' -and $result.diagnostic.primary_code -cne 'conversion_cancelled')) {
                throw 'Invalid cancelled result.'
            }
        }
        'timeout' {
            if ($ExitCode -ne 130 -or $result.code -cnotin @('conversion_timeout', 'conversion_cleanup_failed') -or
                ($result.code -ceq 'conversion_cleanup_failed' -and $result.diagnostic.primary_code -cne 'conversion_timeout')) {
                throw 'Invalid timeout result.'
            }
        }
        'unsupported' {
            if ($ExitCode -ne 7 -or $result.code -cne 'unsupported_document') { throw 'Invalid unsupported-document result.' }
        }
        default { throw 'Unknown engine result category.' }
    }
    if ($result.status -cnotin @('success', 'incomplete') -and ($result.written_count -ne 0 -or $null -ne $result.execution)) {
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
