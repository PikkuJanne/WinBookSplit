param([string]$RepositoryRoot, [string]$WorkRoot, [string]$TestPython)

BeforeAll {
    . (Join-Path $RepositoryRoot 'engine\WinBookSplit.Support.ps1')
    Initialize-WinBookSplitSupport
    $supportRoot = Join-Path $WorkRoot ('support-' + [guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $supportRoot
    $secret = 'PRIVATE_SENTINEL_7c4c_path_title_message_token'
    function New-SupportFixture {
        return [ordered]@{
            protocol='winbooksplit.run'; version=1; run_id=('a'*32); application_version='1.0.0-dev'
            started_utc='2026-10-10T00:00:00.0000000Z'; finished_utc='2026-10-10T00:00:01.0000000Z'; diagnostics_finalized=$true
            runtime_versions=[ordered]@{ powershell='5.1.26100.9444'; python='3.14.8'; pypdf='6.19.0'; calibre=$null }
            settings=[ordered]@{ mode='manual'; input_kind='pdf'; preview=$false; non_interactive=$true; no_pause=$true; keep_converted_pdf=$false
                conversion_timeout=120; process_timeout=300; normalized_inputs=@{ starts=@(1,3); authored_secret=$secret } }
            source_identity=@{ path=('C:\private\'+$secret+'.pdf'); resolved_path=$secret; sha256=$secret; size_bytes=100; binding=$secret }
            plan=@{ total_pages=4; ranges=@(@(0,2),@(2,4)); coverage=@{ complete=$true; covered_pages=4; section_count=2 }
                source_identity=@{ path=$secret }; titles=@($secret); entries=@(@{ title=$secret; filename=$secret });
                output_naming=@{ resolved_base=$secret }; normalized_inputs=@{ text=$secret }; notices=@($secret); warnings=@($secret) }
            engine_result=@{ protocol='winbooksplit.result'; message=$secret; argv=@($secret); diagnostic=@{ message=$secret; password=$secret }; execution=@{ final_directory=$secret; outputs=@($secret) } }
            outcome=@{ status='success'; code='split_complete'; exit_code=0; written_count=2; message=$secret; console_log=$secret; final_directory=$secret; engine_result=@{ secret=$secret } }
            warnings=@{ planning=@(@{ code='invalid_destination'; message=$secret; title=$secret; path=$secret });
                parser=@(@{ code='pypdf_parser_warning'; category='pdf_parser'; severity='WARNING'; message=$secret });
                parser_suppressed_count=3; parser_message_truncated_count=1 }
            log=@{ filename=$secret; max_bytes=50331648; body_limit_bytes=33554432; bytes_written=512; limit_reached=$false }
        }
    }
    function Write-SupportFixture {
        param($Fixture, [string]$Leaf='WinBookSplit_Run.json')
        $directory = Join-Path $supportRoot ([guid]::NewGuid().ToString('N'))
        $null = New-Item -ItemType Directory -Path $directory
        $path = Join-Path $directory $Leaf
        [IO.File]::WriteAllText($path, ($Fixture | ConvertTo-Json -Depth 60), (New-Object Text.UTF8Encoding($false, $true)))
        return $path
    }
    function New-SupportDestination { return (Join-Path $supportRoot (([guid]::NewGuid().ToString('N'))+'.json')) }
    function Assert-SupportRejected {
        param($Fixture)
        $path=Write-SupportFixture $Fixture; $destination=New-SupportDestination
        $hash=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
        { Export-WinBookSplitSupportSummary $path $destination } | Should -Throw
        Test-Path -LiteralPath $destination | Should -BeFalse
        (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash | Should -BeExactly $hash
    }
}

Describe 'Explicit allowlisted support diagnostic export' {
    It 'exports only safe fields and preserves an original containing secrets in every discarded detail' {
        $path=Write-SupportFixture (New-SupportFixture); $destination=New-SupportDestination
        $before=Get-Item -LiteralPath $path; $hash=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
        $summary=Export-WinBookSplitSupportSummary $path $destination
        $text=[IO.File]::ReadAllText($destination,[Text.Encoding]::UTF8)
        $text | Should -Not -Match $secret
        $text | Should -Not -Match '"(sha256|run_id|started_utc|finished_utc|filename|message|argv|normalized_inputs|source_identity|diagnostics_finalized)"\s*:'
        @($summary.Keys) -join ',' | Should -BeExactly 'protocol,version,application_version,runtime_versions,settings,plan,outcome,warnings,log'
        $summary.protocol | Should -BeExactly 'winbooksplit.support'
        $summary.runtime_versions.powershell | Should -BeExactly '5.1'
        $summary.outcome.written_count | Should -Be 2
        $summary.plan.ranges.Count | Should -Be 2
        $summary.plan.coverage.covered_pages | Should -Be 4
        $summary.warnings.planning[0].category | Should -BeExactly 'bookmark'
        $summary.warnings.parser[0].code | Should -BeExactly 'pypdf_parser_warning'
        (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash | Should -BeExactly $hash
        (Get-Item -LiteralPath $path).LastWriteTimeUtc.Ticks | Should -Be $before.LastWriteTimeUtc.Ticks
        $bytes=[IO.File]::ReadAllBytes($destination)
        ($bytes[0] -eq 239 -and $bytes[1] -eq 187 -and $bytes[2] -eq 191) | Should -BeFalse
    }
    It 'exports only fixed PDF fidelity warning tokens without private metadata or annotation text' {
        foreach ($row in @(@('cross_chapter_link_dropped','navigation'),@('navigation_link_dropped','navigation'),
                @('article_navigation_dropped','navigation'),@('annotation_dropped','annotation'),
                @('annotation_relation_dropped','annotation'),@('metadata_omitted','metadata'),@('metadata_normalized','metadata'))) {
            $fixture=New-SupportFixture
            $fixture.warnings.planning=@(@{ code=$row[0]; message=$secret; source_order=0; depth=0; title=$secret; author=$secret })
            $summary=Export-WinBookSplitSupportSummary (Write-SupportFixture $fixture) (New-SupportDestination)
            $summary.warnings.planning.Count | Should -Be 1
            $summary.warnings.planning[0].code | Should -BeExactly $row[0]
            $summary.warnings.planning[0].category | Should -BeExactly $row[1]
            ($summary | ConvertTo-Json -Depth 32 -Compress) | Should -Not -Match $secret
        }
    }
    It 'preserves failed cancelled no-plan and incomplete zero-count outcomes without fabricating success' {
        foreach ($row in @(@('failed','invalid_document',6),@('cancelled','cancelled',130),@('timeout','processor_timeout',130),@('incomplete','console_finalize_failed',6),@('no_plan','no_bookmarks',5),@('invalid_input','invalid_document',6),@('read_error','unreadable_document',6),@('write_error','output_write_failed',6),@('error','source_changed',6),@('unsupported','unsupported_document',7))) {
            $fixture=New-SupportFixture; $fixture.plan=$null
            $fixture.outcome.status=$row[0]; $fixture.outcome.code=$row[1]; $fixture.outcome.exit_code=$row[2]; $fixture.outcome.written_count=0
            $summary=Export-WinBookSplitSupportSummary (Write-SupportFixture $fixture) (New-SupportDestination)
            $summary.outcome.status | Should -BeExactly $row[0]
            $summary.outcome.exit_code | Should -Be $row[2]
            $summary.outcome.written_count | Should -Be 0
            $summary.plan | Should -BeNullOrEmpty
        }
    }
    It 'rejects private strings in runtime versions including apparently numeric versions outside the trusted pins' {
        foreach ($name in @('powershell','python','pypdf','calibre')) {
            foreach ($value in @($secret,'123456.789012.345678')) {
                $fixture=New-SupportFixture; $fixture.runtime_versions[$name]=$value
                Assert-SupportRejected $fixture
            }
        }
    }
    It 'rejects coerced booleans and numbers and secret enum values in the selected settings' {
        foreach ($name in @('preview','non_interactive','no_pause','keep_converted_pdf','conversion_timeout','process_timeout','mode','input_kind')) {
            $fixture=New-SupportFixture; $fixture.settings[$name]=$secret; Assert-SupportRejected $fixture
        }
        foreach ($value in @($true,'120',120.5,-1,0,172801)) {
            $fixture=New-SupportFixture; $fixture.settings.process_timeout=$value; Assert-SupportRejected $fixture
        }
        foreach ($value in @(0,86401)) { $fixture=New-SupportFixture; $fixture.settings.conversion_timeout=$value; Assert-SupportRejected $fixture }
    }
    It 'normalizes portable PowerShell versions to fixed buckets and accepts the actual timeout limits' {
        foreach ($version in @('5.1.19041.1000','5.1.26100.9444','7.4.6','7.6.5')) {
            $fixture=New-SupportFixture; $fixture.runtime_versions.powershell=$version
            $fixture.settings.conversion_timeout=86400; $fixture.settings.process_timeout=172800
            $summary=Export-WinBookSplitSupportSummary (Write-SupportFixture $fixture) (New-SupportDestination)
            $expected='7'; if($version.StartsWith('5.1.')){$expected='5.1'}
            $summary.runtime_versions.powershell | Should -BeExactly $expected
            $summary.settings.process_timeout | Should -Be 172800
        }
        foreach ($version in @('7.99999.0','7.01.0','5.1.0',('7.4.6-'+$secret))) {
            $fixture=New-SupportFixture; $fixture.runtime_versions.powershell=$version; Assert-SupportRejected $fixture
        }
    }
    It 'rejects unknown warning tokens and conflicting parser severities while discarding warning messages' {
        $fixture=New-SupportFixture; $fixture.warnings.planning[0].code=$secret; Assert-SupportRejected $fixture
        foreach ($name in @('code','category','severity')) {
            $fixture=New-SupportFixture; $fixture.warnings.parser[0][$name]=$secret; Assert-SupportRejected $fixture
        }
        $fixture=New-SupportFixture; $fixture.warnings.parser[0].severity='ERROR'; Assert-SupportRejected $fixture
        $fixture=New-SupportFixture; $fixture.warnings.parser_suppressed_count=$secret; Assert-SupportRejected $fixture
    }
    It 'rejects false success and contradictory status code count or exit bindings' {
        foreach ($change in @(@('written_count',0),@('written_count',1),@('exit_code',6),@('code','invalid_document'),@('status','failed'),@('status',$secret))) {
            $fixture=New-SupportFixture; $fixture.outcome[$change[0]]=$change[1]; Assert-SupportRejected $fixture
        }
        $fixture=New-SupportFixture; $fixture.plan=$null; Assert-SupportRejected $fixture
        $fixture=New-SupportFixture; $fixture.outcome.status='cancelled'; $fixture.outcome.code='cancelled'; $fixture.outcome.exit_code=130
        Assert-SupportRejected $fixture
    }
    It 'rejects omissions overlaps empty ranges and invalid coverage in the physical page partition' {
        foreach ($ranges in @(@(@(1,2),@(2,4)),@(@(0,3),@(2,4)),@(@(0,2),@(2,2)),@(@(0,2)))) {
            $fixture=New-SupportFixture; $fixture.plan.ranges=$ranges; Assert-SupportRejected $fixture
        }
        $fixture=New-SupportFixture; $fixture.plan.coverage.complete='true'; Assert-SupportRejected $fixture
        $fixture=New-SupportFixture; $fixture.plan.coverage.section_count=3; Assert-SupportRejected $fixture
        $fixture=New-SupportFixture; $fixture.plan.total_pages=$true; Assert-SupportRejected $fixture
    }
    It 'rejects missing unknown future and unfinished root schemas before creating a summary' {
        $fixture=New-SupportFixture; $fixture.extra=$secret; Assert-SupportRejected $fixture
        $fixture=New-SupportFixture; $fixture.version=2; Assert-SupportRejected $fixture
        $fixture=New-SupportFixture; $fixture.protocol=$secret; Assert-SupportRejected $fixture
        $fixture=New-SupportFixture; $fixture.diagnostics_finalized=$false; Assert-SupportRejected $fixture
        $fixture=New-SupportFixture; $fixture.diagnostics_finalized='true'; Assert-SupportRejected $fixture
        $fixture=New-SupportFixture; $fixture.Remove('diagnostics_finalized'); Assert-SupportRejected $fixture
        $path=Write-SupportFixture (New-SupportFixture) 'WinBookSplit_Run.pending.json'; $destination=New-SupportDestination
        { Export-WinBookSplitSupportSummary $path $destination } | Should -Throw
        Test-Path -LiteralPath $destination | Should -BeFalse
    }
    It 'rejects malformed UTF8 JSON duplicate keys ambiguous case and unpaired Unicode before writing' {
        foreach ($text in @('{','{"protocol":1,"protocol":2}','{"protocol":1,"Protocol":2}','{"x":"\ud800"}','{"x":1} trailing')) {
            $path=Write-SupportFixture (New-SupportFixture)
            [IO.File]::WriteAllText($path,$text,(New-Object Text.UTF8Encoding($false,$true)))
            $destination=New-SupportDestination
            { Export-WinBookSplitSupportSummary $path $destination } | Should -Throw
            Test-Path -LiteralPath $destination | Should -BeFalse
        }
        $path=Write-SupportFixture (New-SupportFixture); [IO.File]::WriteAllBytes($path,[byte[]]@(123,34,120,34,58,34,195,40,34,125))
        $destination=New-SupportDestination
        { Export-WinBookSplitSupportSummary $path $destination } | Should -Throw
        Test-Path -LiteralPath $destination | Should -BeFalse
    }
    It 'rejects manifests above the 32MiB cap without allocating an output' {
        $path=Write-SupportFixture (New-SupportFixture)
        $stream=[IO.File]::Open($path,[IO.FileMode]::Open,[IO.FileAccess]::Write,[IO.FileShare]::None)
        try { $stream.SetLength(33554433) } finally { $stream.Dispose() }
        $destination=New-SupportDestination
        { Export-WinBookSplitSupportSummary $path $destination } | Should -Throw
        Test-Path -LiteralPath $destination | Should -BeFalse
        (Get-Item -LiteralPath $path).Length | Should -Be 33554433
    }
    It 'never overwrites an existing destination or the source through direct or case aliases' {
        $path=Write-SupportFixture (New-SupportFixture); $destination=New-SupportDestination
        [IO.File]::WriteAllText($destination,$secret)
        $sourceHash=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
        $destinationHash=(Get-FileHash -LiteralPath $destination -Algorithm SHA256).Hash
        foreach ($target in @($destination,$path,$path.ToUpperInvariant())) { { Export-WinBookSplitSupportSummary $path $target } | Should -Throw }
        (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash | Should -BeExactly $sourceHash
        (Get-FileHash -LiteralPath $destination -Algorithm SHA256).Hash | Should -BeExactly $destinationHash
    }
    It 'rejects an existing destination hardlink while preserving both original names' {
        $path=Write-SupportFixture (New-SupportFixture); $destination=New-SupportDestination
        $null=New-Item -ItemType HardLink -Path $destination -Target $path
        $hash=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
        { Export-WinBookSplitSupportSummary $path $destination } | Should -Throw
        (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash | Should -BeExactly $hash
        (Get-FileHash -LiteralPath $destination -Algorithm SHA256).Hash | Should -BeExactly $hash
    }
    It 'rejects reparse ancestors of either source or destination without following them' {
        $path=Write-SupportFixture (New-SupportFixture)
        $junction=Join-Path $supportRoot ('junction-'+[guid]::NewGuid().ToString('N'))
        $null=New-Item -ItemType Junction -Path $junction -Target ([IO.Path]::GetDirectoryName($path))
        $destination=New-SupportDestination
        { Export-WinBookSplitSupportSummary (Join-Path $junction 'WinBookSplit_Run.json') $destination } | Should -Throw
        Test-Path -LiteralPath $destination | Should -BeFalse
        $throughJunction=Join-Path $junction 'summary.json'
        { Export-WinBookSplitSupportSummary $path $throughJunction } | Should -Throw
        Test-Path -LiteralPath $throughJunction | Should -BeFalse
    }
    It 'holds real directory read leases that deny replacement until disposed' {
        $path=Write-SupportFixture (New-SupportFixture); $destination=New-SupportDestination
        $parent=[IO.Path]::GetDirectoryName($path); $moved=$parent+'-moved'
        foreach ($target in @($parent,$moved)) { $target.StartsWith($supportRoot+'\',[StringComparison]::OrdinalIgnoreCase) | Should -BeTrue }
        $files=New-Object WinBookSplitSupport.NativeFiles($path,$destination)
        try {
            { [IO.Directory]::Move($parent,$moved) } | Should -Throw
            Test-Path -LiteralPath $parent -PathType Container | Should -BeTrue
            Test-Path -LiteralPath $moved | Should -BeFalse
            $source=$files.OpenSource(); try { $source.Length | Should -BeGreaterThan 0 } finally { $source.Dispose() }
        }
        finally { $files.Dispose() }
        [IO.Directory]::Move($parent,$moved)
        Test-Path -LiteralPath $moved -PathType Container | Should -BeTrue
        [IO.Directory]::Move($moved,$parent)
    }
    It 'rejects relative paths missing parents and alternate data streams with no implicit directories' {
        $path=Write-SupportFixture (New-SupportFixture); $destination=New-SupportDestination
        foreach ($source in @('WinBookSplit_Run.json',($path+':private'))) { { Export-WinBookSplitSupportSummary $source $destination } | Should -Throw }
        foreach ($target in @('summary.json',($destination+':private'),(Join-Path $supportRoot 'absent\summary.json'))) { { Export-WinBookSplitSupportSummary $path $target } | Should -Throw }
        Test-Path -LiteralPath $destination | Should -BeFalse
        Test-Path -LiteralPath (Join-Path $supportRoot 'absent') | Should -BeFalse
    }
    It 'refuses a source with a competing write handle instead of taking an inconsistent snapshot' {
        $path=Write-SupportFixture (New-SupportFixture); $destination=New-SupportDestination
        $stream=[IO.File]::Open($path,[IO.FileMode]::Open,[IO.FileAccess]::Write,[IO.FileShare]::ReadWrite)
        try { { Export-WinBookSplitSupportSummary $path $destination } | Should -Throw } finally { $stream.Dispose() }
        Test-Path -LiteralPath $destination | Should -BeFalse
    }
}
