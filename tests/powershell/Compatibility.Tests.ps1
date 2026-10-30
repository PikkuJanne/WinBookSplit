param([string]$RepositoryRoot, [string]$WorkRoot, [string]$TestPython)

BeforeAll {
    . (Join-Path $RepositoryRoot 'engine\WinBookSplit.Diagnostics.ps1')
    . (Join-Path $RepositoryRoot 'engine\WinBookSplit.Compatibility.ps1')
    . (Join-Path $RepositoryRoot 'engine\WinBookSplit.Runtime.ps1')
    $script:policy = Get-WinBookSplitCompatibilityPolicy
    function Copy-CompatibilityPolicy { return ($script:policy | ConvertTo-Json -Depth 32 | ConvertFrom-Json) }
    function New-CompatibilityRecord {
        param([string]$Version = '3.14.8')
        $parts = @($Version.Split('.') | ForEach-Object { [int]$_ })
        return [ordered]@{ version=$Version; version_info=$parts; releaselevel='final'; serial=0;
            implementation='cpython'; platform='win32'; machine='AMD64'; bits=64; gil_disabled=$false;
            isolated=$true; dont_write_bytecode=$true; executable='C:\Trusted Python\python.exe'; ok=$false }
    }
    function Write-CompatibilityPolicy {
        param($Policy, [AllowEmptyString()][string]$Text)
        $path = Join-Path $TestDrive (([guid]::NewGuid().ToString('N')) + '.json')
        if (-not $PSBoundParameters.ContainsKey('Text')) { $Text = $Policy | ConvertTo-Json -Depth 32 -Compress }
        [IO.File]::WriteAllText($path, $Text, (New-Object Text.UTF8Encoding($false, $true)))
        return $path
    }
}

Describe 'V101 offline policy and numeric maintenance releases' -Tag 'V101-AC-011','V101-AC-012','V101-AC-016' {
    It 'accepts floor, current and synthetic later stable patches numerically' {
        foreach ($version in @('3.15.0','3.15.1','3.15.9','3.15.10','3.15.99','3.14.8','3.14.9','3.14.10','3.14.99','3.13.16','3.13.17','3.13.99')) {
            $result = Test-WinBookSplitInterpreterRecord (New-CompatibilityRecord $version) $script:policy
            $result.Ok | Should -BeTrue
            $result.Reason | Should -BeExactly 'supported'
            $result.Message | Should -Match ([regex]::Escape($version))
            (Test-WinBookSplitPythonVersion $version $script:policy).Reason | Should -BeExactly $result.Reason
        }
    }
    It 'explains security floors and unsupported branches instead of an exact patch pin' {
        foreach ($version in @('3.14.2','3.14.6','3.14.7','3.13.14','3.13.15')) {
            $result = Test-WinBookSplitInterpreterRecord (New-CompatibilityRecord $version) $script:policy
            $result.Ok | Should -BeFalse
            $result.Reason | Should -BeExactly 'python_security_floor'
            $result.Message | Should -Match 'security floor'
            $result.Message | Should -Not -Match 'exactly'
        }
        foreach ($version in @('3.11.16','3.12.14','3.16.0','4.0.0')) {
            (Test-WinBookSplitPythonVersion $version $script:policy).Reason | Should -BeExactly 'unsupported_python'
        }
        (Test-WinBookSplitPythonVersion '3.12.14' $script:policy).Message | Should -Match 'source-only'
    }
    It 'uses policy data for precise blocked patches and compatibility floors' {
        $blocked = Copy-CompatibilityPolicy
        $blockedBranch = @($blocked.python.branches | Where-Object { $_.minor -eq 14 })[0]
        $blockedBranch.blocked_patches = @(8)
        $blockedBranch.recommended_patch = 9
        (Test-WinBookSplitPythonVersion '3.14.8' $blocked).Reason | Should -BeExactly 'python_security_floor'
        (Test-WinBookSplitPythonVersion '3.14.9' $blocked).Ok | Should -BeTrue
        $higher = Copy-CompatibilityPolicy
        $higherBranch = @($higher.python.branches | Where-Object { $_.minor -eq 14 })[0]
        $higherBranch.compatibility_min_patch = 10
        $higherBranch.recommended_patch = 10
        (Test-WinBookSplitPythonVersion '3.14.8' $higher).Reason | Should -BeExactly 'python_compatibility_floor'
        (Test-WinBookSplitPythonVersion '3.14.10' $higher).Ok | Should -BeTrue
    }
    It 'formats supported floors from the policy without requiring a runtime record' {
        Get-WinBookSplitPythonPolicyDescription $script:policy | Should -BeExactly 'regular Windows x64 CPython 3.15.0+ or 3.14.8+ or 3.13.16+; later stable patches accepted'
        $higher = Copy-CompatibilityPolicy
        $higherBranch = @($higher.python.branches | Where-Object { $_.minor -eq 14 })[0]
        $higherBranch.security_min_patch = 9
        $higherBranch.recommended_patch = 9
        Get-WinBookSplitPythonPolicyDescription $higher | Should -Match '3\.14\.9\+'
    }
}

Describe 'V101 malformed policy fails closed' -Tag 'V101-AC-013' {
    It 'loads the whole policy with strict duplicate-safe JSON and preserves its bytes' {
        $path = Join-Path $RepositoryRoot 'engine\WinBookSplit.Compatibility.json'
        $before = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
        $loaded = Get-WinBookSplitCompatibilityPolicy $path
        $loaded.schema_version | Should -Be 1
        $loaded.policy_version | Should -Be 6
        $loaded.python.branches.Count | Should -Be 3
        (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash | Should -BeExactly $before
    }
    It 'binds the returned policy to its exact read bytes without changing the default API' {
        $path = Join-Path $RepositoryRoot 'engine\WinBookSplit.Compatibility.json'
        $before = [IO.File]::ReadAllBytes($path)
        $expected = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
        $identity = Get-WinBookSplitCompatibilityPolicy -PolicyPath $path -WithIdentity
        $identity.Sha256 | Should -BeExactly $expected
        $identity.Sha256 | Should -Match '^[0-9a-f]{64}$'
        ($identity.Policy | ConvertTo-Json -Depth 32 -Compress) | Should -BeExactly ($script:policy | ConvertTo-Json -Depth 32 -Compress)
        [Convert]::ToBase64String([IO.File]::ReadAllBytes($path)) | Should -BeExactly ([Convert]::ToBase64String($before))
        # A valid reformat has the same policy values and a different byte identity.
        $reformatted = Write-CompatibilityPolicy -Policy $script:policy
        $compact = Get-WinBookSplitCompatibilityPolicy -PolicyPath $reformatted -WithIdentity
        $compact.Sha256 | Should -BeExactly ((Get-FileHash -LiteralPath $reformatted -Algorithm SHA256).Hash.ToLowerInvariant())
        $compact.Sha256 | Should -Not -BeExactly $expected
    }
    It 'rejects missing empty oversized invalid UTF8 nonfinite duplicate and case-colliding keys' {
        { Get-WinBookSplitCompatibilityPolicy (Join-Path $TestDrive 'missing.json') } | Should -Throw
        foreach ($text in @('', (' ' * 32769), '[]', '[{}]', '{"schema_version":NaN}', '{"schema_version":1,"schema_version":1}', '{"schema_version":1,"Schema_version":1}', '{"schema_version":1} trailing')) {
            $path = Write-CompatibilityPolicy -Text $text
            { Get-WinBookSplitCompatibilityPolicy $path } | Should -Throw
        }
        $path = Join-Path $TestDrive 'bad-utf8.json'
        [IO.File]::WriteAllBytes($path, [byte[]]@(123,255,125))
        { Get-WinBookSplitCompatibilityPolicy $path } | Should -Throw
        $path = Join-Path $TestDrive 'utf8-bom.json'
        [IO.File]::WriteAllBytes($path, [byte[]](@(239,187,191) + [Text.Encoding]::UTF8.GetBytes(($script:policy | ConvertTo-Json -Depth 32))))
        { Get-WinBookSplitCompatibilityPolicy $path } | Should -Throw
        # PS 5.1 serializes integral doubles as JSON integers. Author the token
        # literally to verify the loader rejects the current policy version as a float on both hosts.
        $token = '"policy_version":' + [string]$script:policy.policy_version
        $text = ($script:policy | ConvertTo-Json -Depth 32 -Compress).Replace($token, ($token + '.0'))
        $text | Should -Match ([regex]::Escape($token + '.0'))
        { Get-WinBookSplitCompatibilityPolicy (Write-CompatibilityPolicy -Text $text) } | Should -Throw
    }
    It 'rejects schema types, bad dates, unknown fields, arrays and malformed identity' {
        $mutations = @(
            { param($p) $p.schema_version=$true }, { param($p) $p.schema_version=2 },
            { param($p) $p.policy_version=0 }, { param($p) $p.policy_version=1.5 },
            { param($p) $p.policy_version=@(1) },
            { param($p) $p.reviewed_on='2026-02-30' }, { param($p) $p.reviewed_on='2026-1-1' },
            { param($p) $p | Add-Member -NotePropertyName unknown -NotePropertyValue 1 },
            { param($p) $p.python.bits=$true }, { param($p) $p.python.bits=@(64) }, { param($p) $p.python.stable_only=1 },
            { param($p) $p.python.gil_disabled=$null }, { param($p) $p.python.implementation=@('cpython') },
            { param($p) $p.python.branches=@() }, { param($p) $p.python.branches[0].minor='14' },
            { param($p) $p.python.branches[0].security_min_patch=-1 },
            { param($p) $p.python.branches[0].security_min_patch=1000001 },
            { param($p) $p.python.branches[0].security_min_patch=9 },
            { param($p) $p.python.branches[0].blocked_patches=@(0) },
            { param($p) $p.python.branches[0].blocked_patches=@(9,9) },
            { param($p) $p.python.branches[0].blocked_patches=@($false) },
            { param($p) $p.python.branches+=@($p.python.branches[0]) },
            { param($p) $p.excluded_python_branches[0].minor=14 },
            { param($p) $p.excluded_python_branches+=@($p.excluded_python_branches[0]) },
            { param($p) $p.pypdf.version='6.20' }, { param($p) $p.pypdf.wheel_sha256=('A'*64) },
            { param($p) $p.calibre.minimum_version=$true }
        )
        foreach ($mutate in $mutations) {
            $policy = Copy-CompatibilityPolicy; & $mutate $policy
            { Get-WinBookSplitCompatibilityPolicy (Write-CompatibilityPolicy $policy) } | Should -Throw
            (Test-WinBookSplitPythonVersion '3.14.8' $policy).Reason | Should -BeExactly 'malformed_policy'
            (Test-WinBookSplitInterpreterRecord (New-CompatibilityRecord) $policy).Reason | Should -BeExactly 'malformed_policy'
        }
    }
    It 'requires unique cited HTTPS sources and valid branch references' {
        $mutations = @(
            { param($p) $p.sources=@() }, { param($p) $p.sources[0].id='1invalid' },
            { param($p) $p.sources+=@($p.sources[0]) }, { param($p) $p.sources[0].url='http://example.invalid' },
            { param($p) $p.sources[0].url='https://user:password@example.invalid/' },
            { param($p) $p.sources[0].url='https://@example.invalid/' },
            { param($p) $p.sources[0].url='https://example.invalid:bad/' },
            { param($p) $p.sources[0].url='https://example.invalid:99999/' },
            { param($p) $p.sources[0].url='https://example.invalid/with space' },
            { param($p) $p.sources[0].relevance='' },
            { param($p) $p.python.branches[0].security_sources=@() },
            { param($p) $p.python.branches[0].security_sources=@('unknown') },
            { param($p) $p.python.branches[0].security_sources=@('CVE12003','CVE12003') }
        )
        foreach ($mutate in $mutations) {
            $policy = Copy-CompatibilityPolicy; & $mutate $policy
            { Get-WinBookSplitCompatibilityPolicy (Write-CompatibilityPolicy $policy) } | Should -Throw
            (Test-WinBookSplitPythonVersion '3.14.8' $policy).Reason | Should -BeExactly 'malformed_policy'
        }
    }
}

Describe 'V101 interpreter identity is independently validated' -Tag 'V101-AC-014','V101-AC-015' {
    It 'requires complete scalar fields numeric cross-checks isolation architecture and GIL' {
        $cases = @(
            @{Field='serial'; Value=$true; Reason='malformed_runtime'}, @{Field='serial'; Value=1; Reason='unstable_python'},
            @{Field='serial'; Value=1.0; Reason='malformed_runtime'}, @{Field='serial'; Value=@(0); Reason='malformed_runtime'},
            @{Field='bits'; Value=$true; Reason='malformed_runtime'}, @{Field='bits'; Value=@(64); Reason='malformed_runtime'},
            @{Field='bits'; Value=32; Reason='unsupported_architecture'}, @{Field='machine'; Value='ARM64'; Reason='unsupported_architecture'},
            @{Field='platform'; Value='linux'; Reason='unsupported_architecture'}, @{Field='implementation'; Value='pypy'; Reason='unsupported_flavor'},
            @{Field='gil_disabled'; Value=$true; Reason='unsupported_flavor'}, @{Field='gil_disabled'; Value=0; Reason='malformed_runtime'},
            @{Field='isolated'; Value=$false; Reason='malformed_runtime'}, @{Field='dont_write_bytecode'; Value=1; Reason='malformed_runtime'},
            @{Field='executable'; Value='python.exe'; Reason='malformed_runtime'}, @{Field='version_info'; Value=@(3,14,$true); Reason='malformed_runtime'},
            @{Field='version_info'; Value=@(3,14,7); Reason='malformed_runtime'}, @{Field='version_info'; Value=@(3,14,8,0); Reason='malformed_runtime'},
            @{Field='version'; Value='3.14.08'; Reason='malformed_runtime'}, @{Field='version'; Value='3.14.8rc1'; Reason='unstable_python'},
            @{Field='releaselevel'; Value='alpha'; Reason='unstable_python'}, @{Field='releaselevel'; Value='FINAL'; Reason='malformed_runtime'}
        )
        foreach ($case in $cases) {
            $record=New-CompatibilityRecord; $record[$case.Field]=$case.Value
            $result=Test-WinBookSplitInterpreterRecord $record $script:policy
            $result.Ok | Should -BeFalse; $result.Reason | Should -BeExactly $case.Reason
        }
        $record=New-CompatibilityRecord; $record.Remove('serial')
        (Test-WinBookSplitInterpreterRecord $record $script:policy).Reason | Should -BeExactly 'malformed_runtime'
        (Test-WinBookSplitInterpreterRecord @() $script:policy).Reason | Should -BeExactly 'malformed_runtime'
    }
    It 'accepts dictionary and PSCustomObject identities without pypdf fields or a claimed ok flag' {
        $record=New-CompatibilityRecord
        (Test-WinBookSplitInterpreterRecord $record $script:policy).Ok | Should -BeTrue
        $converted=$record | ConvertTo-Json -Depth 8 | ConvertFrom-Json
        (Test-WinBookSplitInterpreterRecord $converted $script:policy).Ok | Should -BeTrue
        $record.version='3.14.2'; $record.version_info=@(3,14,2); $record.ok=$true
        (Test-WinBookSplitInterpreterRecord $record $script:policy).Reason | Should -BeExactly 'python_security_floor'
    }
    It 'matches Python decisions and messages across valid malformed and rejected runtime records' {
        $records = @('3.14.2','3.14.6','3.14.7','3.14.8','3.14.9','3.14.10','3.14.99','3.13.14','3.13.15','3.13.16','3.13.17','3.13.99','3.12.14','3.15.0','3.15.1','3.15.10','3.15.99','3.16.0') | ForEach-Object { New-CompatibilityRecord $_ }
        foreach ($case in @(@('serial',$true),@('serial',1),@('bits',32),@('machine','ARM64'),@('implementation','pypy'),@('isolated',$false),@('gil_disabled',$true),@('version','3.14.8rc1'),@('version','3.14.08'),@('version_info',@(3,14,7)),@('releaselevel','alpha'))) {
            $record=New-CompatibilityRecord; $record[$case[0]]=$case[1]; $records+=@($record)
        }
        $inputPath=Join-Path $TestDrive 'parity-records.json'
        [IO.File]::WriteAllText($inputPath, (@{policy=$script:policy; records=$records} | ConvertTo-Json -Depth 32), (New-Object Text.UTF8Encoding($false, $true)))
        $code='import importlib.util,json,sys; s=importlib.util.spec_from_file_location("compat",sys.argv[1]); m=importlib.util.module_from_spec(s); s.loader.exec_module(m); d=json.load(open(sys.argv[2],encoding="utf-8")); print(json.dumps([m.validate_interpreter(r,d["policy"]) for r in d["records"]]))'
        $probe=Invoke-WinBookSplitRuntimeProbe -Path $TestPython -Arguments @('-I','-B','-c',$code,(Join-Path $RepositoryRoot 'engine\winbooksplit_compatibility.py'),$inputPath) -ApplicationRoot $RepositoryRoot
        $probe.ExitCode | Should -Be 0
        $probe.StreamsComplete | Should -BeTrue
        $pythonResults=$probe.Stdout | ConvertFrom-Json
        $pythonResults.Count | Should -Be $records.Count
        for ($index=0; $index -lt $records.Count; $index++) {
            $powershell=Test-WinBookSplitInterpreterRecord $records[$index] $script:policy
            $powershell.Ok | Should -Be $pythonResults[$index].ok
            $powershell.Reason | Should -BeExactly $pythonResults[$index].reason
            $powershell.Message | Should -BeExactly $pythonResults[$index].message
        }
    }
}

Describe 'V102 supplied version fixtures and shared evaluator parity (SYNTHETIC identities)' -Tag 'V102-AC-007','V102-AC-008','V102-AC-009','V102-AC-010','V102-AC-011' {
    It 'runs every supplied fixture through the actual PowerShell and Python evaluators' {
        $fixtures = [IO.File]::ReadAllText((Join-Path $RepositoryRoot 'docs\codex-v1.0.2\RUNTIME_POLICY_CASES.json')) | ConvertFrom-Json
        $fixtures.not_runtime_configuration | Should -BeTrue
        $records = @(foreach ($case in $fixtures.cases) {
            $record = New-CompatibilityRecord
            $record.version = $case.input
            if ($case.input -is [string] -and $case.input -cmatch '\A(?:0|[1-9][0-9]*)\.(?:0|[1-9][0-9]*)\.(?:0|[1-9][0-9]*)\z') {
                $record.version_info = @($case.input.Split('.') | ForEach-Object { [int]$_ })
            }
            $record
        })
        $inputPath = Join-Path $TestDrive 'v102-policy-fixtures.json'
        [IO.File]::WriteAllText($inputPath, (@{policy=$script:policy;records=$records} | ConvertTo-Json -Depth 32), (New-Object Text.UTF8Encoding($false, $true)))
        $code = 'import importlib.util,json,sys; s=importlib.util.spec_from_file_location("compat",sys.argv[1]); m=importlib.util.module_from_spec(s); s.loader.exec_module(m); d=json.load(open(sys.argv[2],encoding="utf-8")); print(json.dumps([m.validate_interpreter(r,d["policy"]) for r in d["records"]]))'
        $probe = Invoke-WinBookSplitRuntimeProbe -Path $TestPython -Arguments @('-I','-B','-c',$code,(Join-Path $RepositoryRoot 'engine\winbooksplit_compatibility.py'),$inputPath) -ApplicationRoot $RepositoryRoot
        $probe.ExitCode | Should -Be 0
        $probe.StreamsComplete | Should -BeTrue
        $pythonResults = $probe.Stdout | ConvertFrom-Json
        $pythonResults.Count | Should -Be $records.Count
        for ($index=0; $index -lt $records.Count; $index++) {
            $case = $fixtures.cases[$index]
            $versionResult = Test-WinBookSplitPythonVersion $case.input $script:policy
            $powershell = Test-WinBookSplitInterpreterRecord $records[$index] $script:policy
            $powershell.Ok | Should -Be $case.expected_ok
            $powershell.Reason | Should -BeExactly $case.expected_reason
            $versionResult.Ok | Should -Be $powershell.Ok
            $versionResult.Reason | Should -BeExactly $powershell.Reason
            $powershell.Ok | Should -Be $pythonResults[$index].ok
            $powershell.Reason | Should -BeExactly $pythonResults[$index].reason
            $powershell.Message | Should -BeExactly $pythonResults[$index].message
            if ($powershell.Reason -ceq 'unsupported_python') {
                $powershell.Message | Should -Not -Match 'insecure|security floor'
            }
        }
    }
}
