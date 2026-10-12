param([string]$RepositoryRoot, [string]$WorkRoot, [string]$TestPython)

BeforeAll {
    . (Join-Path $RepositoryRoot 'engine\WinBookSplit.Diagnostics.ps1')
    . (Join-Path $RepositoryRoot 'engine\WinBookSplit.Runtime.ps1')
    if (-not [IO.File]::Exists($TestPython)) { throw 'Discovery tests require the concrete accepted TestPython.' }
    $script:discoveryPolicy = Get-WinBookSplitCompatibilityPolicy
    $script:discoveryExpectedVersion = (Resolve-WinBookSplitRuntime -ApplicationRoot $RepositoryRoot -PythonPath $TestPython -InterpreterOnly).Version
    $script:nativeDiscoveryProbe = (Get-Command Invoke-WinBookSplitRuntimeProbe).ScriptBlock
    $script:discoveryExecutableFixtures = New-Object 'Collections.Generic.List[object]'
    function New-DiscoveryProbe {
        param([AllowEmptyString()][string]$Stdout = '', [AllowEmptyString()][string]$Stderr = '', [int]$ExitCode = 0)
        return [pscustomobject]@{ Stdout=$Stdout; Stderr=$Stderr; ExitCode=$ExitCode; TimedOut=$false;
            ParentStopped=$true; DescendantsStopped=$true; StreamsComplete=$true; JobAssigned=$true; Cancelled=$false;
            StartError=$null; StopError=$null; StreamError=$null; StdoutTruncated=$false; StderrTruncated=$false }
    }
    function New-DiscoveryExecutable {
        param([string]$Name = 'py.exe')
        # The shared harness runs from WorkRoot and its TestDrive is beneath CWD.
        # Listing fixtures must be outside that correctly excluded dependency origin.
        $parent = [IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($WorkRoot))
        $directory = Join-Path $parent ('discovery-' + [guid]::NewGuid().ToString('N'))
        $null = New-Item -ItemType Directory -Path $directory -ErrorAction Stop
        $owned = [pscustomobject]@{Directory=$directory;File=$null}
        $script:discoveryExecutableFixtures.Add($owned)
        $path = Join-Path $directory $Name
        $stream = [IO.File]::Open($path,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::None)
        try {
            $owned.File=$path
            $bytes=[Text.Encoding]::UTF8.GetBytes('SYNTHETIC listing executable fixture: never executed.')
            $stream.Write($bytes,0,$bytes.Length)
        } finally { $stream.Dispose() }
        Test-WinBookSplitAutomaticPath -Path $path -ApplicationRoot $RepositoryRoot | Should -BeTrue
        return $path
    }
    function New-DiscoveryApplication {
        $app = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $engine = Join-Path $app 'engine'
        $null = New-Item -ItemType Directory -Path $engine
        foreach ($name in @('WinBookSplit.Compatibility.json','winbooksplit_compatibility.py')) {
            Copy-Item -LiteralPath (Join-Path $RepositoryRoot ('engine\' + $name)) -Destination $engine
        }
        return $app
    }
    function New-DiscoveryRankCandidate {
        param([string]$Path, [string]$Version)
        return [pscustomobject]@{ Path=$Path; Version=$Version; Source='SYNTHETIC'; Attempts=@() }
    }
    function New-DiscoveryManagerRegistration {
        $name = 'PythonSoftwareFoundation.PythonManager'
        $version = '26.4.240.0'
        $full = $name + '_' + $version + '_x64__3847v3x7pw1km'
        return [ordered]@{ schema_version=1; packages=@([ordered]@{Name=$name;
            Publisher='CN=Python Software Foundation, O=Python Software Foundation, L=Beaverton, S=Oregon, C=US';
            PublisherId='3847v3x7pw1km';PackageFamilyName=($name+'_3847v3x7pw1km');PackageFullName=$full;
            InstallLocation=([IO.Path]::Combine([Environment]::GetFolderPath([Environment+SpecialFolder]::ProgramFiles),'WindowsApps',$full));
            Architecture='X64';SignatureKind='Developer';Status='Ok';Version=$version;
            IsDevelopmentMode=$false;IsPartiallyStaged=$false;IsFramework=$false;IsResourcePackage=$false}) }
    }
}

AfterAll {
    foreach ($owned in $script:discoveryExecutableFixtures) {
        # Exact authored files only; nonrecursive deletion preserves any foreign entry.
        if ($owned.File -and [IO.File]::Exists($owned.File)) { [IO.File]::Delete($owned.File) }
        if ([IO.Directory]::Exists($owned.Directory)) { [IO.Directory]::Delete($owned.Directory,$false) }
    }
}

Describe 'V101 bounded installed-runtime listing parser (SYNTHETIC records)' -Tag 'V101-AC-017','V101-AC-018','V101-AC-020','V101-AC-023' {
    It 'parses legacy numeric and tagged lines with literal spaces and punctuation' {
        $text = " -V:PythonCore/3.14 * C:\Trusted Python\python.exe`r`n -3.13-64 C:\Other [&!]\python.exe`r`n"
        $paths = @(ConvertFrom-WinBookSplitPythonListing -Probe (New-DiscoveryProbe -Stdout $text) -Kind Legacy)
        $paths.Count | Should -Be 2
        $paths[0] | Should -BeExactly 'C:\Trusted Python\python.exe'
        $paths[1] | Should -BeExactly 'C:\Other [&!]\python.exe'
        # Legacy launchers can emit their installed-runtime inventory on stderr.
        @(ConvertFrom-WinBookSplitPythonListing -Probe (New-DiscoveryProbe -Stderr $text) -Kind Legacy).Count | Should -Be 2
    }
    It 'parses path-only manager format as data without evaluating it' {
        $text = "C:\Trusted Python\python.exe`n\\server\share\Python [&!]\python.exe`n"
        $paths = @(ConvertFrom-WinBookSplitPythonListing -Probe (New-DiscoveryProbe -Stdout $text) -Kind Manager)
        $paths.Count | Should -Be 2
        $paths[1] | Should -BeExactly '\\server\share\Python [&!]\python.exe'
    }
    It 'rejects empty and mixed malformed relative non-executable quoted or NUL records' {
        foreach ($text in @('', "`r`n `t`r`n", 'python.exe', 'C:\Python\python.bat', '"C:\Python\python.exe"',
                "C:\Python\python.exe`nWelcome to Python", ('C:\Python\py' + [char]0 + 'thon.exe'))) {
            { ConvertFrom-WinBookSplitPythonListing -Probe (New-DiscoveryProbe -Stdout $text) -Kind Manager } | Should -Throw
        }
        foreach ($text in @('C:\Python\python.exe', '-V: C:\Python\python.exe', '-3.14 python.exe',
                "-3.14 C:\Python\python.exe`nUnexpected record", ('-V:' + ('x' * 81) + ' C:\Python\python.exe'))) {
            { ConvertFrom-WinBookSplitPythonListing -Probe (New-DiscoveryProbe -Stdout $text) -Kind Legacy } | Should -Throw
        }
        { ConvertFrom-WinBookSplitPythonListing -Probe (New-DiscoveryProbe -Stdout 'C:\Python\python.exe' -Stderr 'unexpected') -Kind Manager } | Should -Throw
    }
    It 'fails at explicit byte line character and candidate limits without work proportional to unbounded data' {
        $cases = @(
            (' ' * 32769),
            ((' ' + "`n") * 129),
            ('C:\' + ('x' * 4097) + '.exe'),
            ((1..33 | ForEach-Object { 'C:\Python' + $_ + '\python.exe' }) -join "`n")
        )
        foreach ($text in $cases) {
            { ConvertFrom-WinBookSplitPythonListing -Probe (New-DiscoveryProbe -Stdout $text) -Kind Manager } | Should -Throw
        }
        $boundary = (1..32 | ForEach-Object { 'C:\Python' + $_ + '\python.exe' }) -join "`n"
        @(ConvertFrom-WinBookSplitPythonListing -Probe (New-DiscoveryProbe -Stdout $boundary) -Kind Manager).Count | Should -Be 32
    }
    It 'rejects incomplete truncated failed timed-out and unsafe process records' {
        foreach ($field in @('JobAssigned','StreamsComplete','ParentStopped','DescendantsStopped')) {
            $probe = New-DiscoveryProbe -Stdout 'C:\Python\python.exe'
            $probe.$field = $false
            { ConvertFrom-WinBookSplitPythonListing -Probe $probe -Kind Manager } | Should -Throw
        }
        foreach ($field in @('TimedOut','Cancelled','StdoutTruncated','StderrTruncated')) {
            $probe = New-DiscoveryProbe -Stdout 'C:\Python\python.exe'
            $probe.$field = $true
            { ConvertFrom-WinBookSplitPythonListing -Probe $probe -Kind Manager } | Should -Throw
        }
        foreach ($field in @('StartError','StopError','StreamError')) {
            $probe = New-DiscoveryProbe -Stdout 'C:\Python\python.exe'
            $probe.$field = 'SYNTHETIC failure'
            { ConvertFrom-WinBookSplitPythonListing -Probe $probe -Kind Manager } | Should -Throw
        }
        { ConvertFrom-WinBookSplitPythonListing -Probe (New-DiscoveryProbe -ExitCode 2) -Kind Legacy } | Should -Throw
    }
}

Describe 'V101 physical manager registration identity (SYNTHETIC registration records)' -Tag 'V101-AC-018','V101-AC-022','V101-AC-023' {
    It 'requires the complete PSF identity and routes trust checks through the physical package paths' {
        $record = New-DiscoveryManagerRegistration
        $script:registeredLocation = $record.packages[0].InstallLocation
        Mock Get-WinBookSplitTrustedPath { return $Path }
        Mock Test-WinBookSplitAutomaticPath { return $true }
        $text = $record | ConvertTo-Json -Depth 6 -Compress
        $paths = @(ConvertFrom-WinBookSplitManagerRegistration -Probe (New-DiscoveryProbe -Stdout $text) -ApplicationRoot $RepositoryRoot)
        $paths.Count | Should -Be 1
        $paths[0] | Should -BeExactly ([IO.Path]::Combine($script:registeredLocation,'pymanager.exe'))
        Should-Invoke Get-WinBookSplitTrustedPath -Times 1 -Exactly -ParameterFilter { $Path -eq $script:registeredLocation -and $Directory }
        Should-Invoke Get-WinBookSplitTrustedPath -Times 1 -Exactly -ParameterFilter { $Path -eq [IO.Path]::Combine($script:registeredLocation,'pymanager.exe') -and $Automatic }
    }
    It 'rejects mismatched publisher package location architecture status flags duplicate or unknown fields' {
        Mock Get-WinBookSplitTrustedPath { return $Path }
        Mock Test-WinBookSplitAutomaticPath { return $true }
        foreach ($field in @('Name','Publisher','PublisherId','PackageFamilyName','PackageFullName','InstallLocation','Architecture','SignatureKind','Status','Version',
                'IsDevelopmentMode','IsPartiallyStaged','IsFramework','IsResourcePackage')) {
            $record = New-DiscoveryManagerRegistration
            if ($record.packages[0][$field] -is [bool]) { $record.packages[0][$field]=$true }
            else { $record.packages[0][$field]='SYNTHETIC untrusted identity' }
            $text = $record | ConvertTo-Json -Depth 6 -Compress
            { ConvertFrom-WinBookSplitManagerRegistration -Probe (New-DiscoveryProbe -Stdout $text) -ApplicationRoot $RepositoryRoot } | Should -Throw
        }
        $duplicate = New-DiscoveryManagerRegistration
        $duplicate.packages = @($duplicate.packages[0],$duplicate.packages[0])
        $text = $duplicate | ConvertTo-Json -Depth 6 -Compress
        { ConvertFrom-WinBookSplitManagerRegistration -Probe (New-DiscoveryProbe -Stdout $text) -ApplicationRoot $RepositoryRoot } | Should -Throw
        $unknown = New-DiscoveryManagerRegistration
        $unknown.packages[0]['extra']='unexpected'
        $text = $unknown | ConvertTo-Json -Depth 6 -Compress
        { ConvertFrom-WinBookSplitManagerRegistration -Probe (New-DiscoveryProbe -Stdout $text) -ApplicationRoot $RepositoryRoot } | Should -Throw
        foreach ($text in @('{"schema_version":true,"packages":[]}', '{"schema_version":1,"schema_version":1,"packages":[]}',
                '{"schema_version":1,"packages":{}}', '{"schema_version":1,"packages":[],"extra":false}', (' ' * 32769))) {
            { ConvertFrom-WinBookSplitManagerRegistration -Probe (New-DiscoveryProbe -Stdout $text) -ApplicationRoot $RepositoryRoot } | Should -Throw
        }
    }
    It 'rejects a valid claimed identity when its physical origin or automatic directory check fails' {
        $text = New-DiscoveryManagerRegistration | ConvertTo-Json -Depth 6 -Compress
        Mock Get-WinBookSplitTrustedPath { throw 'SYNTHETIC reparse or unsafe physical origin.' }
        { ConvertFrom-WinBookSplitManagerRegistration -Probe (New-DiscoveryProbe -Stdout $text) -ApplicationRoot $RepositoryRoot } | Should -Throw
        Mock Get-WinBookSplitTrustedPath { return $Path }
        Mock Test-WinBookSplitAutomaticPath { return $false }
        { ConvertFrom-WinBookSplitManagerRegistration -Probe (New-DiscoveryProbe -Stdout $text) -ApplicationRoot $RepositoryRoot } | Should -Throw
    }
}

Describe 'V101 independent discovery and optional-listing failures' -Tag 'V101-AC-019','V101-AC-020','V101-AC-023','V101-AC-024' {
    BeforeEach {
        Mock Get-WinBookSplitRegisteredManagerPaths { return @() }
    }
    It 'selects an actual accepted PATH interpreter with no launcher available' {
        Mock Get-WinBookSplitPathCandidates { if ($Name -eq 'python.exe') { return @($TestPython) }; return @() }
        $result = Resolve-WinBookSplitRuntime -ApplicationRoot $RepositoryRoot
        $result.Path | Should -BeExactly $TestPython
        $result.Source | Should -BeExactly 'PATH'
        $result.Probe.StreamsComplete | Should -BeTrue
        @($result.Attempts | Where-Object { $_.Source -eq 'selection' -and $_.Accepted }).Count | Should -Be 1
    }
    It 'retains the original unknown -0 exit-2 diagnostics and still selects an actual PATH interpreter' {
        $script:discoveryLauncher = New-DiscoveryExecutable
        Mock Get-WinBookSplitPathCandidates {
            if ($Name -eq 'python.exe') { return @($TestPython) }
            if ($Name -eq 'py.exe') { return @($script:discoveryLauncher) }
            return @()
        }
        Mock Invoke-WinBookSplitRuntimeProbe {
            if ($Path -eq $script:discoveryLauncher) {
                return New-DiscoveryProbe -Stderr "Unknown option: -0`r`nusage: C:\Python314\python.exe [option] ... [-c cmd | -m mod | file | -] [arg] ...`r`nTry 'python -h' for more information.`r`n" -ExitCode 2
            }
            return & $script:nativeDiscoveryProbe -Path $Path -Arguments $Arguments -ApplicationRoot $ApplicationRoot -TimeoutSeconds $TimeoutSeconds
        }
        $result = Resolve-WinBookSplitRuntime -ApplicationRoot $RepositoryRoot
        $result.Path | Should -BeExactly $TestPython
        $result.Version | Should -BeExactly $script:discoveryExpectedVersion
        $attempt = @($result.Attempts | Where-Object { $_.Source -eq 'py_launcher_listing' })
        $attempt.Count | Should -Be 1
        $attempt[0].Accepted | Should -BeFalse
        $attempt[0].Probe.ExitCode | Should -Be 2
        $attempt[0].Probe.Stderr | Should -Match 'Unknown option: -0'
        Should-Invoke Invoke-WinBookSplitRuntimeProbe -Times 1 -Exactly -ParameterFilter { $Path -eq $script:discoveryLauncher -and ($Arguments -join ',') -eq '-0p' }
    }
    It 'treats empty or malformed optional listings as bounded failures alongside a real PATH success' {
        $script:discoveryLauncher = New-DiscoveryExecutable
        Mock Get-WinBookSplitPathCandidates {
            if ($Name -eq 'python.exe') { return @($TestPython) }
            if ($Name -eq 'py.exe') { return @($script:discoveryLauncher) }
            return @()
        }
        foreach ($text in @('', '-3.14 python.exe', ('-3.14 C:\Python\python.exe' + "`nextra"))) {
            $script:discoveryListing = $text
            Mock Invoke-WinBookSplitRuntimeProbe {
                if ($Path -eq $script:discoveryLauncher) { return New-DiscoveryProbe -Stdout $script:discoveryListing }
                return & $script:nativeDiscoveryProbe -Path $Path -Arguments $Arguments -ApplicationRoot $ApplicationRoot -TimeoutSeconds $TimeoutSeconds
            }
            $result = Resolve-WinBookSplitRuntime -ApplicationRoot $RepositoryRoot
            $result.Path | Should -BeExactly $TestPython
            @($result.Attempts | Where-Object { $_.Source -eq 'py_launcher_listing' -and -not $_.Accepted }).Count | Should -Be 1
        }
    }
    It 'keeps safely stopped optional timeouts informative but makes cancellation and unproved shutdown fatal' {
        $script:discoveryLauncher = New-DiscoveryExecutable
        Mock Get-WinBookSplitPathCandidates {
            if ($Name -eq 'python.exe') { return @($TestPython) }
            if ($Name -eq 'py.exe') { return @($script:discoveryLauncher) }
            return @()
        }
        $script:discoveryFailedProbe = New-DiscoveryProbe
        $script:discoveryFailedProbe.TimedOut = $true
        Mock Invoke-WinBookSplitRuntimeProbe {
            if ($Path -eq $script:discoveryLauncher) { return $script:discoveryFailedProbe }
            return & $script:nativeDiscoveryProbe -Path $Path -Arguments $Arguments -ApplicationRoot $ApplicationRoot -TimeoutSeconds $TimeoutSeconds
        }
        (Resolve-WinBookSplitRuntime -ApplicationRoot $RepositoryRoot).Path | Should -BeExactly $TestPython
        $script:discoveryFailedProbe.TimedOut = $false
        $script:discoveryFailedProbe.Cancelled = $true
        { Resolve-WinBookSplitRuntime -ApplicationRoot $RepositoryRoot } | Should -Throw
        $script:discoveryFailedProbe.Cancelled = $false
        $script:discoveryFailedProbe.DescendantsStopped = $false
        { Resolve-WinBookSplitRuntime -ApplicationRoot $RepositoryRoot } | Should -Throw
    }
    It 'returns runtime_not_found with no viable direct or listed interpreter' {
        Mock Get-WinBookSplitPathCandidates { return @() }
        try { $null = Resolve-WinBookSplitRuntime -ApplicationRoot $RepositoryRoot; throw 'Expected missing-runtime rejection.' }
        catch { $_.Exception.Data['Code'] | Should -BeExactly 'runtime_not_found' }
    }
    It 'preserves actual PATH success after a safely stopped listed-interpreter timeout but retains fatal interruption rules' {
        $script:discoveryLauncher = New-DiscoveryExecutable
        $script:listedInterpreter = New-DiscoveryExecutable -Name 'python.exe'
        Mock Get-WinBookSplitPathCandidates {
            if ($Name -eq 'python.exe') { return @($TestPython) }
            if ($Name -eq 'py.exe') { return @($script:discoveryLauncher) }
            return @()
        }
        $script:listedFailure = New-DiscoveryProbe -ExitCode 1
        $script:listedFailure.TimedOut = $true
        Mock Invoke-WinBookSplitRuntimeProbe {
            if ($Path -eq $script:discoveryLauncher) { return New-DiscoveryProbe -Stdout (' -V:3.14 ' + $script:listedInterpreter) }
            if ($Path -eq $script:listedInterpreter) { return $script:listedFailure }
            return & $script:nativeDiscoveryProbe -Path $Path -Arguments $Arguments -ApplicationRoot $ApplicationRoot -TimeoutSeconds $TimeoutSeconds
        }
        $result = Resolve-WinBookSplitRuntime -ApplicationRoot $RepositoryRoot
        $result.Path | Should -BeExactly $TestPython
        $failed = @($result.Attempts | Where-Object { $_.Path -eq $script:listedInterpreter -and -not $_.Accepted })
        $failed.Count | Should -Be 1
        $failed[0].Probe.TimedOut | Should -BeTrue
        $script:listedFailure.Cancelled = $true
        { Resolve-WinBookSplitRuntime -ApplicationRoot $RepositoryRoot } | Should -Throw
        $script:listedFailure.Cancelled = $false
        $script:listedFailure.DescendantsStopped = $false
        { Resolve-WinBookSplitRuntime -ApplicationRoot $RepositoryRoot } | Should -Throw
        $script:listedFailure.DescendantsStopped = $true
        Mock Get-WinBookSplitPathCandidates { if ($Name -eq 'python.exe') { return @($script:listedInterpreter,$TestPython) }; return @() }
        { Resolve-WinBookSplitRuntime -ApplicationRoot $RepositoryRoot } | Should -Throw
    }
    It 'preserves an already accepted PATH interpreter when the optional discovery budget is exhausted' {
        Mock Get-WinBookSplitPathCandidates { if ($Name -eq 'python.exe') { return @($TestPython) }; return @() }
        # SYNTHETIC elapsed budget; the selected interpreter probe is still actual.
        Mock Get-WinBookSplitDiscoveryTimeout { if ($Optional) { return 0 }; return $Requested }
        $result = Resolve-WinBookSplitRuntime -ApplicationRoot $RepositoryRoot
        $result.Path | Should -BeExactly $TestPython
        $result.Source | Should -BeExactly 'PATH'
        Should-Invoke Get-WinBookSplitRegisteredManagerPaths -Times 0 -Exactly
        Should-Invoke Get-WinBookSplitDiscoveryTimeout -Times 1 -Exactly -ParameterFilter { $Optional }
    }
    It 'does not execute an unregistered PATH file merely named pymanager.exe' {
        $script:unverifiedManager = New-DiscoveryExecutable -Name 'pymanager.exe'
        Mock Get-WinBookSplitPathCandidates {
            if ($Name -eq 'python.exe') { return @($TestPython) }
            if ($Name -eq 'pymanager.exe') { return @($script:unverifiedManager) }
            return @()
        }
        Mock Invoke-WinBookSplitRuntimeProbe {
            return & $script:nativeDiscoveryProbe -Path $Path -Arguments $Arguments -ApplicationRoot $ApplicationRoot -TimeoutSeconds $TimeoutSeconds
        }
        $result = Resolve-WinBookSplitRuntime -ApplicationRoot $RepositoryRoot
        $result.Path | Should -BeExactly $TestPython
        @($result.Attempts | Where-Object { $_.Source -eq 'manager_listing' -and -not $_.Accepted }).Count | Should -Be 1
        Should-Invoke Invoke-WinBookSplitRuntimeProbe -Times 0 -Exactly -ParameterFilter { $Path -eq $script:unverifiedManager }
    }
    It 'deduplicates the same concrete interpreter listed by two sources and direct PATH' {
        $script:discoveryLauncher = New-DiscoveryExecutable
        $script:discoveryManager = New-DiscoveryExecutable -Name 'pymanager.exe'
        # Registration is a controlled seam here, not a native package-registration claim.
        Mock Get-WinBookSplitRegisteredManagerPaths { return @($script:discoveryManager) }
        Mock Get-WinBookSplitPathCandidates {
            if ($Name -eq 'python.exe') { return @($TestPython,$TestPython) }
            if ($Name -eq 'py.exe') { return @($script:discoveryLauncher) }
            if ($Name -eq 'pymanager.exe') { return @($script:discoveryManager) }
            return @()
        }
        Mock Invoke-WinBookSplitRuntimeProbe {
            if ($Path -eq $script:discoveryLauncher -or $Path -eq $script:discoveryManager) {
                return New-DiscoveryProbe -Stdout (' -V:3.14 * ' + $TestPython)
            }
            return & $script:nativeDiscoveryProbe -Path $Path -Arguments $Arguments -ApplicationRoot $ApplicationRoot -TimeoutSeconds $TimeoutSeconds
        }
        $result = Resolve-WinBookSplitRuntime -ApplicationRoot $RepositoryRoot
        $result.Path | Should -BeExactly $TestPython
        $result.Source | Should -BeExactly 'PATH'
        Should-Invoke Invoke-WinBookSplitRuntimeProbe -Times 1 -Exactly -ParameterFilter { $Path -eq $TestPython }
        @($result.Attempts | Where-Object { $_.Source -in @('py_launcher_listing','manager_listing') -and $_.Accepted }).Count | Should -Be 2
        @($result.Attempts | Where-Object { $_.Source -eq 'selection' -and $_.Accepted }).Count | Should -Be 1
    }
    It 'queries a controlled registered physical manager with only -0p then validates its actual interpreter' {
        $script:discoveryManager = New-DiscoveryExecutable -Name 'pymanager.exe'
        Mock Get-WinBookSplitRegisteredManagerPaths { return @($script:discoveryManager) }
        Mock Get-WinBookSplitPathCandidates { return @() }
        Mock Invoke-WinBookSplitRuntimeProbe {
            if ($Path -eq $script:discoveryManager) {
                return New-DiscoveryProbe -Stdout (' -V:3.14-64 * ' + $TestPython)
            }
            return & $script:nativeDiscoveryProbe -Path $Path -Arguments $Arguments -ApplicationRoot $ApplicationRoot -TimeoutSeconds $TimeoutSeconds
        }
        $result = Resolve-WinBookSplitRuntime -ApplicationRoot $RepositoryRoot
        $result.Path | Should -BeExactly $TestPython
        $result.Source | Should -BeExactly 'manager'
        Should-Invoke Invoke-WinBookSplitRuntimeProbe -Times 1 -Exactly -ParameterFilter { $Path -eq $script:discoveryManager -and ($Arguments -join ',') -eq '-0p' -and $TimeoutSeconds -le 2 }
        Should-Invoke Invoke-WinBookSplitRuntimeProbe -Times 0 -Exactly -ParameterFilter { $Path -eq $script:discoveryManager -and ($Arguments -join ',') -ne '-0p' }
        $result.Details.isolated | Should -BeTrue
        $result.PypdfVersion | Should -BeExactly '6.20.0'
    }
}

Describe 'V101 authoritative existing environment shapes and native child settings' -Tag 'V101-AC-021','V101-AC-022','V101-AC-023' {
    It 'does not bypass an application .venv that is a file' {
        $app = New-DiscoveryApplication
        [IO.File]::WriteAllText((Join-Path $app '.venv'), 'Owned damaged environment fixture.')
        Mock Resolve-WinBookSplitAutomaticPython { throw 'Automatic fallback must not occur.' }
        try { $null = Resolve-WinBookSplitRuntime -ApplicationRoot $app; throw 'Expected environment rejection.' }
        catch {
            $_.Exception.Data['Code'] | Should -BeExactly 'runtime_invalid'
            $_.Exception.Data['Attempts'][0].Source | Should -BeExactly 'application_venv'
        }
        Should-Invoke Resolve-WinBookSplitAutomaticPython -Times 0 -Exactly
    }
    It 'does not bypass a reparse .venv and preserves its unrelated owned target' {
        $app = New-DiscoveryApplication
        $target = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $null = New-Item -ItemType Directory -Path $target
        $sentinel = Join-Path $target 'preserve.txt'
        [IO.File]::WriteAllText($sentinel, 'Owned junction target sentinel.')
        $venv = Join-Path $app '.venv'
        $null = New-Item -ItemType Junction -Path $venv -Target $target
        Mock Resolve-WinBookSplitAutomaticPython { throw 'Automatic fallback must not occur.' }
        try {
            try { $null = Resolve-WinBookSplitRuntime -ApplicationRoot $app; throw 'Expected environment rejection.' }
            catch {
                $_.Exception.Data['Code'] | Should -BeExactly 'runtime_invalid'
                $_.Exception.Data['Attempts'][0].Source | Should -BeExactly 'application_venv'
            }
            [IO.File]::ReadAllText($sentinel) | Should -BeExactly 'Owned junction target sentinel.'
            Should-Invoke Resolve-WinBookSplitAutomaticPython -Times 0 -Exactly
        } finally { [IO.Directory]::Delete($venv) }
        [IO.File]::ReadAllText($sentinel) | Should -BeExactly 'Owned junction target sentinel.'
    }
    It 'sets install-disabling child values without changing the parent environment' {
        $savedManager = $env:PYTHON_MANAGER_AUTOMATIC_INSTALL
        $savedLauncher = $env:PYLAUNCHER_NO_SEARCH_PATH
        try {
            $env:PYTHON_MANAGER_AUTOMATIC_INSTALL = 'true'
            $env:PYLAUNCHER_NO_SEARCH_PATH = 'unsafe-parent-value'
            $code = 'import json,os; print(json.dumps({k:os.environ.get(k) for k in ("PYTHON_MANAGER_AUTOMATIC_INSTALL","PYLAUNCHER_NO_SEARCH_PATH")}))'
            $probe = Invoke-WinBookSplitRuntimeProbe -Path $TestPython -Arguments @('-I','-B','-c',$code) -ApplicationRoot $RepositoryRoot
            $probe.ExitCode | Should -Be 0
            $probe.JobAssigned | Should -BeTrue
            $probe.StreamsComplete | Should -BeTrue
            $values = $probe.Stdout | ConvertFrom-Json
            $values.PYTHON_MANAGER_AUTOMATIC_INSTALL | Should -BeExactly 'false'
            $values.PYLAUNCHER_NO_SEARCH_PATH | Should -BeExactly '1'
            $env:PYTHON_MANAGER_AUTOMATIC_INSTALL | Should -BeExactly 'true'
            $env:PYLAUNCHER_NO_SEARCH_PATH | Should -BeExactly 'unsafe-parent-value'
        } finally {
            $env:PYTHON_MANAGER_AUTOMATIC_INSTALL = $savedManager
            $env:PYLAUNCHER_NO_SEARCH_PATH = $savedLauncher
        }
    }
}

Describe 'V102 deterministic policy selection (SYNTHETIC later patches)' -Tag 'V101-AC-024','V102-AC-012' {
    It 'prefers the policy branch before any higher patch on another branch in either input order' {
        $preferred = New-DiscoveryRankCandidate 'C:\A\python.exe' '3.15.0'
        $other = New-DiscoveryRankCandidate 'C:\B\python.exe' '3.14.99'
        $older = New-DiscoveryRankCandidate 'C:\C\python.exe' '3.13.99'
        foreach ($candidates in @(@($older,$other,$preferred),@($preferred,$other,$older))) {
            (Select-WinBookSplitPythonCandidate -Candidates $candidates -Policy $script:discoveryPolicy).Path | Should -BeExactly $preferred.Path
        }
    }
    It 'compares maintenance patches 9 and 10 numerically in either order' {
        $nine = New-DiscoveryRankCandidate 'C:\A\python.exe' '3.15.9'
        $ten = New-DiscoveryRankCandidate 'C:\Z\python.exe' '3.15.10'
        foreach ($candidates in @(@($nine,$ten),@($ten,$nine))) {
            (Select-WinBookSplitPythonCandidate -Candidates $candidates -Policy $script:discoveryPolicy).Version | Should -BeExactly '3.15.10'
        }
    }
    It 'uses ordinal canonical path ties even when input order is reversed' {
        $upper = New-DiscoveryRankCandidate 'C:\A\python.exe' '3.14.8'
        $lower = New-DiscoveryRankCandidate 'C:\a\python.exe' '3.14.8'
        foreach ($candidates in @(@($lower,$upper),@($upper,$lower))) {
            (Select-WinBookSplitPythonCandidate -Candidates $candidates -Policy $script:discoveryPolicy).Path | Should -BeExactly $upper.Path
        }
    }
    It 'never ranks below-floor unsupported or prerelease candidates as eligible' {
        $valid = New-DiscoveryRankCandidate 'C:\Z\python.exe' '3.13.16'
        $candidates = @((New-DiscoveryRankCandidate 'C:\A\python.exe' '3.14.2'),
            (New-DiscoveryRankCandidate 'C:\D\python.exe' '3.14.7'),
            (New-DiscoveryRankCandidate 'C:\E\python.exe' '3.13.15'),
            (New-DiscoveryRankCandidate 'C:\B\python.exe' '3.16.0'),
            (New-DiscoveryRankCandidate 'C:\C\python.exe' '3.14.9rc1'),$valid)
        (Select-WinBookSplitPythonCandidate -Candidates $candidates -Policy $script:discoveryPolicy).Path | Should -BeExactly $valid.Path
    }
}
