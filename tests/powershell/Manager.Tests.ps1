param([string]$RepositoryRoot, [string]$WorkRoot)

# Maintained T02 contracts. Registration, transport and alias results here are
# SYNTHETIC TEST INPUTS. They do not establish a native Store launch or release gate.
BeforeAll {
    . (Join-Path $RepositoryRoot 'engine\WinBookSplit.Runtime.ps1')
    $script:managerCases = Get-Content -LiteralPath (Join-Path $RepositoryRoot 'docs\codex-v1.0.3\fixtures\registration-cases.json') -Raw | ConvertFrom-Json

    function New-ManagerProbe {
        param([AllowEmptyString()][string]$Stdout = '', [AllowEmptyString()][string]$Stderr = '', [int]$ExitCode = 0)
        return [pscustomobject]@{ Stdout=$Stdout; Stderr=$Stderr; ExitCode=$ExitCode; TimedOut=$false;
            ParentStopped=$true; DescendantsStopped=$true; StreamsComplete=$true; JobAssigned=$true; Cancelled=$false;
            StartError=$null; StopError=$null; StreamError=$null; StdoutTruncated=$false; StderrTruncated=$false }
    }
    function New-ManagerPackage {
        param([string]$Id = 'REG01')
        $fixture = @($script:managerCases.cases | Where-Object { $_.id -ceq $Id })
        if ($fixture.Count -ne 1 -or $fixture[0].kind -cne 'SYNTHETIC_TEST_INPUT') { throw 'Expected one explicitly synthetic registration fixture.' }
        $package = $fixture[0].registration | ConvertTo-Json -Depth 6 | ConvertFrom-Json
        # Preserve REG11's unbound location. Other fixtures use this host's
        # ProgramFiles rather than assuming an installation drive in a test.
        if ($Id -cne 'REG11') {
            $package.InstallLocation = [IO.Path]::Combine([Environment]::GetFolderPath([Environment+SpecialFolder]::ProgramFiles), 'WindowsApps', $package.PackageFullName)
        }
        return $package
    }
    function New-ManagerJson {
        param([object[]]$Packages = @())
        return ConvertTo-Json -InputObject ([ordered]@{schema_version=1;packages=@($Packages)}) -Depth 7 -Compress
    }
    function Get-ManagerFixtureAlias {
        param($Package)
        return [IO.Path]::Combine([Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData), 'Microsoft', 'WindowsApps', $Package.PackageFamilyName, 'pymanager.exe')
    }
    function Set-ManagerFixtureVersion {
        param($Package, [string]$Version)
        $Package.Version = $Version
        $Package.PackageFullName = $Package.Name + '_' + $Version + '_x64__' + $Package.PublisherId
        $Package.InstallLocation = [IO.Path]::Combine([Environment]::GetFolderPath([Environment+SpecialFolder]::ProgramFiles), 'WindowsApps', $Package.PackageFullName)
    }
    function Get-AppxPackage {
        [CmdletBinding()]
        param([string]$Name, [string]$Publisher, [switch]$AllUsers, [string]$User)
        throw 'The SYNTHETIC Get-AppxPackage positive-control mock was not installed.'
    }
    function Invoke-ManagerQueryFixture {
        param([string[]]$Arguments)
        if ($Arguments.Count -ne 4 -or $Arguments[0] -cne '-NoProfile' -or $Arguments[1] -cne '-NonInteractive' -or
            $Arguments[2] -cne '-EncodedCommand') { throw 'Expected the real bounded registration query argument vector.' }
        $command = [Text.Encoding]::Unicode.GetString([Convert]::FromBase64String($Arguments[3]))
        $queryTokens = $null; $queryErrors = $null
        $queryAst = [Management.Automation.Language.Parser]::ParseInput($command, [ref]$queryTokens, [ref]$queryErrors)
        if (@($queryErrors).Count -ne 0) { throw 'The real registration query contains syntax errors.' }
        $calls = @($queryAst.FindAll({ param($node) $node -is [Management.Automation.Language.CommandAst] -and
            $node.GetCommandName() -ceq 'Get-AppxPackage' }, $true))
        if ($calls.Count -ne 1) { throw 'Expected exactly one current-user package query.' }
        # Execute the emitted query against an in-process labelled inventory.
        # No real Appx query, manager, interpreter or acquisition process runs.
        $packages = @(& ([scriptblock]::Create($calls[0].Extent.Text)))
        return New-ManagerProbe -Stdout (New-ManagerJson -Packages $packages)
    }
}

Describe 'T02 official manager identities (SYNTHETIC fixtures)' -Tag 'T02-Manager','STORE-007','STORE-008','STORE-009','STORE-010' {
    BeforeEach {
        Mock Get-WinBookSplitTrustedPath { return $Path }
        Mock Get-WinBookSplitManagerEntryPoint { return Get-ManagerFixtureAlias -Package $Package }
    }

    It 'accepts every valid fixture and rejects every invalid fixture without treating identity as alias proof' {
        $accepted = 0
        foreach ($fixture in $script:managerCases.cases) {
            $package = New-ManagerPackage -Id $fixture.id
            $json = New-ManagerJson -Packages @($package)
            if ($fixture.expected_identity_accept) {
                $paths = @(ConvertFrom-WinBookSplitManagerRegistration -Probe (New-ManagerProbe -Stdout $json) -ApplicationRoot $RepositoryRoot)
                $paths.Count | Should -Be 1 -Because $fixture.id
                $paths[0] | Should -BeExactly (Get-ManagerFixtureAlias -Package $package) -Because $fixture.id
                $accepted++
            }
            else {
                { ConvertFrom-WinBookSplitManagerRegistration -Probe (New-ManagerProbe -Stdout $json) -ApplicationRoot $RepositoryRoot } |
                    Should -Throw -Because $fixture.id
            }
        }
        $accepted | Should -Be 3
        Should-Invoke Get-WinBookSplitManagerEntryPoint -Times 3 -Exactly
    }

    It 'binds the alias adapter to the complete validated package and application/document exclusions' {
        $script:managerPackage = New-ManagerPackage
        $script:managerDocument = Join-Path $WorkRoot 'authored-input.pdf'
        $json = New-ManagerJson -Packages @($script:managerPackage)
        $paths = @(ConvertFrom-WinBookSplitManagerRegistration -Probe (New-ManagerProbe -Stdout $json) -ApplicationRoot $RepositoryRoot -DocumentPath $script:managerDocument)
        $paths[0] | Should -BeExactly (Get-ManagerFixtureAlias -Package $script:managerPackage)
        Should-Invoke Get-WinBookSplitTrustedPath -Times 1 -Exactly -ParameterFilter {
            $Directory -and $Path -ceq $script:managerPackage.InstallLocation
        }
        Should-Invoke Get-WinBookSplitManagerEntryPoint -Times 1 -Exactly -ParameterFilter {
            $Package.PackageFullName -ceq $script:managerPackage.PackageFullName -and
            $Package.PackageFamilyName -ceq $script:managerPackage.PackageFamilyName -and
            $Package.Publisher -ceq $script:managerPackage.Publisher -and
            $ApplicationRoot -ceq $RepositoryRoot -and $DocumentPath -ceq $script:managerDocument
        }
    }

    It 'retains validated package data only when explicitly requested for supervised launch' {
        $package = New-ManagerPackage
        $json = New-ManagerJson -Packages @($package)
        $registrations = @(ConvertFrom-WinBookSplitManagerRegistration -Probe (New-ManagerProbe -Stdout $json) -ApplicationRoot $RepositoryRoot -WithRegistration)
        $registrations.Count | Should -Be 1
        $registrations[0].Path | Should -BeExactly (Get-ManagerFixtureAlias -Package $package)
        $registrations[0].Package.PackageFullName | Should -BeExactly $package.PackageFullName
        $registrations[0].Package.Publisher | Should -BeExactly $package.Publisher
        $registrations[0].Package.SignatureKind | Should -BeExactly 'Store'
    }

    It 'rejects all six cross-paired publisher family and signature combinations before alias binding' {
        $store = New-ManagerPackage
        $pythonOrg = New-ManagerPackage -Id REG02
        foreach ($publisher in @($store.Publisher,$pythonOrg.Publisher)) {
            foreach ($identity in @($store,$pythonOrg)) {
                foreach ($signature in @('Store','Developer')) {
                    if (($publisher -ceq $store.Publisher -and $identity.PublisherId -ceq $store.PublisherId -and $signature -ceq 'Store') -or
                        ($publisher -ceq $pythonOrg.Publisher -and $identity.PublisherId -ceq $pythonOrg.PublisherId -and $signature -ceq 'Developer')) { continue }
                    $package = New-ManagerPackage
                    $package.Publisher = $publisher
                    $package.PublisherId = $identity.PublisherId
                    $package.PackageFamilyName = $identity.PackageFamilyName
                    $package.SignatureKind = $signature
                    Set-ManagerFixtureVersion -Package $package -Version $package.Version
                    $json = New-ManagerJson -Packages @($package)
                    { ConvertFrom-WinBookSplitManagerRegistration -Probe (New-ManagerProbe -Stdout $json) -ApplicationRoot $RepositoryRoot } | Should -Throw
                }
            }
        }
        Should-Invoke Get-WinBookSplitTrustedPath -Times 0 -Exactly
        Should-Invoke Get-WinBookSplitManagerEntryPoint -Times 0 -Exactly
    }

    It 'accepts canonical manager version variation through the unsigned 16-bit component boundary for both channels' {
        foreach ($id in @('REG01','REG02')) {
            foreach ($version in @('0.0.0.0','1.0.0.0','26.4.240.0','27.1.0.7','65535.65535.65535.65535')) {
                $package = New-ManagerPackage -Id $id
                Set-ManagerFixtureVersion -Package $package -Version $version
                $json = New-ManagerJson -Packages @($package)
                $paths = @(ConvertFrom-WinBookSplitManagerRegistration -Probe (New-ManagerProbe -Stdout $json) -ApplicationRoot $RepositoryRoot)
                $paths.Count | Should -Be 1 -Because ($id + ' version ' + $version)
                $paths[0] | Should -BeExactly (Get-ManagerFixtureAlias -Package $package)
            }
        }
        Should-Invoke Get-WinBookSplitManagerEntryPoint -Times 10 -Exactly
    }

    It 'rejects out-of-range noncanonical or malformed versions even when full name and location agree' {
        foreach ($version in @('65536.0.0.0','0.65536.0.0','0.0.65536.0','0.0.0.65536','99999.1.2.3',
                '01.2.3.4','1.02.3.4','1.2.03.4','1.2.3.04','1.2.3','1.2.3.4.5','-1.2.3.4','1.2.3.4 ',' 1.2.3.4','1e2.2.3.4','1.2.3.4-beta')) {
            $package = New-ManagerPackage
            Set-ManagerFixtureVersion -Package $package -Version $version
            $json = New-ManagerJson -Packages @($package)
            { ConvertFrom-WinBookSplitManagerRegistration -Probe (New-ManagerProbe -Stdout $json) -ApplicationRoot $RepositoryRoot } | Should -Throw -Because $version
        }
        Should-Invoke Get-WinBookSplitTrustedPath -Times 0 -Exactly
        Should-Invoke Get-WinBookSplitManagerEntryPoint -Times 0 -Exactly
    }

    It 'requires every exact identity string field and rejects mismatched full names and families' {
        foreach ($field in @('Name','Publisher','PublisherId','PackageFamilyName','PackageFullName','Architecture','SignatureKind','Status','Version')) {
            foreach ($value in @('', 'SYNTHETIC untrusted identity', $null, $true, 1, @('unexpected'))) {
                $package = New-ManagerPackage
                $package.$field = $value
                $json = New-ManagerJson -Packages @($package)
                { ConvertFrom-WinBookSplitManagerRegistration -Probe (New-ManagerProbe -Stdout $json) -ApplicationRoot $RepositoryRoot } | Should -Throw -Because $field
            }
        }
        Should-Invoke Get-WinBookSplitManagerEntryPoint -Times 0 -Exactly
    }

    It 'rejects true and nonboolean development staging framework or resource flags' {
        foreach ($field in @('IsDevelopmentMode','IsPartiallyStaged','IsFramework','IsResourcePackage')) {
            foreach ($value in @($true, 'false', 0, $null, @($false))) {
                $package = New-ManagerPackage
                $package.$field = $value
                $json = New-ManagerJson -Packages @($package)
                { ConvertFrom-WinBookSplitManagerRegistration -Probe (New-ManagerProbe -Stdout $json) -ApplicationRoot $RepositoryRoot } | Should -Throw -Because $field
            }
        }
        Should-Invoke Get-WinBookSplitTrustedPath -Times 0 -Exactly
        Should-Invoke Get-WinBookSplitManagerEntryPoint -Times 0 -Exactly
    }

    It 'rejects missing and unknown package fields and duplicate package or JSON identities' {
        $original = New-ManagerPackage
        foreach ($field in @($original.PSObject.Properties.Name)) {
            $package = New-ManagerPackage
            $package.PSObject.Properties.Remove($field)
            $json = New-ManagerJson -Packages @($package)
            { ConvertFrom-WinBookSplitManagerRegistration -Probe (New-ManagerProbe -Stdout $json) -ApplicationRoot $RepositoryRoot } | Should -Throw -Because $field
        }
        $package = New-ManagerPackage
        $package | Add-Member -NotePropertyName extra -NotePropertyValue 'unexpected'
        $json = New-ManagerJson -Packages @($package)
        { ConvertFrom-WinBookSplitManagerRegistration -Probe (New-ManagerProbe -Stdout $json) -ApplicationRoot $RepositoryRoot } | Should -Throw
        $package = New-ManagerPackage
        $json = New-ManagerJson -Packages @($package,$package)
        { ConvertFrom-WinBookSplitManagerRegistration -Probe (New-ManagerProbe -Stdout $json) -ApplicationRoot $RepositoryRoot } | Should -Throw
        $json = (New-ManagerJson -Packages @($package)).Replace('"Name":','"Name":"duplicate","Name":')
        { ConvertFrom-WinBookSplitManagerRegistration -Probe (New-ManagerProbe -Stdout $json) -ApplicationRoot $RepositoryRoot } | Should -Throw
        # One first valid record may bind before its duplicate is rejected, but
        # the parser must never return a partial result to its caller.
        Should-Invoke Get-WinBookSplitManagerEntryPoint -Times 1 -Exactly
    }

    It 'rejects malformed oversized or unknown registration schema before alias binding' {
        foreach ($json in @('{"schema_version":true,"packages":[]}', '{"schema_version":2,"packages":[]}',
                '{"schema_version":1,"schema_version":1,"packages":[]}', '{"schema_version":1,"packages":{}}',
                '{"schema_version":1,"packages":null}', '{"schema_version":1,"packages":[],"extra":false}',
                '{"packages":[]}', '{"schema_version":1}', '[]', '{', (' ' * 32769))) {
            { ConvertFrom-WinBookSplitManagerRegistration -Probe (New-ManagerProbe -Stdout $json) -ApplicationRoot $RepositoryRoot } | Should -Throw
        }
        $json = New-ManagerJson -Packages @((1..17 | ForEach-Object { New-ManagerPackage }))
        { ConvertFrom-WinBookSplitManagerRegistration -Probe (New-ManagerProbe -Stdout $json) -ApplicationRoot $RepositoryRoot } | Should -Throw
        Should-Invoke Get-WinBookSplitManagerEntryPoint -Times 0 -Exactly
    }

    It 'rejects an otherwise valid identity if physical registration origin or the narrow alias binding fails' {
        $json = New-ManagerJson -Packages @(New-ManagerPackage)
        Mock Get-WinBookSplitTrustedPath { throw 'SYNTHETIC reparse physical package origin.' }
        { ConvertFrom-WinBookSplitManagerRegistration -Probe (New-ManagerProbe -Stdout $json) -ApplicationRoot $RepositoryRoot } | Should -Throw
        Should-Invoke Get-WinBookSplitManagerEntryPoint -Times 0 -Exactly
        Mock Get-WinBookSplitTrustedPath { return $Path }
        Mock Get-WinBookSplitManagerEntryPoint { throw 'SYNTHETIC package-family alias binding rejected.' }
        { ConvertFrom-WinBookSplitManagerRegistration -Probe (New-ManagerProbe -Stdout $json) -ApplicationRoot $RepositoryRoot } | Should -Throw
        Should-Invoke Get-WinBookSplitManagerEntryPoint -Times 1 -Exactly
    }

    It 'retains complete bounded stream and process-result requirements for registration records' {
        $json = New-ManagerJson -Packages @(New-ManagerPackage)
        foreach ($field in @('JobAssigned','StreamsComplete','ParentStopped','DescendantsStopped')) {
            $probe = New-ManagerProbe -Stdout $json
            $probe.$field = $false
            { ConvertFrom-WinBookSplitManagerRegistration -Probe $probe -ApplicationRoot $RepositoryRoot } | Should -Throw
        }
        foreach ($field in @('TimedOut','Cancelled','StdoutTruncated','StderrTruncated')) {
            $probe = New-ManagerProbe -Stdout $json
            $probe.$field = $true
            { ConvertFrom-WinBookSplitManagerRegistration -Probe $probe -ApplicationRoot $RepositoryRoot } | Should -Throw
        }
        foreach ($field in @('StartError','StopError','StreamError')) {
            $probe = New-ManagerProbe -Stdout $json
            $probe.$field = 'SYNTHETIC transport failure'
            { ConvertFrom-WinBookSplitManagerRegistration -Probe $probe -ApplicationRoot $RepositoryRoot } | Should -Throw
        }
        { ConvertFrom-WinBookSplitManagerRegistration -Probe (New-ManagerProbe -Stdout $json -ExitCode 2) -ApplicationRoot $RepositoryRoot } | Should -Throw
        { ConvertFrom-WinBookSplitManagerRegistration -Probe (New-ManagerProbe -Stdout $json -Stderr 'SYNTHETIC unexpected diagnostics') -ApplicationRoot $RepositoryRoot } | Should -Throw
        Should-Invoke Get-WinBookSplitManagerEntryPoint -Times 0 -Exactly
    }
}

Describe 'T02 emitted current-user registration query (SYNTHETIC inventory)' -Tag 'T02-Manager','STORE-007','STORE-008','STORE-009' {
    BeforeEach {
        $script:managerQueryPackages = @(New-ManagerPackage)
        Mock Get-AppxPackage {
            foreach ($package in $script:managerQueryPackages) {
                if ($package.Name -ceq $Name -and ([string]::IsNullOrEmpty($Publisher) -or $package.Publisher -ceq $Publisher)) { $package }
            }
        }
        Mock Get-WinBookSplitTrustedPath { return $Path }
        Mock Get-WinBookSplitManagerEntryPoint { return Get-ManagerFixtureAlias -Package $Package }
        Mock Invoke-WinBookSplitRuntimeProbe { Invoke-ManagerQueryFixture -Arguments $Arguments }
    }

    It 'passes Store and python.org positive controls through the real exact-name query without a publisher filter or all-users scope' {
        $script:managerQueryPackages = @(New-ManagerPackage; New-ManagerPackage -Id REG02)
        $attempts = New-Object 'Collections.Generic.List[object]'
        $paths = @(Get-WinBookSplitRegisteredManagerPaths -ApplicationRoot $RepositoryRoot -Attempts $attempts -TimeoutSeconds 2.5)
        $paths.Count | Should -Be 2
        $paths[0] | Should -BeExactly (Get-ManagerFixtureAlias -Package $script:managerQueryPackages[0])
        $paths[1] | Should -BeExactly (Get-ManagerFixtureAlias -Package $script:managerQueryPackages[1])
        $attempts.Count | Should -Be 1
        @(($attempts[0].Probe.Stdout | ConvertFrom-Json).packages).Count | Should -Be 2
        Should-Invoke Get-AppxPackage -Times 1 -Exactly -ParameterFilter {
            $Name -ceq 'PythonSoftwareFoundation.PythonManager' -and [string]::IsNullOrEmpty($Publisher) -and
            -not $AllUsers -and [string]::IsNullOrEmpty($User)
        }
        Should-Invoke Invoke-WinBookSplitRuntimeProbe -Times 1 -Exactly -ParameterFilter {
            $TimeoutSeconds -eq 2.5 -and $ApplicationRoot -ceq $RepositoryRoot -and
            $Arguments[0] -ceq '-NoProfile' -and $Arguments[1] -ceq '-NonInteractive' -and $Arguments[2] -ceq '-EncodedCommand'
        }
        Should-Invoke Get-WinBookSplitManagerEntryPoint -Times 2 -Exactly
    }

    It 'returns an empty query without accepting any manager entry point' {
        $script:managerQueryPackages = @()
        $attempts = New-Object 'Collections.Generic.List[object]'
        @(Get-WinBookSplitRegisteredManagerPaths -ApplicationRoot $RepositoryRoot -Attempts $attempts).Count | Should -Be 0
        $attempts.Count | Should -Be 1
        @(($attempts[0].Probe.Stdout | ConvertFrom-Json).packages).Count | Should -Be 0
        Should-Invoke Get-AppxPackage -Times 1 -Exactly
        Should-Invoke Invoke-WinBookSplitRuntimeProbe -Times 1 -Exactly
        Should-Invoke Get-WinBookSplitManagerEntryPoint -Times 0 -Exactly
    }

    It 'records an invalid queried identity separately from a successful empty inventory' {
        $script:managerQueryPackages = @(New-ManagerPackage -Id REG03)
        $attempts = New-Object 'Collections.Generic.List[object]'
        @(Get-WinBookSplitRegisteredManagerPaths -ApplicationRoot $RepositoryRoot -Attempts $attempts).Count | Should -Be 0
        $attempts.Count | Should -Be 1
        $attempts[0].Accepted | Should -BeFalse
        @(($attempts[0].Probe.Stdout | ConvertFrom-Json).packages).Count | Should -Be 1
        Should-Invoke Get-AppxPackage -Times 1 -Exactly
        Should-Invoke Get-WinBookSplitManagerEntryPoint -Times 0 -Exactly
    }

    It 'keeps nonzero and timeout query results bounded failures with the actual probe attached' {
        $script:managerFailedQuery = New-ManagerProbe -Stderr 'SYNTHETIC query failure' -ExitCode 5
        Mock Invoke-WinBookSplitRuntimeProbe { return $script:managerFailedQuery }
        foreach ($timedOut in @($false,$true)) {
            $script:managerFailedQuery.TimedOut = $timedOut
            $attempts = New-Object 'Collections.Generic.List[object]'
            @(Get-WinBookSplitRegisteredManagerPaths -ApplicationRoot $RepositoryRoot -Attempts $attempts -TimeoutSeconds 1).Count | Should -Be 0
            $attempts.Count | Should -Be 1
            $attempts[0].Accepted | Should -BeFalse
            $attempts[0].Probe.ExitCode | Should -Be 5
            $attempts[0].Probe.TimedOut | Should -Be $timedOut
            $attempts[0].Probe.Stderr | Should -BeExactly 'SYNTHETIC query failure'
        }
        Should-Invoke Invoke-WinBookSplitRuntimeProbe -Times 2 -Exactly -ParameterFilter { $TimeoutSeconds -eq 1 }
        Should-Invoke Get-AppxPackage -Times 0 -Exactly
        Should-Invoke Get-WinBookSplitManagerEntryPoint -Times 0 -Exactly
    }
}

Describe 'T02 narrow package-family alias guard (SYNTHETIC authored files)' -Tag 'T02-Manager','STORE-010','STORE-011' {
    It 'asks the ordinary directory guard for only the documented KnownFolder package-family directory' {
        $script:managerPackage = New-ManagerPackage
        $script:managerExpectedDirectory = [IO.Path]::GetDirectoryName((Get-ManagerFixtureAlias -Package $script:managerPackage))
        Mock Get-WinBookSplitTrustedPath { throw 'SYNTHETIC directory guard rejection.' }
        Mock Invoke-WinBookSplitProcess { throw 'A rejected directory must never be executed.' }
        { Get-WinBookSplitManagerEntryPoint -Package $script:managerPackage -ApplicationRoot $RepositoryRoot } |
            Should -Throw -ExpectedMessage '*SYNTHETIC directory guard rejection*'
        Should-Invoke Get-WinBookSplitTrustedPath -Times 1 -Exactly -ParameterFilter { $Directory -and $Path -ceq $script:managerExpectedDirectory }
        Should-Invoke Invoke-WinBookSplitProcess -Times 0 -Exactly
    }

    It 'retains document and unrelated-CWD exclusion before inspecting any alias file' {
        $script:managerPackage = New-ManagerPackage
        $script:managerAlias = Get-ManagerFixtureAlias -Package $script:managerPackage
        $script:managerDocument = Join-Path $WorkRoot 'authored-input.pdf'
        Mock Get-WinBookSplitTrustedPath { return $Path }
        Mock Test-WinBookSplitAutomaticPath { return $false }
        Mock Invoke-WinBookSplitProcess { throw 'An excluded location must never be executed.' }
        { Get-WinBookSplitManagerEntryPoint -Package $script:managerPackage -ApplicationRoot $RepositoryRoot -DocumentPath $script:managerDocument } | Should -Throw
        Should-Invoke Test-WinBookSplitAutomaticPath -Times 1 -Exactly -ParameterFilter {
            $Path -ceq $script:managerPackage.InstallLocation -and $ApplicationRoot -ceq $RepositoryRoot -and $DocumentPath -ceq $script:managerDocument
        }
        Mock Test-WinBookSplitAutomaticPath { return ($Path -cne $script:managerAlias) }
        { Get-WinBookSplitManagerEntryPoint -Package $script:managerPackage -ApplicationRoot $RepositoryRoot -DocumentPath $script:managerDocument } | Should -Throw
        Should-Invoke Test-WinBookSplitAutomaticPath -Times 1 -Exactly -ParameterFilter { $Path -ceq $script:managerAlias }
        Should-Invoke Invoke-WinBookSplitProcess -Times 0 -Exactly
    }

    It 'rejects an ordinary file and hard-linked file as AppExecutionAlias decoys without executing either' {
        $directory = Join-Path $TestDrive ('manager-decoy-' + [guid]::NewGuid().ToString('N'))
        $null = New-Item -ItemType Directory -Path $directory
        $script:managerFixtureDirectory = $directory
        $path = Join-Path $directory 'pymanager.exe'
        $source = Join-Path $directory 'authored-source.txt'
        $text = 'SYNTHETIC non-executable alias decoy; never launched.'
        [IO.File]::WriteAllText($path,$text)
        [IO.File]::WriteAllText($source,$text)
        Mock Get-WinBookSplitTrustedPath { return $script:managerFixtureDirectory }
        Mock Test-WinBookSplitAutomaticPath { return $true }
        Mock Invoke-WinBookSplitProcess { throw 'A synthetic file must never be executed.' }
        { Get-WinBookSplitManagerEntryPoint -Package (New-ManagerPackage) -ApplicationRoot $RepositoryRoot } |
            Should -Throw -ExpectedMessage '*must be an AppExecutionAlias*'
        [IO.File]::ReadAllText($path) | Should -BeExactly $text
        [IO.File]::Delete($path)
        $null = New-Item -ItemType HardLink -Path $path -Target $source
        { Get-WinBookSplitManagerEntryPoint -Package (New-ManagerPackage) -ApplicationRoot $RepositoryRoot } |
            Should -Throw -ExpectedMessage '*must be an AppExecutionAlias*'
        [IO.File]::ReadAllText($source) | Should -BeExactly $text
        Should-Invoke Invoke-WinBookSplitProcess -Times 0 -Exactly
    }

    It 'rejects a junction carrying the alias filename without following it as an AppExecutionAlias' {
        $directory = Join-Path $TestDrive ('manager-junction-' + [guid]::NewGuid().ToString('N'))
        $target = Join-Path $TestDrive ('manager-target-' + [guid]::NewGuid().ToString('N'))
        $null = New-Item -ItemType Directory -Path $directory
        $null = New-Item -ItemType Directory -Path $target
        $path = Join-Path $directory 'pymanager.exe'
        $marker = Join-Path $target 'authored-marker.txt'
        [IO.File]::WriteAllText($marker,'SYNTHETIC junction target; never executed.')
        $script:managerFixtureDirectory = $directory
        $null = New-Item -ItemType Junction -Path $path -Target $target
        Mock Get-WinBookSplitTrustedPath { return $script:managerFixtureDirectory }
        Mock Test-WinBookSplitAutomaticPath { return $true }
        Mock Invoke-WinBookSplitProcess { throw 'A synthetic junction must never be executed.' }
        try {
            { Get-WinBookSplitManagerEntryPoint -Package (New-ManagerPackage) -ApplicationRoot $RepositoryRoot } |
                Should -Throw -ExpectedMessage '*must be an AppExecutionAlias*'
            [IO.File]::ReadAllText($marker) | Should -BeExactly 'SYNTHETIC junction target; never executed.'
            Should-Invoke Invoke-WinBookSplitProcess -Times 0 -Exactly
        }
        finally { [IO.Directory]::Delete($path,$false) }
    }
}

Describe 'T02 manager listing supervisor adapter (SYNTHETIC transport)' -Tag 'T02-Manager','STORE-010','STORE-011','STORE-012' {
    BeforeEach {
        $script:managerPackage = New-ManagerPackage
        $script:managerAlias = Get-ManagerFixtureAlias -Package $script:managerPackage
        $script:managerRegistration = [pscustomobject]@{Path=$script:managerAlias;Package=$script:managerPackage}
        Mock Get-WinBookSplitTrustedPath { return $Path }
        Mock Get-WinBookSplitManagerEntryPoint { return Get-ManagerFixtureAlias -Package $Package }
        Mock Invoke-WinBookSplitProcess { return New-ManagerProbe -Stdout "C:\SYNTHETIC installed runtime\python.exe`r`n" }
    }

    It 'revalidates and passes the exact installed-only argument vector package and physical image to the bounded supervisor' {
        $probe = Invoke-WinBookSplitManagerListing -Registration $script:managerRegistration -ApplicationRoot $RepositoryRoot -TimeoutSeconds 1.5
        $probe.ExitCode | Should -Be 0
        $probe.Stdout | Should -BeExactly "C:\SYNTHETIC installed runtime\python.exe`r`n"
        Should-Invoke Get-WinBookSplitManagerEntryPoint -Times 1 -Exactly
        Should-Invoke Invoke-WinBookSplitProcess -Times 1 -Exactly -ParameterFilter {
            $Path -ceq $script:managerAlias -and ($Arguments -join ',') -ceq 'list,--format=exe' -and
            $WorkingDirectory -ceq $RepositoryRoot -and $TimeoutSeconds -eq 1.5 -and
            $ExpectedPackageFullName -ceq $script:managerPackage.PackageFullName -and
            $ExpectedImagePath -ceq [IO.Path]::Combine($script:managerPackage.InstallLocation,'pymanager.exe')
        }
    }

    It 'does not reach the supervisor after the narrow helper rejects the alias, with a transport positive control' {
        (Invoke-WinBookSplitManagerListing -Registration $script:managerRegistration -ApplicationRoot $RepositoryRoot).ExitCode | Should -Be 0
        Should-Invoke Invoke-WinBookSplitProcess -Times 1 -Exactly
        Mock Get-WinBookSplitManagerEntryPoint { throw 'SYNTHETIC rejected alias binding.' }
        { Invoke-WinBookSplitManagerListing -Registration $script:managerRegistration -ApplicationRoot $RepositoryRoot } |
            Should -Throw -ExpectedMessage '*SYNTHETIC rejected alias binding*'
        Should-Invoke Invoke-WinBookSplitProcess -Times 1 -Exactly
    }

    It 'rejects a changed registered alias or invalid identity before supervisor execution' {
        $script:managerRegistration.Path = 'C:\SYNTHETIC unbound\pymanager.exe'
        { Invoke-WinBookSplitManagerListing -Registration $script:managerRegistration -ApplicationRoot $RepositoryRoot } |
            Should -Throw -ExpectedMessage '*does not match its registration*'
        $script:managerRegistration.Path = $script:managerAlias
        $script:managerRegistration.Package = New-ManagerPackage -Id REG03
        { Invoke-WinBookSplitManagerListing -Registration $script:managerRegistration -ApplicationRoot $RepositoryRoot } | Should -Throw
        Should-Invoke Get-WinBookSplitManagerEntryPoint -Times 1 -Exactly
        Should-Invoke Invoke-WinBookSplitProcess -Times 0 -Exactly
    }
}
