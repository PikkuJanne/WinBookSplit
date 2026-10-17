param([string]$RepositoryRoot, [string]$WorkRoot, [switch]$DesiredStoreBehavior)

# T01 incident reproduction only. All executable files, links, registrations and
# transport results below are SYNTHETIC TEST INPUTS, never native Store evidence.
# The ordinary shell runner uses an explicit container list and does not select
# this file. Run T01-Store for baseline evidence; opt into DesiredStoreBehavior
# and T01-Store-Desired to record the three expected RED repair contracts.
BeforeAll {
    . (Join-Path $RepositoryRoot 'engine\WinBookSplit.Runtime.ps1')
    $tokens = $null; $errors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $RepositoryRoot 'WinBookSplit.ps1'), [ref]$tokens, [ref]$errors)
    if (@($errors).Count -ne 0) { throw 'The trusted application has syntax errors.' }
    $summary = @($ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and
        $node.Name -ceq 'Get-WinBookSplitDependencySummary' }, $true))
    if ($summary.Count -ne 1) { throw 'Expected one trusted dependency summary function.' }
    . ([scriptblock]::Create($summary[0].Extent.Text))
    $script:storeIncident = Get-Content -LiteralPath (Join-Path $RepositoryRoot 'docs\codex-v1.0.3\fixtures\incident.json') -Raw | ConvertFrom-Json
    $script:storeRegistrationCases = Get-Content -LiteralPath (Join-Path $RepositoryRoot 'docs\codex-v1.0.3\fixtures\registration-cases.json') -Raw | ConvertFrom-Json

    function New-StoreProbe {
        param([string]$Stdout)
        return [pscustomobject]@{ Stdout=$Stdout; Stderr=''; ExitCode=0; TimedOut=$false;
            ParentStopped=$true; DescendantsStopped=$true; StreamsComplete=$true; JobAssigned=$true; Cancelled=$false;
            StartError=$null; StopError=$null; StreamError=$null; StdoutTruncated=$false; StderrTruncated=$false }
    }
    function New-StoreRegistration {
        param([ValidateSet('Store','PythonOrg')][string]$Channel = 'Store')
        $id = if ($Channel -eq 'Store') { 'REG01' } else { 'REG02' }
        $case = @($script:storeRegistrationCases.cases | Where-Object { $_.id -ceq $id })[0]
        if ($case.kind -cne 'SYNTHETIC_TEST_INPUT') { throw 'Completed registration fixtures must remain labelled synthetic.' }
        $package = $case.registration | ConvertTo-Json -Depth 6 | ConvertFrom-Json
        # Only the observed tuple is incident evidence. Version/flags, physical
        # location and successful transport are synthetic schema/control inputs.
        $package.InstallLocation = [IO.Path]::Combine([Environment]::GetFolderPath([Environment+SpecialFolder]::ProgramFiles), 'WindowsApps', $package.PackageFullName)
        return $package
    }
    function Get-AppxPackage {
        [CmdletBinding()]
        param([string]$Name, [string]$Publisher)
        throw 'The SYNTHETIC Get-AppxPackage positive-control mock was not installed.'
    }
    function Invoke-StoreQueryFixture {
        param([string[]]$Arguments)
        if ($Arguments.Count -ne 4 -or $Arguments[2] -cne '-EncodedCommand') { throw 'Expected the real bounded registration query arguments.' }
        $command = [Text.Encoding]::Unicode.GetString([Convert]::FromBase64String($Arguments[3]))
        $queryTokens = $null; $queryErrors = $null
        $queryAst = [Management.Automation.Language.Parser]::ParseInput($command, [ref]$queryTokens, [ref]$queryErrors)
        if (@($queryErrors).Count -ne 0) { throw 'The real registration query contains syntax errors.' }
        $calls = @($queryAst.FindAll({ param($node) $node -is [Management.Automation.Language.CommandAst] -and
            $node.GetCommandName() -ceq 'Get-AppxPackage' }, $true))
        if ($calls.Count -ne 1) { throw 'Expected exactly one current-user package query.' }
        # Execute the real query CommandAst against an in-process synthetic
        # package inventory. No Appx import, package query or child launch occurs.
        $packages = @(& ([scriptblock]::Create($calls[0].Extent.Text)))
        return New-StoreProbe -Stdout (ConvertTo-Json -InputObject ([ordered]@{schema_version=1;packages=$packages}) -Depth 6 -Compress)
    }
    function New-StoreFailureFromIncident {
        $attempts = New-Object 'Collections.Generic.List[object]'
        foreach ($attempt in $script:storeIncident.original_attempts) {
            $path = $null; $message = 'Captured registration query completed with no packages.'
            if ($attempt.PSObject.Properties.Name -contains 'path_ref') { $path = $attempt.path_ref }
            if ($attempt.PSObject.Properties.Name -contains 'message') { $message = $attempt.message }
            $attempts.Add([pscustomobject]@{ Path=$path; Source=$attempt.source; Accepted=$attempt.accepted;
                Message=$message; Probe=$null })
        }
        return New-WinBookSplitDependencyError -Code 'runtime_not_found' -Message 'SYNTHETIC replay of captured uncategorized discovery attempts.' -Attempts $attempts
    }
    function New-StoreApplication {
        $app = Join-Path $TestDrive ('store-app-' + [guid]::NewGuid().ToString('N'))
        $engine = Join-Path $app 'engine'
        $null = New-Item -ItemType Directory -Path $engine
        foreach ($name in @('WinBookSplit.Compatibility.json','winbooksplit_compatibility.py')) {
            Copy-Item -LiteralPath (Join-Path $RepositoryRoot ('engine\' + $name)) -Destination $engine
        }
        return $app
    }
}

Describe 'T01 Store discovery baseline (SYNTHETIC replay, not native acceptance)' -Tag 'T01-Store','STORE-002','STORE-003','STORE-004','STORE-005' {
    BeforeEach {
        $script:storeQueryPackages = @(New-StoreRegistration)
        Mock Get-AppxPackage {
            foreach ($package in $script:storeQueryPackages) {
                if ($package.Name -ceq $Name -and ([string]::IsNullOrEmpty($Publisher) -or $package.Publisher -ceq $Publisher)) { $package }
            }
        }
        Mock Invoke-WinBookSplitRuntimeProbe { Invoke-StoreQueryFixture -Arguments $Arguments }
        Mock Find-WinBookSplitSetupBase { return $null }
    }

    It 'replays the emitted query and excludes the incident Store publisher before validation' {
        $attempts = New-Object 'Collections.Generic.List[object]'
        @(Get-WinBookSplitRegisteredManagerPaths -ApplicationRoot $RepositoryRoot -Attempts $attempts).Count | Should -Be 0
        $attempts.Count | Should -Be 1
        $attempts[0].Accepted | Should -BeTrue
        @(($attempts[0].Probe.Stdout | ConvertFrom-Json).packages).Count | Should -Be 0
        Should-Invoke Get-AppxPackage -Times 1 -Exactly -ParameterFilter {
            $Name -ceq 'PythonSoftwareFoundation.PythonManager' -and
            $Publisher -ceq 'CN=Python Software Foundation, O=Python Software Foundation, L=Beaverton, S=Oregon, C=US'
        }
        Should-Invoke Invoke-WinBookSplitRuntimeProbe -Times 1 -Exactly
    }

    It 'has a positive control: the same emitted query and validator accept synthetic python.org registration' {
        $script:storeQueryPackages = @(New-StoreRegistration -Channel PythonOrg)
        Mock Get-WinBookSplitTrustedPath { return $Path }
        Mock Test-WinBookSplitAutomaticPath { return $true }
        $attempts = New-Object 'Collections.Generic.List[object]'
        $paths = @(Get-WinBookSplitRegisteredManagerPaths -ApplicationRoot $RepositoryRoot -Attempts $attempts)
        $paths.Count | Should -Be 1
        $paths[0] | Should -BeExactly ([IO.Path]::Combine($script:storeQueryPackages[0].InstallLocation, 'pymanager.exe'))
        @(($attempts[0].Probe.Stdout | ConvertFrom-Json).packages).Count | Should -Be 1
        Should-Invoke Get-AppxPackage -Times 1 -Exactly
        Should-Invoke Invoke-WinBookSplitRuntimeProbe -Times 1 -Exactly
    }

    It 'still rejects the complete Store tuple downstream when a repaired query supplies it' {
        Mock Get-WinBookSplitTrustedPath { return $Path }
        Mock Test-WinBookSplitAutomaticPath { return $true }
        $json = ConvertTo-Json -InputObject ([ordered]@{schema_version=1;packages=@(New-StoreRegistration)}) -Depth 6 -Compress
        { ConvertFrom-WinBookSplitManagerRegistration -Probe (New-StoreProbe $json) -ApplicationRoot $RepositoryRoot } |
            Should -Throw -ExpectedMessage '*Unrecognized or duplicate Python manager registration identity*'
        Should-Invoke Get-WinBookSplitTrustedPath -Times 0 -Exactly
        Should-Invoke Invoke-WinBookSplitRuntimeProbe -Times 0 -Exactly
    }

    It 'replays all four original rejection paths with actual authored hard-link and reparse fixtures' {
        $app = New-StoreApplication
        $source = Join-Path $TestDrive 'synthetic-runtime-source.exe'
        $script:storePythonAlias = Join-Path $TestDrive 'synthetic-hardlink-python.exe'
        $script:storeManagerAlias = Join-Path $TestDrive 'synthetic-pymanager.exe'
        $target = Join-Path $TestDrive 'synthetic-py-target'
        $junction = Join-Path $TestDrive 'synthetic-py-junction'
        $null = New-Item -ItemType Directory -Path $target
        [IO.File]::WriteAllText($source, 'SYNTHETIC non-executable hard-link target; never launched.')
        [IO.File]::WriteAllText($script:storeManagerAlias, 'SYNTHETIC unregistered manager; never launched.')
        [IO.File]::WriteAllText((Join-Path $target 'py.exe'), 'SYNTHETIC reparse-ancestor target; never launched.')
        $null = New-Item -ItemType HardLink -Path $script:storePythonAlias -Target $source
        $null = New-Item -ItemType Junction -Path $junction -Target $target
        $script:storePyAlias = Join-Path $junction 'py.exe'
        Mock Get-WinBookSplitPathCandidates {
            switch ($Name) {
                'python.exe' { return @($script:storePythonAlias) }
                'pymanager.exe' { return @($script:storeManagerAlias) }
                'py.exe' { return @($script:storePyAlias) }
                default { throw 'Unexpected synthetic candidate query.' }
            }
        }
        $savedLocation = Get-Location
        $failure = $null
        try {
            # Do not make the fixture directory an unrelated-CWD exclusion.
            Set-Location -LiteralPath $RepositoryRoot
            try { $null = Resolve-WinBookSplitRuntime -ApplicationRoot $app -InterpreterOnly } catch { $failure = $_.Exception }
        }
        finally {
            Set-Location -LiteralPath $savedLocation.ProviderPath
            [IO.Directory]::Delete($junction, $false)
        }
        $failure | Should -Not -BeNullOrEmpty
        $failure.Data['Code'] | Should -BeExactly 'runtime_not_found'
        $attempts = @($failure.Data['Attempts'])
        $attempts.Count | Should -Be 4
        ($attempts.Source -join ',') | Should -BeExactly 'PATH,manager_registration,manager_listing,py_launcher_listing'
        $attempts[0].Message | Should -BeExactly $script:storeIncident.original_attempts[0].message
        $attempts[0].Probe | Should -BeNullOrEmpty
        $attempts[1].Accepted | Should -BeTrue
        @(($attempts[1].Probe.Stdout | ConvertFrom-Json).packages).Count | Should -Be 0
        $attempts[2].Message | Should -BeExactly $script:storeIncident.original_attempts[2].message
        $attempts[2].Probe | Should -BeNullOrEmpty
        $attempts[3].Message | Should -BeExactly $script:storeIncident.original_attempts[3].message
        $attempts[3].Probe | Should -BeNullOrEmpty
        [IO.File]::ReadAllText((Join-Path $target 'py.exe')) | Should -BeExactly 'SYNTHETIC reparse-ancestor target; never launched.'
        Should-Invoke Get-AppxPackage -Times 1 -Exactly
        Should-Invoke Invoke-WinBookSplitRuntimeProbe -Times 1 -Exactly -ParameterFilter { $Arguments[2] -ceq '-EncodedCommand' }
    }

    It 'reproduces missing_python from the captured uncategorized rejection records' {
        $summary = Get-WinBookSplitDependencySummary -Failure (New-StoreFailureFromIncident) -ApplicationRoot $RepositoryRoot
        $summary.Reason | Should -BeExactly 'missing_python'
        $summary.Message | Should -BeExactly 'No usable Python interpreter was found.'
        Should-Invoke Find-WinBookSplitSetupBase -Times 1 -Exactly
        Should-Invoke Invoke-WinBookSplitRuntimeProbe -Times 0 -Exactly
    }

    It 'keeps the observed success, failed PS5.1 collector and unknown base pypdf distinct' {
        $script:storeIncident.manager_registration.PackageFamilyName | Should -BeExactly 'PythonSoftwareFoundation.PythonManager_qbz5n2kfra8p0'
        $script:storeIncident.manager_registration.Publisher | Should -BeExactly 'CN=4975D53F-AA7E-49A5-8B49-EA4FDC1BB66B'
        $script:storeIncident.python_probe.exit_code | Should -Be 0
        $script:storeIncident.python_probe.version | Should -BeExactly '3.15.0'
        $script:storeIncident.ps51_python_probe.status | Should -BeExactly 'COLLECTOR_FAILED'
        $script:storeIncident.pypdf_in_base | Should -BeExactly 'UNKNOWN'
        $script:storeIncident.application_venv_exists | Should -BeFalse
    }
}

Describe 'T01 desired Store repair contracts (expected RED until the repair)' -Tag 'T01-Store-Desired' -Skip:(-not $DesiredStoreBehavior) {
    BeforeEach {
        $script:storeQueryPackages = @(New-StoreRegistration)
        Mock Get-AppxPackage {
            foreach ($package in $script:storeQueryPackages) {
                if ($package.Name -ceq $Name -and ([string]::IsNullOrEmpty($Publisher) -or $package.Publisher -ceq $Publisher)) { $package }
            }
        }
        Mock Invoke-WinBookSplitRuntimeProbe { Invoke-StoreQueryFixture -Arguments $Arguments }
        Mock Find-WinBookSplitSetupBase { return $null }
    }

    It 'does not remove an official Store tuple from the emitted current-user query' {
        Mock Get-WinBookSplitTrustedPath { return $Path }
        Mock Test-WinBookSplitAutomaticPath { return $true }
        $attempts = New-Object 'Collections.Generic.List[object]'
        $null = Get-WinBookSplitRegisteredManagerPaths -ApplicationRoot $RepositoryRoot -Attempts $attempts
        Should-Invoke Get-AppxPackage -Times 1 -Exactly
        Should-Invoke Invoke-WinBookSplitRuntimeProbe -Times 1 -Exactly
        @(($attempts[0].Probe.Stdout | ConvertFrom-Json).packages).Count | Should -Be 1
    }

    It 'accepts the complete paired official Store identity once the query returns it' {
        Mock Get-WinBookSplitTrustedPath { return $Path }
        Mock Test-WinBookSplitAutomaticPath { return $true }
        $json = ConvertTo-Json -InputObject ([ordered]@{schema_version=1;packages=@(New-StoreRegistration)}) -Depth 6 -Compress
        @(ConvertFrom-WinBookSplitManagerRegistration -Probe (New-StoreProbe $json) -ApplicationRoot $RepositoryRoot).Count | Should -Be 1
    }

    It 'does not categorize captured discovery rejections as absent Python' {
        # The eventual categorized reason is an implementation choice; the
        # captured candidates must not collapse to the absence category.
        (Get-WinBookSplitDependencySummary -Failure (New-StoreFailureFromIncident) -ApplicationRoot $RepositoryRoot).Reason |
            Should -Not -BeExactly 'missing_python'
    }
}
