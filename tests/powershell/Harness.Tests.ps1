param([string]$RepositoryRoot, [string]$ToolRoot, [string]$WorkRoot)

BeforeAll {
    Import-Module -Name (Join-Path -Path $RepositoryRoot -ChildPath 'tests\powershell\HarnessSupport.psm1') -Force -ErrorAction Stop
    function Invoke-RejectedShellInvocation {
        param([string]$RejectedToolRoot, [string]$RejectedWorkRoot, [string]$RejectedReportPath)
        $wrapper = Join-Path -Path $RepositoryRoot -ChildPath 'tests\Invoke-ShellTests.ps1'
        $values = @($wrapper, $RepositoryRoot, $RejectedToolRoot, $RejectedWorkRoot, $RejectedReportPath) |
            ForEach-Object { "'" + $_.Replace("'", "''") + "'" }
        $command = '$block = [scriptblock]::Create([IO.File]::ReadAllText(' + $values[0] + ')); & $block -RepositoryRoot ' +
            $values[1] + ' -ToolRoot ' + $values[2] + ' -WorkRoot ' + $values[3] + ' -ReportPath ' + $values[4]
        $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($command))
        $hostName = if ($PSVersionTable.PSEdition -eq 'Desktop') { 'powershell.exe' } else { 'pwsh.exe' }
        $startInfo = New-Object -TypeName Diagnostics.ProcessStartInfo
        $startInfo.FileName = Join-Path -Path $PSHOME -ChildPath $hostName
        $startInfo.Arguments = '-NoProfile -NonInteractive -ExecutionPolicy RemoteSigned -EncodedCommand ' + $encoded
        $startInfo.UseShellExecute = $false
        $startInfo.CreateNoWindow = $true
        $startInfo.RedirectStandardOutput = $true
        $startInfo.RedirectStandardError = $true
        $child = New-Object -TypeName Diagnostics.Process
        $child.StartInfo = $startInfo
        try {
            $null = $child.Start()
            $stdout = $child.StandardOutput.ReadToEndAsync()
            $stderr = $child.StandardError.ReadToEndAsync()
            if (-not $child.WaitForExit(15000)) {
                $child.Kill()
                $child.WaitForExit()
                throw 'Rejected-argument test child timed out.'
            }
            $null = $stdout.GetAwaiter().GetResult()
            $null = $stderr.GetAwaiter().GetResult()
            return $child.ExitCode
        }
        finally { $child.Dispose() }
    }
}

Describe 'M0 test harness failure accounting' -Tag 'AC-009' {
    It 'retains an earlier failure after later successful stages' {
        $outcomes = @(
            [pscustomobject]@{ passed = $false; name = 'failed test' },
            [pscustomobject]@{ passed = $true; name = 'successful native command' },
            [pscustomobject]@{ passed = $true; name = 'successful report' }
        )
        Get-TestOutcomeExitCode -Outcomes $outcomes | Should-Be -Expected 1
    }
    It 'rejects an empty run' {
        Get-TestOutcomeExitCode -Outcomes @() | Should-Be -Expected 1
    }
    It 'allows an entirely passing nonempty run' {
        Get-TestOutcomeExitCode -Outcomes @([pscustomobject]@{ passed = $true }) | Should-Be -Expected 0
    }
}

Describe 'M0 exact shell tool discovery' -Tag 'AC-010' {
    It 'selects the absolute Pester pin rather than an inbox module' {
        $manifest = Resolve-TestToolManifest -ToolRoot $ToolRoot -Name Pester
        [IO.Path]::IsPathRooted($manifest) | Should-BeTrue
        $manifest.EndsWith('\Pester\6.2.0\Pester.psd1', [StringComparison]::OrdinalIgnoreCase) | Should-BeTrue
    }
    It 'rejects a relative tool root' {
        { Resolve-TestToolManifest -ToolRoot '.' -Name Pester } | Should-Throw
    }
    It 'does not fall back when the pinned module is absent' {
        { Resolve-TestToolManifest -ToolRoot $TestDrive -Name Pester } | Should-Throw
    }
}

Describe 'M0 isolated literal fixture paths' -Tag 'AC-010' {
    It 'does not write a report when a required directory is missing' {
        $missingToolRoot = Join-Path -Path $TestDrive -ChildPath 'missing tools'
        $unvalidatedReport = Join-Path -Path $TestDrive -ChildPath 'unvalidated-report.json'
        Invoke-RejectedShellInvocation -RejectedToolRoot $missingToolRoot -RejectedWorkRoot $TestDrive -RejectedReportPath $unvalidatedReport | Should-Be -Expected 1
        Test-Path -LiteralPath $unvalidatedReport | Should-BeFalse
    }
    It 'does not write a report outside the supplied owned work directory' {
        $innerWork = Join-Path -Path $TestDrive -ChildPath 'owned-inner-work'
        $null = New-Item -ItemType Directory -Path $innerWork
        $outsideReport = Join-Path -Path $TestDrive -ChildPath 'outside-report.json'
        Invoke-RejectedShellInvocation -RejectedToolRoot $ToolRoot -RejectedWorkRoot $innerWork -RejectedReportPath $outsideReport | Should-Be -Expected 1
        Test-Path -LiteralPath $outsideReport | Should-BeFalse
    }
    It 'creates generated text only below the owned run directory' {
        $testDirectory = [IO.Path]::GetFullPath($TestDrive)
        $ownedDirectory = [IO.Path]::GetFullPath($WorkRoot).TrimEnd('\') + '\'
        $testDirectory.StartsWith($ownedDirectory, [StringComparison]::OrdinalIgnoreCase) | Should-BeTrue
        $fixture = Join-Path -Path $TestDrive -ChildPath 'original [fixture] - pages.txt'
        $sentinel = Join-Path -Path $TestDrive -ChildPath 'neighbor.txt'
        [IO.File]::WriteAllText($sentinel, 'neighbor belongs to this test')
        $sentinelBefore = (Get-FileHash -LiteralPath $sentinel -Algorithm SHA256).Hash
        [IO.File]::WriteAllText($fixture, 'Original synthetic fixture: page-001; page-002')
        [IO.File]::ReadAllText($fixture) | Should-MatchString -Expected 'page-001; page-002'
        (Get-FileHash -LiteralPath $sentinel -Algorithm SHA256).Hash | Should-Be -Expected $sentinelBefore
        # Pester owns TestDrive; its teardown performs fixture cleanup.
    }
}
