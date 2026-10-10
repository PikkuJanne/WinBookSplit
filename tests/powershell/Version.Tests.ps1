param([string]$RepositoryRoot, [string]$WorkRoot, [string]$TestPython)

BeforeAll {
    foreach ($source in @(@('WinBookSplit.ps1', 'Get-WinBookSplitApplicationVersion'),
                         @('engine\WinBookSplit.Support.ps1', 'Get-WinBookSplitSupportApplicationVersion'))) {
        $tokens = $null; $errors = $null
        $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $RepositoryRoot $source[0]), [ref]$tokens, [ref]$errors)
        if (@($errors).Count -ne 0) { throw 'Trusted version consumer has syntax errors.' }
        $name = $source[1]
        $functions = @($ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -ceq $name }, $true))
        if ($functions.Count -ne 1) { throw 'Expected exactly one trusted version reader.' }
        . ([scriptblock]::Create($functions[0].Extent.Text))
    }
    $versionWork = Join-Path $WorkRoot ('version-' + [guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $versionWork
    $versionPath = Join-Path $versionWork 'VERSION'
}

Describe 'Canonical application VERSION' -Tag 'AC-082', 'AC-083' {
    It 'reads the same shipped candidate in Main and support without their runtime helpers' {
        Get-WinBookSplitApplicationVersion -Root $RepositoryRoot | Should -BeExactly '1.0.0'
        Get-WinBookSplitSupportApplicationVersion -Root $RepositoryRoot | Should -BeExactly '1.0.0'
    }
    It 'accepts bounded ASCII numeric versions with one optional LF or CRLF' {
        foreach ($text in @('1.0.0', "1.0.0`n", "1.0.0`r`n", '0.0.0', "1234567890.1234567890.1234567890`r`n")) {
            [IO.File]::WriteAllBytes($versionPath, [Text.Encoding]::ASCII.GetBytes($text))
            Get-WinBookSplitApplicationVersion -Root $versionWork | Should -BeExactly $text.TrimEnd([char[]]"`r`n")
            Get-WinBookSplitSupportApplicationVersion -Root $versionWork | Should -BeExactly $text.TrimEnd([char[]]"`r`n")
        }
    }
    It 'fails closed for missing VERSION despite an obsolete outcome contract beside it' {
        [IO.File]::WriteAllText((Join-Path $versionWork 'WinBookSplit.Outcomes.json'), '{"application_version":"1.0.0"}')
        if ([IO.File]::Exists($versionPath)) { [IO.File]::Delete($versionPath) }
        { Get-WinBookSplitApplicationVersion -Root $versionWork } | Should -Throw '*Missing or malformed shipped application VERSION*'
        { Get-WinBookSplitSupportApplicationVersion -Root $versionWork } | Should -Throw '*Missing or malformed shipped application VERSION*'
    }
    It 'refuses malformed, non-ASCII, extra line endings and oversized data in both consumers' {
        $texts = @('', '1.0.0-dev', 'v1.0.0', '01.0.0', '1.00.0', '1.0.00', '-1.0.0', '1.0', '1.0.0.0',
                   "1.0.0`r", "1.0.0`n`n", "1.0.0`r`n`r`n", ' 1.0.0', '1.0.0 ', "1.0.0`0", ('9' * 100000))
        $vectors = @($texts | ForEach-Object { ,([Text.Encoding]::ASCII.GetBytes($_)) })
        $vectors += ,([byte[]]@(0xef,0xbb,0xbf,0x31,0x2e,0x30,0x2e,0x30))
        $vectors += ,([byte[]]@(0xff,0x2e,0x30,0x2e,0x30))
        foreach ($bytes in $vectors) {
            [IO.File]::WriteAllBytes($versionPath, $bytes)
            { Get-WinBookSplitApplicationVersion -Root $versionWork } | Should -Throw '*Missing or malformed shipped application VERSION*'
            { Get-WinBookSplitSupportApplicationVersion -Root $versionWork } | Should -Throw '*Missing or malformed shipped application VERSION*'
        }
    }
}
