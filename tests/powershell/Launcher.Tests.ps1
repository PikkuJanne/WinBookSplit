param([string]$RepositoryRoot, [string]$WorkRoot)

BeforeAll {
    Set-StrictMode -Version Latest
    . (Join-Path $RepositoryRoot 'engine\WinBookSplit.Diagnostics.ps1')
    $tokens=$null; $errors=$null
    $ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $RepositoryRoot 'WinBookSplit.ps1'), [ref]$tokens, [ref]$errors)
    if (@($errors).Count -ne 0) { throw 'The trusted application has syntax errors.' }
    $functions=@($ast.FindAll({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and
        $node.Name -cin @('New-WinBookSplitFailure', 'Read-WinBookSplitSessionLine')}, $true))
    if ($functions.Count -ne 2) { throw 'Expected trusted failure and console reader functions.' }
    foreach ($function in $functions) { . ([scriptblock]::Create($function.Extent.Text)) }
    $script:launcherUnicode=([string][char]0x65E5)+[char]0x672C
    function Set-AuthoredLauncherAnswers {
        param([object[]]$Answers)
        $script:launcherAnswers=[Collections.Generic.Queue[object]]::new()
        foreach ($answer in $Answers) { $script:launcherAnswers.Enqueue($answer) }
    }
    function Get-AuthoredLauncherAnswer {
        if ($script:launcherAnswers.Count -eq 0) { throw 'Unexpected additional launcher prompt.' }
        return $script:launcherAnswers.Dequeue()
    }
}

Describe 'M3 exact initial method selection' -Tag 'AC-059' {
    It 'accepts only exact trimmed 1 2 M or C without an implicit default' {
        foreach ($case in @(@('1','select','1'), @(' 1 ','select','1'), @('2','select','2'),
                           @("`t2`r`n",'select','2'), @('M','select','manual'), @(' m ','select','manual'),
                           @('C','cancel',$null), @(' c ','cancel',$null))) {
            $decision=Get-WinBookSplitInitialDecision -Choice $case[0]
            $decision.decision | Should-Be -Expected $case[1]
            $decision.mode | Should-Be -Expected $case[2]
        }
        foreach ($choice in @($null, '', ' ', 'maybe', 'manual', 'arbitrary', 'nope', 'Y', 'yes', 'N', 'cancel',
                               'CANCEL', '3', '01', '11', '1,2', '2extra', 'm1', 'mmmm',
                               ([string][char]0xFF11), ([string][char]0xFF12), '$(unused)')) {
            $decision=Get-WinBookSplitInitialDecision -Choice $choice
            $decision.decision | Should-Be -Expected 'invalid'
            ($null -eq $decision.mode) | Should-BeTrue
        }
    }
    It 'reprompts after arbitrary words rather than selecting manual or Level 1' {
        Set-AuthoredLauncherAnswers @('maybe', 'arbitrary', ' 2 ')
        Mock Read-WinBookSplitSessionLine { Get-AuthoredLauncherAnswer }
        Mock Write-Host { }
        $decision=Read-WinBookSplitInitialDecision
        $decision.decision | Should-Be -Expected 'select'
        $decision.mode | Should-Be -Expected '2'
        $script:launcherAnswers.Count | Should-Be -Expected 0
        Should-Invoke -CommandName Read-WinBookSplitSessionLine -Times 3 -Exactly
    }
    It 'reprompts for blank answers and compatibility fallback letters in the initial menu' {
        Set-AuthoredLauncherAnswers @('', ' ', 'Y', 'N', ' m ')
        Mock Read-WinBookSplitSessionLine { Get-AuthoredLauncherAnswer }
        Mock Write-Host { }
        $decision=Read-WinBookSplitInitialDecision
        $decision.decision | Should-Be -Expected 'select'
        $decision.mode | Should-Be -Expected 'manual'
        Should-Invoke -CommandName Read-WinBookSplitSessionLine -Times 5 -Exactly
    }
    It 'returns explicit cancellation after invalid answers without selecting a mode' {
        Set-AuthoredLauncherAnswers @('cancel', ' C ')
        Mock Read-WinBookSplitSessionLine { Get-AuthoredLauncherAnswer }
        Mock Write-Host { }
        $decision=Read-WinBookSplitInitialDecision
        $decision.decision | Should-Be -Expected 'cancel'
        ($null -eq $decision.mode) | Should-BeTrue
        Get-SplitExitCode 'cancelled' | Should-Be -Expected 130
        Should-Invoke -CommandName Read-WinBookSplitSessionLine -Times 2 -Exactly
    }
    It 'maps unavailable interactive input to cancellation rather than an implicit mode' {
        Mock Read-WinBookSplitSessionLine { $null }
        Mock Write-Host { }
        $failure=$null
        try { $null=Read-WinBookSplitInitialDecision } catch { $failure=$_.Exception }
        ($null -ne $failure) | Should-BeTrue
        $failure.Data['Code'] | Should-Be -Expected 'cancelled'
        Get-SplitExitCode $failure.Data['Code'] | Should-Be -Expected 130
        Should-Invoke -CommandName Read-WinBookSplitSessionLine -Times 1 -Exactly
    }
}

Describe 'M3 literal no-input prompt and cancellation' -Tag 'AC-058', 'AC-059' {
    It 'preserves supplied literal input under every noninteractive and preview flag combination without prompting' {
        $literal=Join-Path $WorkRoot "Synthetic [%!] O'Brien & ($script:launcherUnicode) `$(unused).pdf"
        Mock Read-WinBookSplitSessionLine { throw 'A supplied literal input must not prompt.' }
        foreach ($flags in @(@($false,$false), @($true,$false), @($false,$true), @($true,$true))) {
            Resolve-WinBookSplitInputChoice -InputFile $literal -NonInteractive $flags[0] -Preview $flags[1] |
                Should-Be -Expected $literal
        }
        Should-Invoke -CommandName Read-WinBookSplitSessionLine -Times 0 -Exactly
    }
    It 'takes one literal path answer and never evaluates shell syntax or expands environment text' {
        $literal=Join-Path $WorkRoot ('Synthetic [%TEMP%!] & ('+$script:launcherUnicode+') $(throw unused).epub')
        Set-AuthoredLauncherAnswers @($literal)
        Mock Read-WinBookSplitSessionLine { Get-AuthoredLauncherAnswer }
        Mock Write-Host { }
        Resolve-WinBookSplitInputChoice -InputFile '' -NonInteractive $false -Preview $false | Should-Be -Expected $literal
        Should-Invoke -CommandName Read-WinBookSplitSessionLine -Times 1 -Exactly
    }
    It 'accepts paired outer double quotes from a pasted Explorer path without expanding the path' {
        $literal=Join-Path $WorkRoot ('Synthetic [%TEMP%!] & ('+$script:launcherUnicode+') $(unused).azw3')
        Set-AuthoredLauncherAnswers @('"'+$literal+'"')
        Mock Read-WinBookSplitSessionLine { Get-AuthoredLauncherAnswer }
        Mock Write-Host { }
        Resolve-WinBookSplitInputChoice -InputFile '' -NonInteractive $false -Preview $false | Should-Be -Expected $literal
        Should-Invoke -CommandName Read-WinBookSplitSessionLine -Times 1 -Exactly
    }
    It 'returns cancellation code 130 for blank whitespace C or unavailable path answers' {
        foreach ($answer in @('', ' ', 'C', ' c ', $null)) {
            Set-AuthoredLauncherAnswers @($answer)
            Mock Read-WinBookSplitSessionLine { Get-AuthoredLauncherAnswer }
            Mock Write-Host { }
            $failure=$null
            try { $null=Resolve-WinBookSplitInputChoice -InputFile '' -NonInteractive $false -Preview $false }
            catch { $failure=$_.Exception }
            ($null -ne $failure) | Should-BeTrue
            $failure.Data['Code'] | Should-Be -Expected 'cancelled'
            Get-SplitExitCode $failure.Data['Code'] | Should-Be -Expected 130
        }
        Should-Invoke -CommandName Read-WinBookSplitSessionLine -Times 5 -Exactly
    }
    It 'rejects a missing source in NonInteractive or Preview before Read-WinBookSplitSessionLine' {
        Mock Read-WinBookSplitSessionLine { throw 'Scripted requests must never prompt.' }
        foreach ($flags in @(@($true,$false), @($false,$true), @($true,$true))) {
            foreach ($missing in @('', ' ', $null)) {
                $failure=$null
                try { $null=Resolve-WinBookSplitInputChoice -InputFile $missing -NonInteractive $flags[0] -Preview $flags[1] }
                catch { $failure=$_.Exception }
                ($null -ne $failure) | Should-BeTrue
                $failure.Data['Code'] | Should-Be -Expected 'invalid_arguments'
                Get-SplitExitCode $failure.Data['Code'] | Should-Be -Expected 2
            }
        }
        Should-Invoke -CommandName Read-WinBookSplitSessionLine -Times 0 -Exactly
    }
}

Describe 'M3 unchanged exact bookmark fallback compatibility' -Tag 'AC-059' {
    It 'retains exact trimmed M and Y manual retries but rejects arbitrary words' {
        $result=[pscustomobject]@{message='Authored flat PDF fallback';code='no_bookmarks';fallback_modes=@('manual')}
        foreach ($choice in @('M',' m ','Y',' y ')) {
            $decision=Get-SplitDecision -Result $result -Choice $choice
            $decision.decision | Should-Be -Expected 'retry'
            $decision.retry_mode | Should-Be -Expected 'manual'
        }
        foreach ($choice in @('manual','maybe','yes','arbitrary','1','2','M1','Yup')) {
            $decision=Get-SplitDecision -Result $result -Choice $choice
            $decision.decision | Should-Be -Expected 'invalid'
            ($null -eq $decision.retry_mode) | Should-BeTrue
        }
    }
    It 'allows exact Level 1 only when offered and preserves explicit N or C cancellation' {
        $result=[pscustomobject]@{message='Authored Level 2 fallback';code='no_bookmarks_at_level';fallback_modes=@('1','manual')}
        foreach ($choice in @('1',' 1 ')) {
            $decision=Get-SplitDecision -Result $result -Choice $choice
            $decision.decision | Should-Be -Expected 'retry'
            $decision.retry_mode | Should-Be -Expected '1'
        }
        foreach ($choice in @('N',' n ','C',' c ')) {
            $decision=Get-SplitDecision -Result $result -Choice $choice
            $decision.decision | Should-Be -Expected 'cancel'
            ($null -eq $decision.retry_mode) | Should-BeTrue
        }
        foreach ($choice in @('no','cancel','1maybe','CANCEL')) {
            (Get-SplitDecision -Result $result -Choice $choice).decision | Should-Be -Expected 'invalid'
        }
    }
    It 'keeps blank fallback answers pending and never offers retry for failures without fallback modes' {
        $result=[pscustomobject]@{message='Authored fallback';code='no_bookmarks';fallback_modes=@('manual')}
        foreach ($choice in @('', ' ')) { (Get-SplitDecision -Result $result -Choice $choice).decision | Should-Be -Expected 'pending' }
        $result=[pscustomobject]@{message='Authored input failure';code='input_invalid';fallback_modes=@()}
        foreach ($choice in @('', 'M', 'Y', '1', 'C', 'arbitrary')) {
            $decision=Get-SplitDecision -Result $result -Choice $choice
            $decision.decision | Should-Be -Expected 'cancel'
            ($null -eq $decision.retry_mode) | Should-BeTrue
        }
    }
}
