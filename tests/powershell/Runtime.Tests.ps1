param([string]$RepositoryRoot, [string]$WorkRoot, [string]$TestPython)

BeforeAll {
    . (Join-Path $RepositoryRoot 'engine\WinBookSplit.Diagnostics.ps1')
    . (Join-Path $RepositoryRoot 'engine\WinBookSplit.Runtime.ps1')
    if (-not [IO.File]::Exists($TestPython)) { throw 'Runtime tests require the explicit pinned TestPython.' }
    $script:pinned = Resolve-WinBookSplitRuntime -ApplicationRoot $RepositoryRoot -PythonPath $TestPython
    function New-CompleteProbe {
        param([string]$Stdout)
        return [pscustomobject]@{ Stdout=$Stdout; Stderr=''; ExitCode=0; TimedOut=$false;
            ParentStopped=$true; DescendantsStopped=$true; StreamsComplete=$true; JobAssigned=$true; Cancelled=$false;
            StartError=$null; StopError=$null; StreamError=$null; StdoutTruncated=$false; StderrTruncated=$false }
    }
}

Describe 'M2 dependency selection and isolated import' -Tag 'AC-043' {
    It 'validates the exact explicit interpreter, isolated flags and package origin' {
        $script:pinned.Path | Should-Be -Expected $TestPython
        $script:pinned.Version | Should-Be -Expected '3.14.8'
        $script:pinned.Source | Should-Be -Expected 'explicit'
        $script:pinned.PypdfVersion | Should-Be -Expected '6.19.0'
        $script:pinned.PypdfPath | Should-MatchString -Expected '\\Lib\\site-packages\\pypdf\\__init__\.py$'
        ($script:pinned.Arguments -join ',') | Should-Be -Expected '-I,-B'
        $script:pinned.Details.isolated | Should-BeTrue
        $script:pinned.Details.dont_write_bytecode | Should-BeTrue
        $script:pinned.Probe.StreamsComplete | Should-BeTrue
    }
    It 'does not import a book or CWD shadow despite injected PYTHONPATH and PYTHONHOME' {
        $book = Join-Path $TestDrive 'owned book'
        $null = New-Item -ItemType Directory -Path $book
        $marker = Join-Path $book 'shadow-imported.txt'
        $pythonMarker = $marker.Replace('\','\\').Replace("'","\'")
        [IO.File]::WriteAllText((Join-Path $book 'pypdf.py'), "from pathlib import Path; Path('$pythonMarker').write_text('Owned controlled shadow imported'); __version__='6.19.0'")
        $savedPath=$env:PYTHONPATH; $savedHome=$env:PYTHONHOME; $savedLocation=Get-Location
        try {
            $env:PYTHONPATH=$book; $env:PYTHONHOME=$book; Set-Location -LiteralPath $book
            $result=Resolve-WinBookSplitRuntime -ApplicationRoot $RepositoryRoot -PythonPath $TestPython -DocumentPath (Join-Path $book 'owned.pdf')
            $result.PypdfPath | Should-Be -Expected $script:pinned.PypdfPath
            Test-Path -LiteralPath $marker | Should-BeFalse
            Test-Path -LiteralPath (Join-Path $book '__pycache__') | Should-BeFalse
        } finally { $env:PYTHONPATH=$savedPath; $env:PYTHONHOME=$savedHome; Set-Location -LiteralPath $savedLocation.ProviderPath }
    }
    It 'fails a missing or relative explicit path without PATH substitution' {
        Mock Get-WinBookSplitPathCandidates { throw 'Automatic fallback must not occur' }
        foreach ($path in @('python.exe', (Join-Path $TestDrive 'missing-python.exe'))) {
            { Resolve-WinBookSplitRuntime -ApplicationRoot $RepositoryRoot -PythonPath $path } | Should-Throw
        }
        Should-Invoke -CommandName Get-WinBookSplitPathCandidates -Times 0 -Exactly
    }
    It 'fails a non-Python ordinary executable before any usable PATH substitute' {
        $wrong=Join-Path $TestDrive 'not-python.exe'; [IO.File]::WriteAllText($wrong,'Owned non-executable fixture')
        try { Resolve-WinBookSplitRuntime -ApplicationRoot $RepositoryRoot -PythonPath $wrong; throw 'Expected preflight rejection' }
        catch {
            $_.Exception.Data['Code'] | Should-Be -Expected 'runtime_invalid'
            @($_.Exception.Data['Attempts']).Count | Should-Be -Expected 1
            $_.Exception.Message | Should-MatchString -Expected ([regex]::Escape($wrong))
        }
    }
    It 'hard-fails an existing application venv with no Scripts interpreter' {
        $app=Join-Path $TestDrive 'app'; $null=New-Item -ItemType Directory -Path (Join-Path $app '.venv')
        Mock Get-WinBookSplitPathCandidates { throw 'Automatic fallback must not occur' }
        try { Resolve-WinBookSplitRuntime -ApplicationRoot $app; throw 'Expected venv rejection' }
        catch {
            $_.Exception.Data['Code'] | Should-Be -Expected 'runtime_invalid'
            $_.Exception.Data['Attempts'][0].Source | Should-Be -Expected 'application_venv'
            $_.Exception.Message | Should-MatchString -Expected '\.venv'
        }
        Should-Invoke -CommandName Get-WinBookSplitPathCandidates -Times 0 -Exactly
    }
    It 'rejects an incompatible PATH candidate then selects the actual pinned interpreter' {
        $bad=Join-Path $TestDrive 'bad'; $null=New-Item -ItemType Directory -Path $bad
        [IO.File]::WriteAllText((Join-Path $bad 'python.exe'),'Owned non-executable PATH candidate')
        $savedPath=$env:PATH; $savedLocation=Get-Location
        try {
            Set-Location -LiteralPath $RepositoryRoot
            $env:PATH=$bad+';'+[IO.Path]::GetDirectoryName($TestPython)
            $result=Resolve-WinBookSplitRuntime -ApplicationRoot $RepositoryRoot -DocumentPath (Join-Path $TestDrive 'book\owned.pdf')
            $result.Path | Should-Be -Expected $TestPython
            $result.Source | Should-Be -Expected 'PATH'
            @($result.Attempts | Where-Object { -not $_.Accepted }).Count | Should-Be -Expected 1
            $result.Attempts[0].Probe.ParentStopped | Should-BeTrue
        } finally { $env:PATH=$savedPath; Set-Location -LiteralPath $savedLocation.ProviderPath }
    }
    It 'excludes empty, relative, document and unrelated-CWD automatic candidates' {
        $book=Join-Path $TestDrive 'book'; $cwd=Join-Path $TestDrive 'cwd'
        $null=New-Item -ItemType Directory -Path $book; $null=New-Item -ItemType Directory -Path $cwd
        foreach ($directory in @($book,$cwd)) { [IO.File]::WriteAllText((Join-Path $directory 'python.exe'),'Owned excluded decoy') }
        $savedPath=$env:PATH; $savedLocation=Get-Location
        try {
            $env:PATH=';.;'+$book+';'+$cwd+';'+[IO.Path]::GetDirectoryName($TestPython)
            Set-Location -LiteralPath $cwd
            $attempts=New-Object 'System.Collections.Generic.List[object]'
            $candidates=@(Get-WinBookSplitPathCandidates 'python.exe' $RepositoryRoot (Join-Path $book 'owned.pdf') $attempts)
            $candidates.Count | Should-Be -Expected 1
            $candidates[0] | Should-Be -Expected $TestPython
            $attempts.Count | Should-Be -Expected 2
        } finally { $env:PATH=$savedPath; Set-Location -LiteralPath $savedLocation.ProviderPath }
    }
    It 'rejects an actual hard-linked executable alias before automatic probing' {
        $source=Join-Path $TestDrive 'owned-source.exe'; $alias=Join-Path $TestDrive 'owned-alias.exe'
        [IO.File]::WriteAllText($source,'Owned controlled hard-link decoy; never executable')
        $null=New-Item -ItemType HardLink -Path $alias -Target $source
        ([WinBookSplit.Preflight.Probe]::Canonical($alias)).Links | Should-Be -Expected 2
        { Get-WinBookSplitTrustedPath -Path $alias -Automatic } | Should-Throw
        Get-WinBookSplitTrustedPath -Path $alias | Should-Be -Expected $alias
    }
    It 'rejects array or boolean substitutions in the one-record runtime protocol' {
        foreach ($mutation in @('protocol','schema_version','version','bits','root','pypdf_path')) {
            $record=$script:pinned.Details | ConvertTo-Json -Depth 6 | ConvertFrom-Json
            switch ($mutation) {
                'schema_version' { $record.schema_version=$true }
                'bits' { $record.bits=@(64) }
                'root' { $record=@($record) }
                default { $record.$mutation=@($record.$mutation) }
            }
            $script:malformedJson=ConvertTo-Json -InputObject $record -Depth 6 -Compress
            Mock Invoke-WinBookSplitRuntimeProbe { New-CompleteProbe $script:malformedJson }
            { Resolve-WinBookSplitRuntime -ApplicationRoot $RepositoryRoot -PythonPath $TestPython } | Should-Throw
        }
    }
}

Describe 'M2 narrow bounded native preflight' -Tag 'AC-043' {
    It 'marshals literal arguments without shell evaluation' {
        $unicode=([char]0x65E5).ToString()+[char]0x672C
        $values=@('space value', '[%!] & (literal)', "O'Brien", 'a"b', 'trailing\', $unicode)
        $probe=Invoke-WinBookSplitRuntimeProbe -Path $TestPython -ApplicationRoot $RepositoryRoot -Arguments (@('-I','-B','-c','import json,sys; print(json.dumps(sys.argv[1:]))')+$values)
        $probe.ExitCode | Should-Be -Expected 0
        ($probe.Stdout | ConvertFrom-Json) -join '|' | Should-Be -Expected ($values -join '|')
    }
    It 'drains both large streams and retains bounded final tails without newlines' {
        $code=@'
import os, threading
def emit(fd, char):
    for _ in range(48): os.write(fd, char*8192)
    os.write(fd,b'FINAL_'+char)
threads=[threading.Thread(target=emit,args=(1,b'O')),threading.Thread(target=emit,args=(2,b'E'))]
for thread in threads: thread.start()
for thread in threads: thread.join()
'@
        $probe=Invoke-WinBookSplitRuntimeProbe -Path $TestPython -ApplicationRoot $RepositoryRoot -Arguments @('-I','-B','-c',$code)
        $probe.ExitCode | Should-Be -Expected 0
        $probe.StreamsComplete | Should-BeTrue
        $probe.StdoutTruncated | Should-BeTrue
        $probe.StderrTruncated | Should-BeTrue
        $probe.Stdout.EndsWith('FINAL_O') | Should-BeTrue
        $probe.Stderr.EndsWith('FINAL_E') | Should-BeTrue
        $probe.Stdout.Length | Should-Be -Expected 65536
        $probe.Stderr.Length | Should-Be -Expected 65536
    }
    It 'bounds owned-tree timeout and records stopped descendants' {
        $clock=[Diagnostics.Stopwatch]::StartNew()
        $probe=Invoke-WinBookSplitRuntimeProbe -Path $TestPython -ApplicationRoot $RepositoryRoot -Arguments @('-I','-B','-c','import time; time.sleep(30)') -TimeoutSeconds 0.2
        $probe.TimedOut | Should-BeTrue
        $probe.ParentStopped | Should-BeTrue
        $probe.DescendantsStopped | Should-BeTrue
        $probe.JobAssigned | Should-BeTrue
        $probe.StreamsComplete | Should-BeTrue
        ($clock.Elapsed.TotalSeconds -lt 2) | Should-BeTrue
    }
    It 'bounds inherited pipe EOF after the original parent already exited' {
        $code='import subprocess,sys; p=subprocess.Popen([sys.executable,"-I","-B","-c","import time; time.sleep(1.2)"],stdout=sys.stdout,stderr=sys.stderr); print(p.pid,flush=True)'
        $clock=[Diagnostics.Stopwatch]::StartNew()
        $probe=Invoke-WinBookSplitRuntimeProbe -Path $TestPython -ApplicationRoot $RepositoryRoot -Arguments @('-I','-B','-c',$code) -TimeoutSeconds 0.3
        $probe.TimedOut | Should-BeTrue
        $probe.ParentStopped | Should-BeTrue
        $probe.ExitCode | Should-Be -Expected 0
        $probe.StreamsComplete | Should-BeTrue
        $probe.DescendantsStopped | Should-BeTrue
        $probe.JobAssigned | Should-BeTrue
        $childPid=[int]$probe.Stdout.Trim()
        ($null -eq (Get-Process -Id $childPid -ErrorAction SilentlyContinue)) | Should-BeTrue
        ($clock.Elapsed.TotalSeconds -lt 2) | Should-BeTrue
        # The original parent's zero exit does not skip the owned job/EOF deadline.
    }
}

Describe 'M2 converter version preflight' -Tag 'AC-045' {
    It 'rejects a zero-exit non-Calibre executable without PATH fallback' {
        Mock Get-WinBookSplitPathCandidates { throw 'Automatic fallback must not occur' }
        try { Resolve-WinBookSplitConverter -ApplicationRoot $RepositoryRoot -CalibrePath $TestPython; throw 'Expected converter rejection' }
        catch {
            $_.Exception.Data['Code'] | Should-Be -Expected 'converter_invalid'
            $_.Exception.Data['Attempts'][0].Probe.ExitCode | Should-Be -Expected 0
        }
        Should-Invoke -CommandName Get-WinBookSplitPathCandidates -Times 0 -Exactly
    }
    It 'requires the exact converter version even when the version command succeeds' {
        Mock Invoke-WinBookSplitRuntimeProbe { New-CompleteProbe 'ebook-convert.exe (calibre 9.14.0)' }
        { Resolve-WinBookSplitConverter -ApplicationRoot $RepositoryRoot -CalibrePath $TestPython } | Should-Throw
    }
    It 'accepts a supported complete version record and retains its actual probe' {
        Mock Invoke-WinBookSplitRuntimeProbe { New-CompleteProbe 'ebook-convert.exe (calibre 9.15.0)' }
        $result=Resolve-WinBookSplitConverter -ApplicationRoot $RepositoryRoot -CalibrePath $TestPython
        $result.Version | Should-Be -Expected '9.15.0'
        $result.Source | Should-Be -Expected 'explicit'
        $result.Probe.ExitCode | Should-Be -Expected 0
        $result.Attempts[0].Accepted | Should-BeTrue
    }
    It 'rejects extra output after an otherwise supported converter version' {
        Mock Invoke-WinBookSplitRuntimeProbe { New-CompleteProbe "ebook-convert.exe (calibre 9.15.0)`nextra record" }
        { Resolve-WinBookSplitConverter -ApplicationRoot $RepositoryRoot -CalibrePath $TestPython } | Should-Throw
    }
    It 'accepts the exact two-line banner observed from the real pinned converter' {
        Mock Invoke-WinBookSplitRuntimeProbe { New-CompleteProbe "ebook-convert.exe (calibre 9.15.0)`r`nCreated by: Kovid Goyal <kovid@kovidgoyal.net>`r`n" }
        $result=Resolve-WinBookSplitConverter -ApplicationRoot $RepositoryRoot -CalibrePath $TestPython
        $result.Version | Should-Be -Expected '9.15.0'
        $result.Probe.Stdout | Should-Be -Expected "ebook-convert.exe (calibre 9.15.0)`r`nCreated by: Kovid Goyal <kovid@kovidgoyal.net>`r`n"
    }
    It 'rejects unrelated text or an altered creator following the supported version' {
        foreach ($tail in @('Created by: unrelated control', 'Created by: Kovid Goyal <wrong@example.invalid>',
                            "Created by: Kovid Goyal <kovid@kovidgoyal.net>`r`nextra record")) {
            $script:extraBanner="ebook-convert.exe (calibre 9.15.0)`r`n"+$tail+"`r`n"
            Mock Invoke-WinBookSplitRuntimeProbe { New-CompleteProbe $script:extraBanner }
            { Resolve-WinBookSplitConverter -ApplicationRoot $RepositoryRoot -CalibrePath $TestPython } | Should-Throw
        }
    }
}
