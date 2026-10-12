<#
.SYNOPSIS
Explicitly prepare or check the application-local WinBookSplit Python environment.
.DESCRIPTION
Selects a trusted policy-supported base Python without requiring pypdf first.
Accepts regular GIL Windows x64 CPython 3.15.0 and later stable 3.15 patches,
3.14.8+ within 3.14, or 3.13.16+ within 3.13; unknown minors remain unvalidated.
Setup creates a fresh .venv and installs only the separately hash-locked runtime
requirement. Processing never runs installation. No activation, elevation,
global packages/PATH change or global execution-policy change is required.
.PARAMETER PythonPath
Trusted absolute base Python executable. Discovery is read-only when omitted.
.PARAMETER Check
Read-only readiness check of the authoritative environment or trusted processing
runtime; never installs or repairs.
.PARAMETER NonInteractive
Explicitly authorize unattended setup with no prompts; missing prerequisites fail.
.PARAMETER Wheelhouse
Existing ordinary local directory containing the exact reviewed pypdf wheel.
Installation is offline and hash-locked; there is no online fallback.
.PARAMETER Repair
Explicitly preserve a broken .venv in a unique backup before creating a fresh one.
Healthy environments are unchanged. The backup is not overwritten or deleted.
.EXAMPLE
.\Setup-WinBookSplit.bat
.EXAMPLE
.\Setup-WinBookSplit.bat -Check -NonInteractive
.EXAMPLE
& ([System.IO.Path]::Combine([Environment]::SystemDirectory, 'WindowsPowerShell\v1.0\powershell.exe')) -NoProfile -ExecutionPolicy Bypass -File '.\Setup-WinBookSplit.ps1' -PythonPath 'C:\Tools\Python\python.exe' -NonInteractive -Wheelhouse 'C:\Tools\Wheels'
.NOTES
Unsigned; follow enforced organizational script policy. Full setup and runtime
compatibility requirements are in docs/SETUP.md. Output includes a structured
[SETUP] result and a meaningful native exit status.
#>

[CmdletBinding()]
param([string]$PythonPath, [switch]$Check, [switch]$NonInteractive, [string]$Wheelhouse, [switch]$Repair)

# Machine outcomes and path diagnostics use UTF-8 through every launcher.
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
[Console]::InputEncoding = New-Object System.Text.UTF8Encoding($false)

Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'

function Get-WinBookSplitSetupHostObservation {
    # Local, read-only observations after script entry. An unavailable query
    # never changes the setup result and never substitutes a guessed policy.
    $observation=[ordered]@{executable=$null;executable_status='unavailable';
        version=$PSVersionTable.PSVersion.ToString();edition=$PSVersionTable.PSEdition;
        execution_policy=[ordered]@{status='unavailable';effective=$null;scopes=@();reason_code='query_failed'}}
    $process=$null
    try {
        $process=[Diagnostics.Process]::GetCurrentProcess()
        $observation.executable=$process.MainModule.FileName
        $observation.executable_status='observed'
    } catch { Write-Verbose 'The actual setup host executable could not be observed.' }
    finally { if($null -ne $process){$process.Dispose()} }
    try {
        # A PS7 parent can pass its module search path to the inbox BAT child.
        # Load this actual host's built-in security module, without changing it.
        Import-Module -Name (Join-Path $PSHOME 'Modules\Microsoft.PowerShell.Security\Microsoft.PowerShell.Security.psd1') -Scope Local -ErrorAction Stop
        $policies=@(Get-ExecutionPolicy -List -ErrorAction Stop)
        $effective=(Get-ExecutionPolicy -ErrorAction Stop).ToString()
        $scopes=@('MachinePolicy','UserPolicy','Process','CurrentUser','LocalMachine')
        $values=@('Undefined','Restricted','AllSigned','RemoteSigned','Unrestricted','Bypass','Default')
        $records=@(foreach($scope in $scopes){
            $matches=@($policies | Where-Object {$_.Scope.ToString() -ceq $scope})
            if($matches.Count -ne 1 -or $matches[0].ExecutionPolicy.ToString() -cnotin $values){throw 'Incomplete execution-policy observation.'}
            [pscustomobject]@{scope=$scope;policy=$matches[0].ExecutionPolicy.ToString()}
        })
        if($policies.Count -ne $scopes.Count -or $effective -cnotin $values){throw 'Invalid execution-policy observation.'}
        $observation.execution_policy=[ordered]@{status='observed';effective=$effective;scopes=$records;reason_code=$null}
    } catch { Write-Verbose 'Execution-policy observation was unavailable; preserving the setup outcome.' }
    return [pscustomobject]$observation
}

try {
    . (Join-Path $PSScriptRoot 'engine\WinBookSplit.Setup.ps1')
    $result=Invoke-WinBookSplitSetup -ApplicationRoot $PSScriptRoot -PythonPath $PythonPath -Check:$Check -NonInteractive:$NonInteractive -Wheelhouse $Wheelhouse -Repair:$Repair
} catch {
    $result=[pscustomobject]@{protocol='winbooksplit.setup';schema_version=1;status='failed';code='setup_failed';exit_code=6;changed=$false;message=$_.Exception.Message}
}
$result | Add-Member -MemberType NoteProperty -Name host -Value (Get-WinBookSplitSetupHostObservation)
Write-Host $result.message
if ($result.PSObject.Properties.Name -contains 'observed_python' -and $null -ne $result.observed_python) {
    $displayVersion='version unavailable';if($null -ne $result.observed_version){$displayVersion=$result.observed_version}
    Write-Host ('Python selection: '+$result.observed_python+' ('+$displayVersion+'; '+$result.observed_source+')')
}
Write-Host ('[SETUP] '+($result|ConvertTo-Json -Depth 12 -Compress))
$setupExitCode=[int]$result.exit_code
# Only a zero-argument console launch waits for its result to be read. Explicit
# options, redirected automation and cancellation never acquire another prompt.
if ($PSBoundParameters.Count -eq 0 -and $setupExitCode -ne 130 -and
    -not [Console]::IsInputRedirected -and -not [Console]::IsOutputRedirected) {
    try { $null=Read-Host 'Press Enter to close setup' }
    catch { Write-Verbose 'The console wait was unavailable; preserving the setup exit status.' }
}
exit $setupExitCode
