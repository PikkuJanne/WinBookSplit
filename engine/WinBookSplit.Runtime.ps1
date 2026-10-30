# Dependency preflight only. Dot-sourcing performs no discovery or installation.
# Probe transport shares the owned engine job; discovery never owns file cleanup.
. (Join-Path $PSScriptRoot 'WinBookSplit.Process.ps1')
. (Join-Path $PSScriptRoot 'WinBookSplit.Compatibility.ps1')
. (Join-Path $PSScriptRoot 'WinBookSplit.Diagnostics.ps1')
function Assert-WinBookSplitProbeNotInterrupted {
    param($Probe)
    if ($null -ne $Probe -and ($Probe.Cancelled -or $Probe.TimedOut)) {
        $code = 'processor_cancelled'
        $message = 'Dependency preflight was cancelled.'
        if ($Probe.TimedOut) { $code = 'processor_timeout'; $message = 'Dependency preflight timed out.' }
        if ($Probe.ParentStopped -and $Probe.DescendantsStopped -and -not $Probe.StopError) {
            $message += ' Its owned process tree was proved stopped.'
        }
        else { $message += ' Shutdown could not be proved complete: ' + $Probe.StopError }
        $failure = New-Object InvalidOperationException($message)
        $failure.Data['Code'] = $code
        $failure.Data['Probe'] = $Probe
        throw $failure
    }
}

function Test-WinBookSplitInterruption {
    param($Failure)
    return (($Failure.Data.Contains('Code') -and $Failure.Data['Code'] -in @('processor_cancelled', 'processor_timeout')) -or
        ($Failure.Data.Contains('FatalDiscovery') -and $Failure.Data['FatalDiscovery']))
}

function Assert-WinBookSplitDiscoveryStopped {
    param($Probe)
    if ($null -ne $Probe -and (-not $Probe.ParentStopped -or -not $Probe.DescendantsStopped -or $Probe.StopError)) {
        $failure = New-Object InvalidOperationException('Dependency discovery shutdown could not be proved complete: ' + $Probe.StopError)
        $failure.Data['Code'] = 'processor_failed'; $failure.Data['FatalDiscovery'] = $true; $failure.Data['Probe'] = $Probe
        throw $failure
    }
}

function Assert-WinBookSplitListingProbe {
    param($Probe)
    Assert-WinBookSplitDiscoveryStopped $Probe
    if ($null -ne $Probe -and $Probe.Cancelled) { Assert-WinBookSplitProbeNotInterrupted $Probe }
    if ($null -eq $Probe -or $Probe.TimedOut -or $Probe.ExitCode -ne 0 -or -not $Probe.JobAssigned -or
        -not $Probe.StreamsComplete -or $Probe.StartError -or $Probe.StreamError -or
        $Probe.StdoutTruncated -or $Probe.StderrTruncated) { throw 'Read-only Python listing did not complete safely; independent discovery remains available.' }
}

function ConvertFrom-WinBookSplitPythonListing {
    param($Probe, [ValidateSet('Legacy', 'Manager')][string]$Kind)
    Assert-WinBookSplitListingProbe $Probe
    $text = $Probe.Stdout
    if ($Kind -eq 'Legacy') { $text += "`n" + $Probe.Stderr }
    elseif (-not [string]::IsNullOrWhiteSpace($Probe.Stderr)) { throw 'Unexpected manager listing diagnostics.' }
    if ($text -isnot [string] -or [Text.Encoding]::UTF8.GetByteCount($text) -gt 32768) { throw 'Python listing exceeds the 32 KiB limit.' }
    $lines = @($text -split '\r?\n')
    if ($lines.Count -gt 128) { throw 'Python listing exceeds the 128-line limit.' }
    $paths = New-Object 'Collections.Generic.List[string]'
    foreach ($line in $lines) {
        if ($line.Length -gt 4096) { throw 'Python listing line exceeds the 4096-character limit.' }
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $literal = $line.Trim()
        if ($Kind -eq 'Legacy') {
            if ($line -cnotmatch '\A\s*-(?:V:[A-Za-z0-9][A-Za-z0-9_.:/\[\]-]{0,79}|[0-9]{1,3}\.[0-9]{1,3}(?:-[0-9]{1,3})?)\s+(?:\*\s*)?(?<literal>.+?)\s*\z') {
                throw 'Malformed legacy installed-runtime listing.'
            }
            $literal = $Matches['literal']
        }
        if ($literal -notmatch '\A(?:[A-Za-z]:[\\/]|\\\\[^\\]+\\[^\\]+\\).+\.exe\z' -or
            $literal -match '[\x00-\x1f"*?<>|]' -or $literal.Substring(2).Contains(':')) {
            throw 'Installed-runtime listings must contain absolute literal executable paths.'
        }
        $paths.Add($literal)
        if ($paths.Count -gt 32) { throw 'Python listing exceeds the 32-runtime limit.' }
    }
    if ($paths.Count -eq 0) {
        $failure = New-Object InvalidOperationException('Installed-runtime listing is empty.')
        $failure.Data['Reason'] = 'launcher_no_runtimes'
        if ($Kind -eq 'Manager') { $failure.Data['Reason'] = 'manager_no_runtimes' }
        throw $failure
    }
    return $paths.ToArray()
}

function ConvertFrom-WinBookSplitManagerRegistration {
    param($Probe, [string]$ApplicationRoot, [string]$DocumentPath, [switch]$WithRegistration)
    Assert-WinBookSplitListingProbe $Probe
    if (-not [string]::IsNullOrWhiteSpace($Probe.Stderr) -or [Text.Encoding]::UTF8.GetByteCount($Probe.Stdout) -gt 32768) { throw 'Invalid or oversized manager registration result.' }
    if (-not ('WinBookSplitSupport.StrictJson' -as [type])) {
        . (Join-Path $PSScriptRoot 'WinBookSplit.Support.ps1')
        Initialize-WinBookSplitSupport
    }
    $record = [WinBookSplitSupport.StrictJson]::Parse($Probe.Stdout)
    Assert-WinBookSplitCompatibilityKeys $record @('schema_version', 'packages') 'manager registration'
    if (-not (Test-WinBookSplitCompatibilityInteger $record.schema_version) -or $record.schema_version -ne 1 -or
        $record.packages -isnot [array] -or $record.packages.Count -gt 16) { throw 'Invalid manager registration schema or package count.' }
    $name = 'PythonSoftwareFoundation.PythonManager'
    # Each distribution is one complete tuple, never independent allowlists.
    $identities = @(
        @{PublisherId='qbz5n2kfra8p0';Publisher='CN=4975D53F-AA7E-49A5-8B49-EA4FDC1BB66B';SignatureKind='Store'},
        @{PublisherId='3847v3x7pw1km';Publisher='CN=Python Software Foundation, O=Python Software Foundation, L=Beaverton, S=Oregon, C=US';SignatureKind='Developer'}
    )
    $seen = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $paths = New-Object 'Collections.Generic.List[object]'
    foreach ($package in $record.packages) {
        Assert-WinBookSplitCompatibilityKeys $package @('Name','Publisher','PublisherId','PackageFamilyName','PackageFullName','InstallLocation','Architecture','SignatureKind','Status','Version','IsDevelopmentMode','IsPartiallyStaged','IsFramework','IsResourcePackage') 'registered manager package'
        foreach ($field in @('Name','Publisher','PublisherId','PackageFamilyName','PackageFullName','InstallLocation','Architecture','SignatureKind','Status','Version')) {
            if (-not (Test-WinBookSplitCompatibilityText (Get-WinBookSplitCompatibilityField $package $field) 4096)) { throw 'Registered manager identity requires scalar strings.' }
        }
        foreach ($field in @('IsDevelopmentMode','IsPartiallyStaged','IsFramework','IsResourcePackage')) {
            if ((Get-WinBookSplitCompatibilityField $package $field) -isnot [bool] -or (Get-WinBookSplitCompatibilityField $package $field)) { throw 'Registered manager package is not a complete ordinary application.' }
        }
        $identity = @($identities | Where-Object { $_.PublisherId -ceq $package.PublisherId -and
            $_.Publisher -ceq $package.Publisher -and $_.SignatureKind -ceq $package.SignatureKind })
        $version = $null
        if ($package.Version -cnotmatch '\A(?:0|[1-9][0-9]{0,4})(?:\.(?:0|[1-9][0-9]{0,4})){3}\z' -or
            -not [Version]::TryParse($package.Version, [ref]$version) -or
            $version.Major -gt 65535 -or $version.Minor -gt 65535 -or $version.Build -gt 65535 -or $version.Revision -gt 65535) {
            throw 'Invalid Python manager package version.'
        }
        if ($identity.Count -ne 1 -or $package.Name -cne $name -or
            $package.PackageFamilyName -cne ($name + '_' + $package.PublisherId) -or $package.Architecture -cne 'X64' -or $package.Status -cne 'Ok' -or
            $package.PackageFullName -cne ($name + '_' + $package.Version + '_x64__' + $package.PublisherId) -or
            -not $seen.Add($package.PackageFullName)) { throw 'Unrecognized or duplicate Python manager registration identity.' }
        $expected = [IO.Path]::Combine([Environment]::GetFolderPath([Environment+SpecialFolder]::ProgramFiles), 'WindowsApps', $package.PackageFullName)
        $location = Get-WinBookSplitTrustedPath -Path $package.InstallLocation -Directory
        if (-not $location.Equals($expected, [StringComparison]::OrdinalIgnoreCase)) { throw 'Registered Python manager location does not match its physical package identity.' }
        try { $path = Get-WinBookSplitManagerEntryPoint -Package $package -ApplicationRoot $ApplicationRoot -DocumentPath $DocumentPath }
        catch {
            $_.Exception.Data['Reason'] = 'manager_entry_rejected'; $_.Exception.Data['Stage'] = 'entry'
            throw
        }
        if ($WithRegistration) { $paths.Add([pscustomobject]@{Path=$path;Package=$package}) }
        else { $paths.Add($path) }
    }
    return $paths.ToArray()
}

function Get-WinBookSplitRegisteredManagerPaths {
    param([string]$ApplicationRoot, [string]$DocumentPath, $Attempts, [double]$TimeoutSeconds = 5, [switch]$WithRegistration)
    $probe = $null; $path = $null; $queryCompleted = $false
    try {
        # This exact OS-known inbox executable may have legitimate servicing
        # hard links. It is not a PATH-discovered dependency or alias.
        $path = Get-WinBookSplitTrustedPath -Path ([IO.Path]::Combine([Environment]::GetFolderPath([Environment+SpecialFolder]::System), 'WindowsPowerShell', 'v1.0', 'powershell.exe'))
        # Static installed-only current-user query in inbox 5.1 avoids PS7 compatibility
        # remoting and inherited module lookup. No discovered value enters child code.
        $command = @'
$ErrorActionPreference='Stop'
$PSModuleAutoloadingPreference='None'
Import-Module -Name ([IO.Path]::Combine($PSHOME,'Modules','Microsoft.PowerShell.Utility','Microsoft.PowerShell.Utility.psd1')) -ErrorAction Stop
Import-Module -Name ([IO.Path]::Combine($PSHOME,'Modules','Appx','Appx.psd1')) -ErrorAction Stop
$packages=@(Get-AppxPackage -Name 'PythonSoftwareFoundation.PythonManager' -ErrorAction Stop)
if ($packages.Count -gt 16) { throw 'Registered manager count exceeds 16.' }
$records=@(foreach ($package in $packages) {
    [ordered]@{Name=[string]$package.Name;Publisher=[string]$package.Publisher;PublisherId=[string]$package.PublisherId;
        PackageFamilyName=[string]$package.PackageFamilyName;PackageFullName=[string]$package.PackageFullName;
        InstallLocation=[string]$package.InstallLocation;Architecture=[string]$package.Architecture;
        SignatureKind=[string]$package.SignatureKind;Status=[string]$package.Status;Version=[string]$package.Version;
        IsDevelopmentMode=[bool]$package.IsDevelopmentMode;IsPartiallyStaged=[bool]$package.IsPartiallyStaged;
        IsFramework=[bool]$package.IsFramework;IsResourcePackage=[bool]$package.IsResourcePackage}
})
[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)
[Console]::WriteLine((ConvertTo-Json -InputObject ([ordered]@{schema_version=1;packages=$records}) -Depth 4 -Compress))
'@
        $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($command))
        $probe = Invoke-WinBookSplitRuntimeProbe -Path $path -Arguments @('-NoProfile','-NonInteractive','-EncodedCommand',$encoded) -ApplicationRoot $ApplicationRoot -TimeoutSeconds $TimeoutSeconds
        Assert-WinBookSplitListingProbe $probe
        $queryCompleted = $true
        $paths = @(ConvertFrom-WinBookSplitManagerRegistration -Probe $probe -ApplicationRoot $ApplicationRoot -DocumentPath $DocumentPath -WithRegistration:$WithRegistration)
        $reason = 'manager_not_registered'; if ($paths.Count -gt 0) { $reason = 'manager_registered' }
        $Attempts.Add([pscustomobject]@{Path=$path;Source='manager_registration';Stage='registration';Accepted=($paths.Count -gt 0);
            QueryCompleted=$queryCompleted;RegisteredCount=$paths.Count;Reason=$reason;Message='Bounded current-user registered Python manager query; no runtime launch or install.';Probe=$probe})
        return $paths
    } catch {
        if (Test-WinBookSplitInterruption $_.Exception) { throw }
        $reason = 'manager_registration_failed'; if ($queryCompleted) { $reason = 'manager_registration_rejected' }
        $stage = 'registration'
        if ($_.Exception.Data.Contains('Reason')) { $reason = $_.Exception.Data['Reason'] }
        if ($_.Exception.Data.Contains('Stage')) { $stage = $_.Exception.Data['Stage'] }
        $Attempts.Add([pscustomobject]@{Path=$path;Source='manager_registration';Stage=$stage;Accepted=$false;
            QueryCompleted=$queryCompleted;RegisteredCount=0;Reason=$reason;Message=$_.Exception.Message;Probe=$probe})
        return @()
    }
}

function Get-WinBookSplitManagerEntryPoint {
    param($Package, [string]$ApplicationRoot, [string]$DocumentPath)
    # Only the documented family-specific alias is eligible. Its actual package
    # and physical image are proved on the suspended child before any code runs.
    # Ordinary dependency guards remain unchanged, including at the alias parent.
    $directory = [IO.Path]::Combine([Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData),
        'Microsoft', 'WindowsApps', $Package.PackageFamilyName)
    $directory = Get-WinBookSplitTrustedPath -Path $directory -Directory
    $path = [IO.Path]::Combine($directory, 'pymanager.exe')
    foreach ($candidate in @($Package.InstallLocation, $path)) {
        if (-not (Test-WinBookSplitAutomaticPath $candidate $ApplicationRoot $DocumentPath)) {
            throw 'Registered manager is in an excluded document or current directory.'
        }
    }
    Initialize-WinBookSplitRuntimeProbe
    if (-not [WinBookSplit.Preflight.Probe]::IsAppExecutionAlias($path)) {
        throw 'Registered manager entry point must be an AppExecutionAlias.'
    }
    return $path
}

function Invoke-WinBookSplitManagerListing {
    param($Registration, [string]$ApplicationRoot, [string]$DocumentPath,
        [ValidateRange(0.05, 60)][double]$TimeoutSeconds = 2)
    # Revalidate the complete registration and alias immediately before launch.
    # The automatic resolver passes the complete descriptor, never an arbitrary PATH manager.
    $record = [ordered]@{schema_version=1;packages=@($Registration.Package)}
    $validation = [pscustomobject]@{Stdout=($record | ConvertTo-Json -Depth 6 -Compress);Stderr='';ExitCode=0;
        ParentStopped=$true;DescendantsStopped=$true;JobAssigned=$true;StreamsComplete=$true;
        Cancelled=$false;TimedOut=$false;StartError=$null;StopError=$null;StreamError=$null;
        StdoutTruncated=$false;StderrTruncated=$false}
    try { $paths = @(ConvertFrom-WinBookSplitManagerRegistration $validation $ApplicationRoot $DocumentPath) }
    catch {
        if (-not $_.Exception.Data.Contains('Reason')) { $_.Exception.Data['Reason'] = 'manager_registration_rejected'; $_.Exception.Data['Stage'] = 'registration' }
        throw
    }
    if ($paths.Count -ne 1 -or $Registration.Path -cne $paths[0]) {
        $failure = New-Object InvalidOperationException('Manager entry point does not match its registration.')
        $failure.Data['Reason'] = 'manager_entry_rejected'; $failure.Data['Stage'] = 'entry'; throw $failure
    }
    return Invoke-WinBookSplitProcess -Path $paths[0] -Arguments @('list','--format=exe') -WorkingDirectory $ApplicationRoot -TimeoutSeconds $TimeoutSeconds `
        -ExpectedPackageFullName $Registration.Package.PackageFullName -ExpectedImagePath ([IO.Path]::Combine($Registration.Package.InstallLocation,'pymanager.exe'))
}

function Add-WinBookSplitPythonDiscoveryCandidate {
    param([string]$Path, [string]$Source, [string]$ApplicationRoot, [string]$DocumentPath, $Candidates, $Seen, $Attempts,
        [string[]]$ExcludedDirectories = @(), [string[]]$ExcludedDirectoryPrefixes = @())
    try {
        if (-not (Test-WinBookSplitAutomaticPath $Path $ApplicationRoot $DocumentPath)) { throw 'Document-directory or unrelated-CWD candidate excluded.' }
        $path = Get-WinBookSplitTrustedPath -Path $Path -Automatic
        $path = ([WinBookSplit.Preflight.Probe]::Canonical($path)).Path
        foreach ($directory in $ExcludedDirectories) {
            $excluded = [IO.Path]::GetFullPath($directory).TrimEnd('\')
            if ($path.Equals($excluded, [StringComparison]::OrdinalIgnoreCase) -or $path.StartsWith($excluded + '\', [StringComparison]::OrdinalIgnoreCase)) {
                throw 'Application environment excluded from setup base-interpreter discovery.'
            }
        }
        foreach ($prefix in $ExcludedDirectoryPrefixes) {
            if ($path.StartsWith([IO.Path]::GetFullPath($prefix), [StringComparison]::OrdinalIgnoreCase)) {
                throw 'Preserved setup artifact excluded from base-interpreter discovery.'
            }
        }
        if (-not (Test-WinBookSplitAutomaticPath $path $ApplicationRoot $DocumentPath)) { throw 'Canonical candidate is in an excluded document or current directory.' }
        if ($Seen.Contains($path)) { return }
        if ($Candidates.Count -ge 32) { throw 'Automatic discovery exceeds the 32-interpreter limit.' }
        $null = $Seen.Add($path)
        $Candidates.Add([pscustomobject]@{Path=$path;Source=$Source})
    } catch {
        $Attempts.Add([pscustomobject]@{Path=$Path;Source=$Source;Stage='entry';Accepted=$false;Reason='rejected_entry_point';Message=$_.Exception.Message;Probe=$null})
    }
}

function Select-WinBookSplitPythonCandidate {
    param($Candidates, $Policy)
    $selected = $null; $selectedRank = [int]::MaxValue; $selectedPatch = -1
    foreach ($candidate in $Candidates) {
        if (-not (Test-WinBookSplitPythonVersion -Version $candidate.Version -Policy $Policy).Ok) { continue }
        $parts = ConvertTo-WinBookSplitCompatibilityVersion $candidate.Version
        $rank = 0
        foreach ($branch in $Policy.python.branches) {
            if ($parts[0] -eq $branch.major -and $parts[1] -eq $branch.minor) { break }
            $rank++
        }
        if ($null -eq $selected -or $rank -lt $selectedRank -or
            ($rank -eq $selectedRank -and ($parts[2] -gt $selectedPatch -or
                ($parts[2] -eq $selectedPatch -and [StringComparer]::Ordinal.Compare($candidate.Path, $selected.Path) -lt 0)))) {
            $selected = $candidate; $selectedRank = $rank; $selectedPatch = $parts[2]
        }
    }
    return $selected
}

function Get-WinBookSplitDiscoveryTimeout {
    param($Clock, [double]$Requested, [switch]$Optional)
    $remaining = 60.0 - $Clock.Elapsed.TotalSeconds
    if ($remaining -lt 0.05) {
        if ($Optional) { return 0.0 }
        $failure = New-Object InvalidOperationException('Automatic Python discovery exceeded its overall 60-second deadline.')
        $failure.Data['Code'] = 'processor_timeout'; throw $failure
    }
    return [Math]::Min($Requested, $remaining)
}

function Resolve-WinBookSplitAutomaticPython {
    param([string]$ApplicationRoot, [string]$DocumentPath, $Attempts, $Policy,
        [double]$TimeoutSeconds = 10, [switch]$InterpreterOnly, [string[]]$ExcludedDirectories = @(), [string[]]$ExcludedDirectoryPrefixes = @())
    $clock = [Diagnostics.Stopwatch]::StartNew()
    $candidates = New-Object 'Collections.Generic.List[object]'
    $accepted = New-Object 'Collections.Generic.List[object]'
    $seen = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    # Probe independent PATH before optional launcher work so a stopped listing
    # failure cannot consume its opportunity. Selection waits for all candidates.
    $direct = [string[]]@(Get-WinBookSplitPathCandidates 'python.exe' $ApplicationRoot $DocumentPath $Attempts)
    [Array]::Sort($direct, [StringComparer]::Ordinal)
    foreach ($path in $direct) { Add-WinBookSplitPythonDiscoveryCandidate $path 'PATH' $ApplicationRoot $DocumentPath $candidates $seen $Attempts -ExcludedDirectories $ExcludedDirectories -ExcludedDirectoryPrefixes $ExcludedDirectoryPrefixes }
    foreach ($candidate in $candidates) {
        $timeout = Get-WinBookSplitDiscoveryTimeout $clock $TimeoutSeconds
        $result = Test-WinBookSplitPythonCandidate $candidate.Path $candidate.Source $ApplicationRoot $Attempts $timeout -InterpreterOnly:$InterpreterOnly
        if ($result) { $accepted.Add($result) }
    }
    $processed = $candidates.Count
    $listingTimeout = [Math]::Min($TimeoutSeconds, 2.0)
    $timeout = Get-WinBookSplitDiscoveryTimeout $clock ([Math]::Min($TimeoutSeconds, 5.0)) -Optional
    $registered = @()
    if ($timeout -gt 0) { $registered = @(Get-WinBookSplitRegisteredManagerPaths $ApplicationRoot $DocumentPath $Attempts $timeout -WithRegistration) }
    $managers = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($registration in $registered) { $null = $managers.Add($registration.Path) }
    foreach ($path in @(Get-WinBookSplitPathCandidates 'pymanager.exe' $ApplicationRoot $DocumentPath $Attempts)) {
        if (-not $managers.Contains($path)) { $Attempts.Add([pscustomobject]@{Path=$path;Source='manager_listing';Stage='entry';Accepted=$false;Reason='manager_entry_rejected';Message='Unverified PATH manager name excluded; installed-only listing requires verified physical registration.';Probe=$null}) }
    }
    $queries = New-Object 'Collections.Generic.List[object]'
    # Registration order cannot affect the bounded route budget or final selection.
    $managerPaths = [string[]]@($managers)
    [Array]::Sort($managerPaths, [StringComparer]::Ordinal)
    foreach ($path in $managerPaths) {
        $registration = @($registered | Where-Object { $_.Path -ceq $path })[0]
        $queries.Add([pscustomobject]@{Path=$path;Kind='Manager';Source='manager_listing';Registration=$registration})
    }
    $legacy = [string[]]@(Get-WinBookSplitPathCandidates 'py.exe' $ApplicationRoot $DocumentPath $Attempts)
    [Array]::Sort($legacy, [StringComparer]::Ordinal)
    foreach ($path in $legacy) { $queries.Add([pscustomobject]@{Path=$path;Kind='Legacy';Source='py_launcher_listing'}) }
    $launchers = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($query in $queries) {
        $probe = $null; $stage = 'listing'
        try {
            $timeout = Get-WinBookSplitDiscoveryTimeout $clock $listingTimeout -Optional
            if ($timeout -eq 0) {
                $Attempts.Add([pscustomobject]@{Path=$query.Path;Source=$query.Source;Stage='listing';Accepted=$false;Reason='discovery_deadline';Message='Optional listing skipped at the overall discovery deadline; validated independent candidates retained.';Probe=$null})
                break
            }
            if ($launchers.Count -ge 4) { throw 'Read-only discovery exceeds the four-launcher limit.' }
            $path = $query.Path
            if ($query.Kind -eq 'Legacy') {
                $stage = 'entry'
                $path = Get-WinBookSplitTrustedPath $path -Automatic
                if (-not (Test-WinBookSplitAutomaticPath $path $ApplicationRoot $DocumentPath)) { throw 'Automatic launcher is in an excluded directory.' }
            }
            if (-not $launchers.Add($path)) { continue }
            $stage = 'listing'
            if ($query.Kind -eq 'Manager') {
                $probe = Invoke-WinBookSplitManagerListing -Registration $query.Registration -ApplicationRoot $ApplicationRoot -DocumentPath $DocumentPath -TimeoutSeconds $timeout
            } else {
                $probe = Invoke-WinBookSplitRuntimeProbe $path @('-0p') $ApplicationRoot $timeout
            }
            if ($null -ne $probe -and -not [string]::IsNullOrWhiteSpace($probe.StartError)) { $stage = 'entry' }
            $paths = [string[]]@(ConvertFrom-WinBookSplitPythonListing $probe $query.Kind)
            [Array]::Sort($paths, [StringComparer]::Ordinal)
            $Attempts.Add([pscustomobject]@{Path=$path;Source=$query.Source;Stage='listing';Accepted=$true;Reason='installed_runtimes_listed';Message='Read-only installed-runtime listing; no runtime launch/install requested.';Probe=$probe})
            $source = 'py_launcher'; if ($query.Source -eq 'manager_listing') { $source = 'manager' }
            foreach ($candidate in $paths) { Add-WinBookSplitPythonDiscoveryCandidate $candidate $source $ApplicationRoot $DocumentPath $candidates $seen $Attempts -ExcludedDirectories $ExcludedDirectories -ExcludedDirectoryPrefixes $ExcludedDirectoryPrefixes }
        } catch {
            if (Test-WinBookSplitInterruption $_.Exception) { throw }
            $reason = 'launcher_listing_failed'; if ($query.Kind -eq 'Manager') { $reason = 'manager_listing_failed' }
            if ($stage -eq 'entry') {
                $reason = 'rejected_entry_point'
                if ($query.Kind -eq 'Manager') { $reason = 'manager_entry_rejected' }
            }
            if ($_.Exception.Data.Contains('Reason')) { $reason = $_.Exception.Data['Reason'] }
            if ($_.Exception.Data.Contains('Stage')) { $stage = $_.Exception.Data['Stage'] }
            $Attempts.Add([pscustomobject]@{Path=$query.Path;Source=$query.Source;Stage=$stage;Accepted=$false;Reason=$reason;Message=$_.Exception.Message;Probe=$probe})
        }
    }
    for ($index = $processed; $index -lt $candidates.Count; $index++) {
        $candidate = $candidates[$index]
        $timeout = Get-WinBookSplitDiscoveryTimeout $clock $TimeoutSeconds -Optional
        if ($timeout -eq 0) {
            $Attempts.Add([pscustomobject]@{Path=$candidate.Path;Source=$candidate.Source;Stage='probe';Accepted=$false;Reason='discovery_deadline';Message='Optional listed candidate skipped at the overall discovery deadline; validated independent candidates retained.';Probe=$null})
            break
        }
        try {
            $result = Test-WinBookSplitPythonCandidate $candidate.Path $candidate.Source $ApplicationRoot $Attempts $timeout -InterpreterOnly:$InterpreterOnly
            if ($result) { $accepted.Add($result) }
        } catch {
            $failure = $_.Exception
            $probe = $failure.Data['Probe']
            if ($failure.Data['Code'] -eq 'processor_timeout' -and -not $failure.Data['FatalDiscovery'] -and
                $null -ne $probe -and $probe.TimedOut -and -not $probe.Cancelled -and
                $probe.ParentStopped -and $probe.DescendantsStopped -and -not $probe.StopError) {
                $Attempts.Add([pscustomobject]@{Path=$candidate.Path;Source=$candidate.Source;Stage='probe';Accepted=$false;Reason='runtime_probe_failed';Message='Optional listed interpreter timed out; its owned process tree was proved stopped and independent candidates were retained.';Probe=$probe})
            } else { throw }
        }
    }
    $selected = Select-WinBookSplitPythonCandidate -Candidates $accepted -Policy $Policy
    if ($selected) {
        $Attempts.Add([pscustomobject]@{Path=$selected.Path;Source='selection';Stage='selection';Accepted=$true;Reason='supported';Message='Selected by policy branch preference, highest numeric stable patch, then ordinal canonical path.';Probe=$null})
        $selected.Attempts = @($Attempts.ToArray())
    }
    elseif ($candidates.Count -eq 0) {
        $Attempts.Add([pscustomobject]@{Path=$null;Source='discovery';Stage='discovery';Accepted=$false;Reason='no_candidates';Message='No concrete interpreter candidates passed discovery entry checks.';Probe=$null})
    }
    return $selected
}

function Initialize-WinBookSplitRuntimeProbe {
    if ('WinBookSplit.Preflight.Probe' -as [type]) { return }
    Add-Type -TypeDefinition @'
using System;
using System.Collections;
using System.Diagnostics;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
namespace WinBookSplit.Preflight {
    public sealed class Result {
        public int? ExitCode; public int? Pid; public bool TimedOut;
        public bool ParentStopped; public bool? DescendantsStopped = null;
        public bool StreamsComplete; public string StartError;
        public string Stdout; public string Stderr;
        public long StdoutTotalBytes; public long StderrTotalBytes;
        public bool StdoutTruncated; public bool StderrTruncated;
        public string StreamError; public double ElapsedSeconds;
    }
    sealed class Tail {
        internal readonly object Gate = new object();
        internal byte[] Bytes = new byte[65536]; internal int Count;
        internal long Total; internal bool Done; internal string Error;
        internal void Read(Stream stream) {
            try {
                byte[] block = new byte[8192]; int count;
                while ((count = stream.Read(block, 0, block.Length)) != 0) {
                    lock (Gate) {
                        Total += count;
                        int discard = Math.Max(0, Count + count - Bytes.Length);
                        if (discard != 0) { Buffer.BlockCopy(Bytes, discard, Bytes, 0, Count - discard); Count -= discard; }
                        Buffer.BlockCopy(block, 0, Bytes, Count, count); Count += count;
                    }
                }
            } catch (Exception error) { lock (Gate) { Error = error.Message; } }
            finally {
                try { stream.Dispose(); } catch (Exception error) { lock (Gate) { Error = error.Message; } }
                lock (Gate) { Done = true; }
            }
        }
        internal string Text() { lock (Gate) { return Encoding.UTF8.GetString(Bytes, 0, Count); } }
        internal bool Complete() { lock (Gate) { return Done; } }
    }
    sealed class State {
        internal Process Child; internal Tail Out = new Tail(); internal Tail Err = new Tail();
        readonly object Gate = new object(); bool Returning; int Completed; int Started;
        internal void Start(Thread thread) {
            lock (Gate) { Started++; }
            try { thread.Start(); } catch { lock (Gate) { Started--; } throw; }
        }
        internal void Finished() {
            lock (Gate) { Completed++; if (Returning && Completed == Started) { try { Child.Dispose(); } catch { } } }
        }
        internal void Return() { lock (Gate) { Returning = true; if (Completed == Started) Child.Dispose(); } }
    }
    public sealed class PathDetails { public string Path; public uint Links; }
    public static class Probe {
        [StructLayout(LayoutKind.Sequential)] struct FileInfo {
            public uint Attributes, CreationLow, CreationHigh, AccessLow, AccessHigh, WriteLow, WriteHigh;
            public uint Volume, SizeHigh, SizeLow, Links, IndexHigh, IndexLow;
        }
        [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        static extern IntPtr CreateFileW(string name, uint access, uint share, IntPtr security, uint creation, uint flags, IntPtr template);
        [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        static extern uint GetFinalPathNameByHandleW(IntPtr handle, StringBuilder path, uint size, uint flags);
        [DllImport("kernel32.dll", SetLastError=true)] static extern bool GetFileInformationByHandle(IntPtr handle, out FileInfo info);
        [DllImport("kernel32.dll", SetLastError=true)] static extern bool CloseHandle(IntPtr handle);
        [DllImport("kernel32.dll", SetLastError=true)]
        static extern bool DeviceIoControl(IntPtr handle, uint code, IntPtr input, uint inputLength, byte[] output, uint outputLength, out uint returned, IntPtr overlapped);
        public static bool IsAppExecutionAlias(string path) {
            // OPEN_REPARSE_POINT: inspect the file itself, never its redirect.
            IntPtr handle=CreateFileW(path,0,7,IntPtr.Zero,3,0x02200000,IntPtr.Zero);
            if (handle==new IntPtr(-1)) throw new Win32Exception(Marshal.GetLastWin32Error());
            try {
                byte[] data=new byte[16384]; uint count;
                if (!DeviceIoControl(handle,0x000900A8,IntPtr.Zero,0,data,(uint)data.Length,out count,IntPtr.Zero)) {
                    int error=Marshal.GetLastWin32Error();
                    if (error==4390) return false; // ERROR_NOT_A_REPARSE_POINT
                    throw new Win32Exception(error);
                }
                // Only the documented tag is read; no undocumented payload layout.
                return count>=8 && count<=data.Length && BitConverter.ToUInt32(data,0)==0x8000001B &&
                    BitConverter.ToUInt16(data,4)+8==count;
            } finally { if (!CloseHandle(handle)) throw new Win32Exception(Marshal.GetLastWin32Error()); }
        }
        public static PathDetails Canonical(string path) {
            IntPtr handle=CreateFileW(path,0x80,7,IntPtr.Zero,3,0x02200000,IntPtr.Zero);
            if (handle==new IntPtr(-1)) throw new Win32Exception(Marshal.GetLastWin32Error());
            try {
                StringBuilder text=new StringBuilder(32768); FileInfo info;
                uint length=GetFinalPathNameByHandleW(handle,text,(uint)text.Capacity,0);
                if (length==0 || length>=text.Capacity || !GetFileInformationByHandle(handle,out info)) throw new Win32Exception(Marshal.GetLastWin32Error());
                string final=text.ToString();
                if (final.StartsWith(@"\\?\UNC\",StringComparison.OrdinalIgnoreCase)) final=@"\\"+final.Substring(8);
                else if (final.StartsWith(@"\\?\",StringComparison.OrdinalIgnoreCase)) final=final.Substring(4);
                return new PathDetails { Path=final, Links=info.Links };
            } finally { if (!CloseHandle(handle)) throw new Win32Exception(Marshal.GetLastWin32Error()); }
        }
        public static Result Run(string path, string arguments, string cwd, int timeoutMilliseconds) {
            Result result = new Result(); Stopwatch clock = Stopwatch.StartNew();
            State state = new State(); state.Child = new Process(); bool started = false;
            try {
                ProcessStartInfo info = new ProcessStartInfo(path, arguments);
                info.WorkingDirectory = cwd; info.UseShellExecute = false; info.CreateNoWindow = true;
                info.RedirectStandardInput = true; info.RedirectStandardOutput = true; info.RedirectStandardError = true;
                foreach (string name in new ArrayList(info.EnvironmentVariables.Keys)) {
                    if (name.StartsWith("PYTHON", StringComparison.OrdinalIgnoreCase) ||
                        name.StartsWith("PYLAUNCHER", StringComparison.OrdinalIgnoreCase)) info.EnvironmentVariables.Remove(name);
                }
                info.EnvironmentVariables["PYTHONUTF8"] = "1";
                info.EnvironmentVariables["PYTHONIOENCODING"] = "utf-8:backslashreplace";
                info.EnvironmentVariables["PYTHON_MANAGER_AUTOMATIC_INSTALL"] = "false";
                info.EnvironmentVariables["PYLAUNCHER_NO_SEARCH_PATH"] = "1";
                state.Child.StartInfo = info; started = state.Child.Start();
                result.Pid = state.Child.Id; state.Child.StandardInput.Close();
                Thread output = new Thread(delegate() { state.Out.Read(state.Child.StandardOutput.BaseStream); state.Finished(); });
                Thread error = new Thread(delegate() { state.Err.Read(state.Child.StandardError.BaseStream); state.Finished(); });
                output.IsBackground = true; error.IsBackground = true; state.Start(output); state.Start(error);
                while (!(state.Child.HasExited && state.Out.Complete() && state.Err.Complete())) {
                    if (clock.ElapsedMilliseconds >= timeoutMilliseconds) { result.TimedOut = true; break; }
                    Thread.Sleep(5);
                }
                if (result.TimedOut && !state.Child.HasExited) {
                    try { state.Child.Kill(); } catch (Exception failure) { result.StartError = "Parent termination: " + failure.Message; }
                }
                // One bounded shutdown grace; inherited open pipes never turn
                // into an unbounded GetResult, stream Dispose or WaitForExit.
                if (!state.Child.HasExited) state.Child.WaitForExit(250);
                result.ParentStopped = state.Child.HasExited;
                if (result.ParentStopped) result.ExitCode = state.Child.ExitCode;
            } catch (Exception error) {
                result.StartError = error.Message;
                if (started) {
                    try { if (!state.Child.HasExited) state.Child.Kill(); state.Child.WaitForExit(250); result.ParentStopped = state.Child.HasExited; }
                    catch (Exception shutdown) { result.StartError += "; parent shutdown: " + shutdown.Message; }
                } else { result.ParentStopped = true; }
            } finally {
                result.StreamsComplete = state.Out.Complete() && state.Err.Complete();
                result.Stdout = state.Out.Text(); result.Stderr = state.Err.Text();
                lock (state.Out.Gate) { result.StdoutTotalBytes = state.Out.Total; result.StdoutTruncated = state.Out.Total > 65536; }
                lock (state.Err.Gate) { result.StderrTotalBytes = state.Err.Total; result.StderrTruncated = state.Err.Total > 65536; }
                result.StreamError = state.Out.Error ?? state.Err.Error;
                result.ElapsedSeconds = clock.Elapsed.TotalSeconds;
                try { if (started) state.Return(); else state.Child.Dispose(); }
                catch (Exception close) { result.StreamError = "Probe handle close: " + close.Message; }
            }
            return result;
        }
    }
}
'@ -ErrorAction Stop
}

function Invoke-WinBookSplitRuntimeProbe {
    param([string]$Path, [string[]]$Arguments, [string]$ApplicationRoot,
          [ValidateRange(0.05, 60)][double]$TimeoutSeconds = 10)
    return Invoke-WinBookSplitProcess -Path $Path -Arguments $Arguments -WorkingDirectory $ApplicationRoot -TimeoutSeconds $TimeoutSeconds
}

function Get-WinBookSplitTrustedPath {
    param([string]$Path, [switch]$Directory, [switch]$Automatic)
    if ([string]::IsNullOrWhiteSpace($Path) -or $Path -notmatch '^(?:[A-Za-z]:[\\/]|\\\\[^\\]+\\[^\\]+\\)') {
        throw 'Select an absolute literal FileSystem path.'
    }
    $full = [IO.Path]::GetFullPath($Path)
    $attributes = [IO.File]::GetAttributes($full)
    if ([bool]($attributes -band [IO.FileAttributes]::Directory) -ne [bool]$Directory) {
        throw 'The selected path has the wrong file/directory type.'
    }
    $current = $full
    while ($current) {
        if ([IO.File]::GetAttributes($current) -band [IO.FileAttributes]::ReparsePoint) {
            throw 'Dependency paths and their ancestors must not be reparse points.'
        }
        $current = [IO.Path]::GetDirectoryName($current)
    }
    Initialize-WinBookSplitRuntimeProbe
    $canonical = [WinBookSplit.Preflight.Probe]::Canonical($full)
    if (-not $canonical.Path.TrimEnd('\').Equals($full.TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Dependency aliases such as short 8.3 names are ambiguous; select the ordinary long path explicitly.'
    }
    if ($Automatic -and -not $Directory -and $canonical.Links -ne 1) {
        throw 'Automatic dependency discovery refuses hard-linked executable aliases; select a trusted ordinary executable explicitly.'
    }
    return $full
}

function Test-WinBookSplitAutomaticPath {
    param([string]$Path, [string]$ApplicationRoot, [string]$DocumentPath)
    $full = [IO.Path]::GetFullPath($Path)
    $cwd = [IO.Path]::GetFullPath((Get-Location).ProviderPath).TrimEnd('\')
    Initialize-WinBookSplitRuntimeProbe
    $cwd = ([WinBookSplit.Preflight.Probe]::Canonical($cwd)).Path.TrimEnd('\')
    if (-not $cwd.Equals($ApplicationRoot.TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase) -and
        ($full.StartsWith($cwd + '\', [StringComparison]::OrdinalIgnoreCase))) { return $false }
    if (-not [string]::IsNullOrWhiteSpace($DocumentPath)) {
        $bookDirectory = [IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($DocumentPath)).TrimEnd('\')
        if ([IO.Directory]::Exists($bookDirectory)) { $bookDirectory = ([WinBookSplit.Preflight.Probe]::Canonical($bookDirectory)).Path.TrimEnd('\') }
        if ($full.StartsWith($bookDirectory + '\', [StringComparison]::OrdinalIgnoreCase)) { return $false }
    }
    return $true
}

function Get-WinBookSplitPathCandidates {
    param([string]$Name, [string]$ApplicationRoot, [string]$DocumentPath, $Attempts)
    $seen = @{}; $directories = @($env:PATH -split ';')
    if ($directories.Count -gt 256) { throw 'PATH exceeds the bounded 256-directory discovery limit.' }
    foreach ($directory in $directories) {
        $directory = $directory.Trim().Trim('"')
        if ([string]::IsNullOrWhiteSpace($directory) -or $directory -notmatch '^(?:[A-Za-z]:[\\/]|\\\\[^\\]+\\[^\\]+\\)') { continue }
        try {
            $candidate = [IO.Path]::Combine([IO.Path]::GetFullPath($directory), $Name)
            if ($seen.ContainsKey($candidate)) { continue }; $seen[$candidate] = $true
            if (-not (Test-WinBookSplitAutomaticPath $candidate $ApplicationRoot $DocumentPath)) {
                $Attempts.Add([pscustomobject]@{ Path = $candidate; Source = 'PATH'; Stage = 'entry'; Accepted = $false; Reason = 'rejected_entry_point'; Message = 'Document-directory or unrelated-CWD candidate excluded.'; Probe = $null })
                continue
            }
            if ([IO.File]::Exists($candidate)) { $candidate }
        } catch {
            $Attempts.Add([pscustomobject]@{ Path = $directory; Source = 'PATH'; Stage = 'entry'; Accepted = $false; Reason = 'rejected_entry_point'; Message = $_.Exception.Message; Probe = $null })
        }
    }
}

function New-WinBookSplitDependencyError {
    param([string]$Code, [string]$Message, $Attempts, [string]$Reason, $Policy, [string]$ApplicationRoot)
    if ($null -ne $Policy -and $Code -cin @('runtime_invalid', 'runtime_not_found')) {
        $summary = Get-WinBookSplitDiscoverySummary -Attempts $Attempts.ToArray() -Policy $Policy -Code $Code -FailureReason $Reason
        $relevant = $summary.Relevant
        if ($summary.Reason -cin @('missing_pypdf', 'pypdf_version_mismatch', 'pypdf_origin_invalid') -and
            $summary.Stage -ceq 'package' -and $null -ne $summary.ObservedVersion -and
            $null -ne $relevant -and $relevant.IdentityVerified -is [bool] -and $relevant.IdentityVerified -and
            (Test-WinBookSplitPythonVersion -Version $summary.ObservedVersion -Policy $Policy).Ok) {
            $Code = 'setup_required'
            $repair = $summary.Source -ceq 'application_venv'
            $basePath = $null
            if ($summary.Source -ceq 'explicit') {
                $venvPrefix = [IO.Path]::Combine($ApplicationRoot, '.venv') + '\'
                $repair = [IO.Path]::GetFullPath($summary.ObservedPath).StartsWith($venvPrefix, [StringComparison]::OrdinalIgnoreCase)
                if (-not $repair) { $basePath = $summary.ObservedPath }
            }
            $Message = 'Supported Python is available, but the required locked pypdf package is not ready. PowerShell setup: ' +
                (Get-WinBookSplitSetupCommand -ApplicationRoot $ApplicationRoot -PythonPath $basePath -Repair:$repair)
        }
    }
    $error = New-Object InvalidOperationException($Message)
    $error.Data['Code'] = $Code; $error.Data['Attempts'] = @($Attempts.ToArray())
    if (-not [string]::IsNullOrEmpty($Reason)) { $error.Data['Reason'] = $Reason }
    return $error
}

function Test-WinBookSplitPythonCandidate {
    param([string]$Path, [string]$Source, [string]$ApplicationRoot, $Attempts,
          [double]$TimeoutSeconds, [switch]$InterpreterOnly)
    $probe = $null
    $reason = 'rejected_entry_point'; $stage = 'entry'
    $pathAccepted = $false
    $observedVersion = $null; $identityVerified = $false
    try {
        $path = Get-WinBookSplitTrustedPath -Path $Path -Automatic:($Source -in @('PATH', 'py_launcher', 'manager', 'registry'))
        $pathAccepted = $true
        $reason = 'malformed_policy'; $stage = 'identity'
        $policyPath = Get-WinBookSplitTrustedPath ([IO.Path]::Combine($ApplicationRoot, 'engine', 'WinBookSplit.Compatibility.json'))
        $policyIdentity = Get-WinBookSplitCompatibilityPolicy -PolicyPath $policyPath -WithIdentity
        $policy = $policyIdentity.Policy
        $helper = Get-WinBookSplitTrustedPath ([IO.Path]::Combine($ApplicationRoot, 'engine', 'winbooksplit_compatibility.py'))
        $arguments = @('-I', '-B', '-X', 'utf8', $helper, '--probe', '--policy', $policyPath)
        if ($InterpreterOnly) { $arguments += '--interpreter-only' }
        $reason = 'runtime_probe_failed'; $stage = 'probe'
        $probe = Invoke-WinBookSplitRuntimeProbe -Path $path -Arguments $arguments -ApplicationRoot $ApplicationRoot -TimeoutSeconds $TimeoutSeconds
        Assert-WinBookSplitDiscoveryStopped $probe
        Assert-WinBookSplitProbeNotInterrupted $probe
        if ($probe.TimedOut -or $probe.Cancelled -or -not $probe.ParentStopped -or -not $probe.DescendantsStopped -or -not $probe.JobAssigned -or -not $probe.StreamsComplete -or $probe.StartError -or $probe.StopError -or $probe.StreamError -or $probe.StdoutTruncated -or $probe.StderrTruncated) {
            throw 'The bounded interpreter probe did not complete safely.'
        }
        $reason = 'malformed_runtime'; $stage = 'identity'
        $lines = @($probe.Stdout -split '\r?\n' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        if ($lines.Count -ne 1) { throw 'Expected one isolated interpreter JSON record.' }
        if (-not $lines[0].TrimStart().StartsWith('{')) { throw 'The interpreter JSON root must be an object.' }
        # Reuse the strict bounded parser: duplicate keys cannot supply a second identity.
        $strictRecord = [WinBookSplitSupport.StrictJson]::Parse($lines[0])
        if ($strictRecord -isnot [System.Collections.IDictionary]) { throw 'The interpreter JSON root must be an object.' }
        $record = $lines[0] | ConvertFrom-Json -ErrorAction Stop
        foreach ($field in @('protocol', 'version', 'implementation', 'platform', 'machine', 'executable', 'message', 'reason', 'policy_sha256')) {
            if ($record.$field -isnot [string] -or [string]::IsNullOrWhiteSpace($record.$field)) {
                throw ('Interpreter JSON requires a nonempty scalar string: ' + $field)
            }
        }
        if (($record.schema_version -isnot [int] -and $record.schema_version -isnot [long]) -or
            $record.protocol -cne 'winbooksplit.runtime' -or $record.schema_version -ne 1 -or
            $record.ok -isnot [bool] -or
            ($record.policy_version -isnot [int] -and $record.policy_version -isnot [long]) -or
            $record.policy_version -ne $policy['policy_version'] -or
            $record.policy_sha256 -cnotmatch '^[a-f0-9]{64}$' -or
            $record.policy_sha256 -cne $policyIdentity.Sha256) {
            throw 'Invalid interpreter protocol or compatibility policy identity.'
        }
        if ($record.executable -notmatch '\A(?:[A-Za-z]:[\\/]|\\\\[^\\]+\\[^\\]+\\)' -or
            -not ([IO.Path]::GetFullPath($record.executable).Equals($path, [StringComparison]::OrdinalIgnoreCase))) {
            $reason = 'malformed_runtime'
            throw 'Interpreter executable identity does not match the selected trusted path.'
        }
        $decision = Test-WinBookSplitInterpreterRecord -Record $strictRecord -Policy $policy
        $parts = ConvertTo-WinBookSplitCompatibilityVersion $record.version
        if ($null -eq $parts -and $record.version -cmatch '\A(0|[1-9][0-9]{0,6})\.(0|[1-9][0-9]{0,6})\.(0|[1-9][0-9]{0,6})(?:(?:a|b|rc)[0-9]{1,6}|\.dev[0-9]{1,6}|\+)\z') {
            $parts = @([long]$Matches[1], [long]$Matches[2], [long]$Matches[3])
        }
        if ($decision.Reason -cnotin @('malformed_runtime','malformed_policy') -and $null -ne $parts -and
            $record.isolated -is [bool] -and $record.isolated -and $record.dont_write_bytecode -is [bool] -and $record.dont_write_bytecode -and
            $record.version_info -is [array] -and $record.version_info.Count -eq 3 -and
            $parts[0] -eq $record.version_info[0] -and $parts[1] -eq $record.version_info[1] -and $parts[2] -eq $record.version_info[2]) {
            $observedVersion = $record.version; $identityVerified = $true
        }
        if (-not $decision.Ok) {
            $reason = $decision.Reason
            if ($reason -cnotin @('malformed_runtime','malformed_policy')) { $stage = 'compatibility' }
            throw ('Incompatible interpreter: ' + $decision.Message)
        }
        if (-not $record.ok -or $probe.ExitCode -ne 0) {
            $reason = $record.reason
            $stage = 'package'
            if ($reason -cin @('malformed_runtime','malformed_policy')) { $stage = 'identity' }
            throw ('Incompatible interpreter/dependency: ' + $record.message + ' Observed Python ' + $record.version + '.')
        }
        $pypdfVersion, $pypdfPath = $null, $null
        if (-not $InterpreterOnly) {
            $stage = 'package'; $reason = 'pypdf_origin_invalid'
            foreach ($field in @('pypdf_version', 'pypdf_path')) {
                if ($record.$field -isnot [string] -or [string]::IsNullOrWhiteSpace($record.$field)) { throw ('Interpreter JSON requires a nonempty scalar string: ' + $field) }
            }
            if ($record.pypdf_version -cne $policy['pypdf']['version']) { $reason = 'pypdf_version_mismatch'; throw 'The locked pypdf version was not imported.' }
            $pypdfPath = Get-WinBookSplitTrustedPath $record.pypdf_path
            $pypdfVersion = $record.pypdf_version
        }
        $reason = 'supported'
        $Attempts.Add([pscustomobject]@{ Path = $path; Source = $Source; Stage = 'verified'; Accepted = $true; Reason = $reason; Message = $record.message; Probe = $probe;
            ObservedVersion = $observedVersion; IdentityVerified = $identityVerified })
        return [pscustomobject]@{ Path = $path; Version = $record.version; Source = $Source;
            PypdfVersion = $pypdfVersion; PypdfPath = $pypdfPath; PolicyVersion = $record.policy_version;
            PolicySha256 = $record.policy_sha256; InterpreterOnly = [bool]$InterpreterOnly;
            Details = $record; Probe = $probe; Arguments = @('-I', '-B'); Attempts = @($Attempts.ToArray()) }
    } catch {
        if (Test-WinBookSplitInterruption $_.Exception) { throw }
        if (-not $pathAccepted) {
            $pathFailure = $_.Exception
            while ($pathFailure.InnerException) { $pathFailure = $pathFailure.InnerException }
            if ($pathFailure -is [IO.FileNotFoundException] -or $pathFailure -is [IO.DirectoryNotFoundException]) { $reason = 'missing_python' }
        }
        $Attempts.Add([pscustomobject]@{ Path = $Path; Source = $Source; Stage = $stage; Accepted = $false; Reason = $reason; Message = $_.Exception.Message; Probe = $probe;
            ObservedVersion = $observedVersion; IdentityVerified = $identityVerified })
        return $null
    }
}

function Resolve-WinBookSplitRuntime {
    param([Parameter(Mandatory = $true)][string]$ApplicationRoot, [string]$PythonPath,
          [string]$DocumentPath, [ValidateRange(0.05, 60)][double]$TimeoutSeconds = 10, [switch]$InterpreterOnly)
    $root = Get-WinBookSplitTrustedPath -Path $ApplicationRoot -Directory
    $attempts = New-Object 'System.Collections.Generic.List[object]'
    try { $policy = Get-WinBookSplitCompatibilityPolicy -PolicyPath ([IO.Path]::Combine($root, 'engine', 'WinBookSplit.Compatibility.json')) }
    catch { throw (New-WinBookSplitDependencyError 'runtime_invalid' ('Shipped compatibility policy is missing or invalid: ' + $_.Exception.Message) $attempts 'malformed_policy') }
    $failurePolicy = $null
    if (-not $InterpreterOnly) { $failurePolicy = $policy }
    $setup = 'Use ' + (Get-WinBookSplitPythonPolicyDescription -Policy $policy) + ' and locked pypdf ' + $policy['pypdf']['version'] + '. Select a trusted -PythonPath or create the application .venv. PowerShell setup: ' + (Get-WinBookSplitSetupCommand -ApplicationRoot $root)

    if (-not [string]::IsNullOrWhiteSpace($PythonPath)) {
        $selected = Test-WinBookSplitPythonCandidate $PythonPath 'explicit' $root $attempts $TimeoutSeconds -InterpreterOnly:$InterpreterOnly
        if ($selected) { return $selected }
        throw (New-WinBookSplitDependencyError 'runtime_invalid' ("Python preflight failed for '$PythonPath': " + $attempts[-1].Message + ' ' + $setup) $attempts -Policy $failurePolicy -ApplicationRoot $root)
    }
    $venv = [IO.Path]::Combine($root, '.venv')
    $venvExists = $true
    try { $null = [IO.File]::GetAttributes($venv) }
    catch {
        $failure = $_.Exception
        while ($failure.InnerException) { $failure = $failure.InnerException }
        if ($failure -is [IO.FileNotFoundException] -or $failure -is [IO.DirectoryNotFoundException]) { $venvExists = $false }
        else {
            $attempts.Add([pscustomobject]@{Path=$venv;Source='application_venv';Stage='venv';Accepted=$false;Reason='invalid_application_venv';Message=$_.Exception.Message;Probe=$null})
            throw (New-WinBookSplitDependencyError 'runtime_invalid' ('Application .venv presence could not be checked safely. ' + $setup) $attempts)
        }
    }
    if ($venvExists) {
        $candidate = [IO.Path]::Combine($venv, 'Scripts\python.exe')
        try { $null = Get-WinBookSplitTrustedPath -Path $venv -Directory }
        catch {
            $attempts.Add([pscustomobject]@{Path=$venv;Source='application_venv';Stage='venv';Accepted=$false;Reason='invalid_application_venv';Message=$_.Exception.Message;Probe=$null})
            throw (New-WinBookSplitDependencyError 'runtime_invalid' ('Application .venv is not a trusted ordinary directory: ' + $_.Exception.Message + ' ' + $setup) $attempts)
        }
        $selected = Test-WinBookSplitPythonCandidate $candidate 'application_venv' $root $attempts $TimeoutSeconds -InterpreterOnly:$InterpreterOnly
        if ($selected) { return $selected }
        throw (New-WinBookSplitDependencyError 'runtime_invalid' ("Application .venv preflight failed for '$candidate': " + $attempts[-1].Message + ' PowerShell repair: ' + (Get-WinBookSplitSetupCommand -ApplicationRoot $root -Repair)) $attempts -Policy $failurePolicy -ApplicationRoot $root)
    }
    $selected = Resolve-WinBookSplitAutomaticPython -ApplicationRoot $root -DocumentPath $DocumentPath -Attempts $attempts -Policy $policy -TimeoutSeconds $TimeoutSeconds -InterpreterOnly:$InterpreterOnly
    if ($selected) { return $selected }
    throw (New-WinBookSplitDependencyError 'runtime_not_found' ('No compatible trusted Python was found. ' + $setup) $attempts -Policy $failurePolicy -ApplicationRoot $root)
}

function Find-WinBookSplitSetupBase {
    param([string]$ApplicationRoot, $Policy, [string]$PythonPath, [string]$DocumentPath)
    # Separate read-only remedy discovery. Never replaces a processing choice,
    # requires pypdf, or appends these attempts to its authoritative failure.
    $root = Get-WinBookSplitTrustedPath -Path $ApplicationRoot -Directory
    $venv = [IO.Path]::Combine($root, '.venv')
    $prefixes = @([IO.Path]::Combine($root,'.WinBookSplit-venv-backup-'),[IO.Path]::Combine($root,'.WinBookSplit-incomplete-'))
    $attempts = New-Object 'Collections.Generic.List[object]'
    if (-not [string]::IsNullOrWhiteSpace($PythonPath)) {
        try {
            $explicit = Get-WinBookSplitTrustedPath $PythonPath
            $excluded = $explicit.StartsWith($venv+'\',[StringComparison]::OrdinalIgnoreCase)
            foreach ($prefix in $prefixes) { $excluded = $excluded -or $explicit.StartsWith($prefix,[StringComparison]::OrdinalIgnoreCase) }
            if (-not $excluded) {
                $base = Test-WinBookSplitPythonCandidate $explicit 'explicit' $root $attempts 5 -InterpreterOnly
                if ($base) { return $base }
            }
        } catch { if (Test-WinBookSplitInterruption $_.Exception) { throw } }
    }
    return (Resolve-WinBookSplitAutomaticPython -ApplicationRoot $root -DocumentPath $DocumentPath -Attempts $attempts -Policy $Policy -TimeoutSeconds 2 -InterpreterOnly -ExcludedDirectories @($venv) -ExcludedDirectoryPrefixes $prefixes)
}

function Test-WinBookSplitConverterCandidate {
    param([string]$Path, [string]$Source, [string]$ApplicationRoot, $Attempts, [double]$TimeoutSeconds)
    $probe = $null
    try {
        $path = Get-WinBookSplitTrustedPath $Path -Automatic:($Source -in @('PATH', 'known_location'))
        $policy = Get-WinBookSplitCompatibilityPolicy -PolicyPath ([IO.Path]::Combine($ApplicationRoot, 'engine', 'WinBookSplit.Compatibility.json'))
        $probe = Invoke-WinBookSplitRuntimeProbe $path @('--version') $ApplicationRoot -TimeoutSeconds $TimeoutSeconds
        Assert-WinBookSplitDiscoveryStopped $probe
        Assert-WinBookSplitProbeNotInterrupted $probe
        if ($probe.ExitCode -ne 0 -or $probe.TimedOut -or $probe.Cancelled -or -not $probe.ParentStopped -or -not $probe.DescendantsStopped -or -not $probe.JobAssigned -or -not $probe.StreamsComplete -or
            $probe.StartError -or $probe.StopError -or $probe.StreamError -or $probe.StdoutTruncated -or $probe.StderrTruncated -or
            $probe.Stdout -cnotmatch '\Aebook-convert(?:\.exe)? \(calibre ([0-9]+\.[0-9]+\.[0-9]+)\)(?:\r?\nCreated by: Kovid Goyal <kovid@kovidgoyal\.net>)?(?:\r?\n)?\z') {
            throw 'The converter did not return a complete ebook-convert version probe.'
        }
        $version = $Matches[1]
        if (-not (Test-WinBookSplitCalibreVersion -Version $version -Policy $policy)) { throw ('Expected stable Calibre ' + $policy['calibre']['minimum_version'] + ' or newer; observed ' + $version + '.') }
        $Attempts.Add([pscustomobject]@{ Path = $path; Source = $Source; Accepted = $true; Message = 'Supported converter version.'; Probe = $probe })
        return [pscustomobject]@{ Path = $path; Version = $version; Source = $Source; Probe = $probe; Attempts = @($Attempts.ToArray()) }
    } catch {
        if (Test-WinBookSplitInterruption $_.Exception) { throw }
        $Attempts.Add([pscustomobject]@{ Path = $Path; Source = $Source; Accepted = $false; Message = $_.Exception.Message; Probe = $probe })
        return $null
    }
}

function Resolve-WinBookSplitConverter {
    param([Parameter(Mandatory = $true)][string]$ApplicationRoot, [string]$CalibrePath,
          [string]$DocumentPath, [ValidateRange(0.05, 60)][double]$TimeoutSeconds = 10)
    $root = Get-WinBookSplitTrustedPath -Path $ApplicationRoot -Directory
    $attempts = New-Object 'System.Collections.Generic.List[object]'
    try { $policy = Get-WinBookSplitCompatibilityPolicy -PolicyPath ([IO.Path]::Combine($root, 'engine', 'WinBookSplit.Compatibility.json')) }
    catch { throw (New-WinBookSplitDependencyError 'converter_invalid' ('Compatibility policy could not be loaded: ' + $_.Exception.Message) $attempts 'malformed_policy') }
    $guidance = 'Select trusted stable Calibre ' + $policy['calibre']['minimum_version'] + ' or newer ebook-convert.exe with -CalibrePath, a trusted PATH entry, or a standard install; PDF-only processing does not require Calibre.'
    if (-not [string]::IsNullOrWhiteSpace($CalibrePath)) {
        $selected = Test-WinBookSplitConverterCandidate $CalibrePath 'explicit' $root $attempts $TimeoutSeconds
        if ($selected) { return $selected }
        throw (New-WinBookSplitDependencyError 'converter_invalid' ("Converter preflight failed for '$CalibrePath': " + $attempts[-1].Message + ' ' + $guidance) $attempts)
    }
    foreach ($candidate in @(Get-WinBookSplitPathCandidates 'ebook-convert.exe' $root $DocumentPath $attempts)) {
        $selected = Test-WinBookSplitConverterCandidate $candidate 'PATH' $root $attempts $TimeoutSeconds
        if ($selected) { return $selected }
    }
    foreach ($candidate in @('C:\Program Files\Calibre2\ebook-convert.exe', 'C:\Program Files (x86)\Calibre2\ebook-convert.exe',
                            [IO.Path]::Combine($env:LOCALAPPDATA, 'Programs\Calibre\ebook-convert.exe'))) {
        if ([IO.File]::Exists($candidate) -and (Test-WinBookSplitAutomaticPath $candidate $root $DocumentPath)) {
            $selected = Test-WinBookSplitConverterCandidate $candidate 'known_location' $root $attempts $TimeoutSeconds
            if ($selected) { return $selected }
        }
    }
    throw (New-WinBookSplitDependencyError 'converter_not_found' ('No compatible trusted converter was found. ' + $guidance) $attempts)
}
