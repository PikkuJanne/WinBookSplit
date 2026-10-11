# Offline compatibility validation. Dot-sourcing never probes or installs tools.
function Test-WinBookSplitCompatibilityInteger {
    param($Value)
    $integer = ($Value -is [int] -or $Value -is [long])
    if (Get-Command Test-SplitInteger -CommandType Function -ErrorAction SilentlyContinue) {
        $integer = Test-SplitInteger $Value
    }
    return ($integer -and $Value -ge 0 -and $Value -le 1000000)
}

function Get-WinBookSplitCompatibilityKeys {
    param($Object)
    if ($Object -is [Collections.IDictionary]) { return ,@($Object.Keys) }
    if ($Object -is [pscustomobject]) { return ,@($Object.PSObject.Properties.Name) }
    throw 'Compatibility policy requires JSON objects.'
}

function Get-WinBookSplitCompatibilityField {
    param($Object, [string]$Name)
    if ($Object -is [Collections.IDictionary]) { return ,$Object[$Name] }
    if ($Object -is [pscustomobject]) { return ,$Object.PSObject.Properties[$Name].Value }
    throw 'Compatibility policy requires JSON objects.'
}

function Assert-WinBookSplitCompatibilityKeys {
    param($Object, [string[]]$Keys, [string]$Label)
    $actual = Get-WinBookSplitCompatibilityKeys $Object
    if ($actual.Count -ne $Keys.Count) { throw ('Invalid ' + $Label + ' object or fields.') }
    foreach ($key in $Keys) {
        if ($actual -cnotcontains $key) { throw ('Invalid ' + $Label + ' object or fields.') }
    }
}

function Test-WinBookSplitCompatibilityText {
    param($Value, [int]$Maximum = 1024)
    return ($Value -is [string] -and $Value.Length -gt 0 -and $Value.Length -le $Maximum -and
            -not [string]::IsNullOrWhiteSpace($Value) -and $Value.IndexOf([char]0) -lt 0)
}

function ConvertTo-WinBookSplitCompatibilityVersion {
    param($Value)
    if ($Value -isnot [string] -or $Value -cnotmatch '\A(0|[1-9][0-9]{0,6})\.(0|[1-9][0-9]{0,6})\.(0|[1-9][0-9]{0,6})\z') { return $null }
    $parts = @([long]$Matches[1], [long]$Matches[2], [long]$Matches[3])
    foreach ($part in $parts) { if (-not (Test-WinBookSplitCompatibilityInteger $part)) { return $null } }
    return ,$parts
}

function Assert-WinBookSplitCompatibilityPolicy {
    param($Policy)
    Assert-WinBookSplitCompatibilityKeys $Policy @('schema_version','policy_version','reviewed_on','python','pypdf','calibre','excluded_python_branches','sources') 'policy'
    if (-not (Test-WinBookSplitCompatibilityInteger $Policy.schema_version) -or $Policy.schema_version -ne 1) { throw 'Unsupported compatibility policy schema.' }
    if (-not (Test-WinBookSplitCompatibilityInteger $Policy.policy_version) -or $Policy.policy_version -le 0) { throw 'Invalid compatibility policy version.' }
    $reviewed = [datetime]::MinValue
    if ($Policy.reviewed_on -isnot [string] -or $Policy.reviewed_on -cnotmatch '\A[0-9]{4}-[0-9]{2}-[0-9]{2}\z' -or
        -not [datetime]::TryParseExact($Policy.reviewed_on, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::None, [ref]$reviewed)) { throw 'Invalid compatibility policy review date.' }
    if ($Policy.sources -isnot [array] -or $Policy.sources.Count -lt 1 -or $Policy.sources.Count -gt 64) { throw 'Invalid policy sources.' }
    $sourceIds = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
    foreach ($source in $Policy.sources) {
        Assert-WinBookSplitCompatibilityKeys $source @('id','url','relevance') 'source'
        if ($source.id -isnot [string] -or $source.id -cnotmatch '\A[A-Za-z][A-Za-z0-9_-]{0,63}\z' -or -not $sourceIds.Add($source.id)) { throw 'Invalid or duplicate policy source ID.' }
        if (-not (Test-WinBookSplitCompatibilityText $source.url 2048) -or -not (Test-WinBookSplitCompatibilityText $source.relevance)) { throw 'Invalid policy source text.' }
        $uri = $null
        if (-not [Uri]::TryCreate($source.url, [UriKind]::Absolute, [ref]$uri) -or $uri.Scheme -cne 'https' -or
            [string]::IsNullOrWhiteSpace($uri.Host) -or $uri.UserInfo -ne '' -or $source.url -match '\Ahttps://[^/?#]*@' -or
            $source.url -match '\s') { throw 'Policy sources require absolute HTTPS URLs without credentials.' }
    }
    $python = $Policy.python
    Assert-WinBookSplitCompatibilityKeys $python @('implementation','platform','machine','bits','stable_only','gil_disabled','branches') 'Python policy'
    if ($python.implementation -isnot [string] -or $python.implementation -cne 'cpython' -or $python.platform -isnot [string] -or $python.platform -cne 'win32' -or
        $python.machine -isnot [string] -or $python.machine -cne 'AMD64' -or -not (Test-WinBookSplitCompatibilityInteger $python.bits) -or $python.bits -ne 64 -or
        $python.stable_only -isnot [bool] -or -not $python.stable_only -or $python.gil_disabled -isnot [bool] -or $python.gil_disabled) { throw 'Invalid supported Python identity.' }
    if ($python.branches -isnot [array] -or $python.branches.Count -lt 1 -or $python.branches.Count -gt 8) { throw 'Invalid Python branch list.' }
    $branchIds = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
    foreach ($branch in $python.branches) {
        Assert-WinBookSplitCompatibilityKeys $branch @('major','minor','compatibility_min_patch','security_min_patch','recommended_patch','blocked_patches','security_sources') 'Python branch'
        foreach ($name in @('major','minor','compatibility_min_patch','security_min_patch','recommended_patch')) {
            if (-not (Test-WinBookSplitCompatibilityInteger (Get-WinBookSplitCompatibilityField $branch $name))) { throw 'Python branch versions require bounded nonnegative integers.' }
        }
        $identity = [string]$branch.major + '.' + [string]$branch.minor
        if ($branch.major -ne 3 -or -not $branchIds.Add($identity)) { throw 'Invalid or duplicate Python branch.' }
        if ($branch.blocked_patches -isnot [array] -or $branch.blocked_patches.Count -gt 32) { throw 'Invalid or duplicate blocked Python patches.' }
        $blocked = New-Object 'Collections.Generic.HashSet[long]'
        foreach ($patch in $branch.blocked_patches) {
            if (-not (Test-WinBookSplitCompatibilityInteger $patch) -or -not $blocked.Add([long]$patch)) { throw 'Invalid or duplicate blocked Python patches.' }
        }
        if ($branch.recommended_patch -lt [Math]::Max([long]$branch.compatibility_min_patch, [long]$branch.security_min_patch) -or
            $blocked.Contains([long]$branch.recommended_patch)) { throw 'Recommended patch contradicts branch floors or exclusions.' }
        if ($branch.security_sources -isnot [array] -or $branch.security_sources.Count -lt 1 -or $branch.security_sources.Count -gt 64) { throw 'Missing, unknown or duplicate Python security source reference.' }
        $references = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
        foreach ($reference in $branch.security_sources) {
            if ($reference -isnot [string] -or -not $sourceIds.Contains($reference) -or -not $references.Add($reference)) { throw 'Missing, unknown or duplicate Python security source reference.' }
        }
    }
    if ($Policy.excluded_python_branches -isnot [array] -or $Policy.excluded_python_branches.Count -gt 16) { throw 'Invalid excluded Python branches.' }
    $excludedIds = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
    foreach ($branch in $Policy.excluded_python_branches) {
        Assert-WinBookSplitCompatibilityKeys $branch @('major','minor','reason') 'excluded Python branch'
        if (-not (Test-WinBookSplitCompatibilityInteger $branch.major) -or -not (Test-WinBookSplitCompatibilityInteger $branch.minor) -or
            -not (Test-WinBookSplitCompatibilityText $branch.reason)) { throw 'Invalid excluded Python branch identity or reason.' }
        $identity = [string]$branch.major + '.' + [string]$branch.minor
        if ($branchIds.Contains($identity) -or -not $excludedIds.Add($identity)) { throw 'Duplicate or supported-and-excluded Python branch.' }
    }
    Assert-WinBookSplitCompatibilityKeys $Policy.pypdf @('version','wheel_sha256') 'pypdf policy'
    if ($null -eq (ConvertTo-WinBookSplitCompatibilityVersion $Policy.pypdf.version)) { throw 'Invalid locked pypdf version.' }
    if ($Policy.pypdf.wheel_sha256 -isnot [string] -or $Policy.pypdf.wheel_sha256 -cnotmatch '\A[0-9a-f]{64}\z') { throw 'Invalid locked pypdf wheel SHA256.' }
    Assert-WinBookSplitCompatibilityKeys $Policy.calibre @('version') 'Calibre policy'
    if ($null -eq (ConvertTo-WinBookSplitCompatibilityVersion $Policy.calibre.version)) { throw 'Invalid locked Calibre version.' }
}

function Get-WinBookSplitCompatibilityPolicy {
    param([string]$PolicyPath = (Join-Path $PSScriptRoot 'WinBookSplit.Compatibility.json'), [switch]$WithIdentity)
    $stream = $null
    try {
        if (-not ('WinBookSplitSupport.StrictJson' -as [type])) {
            if (-not (Get-Command Initialize-WinBookSplitSupport -CommandType Function -ErrorAction SilentlyContinue)) { . (Join-Path $PSScriptRoot 'WinBookSplit.Support.ps1') }
            Initialize-WinBookSplitSupport
        }
        $stream = New-Object IO.FileStream($PolicyPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
        if ($stream.Length -lt 1 -or $stream.Length -gt 32768) { throw 'Compatibility policy is empty or exceeds 32 KiB.' }
        $bytes = New-Object byte[] ([int]$stream.Length)
        $offset = 0
        while ($offset -lt $bytes.Length) {
            $count = $stream.Read($bytes, $offset, $bytes.Length - $offset)
            if ($count -eq 0) { throw 'Compatibility policy read incomplete.' }
            $offset += $count
        }
        if ($stream.ReadByte() -ne -1) { throw 'Compatibility policy read incomplete.' }
        $text = (New-Object Text.UTF8Encoding($false, $true)).GetString($bytes)
        $policy = [WinBookSplitSupport.StrictJson]::Parse($text)
        Assert-WinBookSplitCompatibilityPolicy $policy
        if ($WithIdentity) {
            $algorithm = [Security.Cryptography.SHA256]::Create()
            try { $hash = [BitConverter]::ToString($algorithm.ComputeHash($bytes)).Replace('-', '').ToLowerInvariant() }
            finally { $algorithm.Dispose() }
            return [pscustomobject]@{ Policy=$policy; Sha256=$hash }
        }
        return ,$policy
    } catch {
        $failure = New-Object IO.InvalidDataException('Missing or malformed shipped compatibility policy: ' + $_.Exception.Message)
        $failure.Data['Reason'] = 'malformed_policy'
        throw $failure
    } finally { if ($null -ne $stream) { $stream.Dispose() } }
}

function New-WinBookSplitCompatibilityResult {
    param([bool]$Ok, [string]$Reason, [string]$Message)
    return [pscustomobject]@{ Ok=$Ok; Reason=$Reason; Message=$Message }
}

function Test-WinBookSplitPythonVersion {
    param($Version, $Policy)
    try { Assert-WinBookSplitCompatibilityPolicy $Policy }
    catch { return (New-WinBookSplitCompatibilityResult $false 'malformed_policy' $_.Exception.Message) }
    $parts = ConvertTo-WinBookSplitCompatibilityVersion $Version
    if ($null -eq $parts) {
        if ($Version -is [string] -and $Version -cmatch '\A(?:0|[1-9][0-9]{0,6})\.(?:0|[1-9][0-9]{0,6})\.(?:0|[1-9][0-9]{0,6})(?:(?:a|b|rc)[0-9]+|\.dev[0-9]+|\+)\z') {
            return (New-WinBookSplitCompatibilityResult $false 'unstable_python' 'A stable final Python release is required.')
        }
        return (New-WinBookSplitCompatibilityResult $false 'malformed_runtime' 'Runtime version must contain three bounded integers.')
    }
    $branch = $null
    foreach ($candidate in $Policy.python.branches) { if ($candidate.major -eq $parts[0] -and $candidate.minor -eq $parts[1]) { $branch=$candidate; break } }
    $observed = $parts -join '.'
    if ($null -eq $branch) {
        $detail = 'This Python minor branch is outside the reviewed compatibility policy.'
        foreach ($exclusion in $Policy.excluded_python_branches) { if ($exclusion.major -eq $parts[0] -and $exclusion.minor -eq $parts[1]) { $detail=$exclusion.reason; break } }
        return (New-WinBookSplitCompatibilityResult $false 'unsupported_python' ('Python ' + $observed + ': ' + $detail))
    }
    $branchName = [string]$branch.major + '.' + [string]$branch.minor
    if ($parts[2] -lt $branch.security_min_patch -or $branch.blocked_patches -contains $parts[2]) {
        return (New-WinBookSplitCompatibilityResult $false 'python_security_floor' ('Python ' + $observed + ' is excluded by the reviewed security policy; the branch security floor is ' + $branchName + '.' + $branch.security_min_patch + '. See ' + ($branch.security_sources -join ', ') + '.'))
    }
    if ($parts[2] -lt $branch.compatibility_min_patch) {
        return (New-WinBookSplitCompatibilityResult $false 'python_compatibility_floor' ('Python ' + $observed + ' is below compatibility floor ' + $branchName + '.' + $branch.compatibility_min_patch + '.'))
    }
    return (New-WinBookSplitCompatibilityResult $true 'supported' ('Python ' + $observed + ' is accepted by compatibility policy ' + $Policy.policy_version + '.'))
}

function Test-WinBookSplitInterpreterRecord {
    param($Record, $Policy)
    try { Assert-WinBookSplitCompatibilityPolicy $Policy }
    catch { return (New-WinBookSplitCompatibilityResult $false 'malformed_policy' $_.Exception.Message) }
    try { $keys = Get-WinBookSplitCompatibilityKeys $Record }
    catch { return (New-WinBookSplitCompatibilityResult $false 'malformed_runtime' 'Runtime identity is missing required fields.') }
    foreach ($name in @('version','version_info','releaselevel','serial','implementation','platform','machine','bits','gil_disabled','isolated','dont_write_bytecode','executable')) {
        if ($keys -cnotcontains $name) { return (New-WinBookSplitCompatibilityResult $false 'malformed_runtime' 'Runtime identity is missing required fields.') }
    }
    $valid = ($Record.version_info -is [array] -and $Record.version_info.Count -eq 3 -and
              (Test-WinBookSplitCompatibilityInteger $Record.serial) -and (Test-WinBookSplitCompatibilityInteger $Record.bits))
    if ($valid) { foreach ($part in $Record.version_info) { if (-not (Test-WinBookSplitCompatibilityInteger $part)) { $valid=$false } } }
    foreach ($name in @('gil_disabled','isolated','dont_write_bytecode')) { if ((Get-WinBookSplitCompatibilityField $Record $name) -isnot [bool]) { $valid=$false } }
    foreach ($name in @('implementation','platform','machine','releaselevel','executable')) {
        $limit = 128; if ($name -ceq 'executable') { $limit=32768 }
        if (-not (Test-WinBookSplitCompatibilityText (Get-WinBookSplitCompatibilityField $Record $name) $limit)) { $valid=$false }
    }
    if (-not $valid -or $Record.executable -cnotmatch '\A(?:[A-Za-z]:[\\/]|\\\\[^\\]+\\[^\\]+\\)') { return (New-WinBookSplitCompatibilityResult $false 'malformed_runtime' 'Runtime identity has malformed scalar fields.') }
    if ($Record.releaselevel -cnotin @('final','alpha','beta','candidate')) { return (New-WinBookSplitCompatibilityResult $false 'malformed_runtime' 'Runtime release level is malformed.') }
    $parts = ConvertTo-WinBookSplitCompatibilityVersion $Record.version
    if ($null -eq $parts) { return (Test-WinBookSplitPythonVersion $Record.version $Policy) }
    for ($index=0; $index -lt 3; $index++) { if ($parts[$index] -ne $Record.version_info[$index]) { return (New-WinBookSplitCompatibilityResult $false 'malformed_runtime' 'Runtime version and version_info disagree.') } }
    if ($Record.releaselevel -cne 'final' -or $Record.serial -ne 0) { return (New-WinBookSplitCompatibilityResult $false 'unstable_python' 'A stable final Python release is required.') }
    if (-not $Record.isolated -or -not $Record.dont_write_bytecode) { return (New-WinBookSplitCompatibilityResult $false 'malformed_runtime' 'Isolated mode and no-bytecode mode are required.') }
    if ($Record.implementation -cne $Policy.python.implementation -or $Record.gil_disabled) { return (New-WinBookSplitCompatibilityResult $false 'unsupported_flavor' 'Regular GIL-enabled CPython is required.') }
    if ($Record.platform -cne $Policy.python.platform -or $Record.machine -cne $Policy.python.machine -or $Record.bits -ne $Policy.python.bits) { return (New-WinBookSplitCompatibilityResult $false 'unsupported_architecture' 'Regular Windows x64 Python is required.') }
    return (Test-WinBookSplitPythonVersion $Record.version $Policy)
}

function Get-WinBookSplitPythonPolicyDescription {
    param($Policy)
    Assert-WinBookSplitCompatibilityPolicy $Policy
    $branches = foreach ($branch in $Policy.python.branches) {
        [string]$branch.major + '.' + [string]$branch.minor + '.' + [Math]::Max([long]$branch.compatibility_min_patch, [long]$branch.security_min_patch) + '+'
    }
    return ('regular Windows x64 CPython ' + ($branches -join ' or ') + '; later stable patches accepted')
}
