# Literal FileSystem validation. Dot-sourcing does not read documents or create output.
function Resolve-WinBookSplitInput {
    param([AllowEmptyString()][string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { throw 'Choose one existing PDF, EPUB, or AZW3 file.' }
    $resolved = Resolve-Path -LiteralPath $Path -ErrorAction Stop
    if ($resolved.Provider.Name -ne 'FileSystem') { throw 'The input must be a FileSystem file.' }
    $item = Get-Item -LiteralPath $resolved.ProviderPath -Force -ErrorAction Stop
    if ($item -isnot [IO.FileInfo]) { throw 'The input must be a file, not a directory.' }
    if ($item.Extension.ToLowerInvariant() -notin @('.pdf', '.epub', '.azw3')) {
        throw 'Choose a PDF, EPUB, or AZW3 file.'
    }
    $stream = $null
    try {
        $stream = [IO.File]::Open($resolved.ProviderPath, [IO.FileMode]::Open, [IO.FileAccess]::Read,
            ([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
    }
    catch { throw ('Cannot read the input file: ' + $_.Exception.Message) }
    finally { if ($null -ne $stream) { $stream.Dispose() } }
    return $item
}

function Resolve-WinBookSplitOutputBase {
    param([AllowEmptyString()][string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { throw 'Choose an existing output base directory.' }
    $resolved = Resolve-Path -LiteralPath $Path -ErrorAction Stop
    if ($resolved.Provider.Name -ne 'FileSystem') { throw 'The output base must be a FileSystem directory.' }
    $item = Get-Item -LiteralPath $resolved.ProviderPath -Force -ErrorAction Stop
    if ($item -isnot [IO.DirectoryInfo]) { throw 'The output base must be an existing directory.' }
    # Reject already observed junction/symlink ancestors before creating console files.
    for ($directory = $item; $null -ne $directory; $directory = $directory.Parent) {
        $observed = Get-Item -LiteralPath $directory.FullName -Force -ErrorAction Stop
        if (($observed.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw 'The output base and its ancestors must not be junctions or symbolic links.'
        }
    }
    return $item.FullName
}

function New-WinBookSplitConsoleDirectory {
    param([string]$Base)
    # String.Length counts UTF-16 units on both supported hosts. Stay inside the
    # conservative Win32 file/directory limits without changing machine policy.
    $example = [IO.Path]::Combine($Base, ('.WinBookSplit-console-' + ('0' * 32)))
    if ($example.Length -gt 247 -or
        [IO.Path]::Combine($example, '.WinBookSplit-console-owner.json').Length -gt 259) {
        throw 'The output base is too long for console records. Choose a shorter output base.'
    }
    if (-not ('WinBookSplitPaths.NativeDirectory' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
namespace WinBookSplitPaths {
    public static class NativeDirectory {
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool CreateDirectoryW(string path, IntPtr security);
        public static void CreateExclusive(string path) {
            if (!CreateDirectoryW(path, IntPtr.Zero))
                throw new Win32Exception(Marshal.GetLastWin32Error());
        }
    }
}
'@
    }
    for ($attempt = 0; $attempt -lt 100; $attempt++) {
        $id = [guid]::NewGuid().ToString('N')
        $path = [IO.Path]::Combine($Base, ('.WinBookSplit-console-' + $id))
        try {
            [WinBookSplitPaths.NativeDirectory]::CreateExclusive($path)
            return [pscustomobject]@{ run_id = $id; path = $path }
        }
        catch {
            $nativeError = $_.Exception
            while ($null -ne $nativeError.InnerException) { $nativeError = $nativeError.InnerException }
            if ($nativeError -isnot [ComponentModel.Win32Exception] -or $nativeError.NativeErrorCode -ne 183) {
                throw ('Cannot create a console directory in the output base: ' + $nativeError.Message)
            }
        }
    }
    throw 'Cannot reserve a unique console directory. Choose another output base.'
}
