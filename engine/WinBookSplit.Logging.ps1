# Local run records only. No upload, environment dump or source-file writes.
function Initialize-WinBookSplitLogTypes {
    if ('WinBookSplitLogging.BoundedWriter' -as [type]) { return }
    Add-Type -TypeDefinition @'
using System;
using System.IO;
using System.Text;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
namespace WinBookSplitLogging {
    public sealed class PathLease : IDisposable {
        [StructLayout(LayoutKind.Sequential)]
        private struct Information {
            public uint Attributes;
            public System.Runtime.InteropServices.ComTypes.FILETIME Creation, Access, Write;
            public uint Volume, SizeHigh, SizeLow, Links, IndexHigh, IndexLow;
        }
        [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        private static extern SafeFileHandle CreateFileW(string path, uint access, uint sharing,
            IntPtr security, uint creation, uint flags, IntPtr template);
        [DllImport("kernel32.dll", SetLastError=true)]
        private static extern bool GetFileInformationByHandle(SafeFileHandle handle, out Information info);
        [DllImport("kernel32.dll", SetLastError=true)]
        private static extern bool SetFileInformationByHandle(SafeFileHandle handle, int kind, IntPtr info, uint size);
        [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)]
        private struct RenameInformation { public uint Replace; public IntPtr Root; public uint Length; public char Name; }
        private SafeFileHandle handle;
        private string path;
        private readonly bool directory;
        private readonly bool deleteAccess;
        private readonly uint access, sharing;
        private readonly Information identity;
        public PathLease(string path, bool directory) : this(path, directory, false, false) { }
        public PathLease(string path, bool directory, bool deleteAccess, bool strict) {
            this.path = Path.GetFullPath(path); this.directory = directory;
            this.deleteAccess = deleteAccess;
            // Metadata-only opens do not participate in Windows share checks.
            // Retain actual read access so exclusion of DELETE is enforced.
            access = 0x80000000u | (deleteAccess ? 0x10000u : 0u);
            sharing = strict ? 1u : 3u;
            handle = Open(this.path, access, sharing);
            try { identity = Read(handle); AssertUnchanged(); }
            catch { handle.Dispose(); throw; }
        }
        private SafeFileHandle Open(string value, uint desiredAccess, uint shareMode) {
            var result = CreateFileW(value, desiredAccess, shareMode, IntPtr.Zero, 3, 0x02200000, IntPtr.Zero);
            if (result.IsInvalid) { result.Dispose(); throw new Win32Exception(Marshal.GetLastWin32Error()); }
            return result;
        }
        private Information Read(SafeFileHandle value) {
            Information info;
            if (!GetFileInformationByHandle(value, out info)) throw new Win32Exception(Marshal.GetLastWin32Error());
            if ((info.Attributes & 0x400) != 0 || ((info.Attributes & 0x10) != 0) != directory ||
                (!directory && info.Links != 1)) throw new IOException("Run records require ordinary, unlinked files and directories.");
            return info;
        }
        public void AssertUnchanged() {
            if (handle == null || handle.IsClosed) throw new IOException("The run-record path lease is closed.");
            var held = Read(handle);
            using (var current = Open(path, 0x80, 7)) {
                var named = Read(current);
                if (held.Volume != identity.Volume || held.IndexHigh != identity.IndexHigh || held.IndexLow != identity.IndexLow ||
                    named.Volume != identity.Volume || named.IndexHigh != identity.IndexHigh || named.IndexLow != identity.IndexLow)
                    throw new IOException("The run-record path identity changed.");
            }
        }
        public void AssertSameHandle(SafeFileHandle value) {
            var observed = Read(value);
            if (observed.Volume != identity.Volume || observed.IndexHigh != identity.IndexHigh || observed.IndexLow != identity.IndexLow)
                throw new IOException("The created record no longer names its held file.");
        }
        public void AssertSameLease(PathLease other) {
            AssertUnchanged(); other.AssertUnchanged();
            if (other.identity.Volume != identity.Volume || other.identity.IndexHigh != identity.IndexHigh || other.identity.IndexLow != identity.IndexLow)
                throw new IOException("The run record changed while acquiring its read-only seal.");
        }
        public void PublishTo(string target) {
            target = Path.GetFullPath(target);
            if (!deleteAccess || directory || !String.Equals(Path.GetDirectoryName(path), Path.GetDirectoryName(target), StringComparison.OrdinalIgnoreCase))
                throw new IOException("Record publication requires its held owned sibling file.");
            AssertUnchanged();
            var name = Encoding.Unicode.GetBytes(target);
            int offset = Marshal.OffsetOf(typeof(RenameInformation), "Name").ToInt32();
            // Keep the counted UTF-16 name and an explicit zero terminator;
            // Windows also reads the absolute name as a terminated string.
            int size = checked(offset + name.Length + 2);
            var buffer = Marshal.AllocHGlobal(size);
            try {
                Marshal.Copy(new byte[size], 0, buffer, size);
                var info = new RenameInformation { Replace=0, Root=IntPtr.Zero, Length=(uint)name.Length, Name='\0' };
                Marshal.StructureToPtr(info, buffer, false); Marshal.Copy(name, 0, IntPtr.Add(buffer, offset), name.Length);
                if (!SetFileInformationByHandle(handle, 3, buffer, (uint)size)) throw new Win32Exception(Marshal.GetLastWin32Error());
                path = target;
            }
            finally { Marshal.FreeHGlobal(buffer); }
        }
        public void MarkIncomplete(byte[] data) {
            if (!deleteAccess || directory) throw new IOException("Incomplete evidence requires its owned pending file.");
            AssertUnchanged();
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Write, FileShare.Read | FileShare.Delete)) {
                AssertSameHandle(stream.SafeFileHandle); stream.SetLength(0);
                stream.Write(data, 0, data.Length); stream.Flush(true);
            }
        }
        public void Dispose() { if (handle != null) { handle.Dispose(); handle = null; } }
    }
    public sealed class BoundedWriter : TextWriter {
        private readonly Stream stream;
        private readonly string path;
        private readonly UTF8Encoding utf8 = new UTF8Encoding(false, true);
        private readonly PathLease fileLease;
        private bool closed;
        public long BytesWritten { get; private set; }
        public long BodyLimitBytes { get; private set; }
        public long MaxBytes { get; private set; }
        public bool LimitReached { get; private set; }
        public override Encoding Encoding { get { return utf8; } }
        public BoundedWriter(Stream stream, string path, long bodyLimit, long finalReserve) {
            if (bodyLimit < 1 || finalReserve < 1) throw new ArgumentOutOfRangeException("bodyLimit");
            this.stream = stream; this.path = path;
            BodyLimitBytes = bodyLimit; MaxBytes = checked(bodyLimit + finalReserve);
            if (!String.IsNullOrEmpty(path)) {
                fileLease = new PathLease(path, false);
                fileLease.AssertSameHandle(((FileStream)stream).SafeFileHandle);
            }
        }
        private void Put(string value, long limit) {
            if (closed) throw new ObjectDisposedException("BoundedWriter");
            var data = utf8.GetBytes(value ?? "");
            if (data.Length > limit - BytesWritten) {
                LimitReached = true;
                throw new IOException("The bounded console log limit was reached; further records were refused.");
            }
            if (fileLease != null) fileLease.AssertUnchanged();
            stream.Write(data, 0, data.Length); BytesWritten += data.Length; stream.Flush();
        }
        public override void Write(string value) { Put(value, BodyLimitBytes); }
        public override void WriteLine(string value) { Put((value ?? "") + "\r\n", BodyLimitBytes); }
        public override void WriteLine() { Put("\r\n", BodyLimitBytes); }
        public void WriteFinal(string value) { Put((value ?? "") + "\r\n", MaxBytes); }
        public override void Flush() { if (!closed) stream.Flush(); }
        public void AppendCorrection(string value) {
            if (path == null) return;
            if (!closed) { WriteFinal(value); return; }
            fileLease.AssertUnchanged();
            using (var target = new FileStream(path, FileMode.Open, FileAccess.Write, FileShare.Read)) {
                if (target.Length != BytesWritten) throw new IOException("The closed log length changed; correction was refused.");
                var data = utf8.GetBytes(value + "\r\n");
                if (data.Length > MaxBytes - BytesWritten) throw new IOException("The final log reserve was exhausted.");
                target.Seek(0, SeekOrigin.End); target.Write(data, 0, data.Length); target.Flush(true);
                BytesWritten += data.Length;
            }
        }
        protected override void Dispose(bool disposing) {
            if (disposing && !closed) { closed = true; stream.Dispose(); }
            base.Dispose(disposing);
        }
        public void CloseLease() { if (fileLease != null) fileLease.Dispose(); }
    }
}
'@
}

function New-WinBookSplitLogWriter {
    param([IO.Stream]$Stream, [AllowNull()][string]$Path,
        [long]$BodyLimitBytes = 33554432, [long]$FinalReserveBytes = 16777216)
    Initialize-WinBookSplitLogTypes
    return [WinBookSplitLogging.BoundedWriter]::new($Stream, $Path, $BodyLimitBytes, $FinalReserveBytes)
}

function New-WinBookSplitRecordLeases {
    param([string]$Directory, [bool]$IncludeMarker = $true)
    Initialize-WinBookSplitLogTypes
    $paths = [Collections.Generic.List[string]]::new()
    for ($entry = [IO.DirectoryInfo]::new($Directory); $null -ne $entry; $entry = $entry.Parent) {
        $paths.Insert(0, $entry.FullName)
    }
    $leases = [Collections.Generic.List[object]]::new()
    try {
        foreach ($path in $paths) { $leases.Add([WinBookSplitLogging.PathLease]::new($path, $true)) }
        if ($IncludeMarker) { $leases.Add([WinBookSplitLogging.PathLease]::new([IO.Path]::Combine($Directory, '.WinBookSplit-console-owner.json'), $false, $false, $true)) }
        return ,$leases
    }
    catch { foreach ($lease in $leases) { $lease.Dispose() }; throw }
}

function Write-WinBookSplitRunManifest {
    param([string]$Directory, $Manifest, $Leases)
    # Publication happens only after UTF-8 write/flush/close. A pending file is
    # incomplete evidence, never a finalized success manifest. No replacement.
    $bytes = [Text.UTF8Encoding]::new($false, $true).GetBytes(($Manifest | ConvertTo-Json -Depth 100 -Compress) + "`n")
    if ($bytes.Length -gt 33554432) { throw 'The run manifest exceeds its 32 MiB bound.' }
    foreach ($lease in $Leases) { $lease.AssertUnchanged() }
    $pending = [IO.Path]::Combine($Directory, 'WinBookSplit_Run.pending.json')
    $target = [IO.Path]::Combine($Directory, 'WinBookSplit_Run.json')
    $stream = [IO.File]::Open($pending, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, ([IO.FileShare]::Read -bor [IO.FileShare]::Delete))
    $pendingLease = $null
    try {
        $pendingLease = [WinBookSplitLogging.PathLease]::new($pending, $false, $true, $false)
        $pendingLease.AssertSameHandle($stream.SafeFileHandle)
        $stream.Write($bytes, 0, $bytes.Length); $stream.Flush($true); $stream.Dispose(); $stream = $null
        foreach ($lease in $Leases) { $lease.AssertUnchanged() }
        $pendingLease.PublishTo($target)
    }
    catch {
        $primaryError = $_
        if ($null -ne $pendingLease) {
            try {
                if ($null -ne $stream) { $stream.Dispose(); $stream = $null }
                $incomplete = $Manifest | ConvertTo-Json -Depth 100 -Compress | ConvertFrom-Json
                $incomplete | Add-Member -MemberType NoteProperty -Name diagnostics_finalized -Value $false -Force
                $pendingLease.MarkIncomplete([Text.UTF8Encoding]::new($false, $true).GetBytes(($incomplete | ConvertTo-Json -Depth 100 -Compress) + "`n"))
            }
            catch { throw ('Run manifest finalization failed: ' + $primaryError.Exception.Message + '; pending record could not be marked incomplete: ' + $_.Exception.Message) }
        }
        throw $primaryError
    }
    finally { if ($null -ne $stream) { $stream.Dispose() }; if ($null -ne $pendingLease) { $pendingLease.Dispose() } }
}
