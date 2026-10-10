# Explicit local support export. Free-text diagnostics never enter the summary.
function Initialize-WinBookSplitSupport {
    if ('WinBookSplitSupport.NativeFiles' -as [type]) { return }
    Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32.SafeHandles;
namespace WinBookSplitSupport {
    public static class StrictJson {
        sealed class Reader {
            readonly string text; int at, nodes;
            internal Reader(string value) { text = value; }
            void Space() { while (at < text.Length && (text[at] == ' ' || text[at] == '\r' || text[at] == '\n' || text[at] == '\t')) at++; }
            void Need(bool condition) { if (!condition) throw new InvalidDataException("Invalid diagnostic JSON."); }
            char Take() { Need(at < text.Length); return text[at++]; }
            string String() {
                Need(Take() == '"'); StringBuilder result = new StringBuilder();
                while (true) {
                    char value = Take(); if (value == '"') break;
                    Need(value >= 32);
                    if (value == '\\') {
                        char escaped = Take();
                        if (escaped == '"' || escaped == '\\' || escaped == '/') value = escaped;
                        else if (escaped == 'b') value = '\b'; else if (escaped == 'f') value = '\f';
                        else if (escaped == 'n') value = '\n'; else if (escaped == 'r') value = '\r'; else if (escaped == 't') value = '\t';
                        else if (escaped == 'u') {
                            Need(at + 4 <= text.Length); int number;
                            Need(Int32.TryParse(text.Substring(at, 4), NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out number));
                            value = (char)number; at += 4;
                        } else throw new InvalidDataException("Invalid diagnostic JSON.");
                    }
                    result.Append(value);
                }
                string answer = result.ToString();
                for (int index = 0; index < answer.Length; index++) {
                    if (Char.IsHighSurrogate(answer[index])) { Need(index + 1 < answer.Length && Char.IsLowSurrogate(answer[index + 1])); index++; }
                    else Need(!Char.IsLowSurrogate(answer[index]));
                }
                return answer;
            }
            object Value(int depth) {
                Need(depth <= 32 && ++nodes <= 500000); Space(); Need(at < text.Length); char first = text[at];
                if (first == '"') return String();
                if (first == '{') {
                    at++; Space(); Dictionary<string, object> result = new Dictionary<string, object>(StringComparer.Ordinal);
                    HashSet<string> keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    if (at < text.Length && text[at] == '}') { at++; return result; }
                    while (true) {
                        Space(); string key = String(); Need(keys.Add(key)); Space(); Need(Take() == ':'); result.Add(key, Value(depth + 1));
                        Space(); char separator = Take(); if (separator == '}') return result; Need(separator == ',');
                    }
                }
                if (first == '[') {
                    at++; Space(); List<object> result = new List<object>();
                    if (at < text.Length && text[at] == ']') { at++; return result.ToArray(); }
                    while (true) { result.Add(Value(depth + 1)); Space(); char separator = Take(); if (separator == ']') return result.ToArray(); Need(separator == ','); }
                }
                foreach (string literal in new string[] { "true", "false", "null" }) {
                    if (at + literal.Length <= text.Length && text.Substring(at, literal.Length) == literal) {
                        at += literal.Length; if (literal == "null") return null; return literal == "true";
                    }
                }
                int start = at; if (text[at] == '-') at++; Need(at < text.Length);
                if (text[at] == '0') at++; else { Need(text[at] >= '1' && text[at] <= '9'); while (at < text.Length && text[at] >= '0' && text[at] <= '9') at++; }
                bool integer = true;
                if (at < text.Length && text[at] == '.') { integer = false; at++; int begin = at; while (at < text.Length && text[at] >= '0' && text[at] <= '9') at++; Need(at > begin); }
                if (at < text.Length && (text[at] == 'e' || text[at] == 'E')) { integer = false; at++; if (at < text.Length && (text[at] == '+' || text[at] == '-')) at++; int begin = at; while (at < text.Length && text[at] >= '0' && text[at] <= '9') at++; Need(at > begin); }
                string numeric = text.Substring(start, at - start);
                if (integer) { long number; Need(Int64.TryParse(numeric, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out number)); return number; }
                decimal fraction; Need(Decimal.TryParse(numeric, NumberStyles.Float, CultureInfo.InvariantCulture, out fraction)); return fraction;
            }
            internal object Read() { object result = Value(0); Space(); Need(at == text.Length); return result; }
        }
        public static object Parse(string text) { return new Reader(text).Read(); }
        public static object Read(Stream stream) {
            if (stream.Length < 2 || stream.Length > 33554432) throw new InvalidDataException("Diagnostic manifest size invalid.");
            byte[] bytes = new byte[(int)stream.Length]; int offset = 0, count;
            while (offset < bytes.Length && (count = stream.Read(bytes, offset, bytes.Length - offset)) != 0) offset += count;
            if (offset != bytes.Length || stream.ReadByte() != -1) throw new InvalidDataException("Diagnostic manifest read incomplete.");
            int begin = bytes.Length >= 3 && bytes[0] == 239 && bytes[1] == 187 && bytes[2] == 191 ? 3 : 0;
            return Parse(new UTF8Encoding(false, true).GetString(bytes, begin, bytes.Length - begin));
        }
    }
    public sealed class NativeFiles : IDisposable {
        [StructLayout(LayoutKind.Sequential)] struct Info { public uint Attributes; public System.Runtime.InteropServices.ComTypes.FILETIME Creation, Access, Write; public uint Volume, SizeHigh, SizeLow, Links, IndexHigh, IndexLow; }
        [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)] static extern SafeFileHandle CreateFileW(string path, uint access, uint share, IntPtr security, uint creation, uint flags, IntPtr template);
        [DllImport("kernel32.dll", SetLastError=true)] static extern bool GetFileInformationByHandle(SafeFileHandle file, out Info info);
        [DllImport("kernel32.dll", SetLastError=true)] static extern uint GetFileType(SafeFileHandle file);
        sealed class DirectoryLease { internal string Path; internal SafeFileHandle Handle; internal Info Identity; }
        readonly List<DirectoryLease> directories = new List<DirectoryLease>();
        readonly HashSet<string> guarded = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public readonly string Source, Destination;
        static string Normalize(string path) {
            if (String.IsNullOrWhiteSpace(path) || path.Length < 4 || !Char.IsLetter(path[0]) || path[1] != ':' || (path[2] != '\\' && path[2] != '/')) throw new InvalidDataException("Use absolute local file paths.");
            string value = Path.GetFullPath(path).Replace('/', '\\');
            if (value.Length > 259 || value.Substring(2).IndexOf(':') >= 0) throw new InvalidDataException("Unsafe diagnostic file path.");
            foreach (string part in value.Substring(3).Split('\\')) {
                if (part.Length == 0 || part.EndsWith(".", StringComparison.Ordinal) || part.EndsWith(" ", StringComparison.Ordinal)) throw new InvalidDataException("Unsafe diagnostic file path.");
                string stem = part.Split('.')[0].ToUpperInvariant();
                if (stem == "CON" || stem == "PRN" || stem == "AUX" || stem == "NUL" || (stem.Length == 4 && (stem.StartsWith("COM") || stem.StartsWith("LPT")) && stem[3] >= '1' && stem[3] <= '9')) throw new InvalidDataException("Unsafe diagnostic file path.");
            }
            return value;
        }
        void Ancestors(string path) {
            Stack<string> pending = new Stack<string>(); string current = Path.GetDirectoryName(path);
            while (current != null) { pending.Push(current); DirectoryInfo parent = Directory.GetParent(current); current = parent == null ? null : parent.FullName; }
            while (pending.Count != 0) {
                string directory = pending.Pop(); if (!guarded.Add(directory)) continue;
                // Real read access makes exclusion of DELETE participate in Windows share checks.
                SafeFileHandle handle = CreateFileW(directory, 0x80000000, 3, IntPtr.Zero, 3, 0x02200000, IntPtr.Zero); Info info;
                if (handle.IsInvalid || GetFileType(handle) != 1 || !GetFileInformationByHandle(handle, out info) || (info.Attributes & 16) == 0 || (info.Attributes & 1024) != 0) { handle.Dispose(); throw new IOException("Unsafe diagnostic directory."); }
                directories.Add(new DirectoryLease { Path=directory, Handle=handle, Identity=info });
            }
        }
        static bool Same(Info left, Info right) { return left.Volume == right.Volume && left.IndexHigh == right.IndexHigh && left.IndexLow == right.IndexLow; }
        void AssertAncestors() {
            foreach (DirectoryLease lease in directories) {
                Info held, named;
                if (lease.Handle.IsClosed || !GetFileInformationByHandle(lease.Handle, out held) || !Same(held, lease.Identity) || (held.Attributes & 16) == 0 || (held.Attributes & 1024) != 0) throw new IOException("Diagnostic directory identity changed.");
                using (SafeFileHandle current = CreateFileW(lease.Path, 128, 7, IntPtr.Zero, 3, 0x02200000, IntPtr.Zero)) {
                    if (current.IsInvalid || GetFileType(current) != 1 || !GetFileInformationByHandle(current, out named) || (named.Attributes & 16) == 0 || (named.Attributes & 1024) != 0 || !Same(named, lease.Identity)) throw new IOException("Diagnostic directory identity changed.");
                }
            }
        }
        public NativeFiles(string source, string destination) {
            Source = Normalize(source); Destination = Normalize(destination);
            if (String.Equals(Source, Destination, StringComparison.OrdinalIgnoreCase)) throw new IOException("Diagnostic destination aliases source.");
            try { Ancestors(Source); Ancestors(Destination); } catch { Dispose(); throw; }
        }
        public FileStream OpenSource() {
            AssertAncestors();
            SafeFileHandle handle = CreateFileW(Source, 0x80000000, 1, IntPtr.Zero, 3, 0x00200000, IntPtr.Zero); Info info;
            if (handle.IsInvalid || GetFileType(handle) != 1 || !GetFileInformationByHandle(handle, out info) || (info.Attributes & (16 | 1024)) != 0 || info.Links != 1) { handle.Dispose(); throw new IOException("Unsafe diagnostic source."); }
            return new FileStream(handle, FileAccess.Read);
        }
        public void WriteNew(byte[] bytes) {
            if (bytes == null || bytes.Length > 1048576) throw new InvalidDataException("Support summary size invalid.");
            AssertAncestors();
            SafeFileHandle handle = CreateFileW(Destination, 0x40000000, 0, IntPtr.Zero, 1, 0x00200080, IntPtr.Zero); Info info;
            if (handle.IsInvalid || GetFileType(handle) != 1 || !GetFileInformationByHandle(handle, out info) || (info.Attributes & (16 | 1024)) != 0) { handle.Dispose(); throw new IOException("Cannot create diagnostic export."); }
            using (FileStream stream = new FileStream(handle, FileAccess.Write)) { stream.Write(bytes, 0, bytes.Length); stream.Flush(true); }
        }
        public void Dispose() { for (int index = directories.Count - 1; index >= 0; index--) directories[index].Handle.Dispose(); directories.Clear(); }
    }
}
'@
}

function Get-WinBookSplitSupportField {
    param($Object, [string]$Name)
    if ($Object -isnot [Collections.IDictionary] -or -not $Object.ContainsKey($Name)) { throw 'Invalid diagnostic manifest field.' }
    return ,$Object[$Name]
}

function Assert-WinBookSplitSupportKeys {
    param($Object, [string[]]$Required, [string[]]$Allowed)
    if ($Object -isnot [Collections.IDictionary]) { throw 'Invalid diagnostic manifest object.' }
    foreach ($name in $Required) { if (-not $Object.ContainsKey($name)) { throw 'Missing diagnostic manifest field.' } }
    foreach ($name in $Object.Keys) { if ($name -cnotin $Allowed) { throw 'Unknown diagnostic manifest field.' } }
}

function Get-WinBookSplitSupportInteger {
    param($Value, [long]$Maximum, [long]$Minimum = 0)
    if ($Value -isnot [long] -and $Value -isnot [int]) { throw 'Invalid diagnostic number.' }
    if ($Value -lt $Minimum -or $Value -gt $Maximum) { throw 'Diagnostic number outside supported bounds.' }
    return [long]$Value
}

function Get-WinBookSplitSupportToken {
    param($Value, [string[]]$Allowed)
    if ($Value -isnot [string] -or $Value -cnotin $Allowed) { throw 'Invalid diagnostic token.' }
    return [string]$Value
}

function ConvertTo-WinBookSplitSupportSummary {
    param([Parameter(Mandatory = $true)]$Manifest)
    Initialize-WinBookSplitSupport
    $rootKeys = @('protocol','version','run_id','application_version','started_utc','finished_utc','diagnostics_finalized','runtime_versions','settings','source_identity','plan','engine_result','outcome','warnings','log')
    Assert-WinBookSplitSupportKeys $Manifest $rootKeys $rootKeys
    if ((Get-WinBookSplitSupportToken (Get-WinBookSplitSupportField $Manifest 'protocol') @('winbooksplit.run')) -cne 'winbooksplit.run' -or
        (Get-WinBookSplitSupportInteger (Get-WinBookSplitSupportField $Manifest 'version') 1 1) -ne 1) { throw 'Unsupported diagnostic manifest protocol.' }
    if ($Manifest['diagnostics_finalized'] -isnot [bool] -or $Manifest['diagnostics_finalized'] -ne $true) { throw 'Diagnostic manifest is not finalized.' }
    $contractPath = Join-Path $PSScriptRoot 'WinBookSplit.Outcomes.json'
    $contract = [WinBookSplitSupport.StrictJson]::Parse([IO.File]::ReadAllText($contractPath, [Text.Encoding]::UTF8))
    $canonicalVersion = Get-WinBookSplitSupportToken $contract['application_version'] @('1.0.0-dev','1.0.0')
    $null = Get-WinBookSplitSupportToken (Get-WinBookSplitSupportField $Manifest 'application_version') @($canonicalVersion)
    $runId = Get-WinBookSplitSupportField $Manifest 'run_id'
    if ($runId -isnot [string] -or $runId -cnotmatch '^[a-f0-9]{32}$') { throw 'Invalid diagnostic run identity.' }
    $runtime = Get-WinBookSplitSupportField $Manifest 'runtime_versions'
    Assert-WinBookSplitSupportKeys $runtime @('powershell','python','pypdf','calibre') @('powershell','python','pypdf','calibre')
    $versions = [ordered]@{}
    $pins = @{ python=@('3.14.8'); pypdf=@('6.19.0'); calibre=@('9.15.0') }
    foreach ($name in @('powershell','python','pypdf','calibre')) {
        $value = Get-WinBookSplitSupportField $runtime $name
        $versions[$name] = $null
        if ($null -ne $value) {
            if ($name -ceq 'powershell') {
                if ($value -isnot [string] -or $value -cnotmatch '^(5\.1\.[0-9]{1,5}\.[0-9]{1,5}|7\.[0-9]{1,5}\.[0-9]{1,5}(\.[0-9]{1,5})?)$') { throw 'Invalid diagnostic runtime version.' }
                foreach ($component in $value.Split('.')) { if ([long]$component -gt 65535 -or ($component.Length -gt 1 -and $component[0] -eq '0')) { throw 'Invalid diagnostic runtime version.' } }
                if ($value.StartsWith('5.1.',[StringComparison]::Ordinal)) { $versions[$name]='5.1' } else { $versions[$name]='7' }
            }
            else { $versions[$name] = Get-WinBookSplitSupportToken $value $pins[$name] }
        }
    }
    $settings = Get-WinBookSplitSupportField $Manifest 'settings'
    $settingKeys = @('mode','input_kind','preview','non_interactive','no_pause','keep_converted_pdf','conversion_timeout','process_timeout','normalized_inputs')
    Assert-WinBookSplitSupportKeys $settings $settingKeys $settingKeys
    $cleanSettings = [ordered]@{ mode=(Get-WinBookSplitSupportToken $settings['mode'] @('','manual','1','2')); input_kind=(Get-WinBookSplitSupportToken $settings['input_kind'] @('pdf','epub','azw3')) }
    foreach ($name in @('preview','non_interactive','no_pause','keep_converted_pdf')) {
        if ($settings[$name] -isnot [bool]) { throw 'Invalid diagnostic Boolean.' }
        $cleanSettings[$name] = [bool]$settings[$name]
    }
    $cleanSettings['conversion_timeout'] = Get-WinBookSplitSupportInteger $settings['conversion_timeout'] 86400 1
    $cleanSettings['process_timeout'] = Get-WinBookSplitSupportInteger $settings['process_timeout'] 172800 1
    $plan = Get-WinBookSplitSupportField $Manifest 'plan'
    $cleanPlan = $null
    if ($null -ne $plan) {
        $pages = Get-WinBookSplitSupportInteger (Get-WinBookSplitSupportField $plan 'total_pages') 1000000 1
        $ranges = Get-WinBookSplitSupportField $plan 'ranges'
        if ($ranges -isnot [array] -or $ranges.Count -lt 1 -or $ranges.Count -gt 4096) { throw 'Invalid diagnostic ranges.' }
        $cursor = 0L; $cleanRanges = @()
        foreach ($range in $ranges) {
            if ($range -isnot [array] -or $range.Count -ne 2) { throw 'Invalid diagnostic range.' }
            $start = Get-WinBookSplitSupportInteger $range[0] $pages
            $end = Get-WinBookSplitSupportInteger $range[1] $pages 1
            if ($start -ne $cursor -or $end -le $start) { throw 'Incomplete diagnostic page coverage.' }
            $cleanRanges += ,@($start,$end); $cursor = $end
        }
        $coverage = Get-WinBookSplitSupportField $plan 'coverage'
        if ($cursor -ne $pages -or $coverage['complete'] -isnot [bool] -or $coverage['complete'] -ne $true -or
            (Get-WinBookSplitSupportInteger (Get-WinBookSplitSupportField $coverage 'covered_pages') 1000000) -ne $pages -or
            (Get-WinBookSplitSupportInteger (Get-WinBookSplitSupportField $coverage 'section_count') 4096) -ne $ranges.Count) { throw 'Invalid diagnostic coverage.' }
        $cleanPlan = [ordered]@{ total_pages=$pages; ranges=$cleanRanges; coverage=[ordered]@{ complete=$true; covered_pages=$pages; section_count=$ranges.Count } }
    }
    $outcome = Get-WinBookSplitSupportField $Manifest 'outcome'
    $status = Get-WinBookSplitSupportToken (Get-WinBookSplitSupportField $outcome 'status') @('success','preview','failed','incomplete','no_plan','invalid_input','read_error','write_error','error','unsupported','cancelled','timeout')
    $code = Get-WinBookSplitSupportToken (Get-WinBookSplitSupportField $outcome 'code') @($contract['codes'].Keys)
    $exitCode = Get-WinBookSplitSupportInteger (Get-WinBookSplitSupportField $outcome 'exit_code') 130
    $written = Get-WinBookSplitSupportInteger (Get-WinBookSplitSupportField $outcome 'written_count') 4096
    $mappedExit = [long]$contract['codes'][$code]
    if ($code -ceq 'conversion_cleanup_failed' -and $exitCode -eq 130) {
        $engine = Get-WinBookSplitSupportField $Manifest 'engine_result'
        $primary = Get-WinBookSplitSupportField (Get-WinBookSplitSupportField $engine 'diagnostic') 'primary_code'
        $null = Get-WinBookSplitSupportToken $primary @('conversion_cancelled','conversion_timeout')
        $mappedExit = 130
    }
    if ($exitCode -ne $mappedExit -or ($status -ceq 'success' -and ($code -cne 'split_complete' -or $written -lt 1 -or $null -eq $cleanPlan -or $written -ne $cleanPlan.coverage.section_count -or $cleanSettings.preview)) -or
        ($status -ceq 'preview' -and ($code -cne 'preview_complete' -or $written -ne 0 -or -not $cleanSettings.preview)) -or
        ($status -ceq 'no_plan' -and ($exitCode -ne 5 -or $code -cnotin @('no_bookmarks','no_usable_bookmarks','no_bookmarks_at_level'))) -or
        ($status -ceq 'invalid_input' -and $code -cnotin @('input_invalid','invalid_document','invalid_outline','invalid_start_pages','invalid_mode','invalid_arguments')) -or
        ($status -ceq 'read_error' -and $code -cne 'unreadable_document') -or
        ($status -ceq 'write_error' -and $code -cne 'output_write_failed') -or
        ($status -ceq 'unsupported' -and $code -cne 'unsupported_document') -or
        ($status -ceq 'incomplete' -and $code -cnotin @('console_finalize_failed','output_handle_close_failed')) -or
        ($status -cin @('failed','incomplete','invalid_input','read_error','write_error','error','unsupported') -and $exitCode -cin @(0,130)) -or
        ($status -cin @('cancelled','timeout') -and ($exitCode -ne 130 -or $written -ne 0)) -or
        ($status -cne 'success' -and $written -ne 0)) { throw 'Contradictory diagnostic outcome.' }
    $warnings = Get-WinBookSplitSupportField $Manifest 'warnings'
    $warningKeys = @('planning','parser','parser_suppressed_count','parser_message_truncated_count')
    Assert-WinBookSplitSupportKeys $warnings $warningKeys $warningKeys
    $planning = Get-WinBookSplitSupportField $warnings 'planning'; $parser = Get-WinBookSplitSupportField $warnings 'parser'
    if ($planning -isnot [array] -or $planning.Count -gt 4096 -or $parser -isnot [array] -or $parser.Count -gt 64) { throw 'Invalid diagnostic warning collection.' }
    $cleanPlanning = @(); $cleanParser = @()
    $bookmarkCodes = @('invalid_destination','external_destination','destination_error','duplicate_destination','outline_reordered','duplicate_parent_subtree','invalid_parent_subtree','child_outside_parent','malformed_outline','outline_cycle','outline_limit','output_handle_close_failed')
    foreach ($warning in $planning) {
        $warningCode = Get-WinBookSplitSupportToken (Get-WinBookSplitSupportField $warning 'code') $bookmarkCodes
        $category = 'bookmark'; if ($warningCode -ceq 'output_handle_close_failed') { $category = 'output' }
        $cleanPlanning += [ordered]@{ category=$category; code=$warningCode }
    }
    foreach ($warning in $parser) {
        $warningCode = Get-WinBookSplitSupportToken (Get-WinBookSplitSupportField $warning 'code') @('pypdf_parser_warning','pypdf_parser_error')
        $category = Get-WinBookSplitSupportToken (Get-WinBookSplitSupportField $warning 'category') @('pdf_parser')
        $severity = Get-WinBookSplitSupportToken (Get-WinBookSplitSupportField $warning 'severity') @('WARNING','ERROR')
        if (($warningCode -ceq 'pypdf_parser_warning' -and $severity -cne 'WARNING') -or ($warningCode -ceq 'pypdf_parser_error' -and $severity -cne 'ERROR')) { throw 'Contradictory parser warning category.' }
        $cleanParser += [ordered]@{ category=$category; code=$warningCode; severity=$severity }
    }
    $cleanWarnings = [ordered]@{ planning=$cleanPlanning; parser=$cleanParser;
        parser_suppressed_count=(Get-WinBookSplitSupportInteger $warnings['parser_suppressed_count'] 2147483647);
        parser_message_truncated_count=(Get-WinBookSplitSupportInteger $warnings['parser_message_truncated_count'] 2147483647) }
    $log = Get-WinBookSplitSupportField $Manifest 'log'
    $max = Get-WinBookSplitSupportInteger (Get-WinBookSplitSupportField $log 'max_bytes') 50331648 1
    $body = Get-WinBookSplitSupportInteger (Get-WinBookSplitSupportField $log 'body_limit_bytes') 33554432 1
    $bytes = Get-WinBookSplitSupportInteger (Get-WinBookSplitSupportField $log 'bytes_written') $max
    if ($body -gt $max -or $log['limit_reached'] -isnot [bool]) { throw 'Invalid diagnostic log bounds.' }
    return [ordered]@{ protocol='winbooksplit.support'; version=1; application_version=$canonicalVersion; runtime_versions=$versions;
        settings=$cleanSettings; plan=$cleanPlan; outcome=[ordered]@{ status=$status; code=$code; exit_code=$exitCode; written_count=$written };
        warnings=$cleanWarnings; log=[ordered]@{ max_bytes=$max; body_limit_bytes=$body; bytes_written=$bytes; limit_reached=[bool]$log['limit_reached'] } }
}

function Export-WinBookSplitSupportSummary {
    param([AllowEmptyString()][string]$ManifestPath, [AllowEmptyString()][string]$OutputPath)
    Initialize-WinBookSplitSupport
    $files = $null; $sourceStream = $null
    try {
        $files = New-Object WinBookSplitSupport.NativeFiles($ManifestPath, $OutputPath)
        if ([IO.Path]::GetFileName($files.Source) -cne 'WinBookSplit_Run.json') { throw 'Use the finalized run manifest.' }
        $sourceStream = $files.OpenSource()
        $manifest = [WinBookSplitSupport.StrictJson]::Read($sourceStream)
        $summary = ConvertTo-WinBookSplitSupportSummary $manifest
        $text = ($summary | ConvertTo-Json -Depth 12) + [Environment]::NewLine
        $bytes = (New-Object Text.UTF8Encoding($false, $true)).GetBytes($text)
        $files.WriteNew($bytes)
        return $summary
    }
    finally {
        if ($null -ne $sourceStream) { $sourceStream.Dispose() }
        if ($null -ne $files) { $files.Dispose() }
    }
}
