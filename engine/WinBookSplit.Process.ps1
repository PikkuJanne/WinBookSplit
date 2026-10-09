# Narrow Windows engine transport. Dot-sourcing does not launch an executable.
# The existing result validator decides the outcome; human tails never do.
function Initialize-WinBookSplitProcess {
    if ('WinBookSplit.Supervision.Child' -as [type]) { return }
    Add-Type -TypeDefinition @'
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using Microsoft.Win32.SafeHandles;
namespace WinBookSplit.Supervision {
    public sealed class Result {
        public int? Pid, ExitCode;
        public bool TimedOut, Cancelled, ParentStopped, DescendantsStopped, StreamsComplete, JobAssigned;
        public string StartError, StopError, StreamError, ResultError, Stdout, Stderr, HumanStdout;
        public long StdoutTotalBytes, StderrTotalBytes;
        public bool StdoutTruncated, StderrTruncated;
        public string[] ResultRecords = new string[0];
        public double ElapsedSeconds;
    }
    sealed class Drain {
        readonly object gate = new object();
        readonly byte[] tail; readonly int resultLimit; readonly bool frames;
        readonly List<string> records = new List<string>();
        readonly List<long> recordStarts = new List<long>(), recordEnds = new List<long>();
        readonly MemoryStream line = new MemoryStream();
        readonly Decoder validator = new UTF8Encoding(false, true).GetDecoder();
        int count, position, recordBytes; bool candidate, startedLine, overflow;
        long total, lineStart; bool done, eof; string error, resultError;
        internal Drain(int limit, int resultLimit, bool frames) { tail = new byte[limit]; this.resultLimit = resultLimit; this.frames = frames; }
        void EndLine(long end) {
            if (candidate && !overflow) {
                byte[] bytes = line.ToArray();
                int length = bytes.Length;
                if (length != 0 && bytes[length-1] == 13) length--;
                if (records.Count >= 2) resultError = "More than two structured result candidates were emitted.";
                else if (recordBytes + length > resultLimit) resultError = "Structured result records exceed the bounded result limit.";
                else if (resultError == null) {
                    records.Add(new UTF8Encoding(false, true).GetString(bytes, 0, length)); recordBytes += length;
                    recordStarts.Add(lineStart); recordEnds.Add(end);
                }
            }
            line.SetLength(0); candidate = false; startedLine = false; overflow = false; lineStart=end;
        }
        void Frame(byte value, long offset) {
            if (value == 10) { EndLine(offset+1); return; }
            if (!startedLine && value != 32 && value != 9 && value != 13) { startedLine = true; candidate = value == 123; if (!candidate) line.SetLength(0); }
            if ((!startedLine || candidate) && !overflow && resultError == null) {
                if (line.Length >= resultLimit) {
                    if (candidate) resultError = "A structured result line exceeds the bounded result limit.";
                    overflow = true; line.SetLength(0);
                } else line.WriteByte(value);
            }
        }
        internal void Read(Stream stream) {
            try {
                byte[] block = new byte[8192]; char[] chars = new char[8194]; int length;
                while ((length = stream.Read(block, 0, block.Length)) != 0) {
                    lock (gate) {
                        // Validate the whole stream incrementally, independently of a ring cut.
                        validator.GetChars(block, 0, length, chars, 0, false);
                        total += length;
                        for (int i = 0; i < length; i++) {
                            tail[position] = block[i]; position = (position + 1) % tail.Length;
                            if (count < tail.Length) count++;
                            if (frames) Frame(block[i], total-length+i);
                        }
                    }
                }
                lock (gate) { validator.GetChars(new byte[0], 0, 0, new char[2], 0, true); if (frames) EndLine(total); eof = true; }
            } catch (Exception failure) { lock (gate) { error = failure.Message; } }
            finally {
                try { stream.Dispose(); } catch (Exception failure) { lock (gate) { error = failure.Message; } }
                lock (gate) { line.Dispose(); done = true; }
            }
        }
        internal bool Done { get { lock (gate) { return done; } } }
        internal void Copy(Result result, bool stdout) {
            lock (gate) {
                byte[] bytes = new byte[count]; int first = count == tail.Length ? position : 0;
                int beforeWrap = Math.Min(count, tail.Length - first);
                Buffer.BlockCopy(tail, first, bytes, 0, beforeWrap);
                Buffer.BlockCopy(tail, 0, bytes, beforeWrap, count - beforeWrap);
                int skip = 0;
                // A bounded tail may begin inside one valid multibyte character.
                if (total > tail.Length) while (skip < bytes.Length && (bytes[skip] & 0xc0) == 0x80) skip++;
                string text;
                try { text = new UTF8Encoding(false, true).GetString(bytes, skip, bytes.Length - skip); }
                catch (Exception failure) { text = ""; error = error ?? failure.Message; }
                if (stdout) {
                    result.Stdout = text; result.StdoutTotalBytes = total; result.StdoutTruncated = total > tail.Length;
                    result.ResultRecords = records.ToArray(); result.ResultError = resultError;
                    using (MemoryStream human = new MemoryStream()) {
                        long absoluteStart=total-count;
                        for (int i=skip; i<bytes.Length; i++) {
                            bool machine=false; long offset=absoluteStart+i;
                            for (int r=0; r<recordStarts.Count; r++) if (offset>=recordStarts[r] && offset<recordEnds[r]) { machine=true; break; }
                            if (!machine) human.WriteByte(bytes[i]);
                        }
                        try { result.HumanStdout=new UTF8Encoding(false,true).GetString(human.ToArray()); }
                        catch (Exception failure) { result.HumanStdout=""; error=error??failure.Message; }
                    }
                } else { result.Stderr = text; result.StderrTotalBytes = total; result.StderrTruncated = total > tail.Length; }
                if (error != null) result.StreamError = (result.StreamError == null ? "" : result.StreamError + "; ") + (stdout ? "stdout: " : "stderr: ") + error;
            }
        }
        internal bool Eof { get { lock (gate) { return eof; } } }
    }
    public static class Child {
        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        delegate bool ConsoleControl(uint kind);
        // If Windows refuses removal, retain the registered callback until host
        // exit rather than leave a native pointer to a collected delegate.
        static readonly List<ConsoleControl> retainedControlHandlers=new List<ConsoleControl>();
        [StructLayout(LayoutKind.Sequential)] struct Security { public int Size; public IntPtr Descriptor; public int Inherit; }
        [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)] struct Startup {
            public int Size; public string Reserved, Desktop, Title; public int X,Y,XSize,YSize,XChars,YChars,Fill,Flags;
            public short Show, ReservedSize; public IntPtr ReservedBytes, Input, Output, Error;
        }
        [StructLayout(LayoutKind.Sequential)] struct StartupEx { public Startup Info; public IntPtr Attributes; }
        [StructLayout(LayoutKind.Sequential)] struct ProcessInfo { public IntPtr Process, Thread; public uint ProcessId, ThreadId; }
        [StructLayout(LayoutKind.Sequential)] struct BasicLimits {
            public long ProcessTime, JobTime; public uint Flags; public UIntPtr Minimum, Maximum; public uint ActiveLimit;
            public UIntPtr Affinity; public uint Priority, Scheduling;
        }
        [StructLayout(LayoutKind.Sequential)] struct IoCounters { public ulong ReadOperations,WriteOperations,OtherOperations,ReadBytes,WriteBytes,OtherBytes; }
        [StructLayout(LayoutKind.Sequential)] struct ExtendedLimits { public BasicLimits Basic; public IoCounters Io; public UIntPtr ProcessMemory,JobMemory,PeakProcess,PeakJob; }
        [StructLayout(LayoutKind.Sequential)] struct Accounting {
            public long User,Kernel,PeriodUser,PeriodKernel; public uint Faults,Total,Active,Terminated;
        }
        [DllImport("kernel32.dll",SetLastError=true)] static extern bool CreatePipe(out IntPtr read,out IntPtr write,ref Security security,uint size);
        [DllImport("kernel32.dll",SetLastError=true)] static extern bool SetHandleInformation(IntPtr handle,uint mask,uint flags);
        [DllImport("kernel32.dll",CharSet=CharSet.Unicode,SetLastError=true)] static extern IntPtr CreateFileW(string path,uint access,uint share,ref Security security,uint creation,uint flags,IntPtr template);
        [DllImport("kernel32.dll",CharSet=CharSet.Unicode,SetLastError=true)] static extern bool CreateProcessW(string executable,StringBuilder command,IntPtr processSecurity,IntPtr threadSecurity,bool inherit,uint flags,IntPtr environment,string cwd,ref StartupEx startup,out ProcessInfo process);
        [DllImport("kernel32.dll",SetLastError=true)] static extern bool InitializeProcThreadAttributeList(IntPtr list,int count,uint flags,ref IntPtr size);
        [DllImport("kernel32.dll",SetLastError=true)] static extern bool UpdateProcThreadAttribute(IntPtr list,uint flags,IntPtr attribute,IntPtr value,IntPtr size,IntPtr previous,IntPtr returned);
        [DllImport("kernel32.dll")] static extern void DeleteProcThreadAttributeList(IntPtr list);
        [DllImport("kernel32.dll",CharSet=CharSet.Unicode,SetLastError=true)] static extern IntPtr CreateJobObjectW(IntPtr security,string name);
        [DllImport("kernel32.dll",SetLastError=true)] static extern bool SetInformationJobObject(IntPtr job,int type,ref ExtendedLimits limits,uint size);
        [DllImport("kernel32.dll",SetLastError=true)] static extern bool QueryInformationJobObject(IntPtr job,int type,out Accounting accounting,uint size,IntPtr returned);
        [DllImport("kernel32.dll",SetLastError=true)] static extern bool AssignProcessToJobObject(IntPtr job,IntPtr process);
        [DllImport("kernel32.dll",SetLastError=true)] static extern bool IsProcessInJob(IntPtr process,IntPtr job,out bool belongs);
        [DllImport("kernel32.dll",SetLastError=true)] static extern bool TerminateJobObject(IntPtr job,uint code);
        [DllImport("kernel32.dll",SetLastError=true)] static extern bool TerminateProcess(IntPtr process,uint code);
        [DllImport("kernel32.dll",SetLastError=true)] static extern uint ResumeThread(IntPtr thread);
        [DllImport("kernel32.dll",SetLastError=true)] static extern uint WaitForSingleObject(IntPtr handle,uint timeout);
        [DllImport("kernel32.dll",SetLastError=true)] static extern bool GetExitCodeProcess(IntPtr process,out uint code);
        [DllImport("kernel32.dll",SetLastError=true)] static extern bool CloseHandle(IntPtr handle);
        [DllImport("kernel32.dll",SetLastError=true)] static extern bool SetConsoleCtrlHandler(ConsoleControl handler,bool add);
        static Win32Exception Failure(string action) { return new Win32Exception(Marshal.GetLastWin32Error(), action); }
        static uint Active(IntPtr job) {
            Accounting value;
            if (!QueryInformationJobObject(job,1,out value,(uint)Marshal.SizeOf(typeof(Accounting)),IntPtr.Zero)) throw Failure("Query the owned engine job");
            return value.Active;
        }
        static void Close(ref IntPtr handle,Result result) {
            if (handle == IntPtr.Zero || handle == new IntPtr(-1)) return;
            if (!CloseHandle(handle)) StopError(result,Failure("Close an owned engine handle").Message);
            handle = IntPtr.Zero;
        }
        static void StopError(Result result,string message) { result.StopError=(result.StopError==null ? "" : result.StopError+"; ")+message; }
        static FileStream PipeStream(ref IntPtr handle) {
            SafeFileHandle owned=new SafeFileHandle(handle,true); handle=IntPtr.Zero;
            try { return new FileStream(owned,FileAccess.Read,8192,false); }
            catch { owned.Dispose(); throw; }
        }
        static void Join(Thread thread,bool started,Result result) {
            try { if (started) thread.Join(1000); }
            catch (Exception failure) { StopError(result,"Join an engine drain: "+failure.Message); }
        }
        static void DisposeUnstarted(FileStream stream,bool started,Result result) {
            try { if (!started && stream!=null) stream.Dispose(); }
            catch (Exception failure) { StopError(result,"Close an unstarted engine drain: "+failure.Message); }
        }
        static IntPtr EnvironmentBlock() {
            SortedDictionary<string,string> values = new SortedDictionary<string,string>(StringComparer.OrdinalIgnoreCase);
            foreach (DictionaryEntry item in Environment.GetEnvironmentVariables()) {
                string name = (string)item.Key;
                if (!name.StartsWith("PYTHON",StringComparison.OrdinalIgnoreCase) && !name.StartsWith("PYLAUNCHER",StringComparison.OrdinalIgnoreCase)) values[name] = (string)item.Value;
            }
            values["PYTHONUTF8"]="1"; values["PYTHONIOENCODING"]="utf-8:backslashreplace";
            values["PYTHON_MANAGER_AUTOMATIC_INSTALL"]="false"; values["PYLAUNCHER_NO_SEARCH_PATH"]="1";
            StringBuilder block = new StringBuilder(); foreach (KeyValuePair<string,string> item in values) block.Append(item.Key).Append('=').Append(item.Value).Append('\0'); block.Append('\0');
            return Marshal.StringToHGlobalUni(block.ToString());
        }
        public static Result Run(string path,string command,string cwd,int timeout,int captureLimit,int resultLimit,CancellationToken cancellation) {
            Result result = new Result(); Stopwatch clock = Stopwatch.StartNew();
            IntPtr job=IntPtr.Zero, input=IntPtr.Zero, outputRead=IntPtr.Zero, outputWrite=IntPtr.Zero,errorRead=IntPtr.Zero,errorWrite=IntPtr.Zero, env=IntPtr.Zero,attributes=IntPtr.Zero,handles=IntPtr.Zero;
            ProcessInfo process=new ProcessInfo(); bool created=false, attributeReady=false; int consoleCancelled=0;
            Drain output=new Drain(captureLimit,resultLimit,true), error=new Drain(captureLimit,resultLimit,false);
            Thread outThread=null,errThread=null; FileStream outStream=null,errStream=null; bool outStarted=false,errStarted=false;
            ConsoleControl cancelHandler=delegate(uint kind) {
                if (kind!=0 && kind!=1) return false;
                Interlocked.Exchange(ref consoleCancelled,1); return true;
            };
            bool handlerAdded=false;
            try {
                if (Environment.OSVersion.Platform != PlatformID.Win32NT) throw new PlatformNotSupportedException("Engine supervision requires Windows.");
                if (cancellation.IsCancellationRequested) { result.Cancelled=true; return result; }
                // Register before creating any owned child. This native last
                // handler consumes Ctrl+C/Break before PowerShell can schedule
                // a stop of the enclosing script pipeline.
                if (!SetConsoleCtrlHandler(cancelHandler,true)) throw Failure("Register owned-process console cancellation");
                handlerAdded=true;
                Security security=new Security { Size=Marshal.SizeOf(typeof(Security)), Inherit=1 };
                job=CreateJobObjectW(IntPtr.Zero,null); if (job==IntPtr.Zero) throw Failure("Create the owned engine job");
                ExtendedLimits limits=new ExtendedLimits(); limits.Basic.Flags=0x2000;
                if (!SetInformationJobObject(job,9,ref limits,(uint)Marshal.SizeOf(typeof(ExtendedLimits)))) throw Failure("Set kill-on-close engine job limits");
                if (!CreatePipe(out outputRead,out outputWrite,ref security,0) || !SetHandleInformation(outputRead,1,0)) throw Failure("Create the stdout pipe");
                if (!CreatePipe(out errorRead,out errorWrite,ref security,0) || !SetHandleInformation(errorRead,1,0)) throw Failure("Create the stderr pipe");
                input=CreateFileW("NUL",0x80000000,3,ref security,3,0,IntPtr.Zero); if (input==new IntPtr(-1)) throw Failure("Open the engine stdin null device");
                IntPtr size=IntPtr.Zero; InitializeProcThreadAttributeList(IntPtr.Zero,1,0,ref size);
                if (size==IntPtr.Zero) throw Failure("Size the restricted engine handle list");
                attributes=Marshal.AllocHGlobal(size); if (!InitializeProcThreadAttributeList(attributes,1,0,ref size)) throw Failure("Initialize the restricted engine handle list"); attributeReady=true;
                handles=Marshal.AllocHGlobal(3*IntPtr.Size); Marshal.WriteIntPtr(handles,0,input); Marshal.WriteIntPtr(handles,IntPtr.Size,outputWrite); Marshal.WriteIntPtr(handles,2*IntPtr.Size,errorWrite);
                if (!UpdateProcThreadAttribute(attributes,0,new IntPtr(0x20002),handles,new IntPtr(3*IntPtr.Size),IntPtr.Zero,IntPtr.Zero)) throw Failure("Restrict engine inherited handles");
                StartupEx startup=new StartupEx(); startup.Info.Size=Marshal.SizeOf(typeof(StartupEx)); startup.Info.Flags=0x100; startup.Info.Input=input; startup.Info.Output=outputWrite; startup.Info.Error=errorWrite; startup.Attributes=attributes;
                env=EnvironmentBlock();
                if (!CreateProcessW(path,new StringBuilder(command),IntPtr.Zero,IntPtr.Zero,true,0x08080404,env,cwd,ref startup,out process)) throw Failure("Start the suspended engine executable");
                created=true; result.Pid=(int)process.ProcessId;
                if (!AssignProcessToJobObject(job,process.Process)) throw Failure("Assign the suspended engine to its owned job");
                result.JobAssigned=true; bool belongs;
                if (!IsProcessInJob(process.Process,job,out belongs) || !belongs) throw Failure("Verify engine job ownership");
                Close(ref outputWrite,result); Close(ref errorWrite,result); Close(ref input,result);
                if (result.StopError!=null) throw new IOException(result.StopError);
                outStream=PipeStream(ref outputRead); errStream=PipeStream(ref errorRead);
                outThread=new Thread(delegate() { output.Read(outStream); }); outThread.IsBackground=true; outThread.Start(); outStarted=true;
                errThread=new Thread(delegate() { error.Read(errStream); }); errThread.IsBackground=true; errThread.Start(); errStarted=true;
                if (cancellation.IsCancellationRequested || Interlocked.CompareExchange(ref consoleCancelled,0,0)!=0) { result.Cancelled=true; return result; }
                if (ResumeThread(process.Thread)!=1) throw Failure("Resume the assigned engine primary thread");
                Close(ref process.Thread,result);
                while (true) {
                    result.ParentStopped=WaitForSingleObject(process.Process,0)==0;
                    result.DescendantsStopped=Active(job)==0;
                    if (result.ParentStopped && result.DescendantsStopped && output.Done && error.Done) break;
                    if (cancellation.IsCancellationRequested || Interlocked.CompareExchange(ref consoleCancelled,0,0)!=0) { result.Cancelled=true; break; }
                    if (clock.ElapsedMilliseconds>=timeout) { result.TimedOut=true; break; }
                    Thread.Sleep(5);
                }
            } catch (Exception failure) { result.StartError=failure.Message; }
            finally {
                // Failure/timeout/cancel shuts down only the handles created by this run.
                if (created && !(result.ParentStopped && result.DescendantsStopped)) {
                    try {
                        if (result.JobAssigned) { if (!TerminateJobObject(job,1)) throw Failure("Terminate the owned engine job"); }
                        else if (!TerminateProcess(process.Process,1)) throw Failure("Terminate the still-suspended engine");
                        Stopwatch grace=Stopwatch.StartNew();
                        do {
                            result.ParentStopped=WaitForSingleObject(process.Process,0)==0;
                            result.DescendantsStopped=result.JobAssigned ? Active(job)==0 : result.ParentStopped;
                            if (result.ParentStopped && result.DescendantsStopped) break;
                            Thread.Sleep(5);
                        } while (grace.ElapsedMilliseconds<2000);
                        if (!result.ParentStopped || !result.DescendantsStopped) throw new IOException("The owned engine tree could not be proved stopped.");
                    } catch (Exception failure) { StopError(result,failure.Message); }
                } else if (!created) { result.ParentStopped=true; result.DescendantsStopped=true; }
                Close(ref outputWrite,result); Close(ref errorWrite,result); Close(ref input,result);
                Join(outThread,outStarted,result); Join(errThread,errStarted,result);
                // Normally the stopped non-breakaway job closes every inherited writer.
                // Incomplete drains remain explicit; no unbounded cross-thread Dispose.
                result.StreamsComplete=!created || (output.Eof && error.Eof);
                if (created) {
                    uint code; if (result.ParentStopped && GetExitCodeProcess(process.Process,out code)) result.ExitCode=unchecked((int)code);
                    else if (result.ParentStopped) StopError(result,Failure("Read the engine exit code").Message);
                }
                try { output.Copy(result,true); error.Copy(result,false); }
                catch (Exception failure) { result.StreamError="Capture engine diagnostics: "+failure.Message; }
                finally {
                    DisposeUnstarted(outStream,outStarted,result); DisposeUnstarted(errStream,errStarted,result);
                    Close(ref outputRead,result); Close(ref errorRead,result); Close(ref process.Thread,result); Close(ref process.Process,result); Close(ref job,result);
                    if (attributeReady) DeleteProcThreadAttributeList(attributes);
                    if (attributes!=IntPtr.Zero) Marshal.FreeHGlobal(attributes); if (handles!=IntPtr.Zero) Marshal.FreeHGlobal(handles); if (env!=IntPtr.Zero) Marshal.FreeHGlobal(env);
                    if (handlerAdded && !SetConsoleCtrlHandler(cancelHandler,false)) {
                        StopError(result,Failure("Remove owned-process console cancellation").Message);
                        lock (retainedControlHandlers) retainedControlHandlers.Add(cancelHandler);
                    }
                    GC.KeepAlive(cancelHandler);
                    result.ElapsedSeconds=clock.Elapsed.TotalSeconds;
                }
            }
            return result;
        }
    }
}
'@ -ErrorAction Stop
}

function Invoke-WinBookSplitProcess {
    param([Parameter(Mandatory = $true)][string]$Path,
          [Parameter(Mandatory = $true)][AllowEmptyString()][AllowEmptyCollection()][string[]]$Arguments,
          [Parameter(Mandatory = $true)][string]$WorkingDirectory,
          [ValidateRange(0.05, 172800)][double]$TimeoutSeconds = 3600,
          [ValidateRange(1024, 1048576)][int]$CaptureLimitBytes = 65536,
          [ValidateRange(1024, 33554432)][int]$ResultLimitBytes = 8388608,
          [Threading.CancellationToken]$CancellationToken = [Threading.CancellationToken]::None)
    if ($Path -notmatch '^(?:[A-Za-z]:[\\/]|\\\\[^\\]+\\[^\\]+\\)' -or $Path.IndexOf([char]0) -ge 0 -or
        $WorkingDirectory -notmatch '^(?:[A-Za-z]:[\\/]|\\\\[^\\]+\\[^\\]+\\)' -or $WorkingDirectory.IndexOf([char]0) -ge 0) {
        throw 'Engine launch requires absolute literal executable and working-directory paths.'
    }
    foreach ($argument in $Arguments) {
        if ($null -eq $argument -or $argument.IndexOf([char]0) -ge 0) { throw 'Native arguments must be literal strings without NUL.' }
    }
    Initialize-WinBookSplitProcess
    $command = (ConvertTo-NativeArgument -Value $Path) + ' ' + (($Arguments | ForEach-Object { ConvertTo-NativeArgument -Value $_ }) -join ' ')
    return [WinBookSplit.Supervision.Child]::Run($Path, $command, $WorkingDirectory,
        [int]($TimeoutSeconds * 1000), $CaptureLimitBytes, $ResultLimitBytes, $CancellationToken)
}
