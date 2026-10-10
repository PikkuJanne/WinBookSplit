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
        public string StartError, StopError, StreamError, ResultError, InteractionError, InputError, Stdout, Stderr, HumanStdout;
        public long StdoutTotalBytes, StderrTotalBytes;
        public long InteractionCount, InteractionTotalBytes, QueuedReplyCount, ReplyCount, ReplyTotalBytes;
        public bool StdoutTruncated, StderrTruncated;
        public bool InputWriterStopped = true;
        public string[] ResultRecords = new string[0];
        public double ElapsedSeconds;
    }
    sealed class Drain {
        readonly object gate = new object();
        const string interactionPrefix = "[WBS-INTERACTION] ";
        readonly byte[] tail; readonly int resultLimit; readonly bool frames;
        readonly Session session;
        readonly List<string> records = new List<string>();
        readonly List<long> recordStarts = new List<long>(), recordEnds = new List<long>();
        readonly MemoryStream line = new MemoryStream();
        readonly Decoder validator = new UTF8Encoding(false, true).GetDecoder();
        int count, position, recordBytes, prefixPosition; bool candidate, startedLine, overflow, interaction, prefixPossible = true;
        long total, lineStart, interactionCount, interactionBytes; bool done, eof; string error, resultError, interactionError;
        internal Drain(int limit, int resultLimit, bool frames, Session session) {
            tail = new byte[limit]; this.resultLimit = resultLimit; this.frames = frames; this.session = session;
        }
        void Exclude(long end) {
            // Only spans intersecting the retained tail are needed. An event
            // cannot grow this ledger after it leaves that bounded ring.
            while (recordEnds.Count != 0 && recordEnds[0] <= total-tail.Length) {
                recordStarts.RemoveAt(0); recordEnds.RemoveAt(0);
            }
            recordStarts.Add(lineStart); recordEnds.Add(end);
        }
        void EndLine(long end) {
            if (interaction) {
                Exclude(end);
                if (!overflow && interactionError == null) {
                    byte[] bytes = line.ToArray();
                    int length = bytes.Length;
                    if (length != 0 && bytes[length-1] == 13) length--;
                    int payloadLength = length-interactionPrefix.Length;
                    if (payloadLength > Session.EventLimitBytes)
                        interactionError = "An interaction event exceeds the bounded event limit.";
                    else {
                        interactionCount++; interactionBytes += payloadLength;
                        string value = new UTF8Encoding(false, true).GetString(bytes, interactionPrefix.Length, payloadLength);
                        if (session == null) interactionError = "Interaction events require a supervised session.";
                        else {
                            string failure = session.EnqueueRequest(value, payloadLength);
                            if (failure != null) interactionError = failure;
                        }
                    }
                }
            }
            else if (candidate && !overflow) {
                byte[] bytes = line.ToArray();
                int length = bytes.Length;
                if (length != 0 && bytes[length-1] == 13) length--;
                if (records.Count >= 2) resultError = "More than two structured result candidates were emitted.";
                else if (recordBytes + length > resultLimit) resultError = "Structured result records exceed the bounded result limit.";
                else if (resultError == null) {
                    records.Add(new UTF8Encoding(false, true).GetString(bytes, 0, length)); recordBytes += length;
                    Exclude(end);
                }
            }
            line.SetLength(0); candidate = false; startedLine = false; overflow = false;
            interaction = false; prefixPossible = true; prefixPosition = 0; lineStart=end;
        }
        void Frame(byte value, long offset) {
            if (value == 10) { EndLine(offset+1); return; }
            if (prefixPossible) {
                if (value == (byte)interactionPrefix[prefixPosition]) {
                    prefixPosition++;
                    if (prefixPosition == interactionPrefix.Length) { interaction = true; prefixPossible = false; }
                } else prefixPossible = false;
            }
            if (!startedLine && value != 32 && value != 9 && value != 13) {
                startedLine = true; candidate = value == 123;
            }
            bool retained = !startedLine || candidate || interaction || prefixPossible;
            if (!retained) line.SetLength(0);
            if (retained && !overflow && (interaction || prefixPossible || resultError == null)) {
                // Reserve one bounded byte for the optional physical CR in
                // CRLF; EndLine enforces the JSON payload limit after stripping it.
                int limit = interaction || prefixPossible ? Session.EventLimitBytes+interactionPrefix.Length+1 : resultLimit;
                if (line.Length >= limit) {
                    if (interaction) interactionError = "An interaction event exceeds the bounded event limit.";
                    else if (candidate) resultError = "A structured result line exceeds the bounded result limit.";
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
                    result.InteractionError = interactionError;
                    result.InteractionCount = interactionCount; result.InteractionTotalBytes = interactionBytes;
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
        internal bool InteractionFailed { get { lock (gate) { return interactionError != null; } } }
    }
    public sealed class Session : IDisposable {
        internal const int EventLimitBytes = 8388608;
        public const int ReplyLimitBytes = 65536;
        const int QueueLimit = 8;
        readonly object gate = new object();
        readonly Queue<string> requests = new Queue<string>();
        readonly Queue<byte[]> replies = new Queue<byte[]>();
        readonly Thread supervisor;
        Thread writer; FileStream inputStream;
        int requestBytes, replyBytes, cancelRequested;
        long queuedReplyCount, replyCount, replyTotalBytes;
        bool writerStop, disposed; volatile bool completed;
        string inputError; Result result;
        internal Session(string path,string command,string cwd,int timeout,int captureLimit,int resultLimit,CancellationToken cancellation) {
            supervisor = new Thread(delegate() {
                Result final;
                try { final = Child.RunSession(path,command,cwd,timeout,captureLimit,resultLimit,cancellation,this); }
                catch (Exception failure) { final = new Result { StartError = "Session supervisor: "+failure.Message, InputWriterStopped = false }; }
                lock (gate) { result = final; completed = true; Monitor.PulseAll(gate); }
            });
            supervisor.IsBackground = true; supervisor.Start();
        }
        public bool Completed { get { return completed; } }
        public Result Result {
            get { lock (gate) { if (!completed) throw new InvalidOperationException("The session has not completed."); return result; } }
        }
        internal bool CancelRequested { get { return Interlocked.CompareExchange(ref cancelRequested,0,0) != 0; } }
        internal bool InputFailed { get { lock (gate) { return inputError != null; } } }
        internal string EnqueueRequest(string value,int bytes) {
            lock (gate) {
                if (requests.Count >= QueueLimit || requestBytes+bytes > EventLimitBytes)
                    return "Interaction requests exceed the bounded queue.";
                requests.Enqueue(value); requestBytes += bytes; return null;
            }
        }
        public string TryReadRequest() {
            lock (gate) {
                if (requests.Count == 0) return null;
                string value = requests.Dequeue();
                requestBytes -= new UTF8Encoding(false,true).GetByteCount(value);
                return value;
            }
        }
        public void SendReply(string value) {
            if (value == null || value.Length == 0 || value.IndexOf('\0') >= 0 || value.IndexOf('\r') >= 0 || value.IndexOf('\n') >= 0)
                throw new ArgumentException("A session reply must be one nonempty literal JSON line.");
            byte[] payload = new UTF8Encoding(false,true).GetBytes(value);
            if (payload.Length > ReplyLimitBytes) throw new ArgumentException("The session reply exceeds 65536 UTF-8 bytes.");
            byte[] line = new byte[payload.Length+1]; Buffer.BlockCopy(payload,0,line,0,payload.Length); line[line.Length-1]=10;
            lock (gate) {
                if (completed || disposed || writerStop || CancelRequested || inputError != null)
                    throw new InvalidOperationException("The supervised session no longer accepts replies.");
                if (replies.Count >= QueueLimit || replyBytes+payload.Length > ReplyLimitBytes)
                    throw new InvalidOperationException("Session replies exceed the bounded queue.");
                replies.Enqueue(line); replyBytes += payload.Length; queuedReplyCount++; Monitor.PulseAll(gate);
            }
        }
        internal void StartWriter(FileStream stream) {
            inputStream = stream;
            writer = new Thread(delegate() {
                try {
                    while (true) {
                        byte[] line;
                        lock (gate) {
                            while (!writerStop && replies.Count == 0) Monitor.Wait(gate);
                            if (writerStop) break;
                            line = replies.Dequeue(); replyBytes -= line.Length-1;
                        }
                        inputStream.Write(line,0,line.Length); inputStream.Flush();
                        lock (gate) { replyCount++; replyTotalBytes += line.Length; }
                    }
                }
                catch (Exception failure) { lock (gate) { if (!writerStop && !CancelRequested) inputError = failure.Message; } }
                finally {
                    try { inputStream.Dispose(); }
                    catch (Exception failure) { lock (gate) { inputError = inputError ?? failure.Message; } }
                }
            });
            writer.IsBackground = true;
            try { writer.Start(); }
            catch { inputStream.Dispose(); throw; }
        }
        internal void RequestWriterStop() {
            lock (gate) { writerStop = true; replies.Clear(); replyBytes = 0; Monitor.PulseAll(gate); }
        }
        internal void StopWriter(Result final) {
            RequestWriterStop();
            if (writer != null && writer.IsAlive) writer.Join(1000);
            lock (gate) {
                final.InputWriterStopped = writer == null || !writer.IsAlive;
                final.InputError = inputError; final.QueuedReplyCount = queuedReplyCount;
                final.ReplyCount = replyCount; final.ReplyTotalBytes = replyTotalBytes;
                if (!final.InputWriterStopped) final.InputError = (inputError == null ? "" : inputError+"; ")+"The session stdin writer was not proved stopped.";
            }
        }
        public void Cancel() { Interlocked.Exchange(ref cancelRequested,1); }
        public void Dispose() {
            if (!completed) { Cancel(); if (!supervisor.Join(6000)) throw new IOException("Session supervision did not complete; retain owned work."); }
            lock (gate) { disposed = true; requests.Clear(); requestBytes = 0; }
        }
    }
    public sealed class ConsoleLine {
        static readonly object readersGate = new object();
        static readonly Dictionary<TextReader,ConsoleLine> readers = new Dictionary<TextReader,ConsoleLine>();
        readonly object gate = new object(); readonly Thread worker;
        string value, error; volatile bool completed, abandoned;
        ConsoleLine(TextReader reader) {
            worker = new Thread(delegate() {
                try { string line = reader.ReadLine(); lock (gate) { value = line; } }
                catch (Exception failure) { lock (gate) { error = failure.Message; } }
                finally {
                    completed = true;
                    lock (readersGate) {
                        ConsoleLine current;
                        if (readers.TryGetValue(reader,out current) && Object.ReferenceEquals(current,this)) readers.Remove(reader);
                    }
                }
            });
            worker.IsBackground = true;
        }
        public static ConsoleLine Start(TextReader reader) {
            if (reader == null) throw new ArgumentNullException("reader");
            lock (readersGate) {
                ConsoleLine current;
                if (readers.TryGetValue(reader,out current) && !current.Completed)
                    throw new InvalidOperationException("A console line is already pending; an abandoned read requires immediate application exit.");
                ConsoleLine line = new ConsoleLine(reader); readers[reader] = line;
                try { line.worker.Start(); } catch { readers.Remove(reader); throw; }
                return line;
            }
        }
        public bool Completed { get { return completed; } }
        public bool AbandonedForExit { get { return abandoned; } }
        public string Value { get { lock (gate) { if (!completed) throw new InvalidOperationException("The console line is pending."); return value; } } }
        public string Error { get { lock (gate) { return error; } } }
        // A blocked console ReadLine cannot safely be aborted or disposed.
        // The caller may abandon it only while immediately exiting the app.
        public void AbandonForExit() { abandoned = true; }
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
        static FileStream PipeWriter(ref IntPtr handle) {
            SafeFileHandle owned=new SafeFileHandle(handle,true); handle=IntPtr.Zero;
            try { return new FileStream(owned,FileAccess.Write,4096,false); }
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
            return RunCore(path,command,cwd,timeout,captureLimit,resultLimit,cancellation,null);
        }
        public static Session StartSession(string path,string command,string cwd,int timeout,int captureLimit,int resultLimit,CancellationToken cancellation) {
            return new Session(path,command,cwd,timeout,captureLimit,resultLimit,cancellation);
        }
        internal static Result RunSession(string path,string command,string cwd,int timeout,int captureLimit,int resultLimit,CancellationToken cancellation,Session session) {
            return RunCore(path,command,cwd,timeout,captureLimit,resultLimit,cancellation,session);
        }
        // Ordinary and interactive invocations share the same owned job,
        // restricted startup handles, drains, deadline and shutdown proofs.
        static Result RunCore(string path,string command,string cwd,int timeout,int captureLimit,int resultLimit,CancellationToken cancellation,Session session) {
            Result result = new Result(); Stopwatch clock = Stopwatch.StartNew();
            IntPtr job=IntPtr.Zero, input=IntPtr.Zero, inputWrite=IntPtr.Zero, outputRead=IntPtr.Zero, outputWrite=IntPtr.Zero,errorRead=IntPtr.Zero,errorWrite=IntPtr.Zero, env=IntPtr.Zero,attributes=IntPtr.Zero,handles=IntPtr.Zero;
            ProcessInfo process=new ProcessInfo(); bool created=false, attributeReady=false; int consoleCancelled=0;
            Drain output=new Drain(captureLimit,resultLimit,true,session), error=new Drain(captureLimit,resultLimit,false,null);
            Thread outThread=null,errThread=null; FileStream outStream=null,errStream=null; bool outStarted=false,errStarted=false;
            ConsoleControl cancelHandler=delegate(uint kind) {
                if (kind!=0 && kind!=1) return false;
                Interlocked.Exchange(ref consoleCancelled,1); return true;
            };
            bool handlerAdded=false;
            try {
                if (Environment.OSVersion.Platform != PlatformID.Win32NT) throw new PlatformNotSupportedException("Engine supervision requires Windows.");
                if (cancellation.IsCancellationRequested || (session != null && session.CancelRequested)) { result.Cancelled=true; return result; }
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
                if (session == null) {
                    input=CreateFileW("NUL",0x80000000,3,ref security,3,0,IntPtr.Zero); if (input==new IntPtr(-1)) throw Failure("Open the engine stdin null device");
                } else if (!CreatePipe(out input,out inputWrite,ref security,0) || !SetHandleInformation(inputWrite,1,0)) {
                    throw Failure("Create restricted interactive stdin pipe");
                }
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
                if (session != null) session.StartWriter(PipeWriter(ref inputWrite));
                if (cancellation.IsCancellationRequested || Interlocked.CompareExchange(ref consoleCancelled,0,0)!=0 || (session != null && session.CancelRequested)) { result.Cancelled=true; return result; }
                if (ResumeThread(process.Thread)!=1) throw Failure("Resume the assigned engine primary thread");
                Close(ref process.Thread,result);
                while (true) {
                    result.ParentStopped=WaitForSingleObject(process.Process,0)==0;
                    result.DescendantsStopped=Active(job)==0;
                    if (result.ParentStopped && result.DescendantsStopped && output.Done && error.Done) break;
                    if (cancellation.IsCancellationRequested || Interlocked.CompareExchange(ref consoleCancelled,0,0)!=0 || (session != null && session.CancelRequested)) { result.Cancelled=true; break; }
                    if (session != null && (output.InteractionFailed || session.InputFailed)) throw new IOException("The bounded session interaction channel failed.");
                    if (clock.ElapsedMilliseconds>=timeout) { result.TimedOut=true; break; }
                    Thread.Sleep(5);
                }
            } catch (Exception failure) { result.StartError=failure.Message; }
            finally {
                if (session != null) session.RequestWriterStop();
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
                Close(ref outputWrite,result); Close(ref errorWrite,result); Close(ref input,result); Close(ref inputWrite,result);
                if (session != null) {
                    try { session.StopWriter(result); }
                    catch (Exception failure) { StopError(result,"Stop the session stdin writer: "+failure.Message); }
                }
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

# The session owns only managed/native transport. JSON schema, nonce, sequence
# and application-stage validation stay in the caller and the shipped engine.
function Start-WinBookSplitProcessSession {
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
    return [WinBookSplit.Supervision.Child]::StartSession($Path, $command, $WorkingDirectory,
        [int]($TimeoutSeconds * 1000), $CaptureLimitBytes, $ResultLimitBytes, $CancellationToken)
}

function Start-WinBookSplitConsoleLine {
    param([IO.TextReader]$Reader = [Console]::In)
    Initialize-WinBookSplitProcess
    return [WinBookSplit.Supervision.ConsoleLine]::Start($Reader)
}
