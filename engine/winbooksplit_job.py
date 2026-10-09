"""Narrow Windows converter job: assign a suspended child before it can run.

The caller drains the binary pipes and owns its timeout/cancellation policy.
This module owns only the launched process tree, never a process found by name.
Job assignment or primary-thread verification failure has no unguarded fallback.
Windows 11 and the project's pinned regular CPython runtime are the test target.
"""

import ctypes
from ctypes import wintypes
import os
import subprocess
import time


_CREATE_SUSPENDED = 0x00000004
_KILL_ON_JOB_CLOSE = 0x00002000
_THREAD_SNAPSHOT = 0x00000004
_THREAD_ACCESS = 0x00000002 | 0x00000800  # SUSPEND_RESUME | QUERY_LIMITED_INFORMATION
_PROCESS_ACCESS = 0x00000100 | 0x00000001 | 0x00001000 | 0x00100000
_INVALID_HANDLE = ctypes.c_void_p(-1).value


class _BasicLimits(ctypes.Structure):
    _fields_ = [("process_time", ctypes.c_longlong), ("job_time", ctypes.c_longlong),
                ("flags", wintypes.DWORD), ("minimum_working_set", ctypes.c_size_t),
                ("maximum_working_set", ctypes.c_size_t), ("active_limit", wintypes.DWORD),
                ("affinity", ctypes.c_size_t), ("priority", wintypes.DWORD),
                ("scheduling", wintypes.DWORD)]


class _IoCounters(ctypes.Structure):
    _fields_ = [(name, ctypes.c_ulonglong) for name in
                ("read_operations", "write_operations", "other_operations",
                 "read_bytes", "write_bytes", "other_bytes")]


class _ExtendedLimits(ctypes.Structure):
    _fields_ = [("basic", _BasicLimits), ("io", _IoCounters),
                ("process_memory", ctypes.c_size_t), ("job_memory", ctypes.c_size_t),
                ("peak_process_memory", ctypes.c_size_t), ("peak_job_memory", ctypes.c_size_t)]


class _Accounting(ctypes.Structure):
    _fields_ = [("user_time", ctypes.c_longlong), ("kernel_time", ctypes.c_longlong),
                ("period_user_time", ctypes.c_longlong), ("period_kernel_time", ctypes.c_longlong),
                ("page_faults", wintypes.DWORD), ("total_processes", wintypes.DWORD),
                ("active_processes", wintypes.DWORD), ("terminated_processes", wintypes.DWORD)]


class _ThreadEntry(ctypes.Structure):
    _fields_ = [("size", wintypes.DWORD), ("usage", wintypes.DWORD),
                ("thread_id", wintypes.DWORD), ("owner_pid", wintypes.DWORD),
                ("base_priority", wintypes.LONG), ("delta_priority", wintypes.LONG),
                ("flags", wintypes.DWORD)]


def _windows_api():
    if os.name != "nt":
        raise OSError("Converter job ownership requires the supported Windows runtime.")
    api = ctypes.WinDLL("kernel32", use_last_error=True)
    declarations = {
        "CreateJobObjectW": ([ctypes.c_void_p, wintypes.LPCWSTR], wintypes.HANDLE),
        "SetInformationJobObject": ([wintypes.HANDLE, ctypes.c_int, ctypes.c_void_p, wintypes.DWORD], wintypes.BOOL),
        "QueryInformationJobObject": ([wintypes.HANDLE, ctypes.c_int, ctypes.c_void_p, wintypes.DWORD, ctypes.c_void_p], wintypes.BOOL),
        "AssignProcessToJobObject": ([wintypes.HANDLE, wintypes.HANDLE], wintypes.BOOL),
        "IsProcessInJob": ([wintypes.HANDLE, wintypes.HANDLE, ctypes.POINTER(wintypes.BOOL)], wintypes.BOOL),
        "TerminateJobObject": ([wintypes.HANDLE, wintypes.UINT], wintypes.BOOL),
        "OpenProcess": ([wintypes.DWORD, wintypes.BOOL, wintypes.DWORD], wintypes.HANDLE),
        "GetProcessId": ([wintypes.HANDLE], wintypes.DWORD),
        "CreateToolhelp32Snapshot": ([wintypes.DWORD, wintypes.DWORD], wintypes.HANDLE),
        "Thread32First": ([wintypes.HANDLE, ctypes.POINTER(_ThreadEntry)], wintypes.BOOL),
        "Thread32Next": ([wintypes.HANDLE, ctypes.POINTER(_ThreadEntry)], wintypes.BOOL),
        "OpenThread": ([wintypes.DWORD, wintypes.BOOL, wintypes.DWORD], wintypes.HANDLE),
        "GetProcessIdOfThread": ([wintypes.HANDLE], wintypes.DWORD),
        "ResumeThread": ([wintypes.HANDLE], wintypes.DWORD),
        "CloseHandle": ([wintypes.HANDLE], wintypes.BOOL),
    }
    for name, (arguments, result) in declarations.items():
        function = getattr(api, name)
        function.argtypes, function.restype = arguments, result
    return api


def _native_error(action):
    error = ctypes.WinError(ctypes.get_last_error())
    error.add_note(action)
    return error


class JobProcess:
    """A Popen plus one unnamed, non-breakaway, kill-on-close Windows job."""

    def __init__(self, argv, *, cwd, env):
        if not isinstance(argv, (list, tuple)) or not argv or not argv[0] \
                or any(not isinstance(argument, str) or "\x00" in argument for argument in argv) \
                or not os.path.isabs(argv[0]):
            raise ValueError("Converter launch requires literal argv and an absolute executable.")
        self.argv = list(argv)
        self.process = None
        self._api = _windows_api()
        self._job = None
        self._process_handle = None
        self._assigned = False
        self._tree_stopped = False
        self._closed = False
        self.primary_thread_id = None
        self.assigned_before_resume = False
        try:
            self._job = self._api.CreateJobObjectW(None, None)
            if not self._job:
                raise _native_error("Create the owned converter job")
            limits = _ExtendedLimits()
            limits.basic.flags = _KILL_ON_JOB_CLOSE
            if not self._api.SetInformationJobObject(self._job, 9, ctypes.byref(limits), ctypes.sizeof(limits)):
                raise _native_error("Set kill-on-close converter job limits")
            self.process = subprocess.Popen(self.argv, cwd=cwd, env=env, shell=False,
                                            stdin=subprocess.DEVNULL, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                            bufsize=0,
                                            creationflags=_CREATE_SUSPENDED | subprocess.CREATE_NO_WINDOW)
            self._process_handle = self._api.OpenProcess(_PROCESS_ACCESS, False, self.process.pid)
            if not self._process_handle:
                raise _native_error("Open the owned suspended converter process")
            # The original Popen still holds the created process handle. If it
            # exited before OpenProcess, refuse any possibly reused PID handle.
            if self.process.poll() is not None or self._api.GetProcessId(self._process_handle) != self.process.pid:
                raise OSError("The created converter exited before job assignment.")
            if not self._api.AssignProcessToJobObject(self._job, self._process_handle):
                raise _native_error("Assign the suspended converter to its owned job")
            self._assigned = True
            belongs = wintypes.BOOL()
            if not self._api.IsProcessInJob(self._process_handle, self._job, ctypes.byref(belongs)) or not belongs.value:
                raise OSError("Cannot verify the converter's owned job membership.")
            self._resume_primary_thread()
            self.assigned_before_resume = True
        except BaseException as error:
            try:
                self.close()
            except BaseException as close_error:
                error.add_note("Converter startup cleanup: " + str(close_error))
            error.converter_tree_stopped = self._tree_stopped
            raise

    def _resume_primary_thread(self):
        snapshot = self._api.CreateToolhelp32Snapshot(_THREAD_SNAPSHOT, 0)
        if not snapshot or snapshot == _INVALID_HANDLE:
            raise _native_error("Snapshot the suspended converter's primary thread")
        thread_ids = []
        try:
            entry = _ThreadEntry()
            entry.size = ctypes.sizeof(entry)
            available = self._api.Thread32First(snapshot, ctypes.byref(entry))
            seen = 0
            while available:
                if entry.size < _ThreadEntry.owner_pid.offset + ctypes.sizeof(wintypes.DWORD):
                    raise OSError("The thread snapshot lacks its owner identity.")
                if entry.owner_pid == self.process.pid:
                    thread_ids.append(entry.thread_id)
                seen += 1
                if seen > 100000:
                    raise OSError("The thread snapshot exceeds the bounded startup scan.")
                entry.size = ctypes.sizeof(entry)
                available = self._api.Thread32Next(snapshot, ctypes.byref(entry))
            if ctypes.get_last_error() != 18:  # ERROR_NO_MORE_FILES
                raise _native_error("Enumerate the suspended converter's primary thread")
        finally:
            if not self._api.CloseHandle(snapshot):
                raise _native_error("Close the converter thread snapshot")
        if len(thread_ids) != 1 or self.process.poll() is not None:
            raise OSError("The suspended converter must have exactly one owned primary thread.")
        thread = self._api.OpenThread(_THREAD_ACCESS, False, thread_ids[0])
        if not thread:
            raise _native_error("Open the owned converter primary thread")
        try:
            if self._api.GetProcessIdOfThread(thread) != self.process.pid or self.process.poll() is not None:
                raise OSError("The suspended primary thread changed owner.")
            previous_count = self._api.ResumeThread(thread)
            if previous_count == 0xFFFFFFFF:
                raise _native_error("Resume the assigned converter primary thread")
            if previous_count != 1:
                raise OSError("The converter primary thread had an unexpected suspend count.")
            self.primary_thread_id = thread_ids[0]
        finally:
            if not self._api.CloseHandle(thread):
                raise _native_error("Close the converter primary thread handle")

    @property
    def stdout(self):
        return self.process.stdout

    @property
    def stderr(self):
        return self.process.stderr

    @property
    def pid(self):
        return self.process.pid

    @property
    def returncode(self):
        return self.process.returncode

    def poll(self):
        return self.process.poll()

    def wait(self, timeout=None):
        return self.process.wait(timeout=timeout)

    def active_process_count(self):
        if self._job is None:
            if self._tree_stopped:
                return 0
            raise OSError("The converter job is closed without a stopped-tree observation.")
        accounting = _Accounting()
        if not self._api.QueryInformationJobObject(self._job, 1, ctypes.byref(accounting), ctypes.sizeof(accounting), None):
            raise _native_error("Query owned converter job membership")
        return accounting.active_processes

    def terminate_tree(self, exit_code=1):
        if self._job is None:
            if self._tree_stopped:
                return
            raise OSError("Cannot terminate a closed converter job.")
        if not self._api.TerminateJobObject(self._job, exit_code):
            raise _native_error("Terminate only the owned converter job")

    def wait_tree(self, timeout=None):
        deadline = None if timeout is None else time.monotonic() + timeout
        while self.active_process_count():
            if deadline is not None and time.monotonic() >= deadline:
                raise subprocess.TimeoutExpired(self.argv, timeout)
            time.sleep(0.01)
        self._tree_stopped = True
        return True

    def close(self):
        """Attempt shutdown and every handle close; report any unproved stop."""
        if self._closed:
            return
        errors = []
        try:
            if self.process is not None:
                if self._assigned:
                    if not self._tree_stopped:
                        if self.active_process_count():
                            self.terminate_tree()
                        self.wait_tree(timeout=10)
                elif self.process.poll() is None:
                    self.process.terminate()  # Original Popen handle, never a discovered PID.
        except BaseException as error:
            errors.append(error)
        if self._assigned and not self._tree_stopped and self._job is not None:
            # An unproved termination/query must not leave a reader waiting on
            # live descendants while close tries to acquire a pipe read lock.
            # Last-handle close invokes the job's kill-on-close fallback. Lost
            # query access is not a stopped-tree proof; the caller preserves its
            # workspace even if the original parent subsequently exits.
            try:
                if not self._api.CloseHandle(self._job):
                    raise _native_error("Close the unproved converter job")
                self._job = None
            except BaseException as error:
                errors.append(error)
        if self.process is None:
            self._tree_stopped = True
        else:
            try:
                self.process.wait(timeout=10)
                if not self._assigned:
                    self._tree_stopped = True  # Suspended startup never ran or created descendants.
            except BaseException as error:
                errors.append(error)
        if self.process is not None:
            for stream in (self.process.stdout, self.process.stderr):
                try:
                    if stream is not None:
                        stream.close()
                except BaseException as error:
                    errors.append(error)
            # Popen has no public Windows-handle close method. Its owned Handle
            # Close operation is used only after wait in the pinned CPython;
            # process control otherwise uses documented public/native APIs.
            try:
                if self.process.returncode is not None:
                    self.process._handle.Close()
            except BaseException as error:
                errors.append(error)
        for name in ("_process_handle", "_job"):
            handle = getattr(self, name)
            if handle is not None:
                try:
                    if not self._api.CloseHandle(handle):
                        raise _native_error("Close owned converter " + name)
                    setattr(self, name, None)
                except BaseException as error:
                    errors.append(error)
        self._closed = not errors
        if errors:
            raise OSError("Converter job shutdown/close failed: " + "; ".join(str(error) for error in errors))

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        try:
            self.close()
        except BaseException as error:
            if exc_value is None:
                raise
            exc_value.add_note("Converter job close: " + str(error))


def launch_job(argv, *, cwd, env):
    """Return a started context-managed owned converter with binary pipes."""
    return JobProcess(argv, cwd=cwd, env=env)
