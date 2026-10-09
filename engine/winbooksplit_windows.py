"""Small Windows directory locks for owned output publication.

The caller must guard existing ancestors from root to leaf before using paths
beneath them. This helper never allocates directories or deletes their contents.
"""

import ctypes
from ctypes import wintypes
import os


_GENERIC_READ = 0x80000000
_READ_ATTRIBUTES = 0x0080
_DELETE = 0x00010000
_SHARE_READ = 0x00000001
_OPEN_EXISTING = 3
_BACKUP_SEMANTICS = 0x02000000
_OPEN_REPARSE_POINT = 0x00200000
_DIRECTORY = 0x0010
_REPARSE_POINT = 0x0400
_FILE_RENAME_INFO = 3
_FILE_ID_INFO = 18


class _HandleInformation(ctypes.Structure):
    _fields_ = [("attributes", wintypes.DWORD),
                ("creation", wintypes.FILETIME), ("access", wintypes.FILETIME),
                ("write", wintypes.FILETIME), ("volume", wintypes.DWORD),
                ("size_high", wintypes.DWORD), ("size_low", wintypes.DWORD),
                ("links", wintypes.DWORD), ("index_high", wintypes.DWORD),
                ("index_low", wintypes.DWORD)]


class _FileIdentity(ctypes.Structure):
    _fields_ = [("volume", ctypes.c_ulonglong), ("file_id", ctypes.c_ubyte * 16)]


class _RenameInformation(ctypes.Structure):
    # The BOOLEAN/DWORD union occupies four bytes in FILE_RENAME_INFO.
    _fields_ = [("replace_if_exists", wintypes.DWORD),
                ("root_directory", wintypes.HANDLE),
                ("name_length", wintypes.DWORD), ("name", wintypes.WCHAR * 1)]


def _windows_api():
    if os.name != "nt":
        raise OSError("Owned directory publication requires Windows.")
    api = ctypes.WinDLL("kernel32", use_last_error=True)
    api.CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
                               ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD,
                               wintypes.HANDLE]
    api.CreateFileW.restype = wintypes.HANDLE
    api.GetFileInformationByHandle.argtypes = [wintypes.HANDLE,
                                               ctypes.POINTER(_HandleInformation)]
    api.GetFileInformationByHandle.restype = wintypes.BOOL
    api.GetFileInformationByHandleEx.argtypes = [wintypes.HANDLE, ctypes.c_int,
                                                ctypes.c_void_p, wintypes.DWORD]
    api.GetFileInformationByHandleEx.restype = wintypes.BOOL
    api.SetFileInformationByHandle.argtypes = [wintypes.HANDLE, ctypes.c_int,
                                              ctypes.c_void_p, wintypes.DWORD]
    api.SetFileInformationByHandle.restype = wintypes.BOOL
    api.CloseHandle.argtypes = [wintypes.HANDLE]
    api.CloseHandle.restype = wintypes.BOOL
    return api


class _PathGuard:
    """Restrict data-write/delete/rename opens on an ordinary filesystem object.

    Compatible metadata access can still alter an empty directory's reparse
    attributes; callers must check them before path-based creation. This is not
    an atomic check/create or a sandbox against another same-account process.
    delete_access is only for held objects the caller owns and will rename/delete.
    Opening may fail on unsupported filesystems or incompatible existing handles;
    callers must fail closed rather than retrying with weaker sharing flags.
    """

    _directory = True
    _share_mode = _SHARE_READ

    def __init__(self, path, delete_access=False):
        self.path = os.path.abspath(os.fspath(path))
        self.delete_access = delete_access
        self._api = _windows_api()
        self._handle = None
        handle = self._open(self.path, _GENERIC_READ | (_DELETE if delete_access else 0),
                            self._share_mode)
        try:
            self.identity = self._information(handle)
        except BaseException:
            self._api.CloseHandle(handle)
            raise
        self._handle = handle
        try:
            self.assert_unchanged()
            observed = os.lstat(self.path)
            self.stat_identity = (observed.st_dev, observed.st_ino)
        except BaseException:
            self.close()
            raise

    def _open(self, path, access, sharing):
        handle = self._api.CreateFileW(path, access, sharing, None, _OPEN_EXISTING,
                                       _BACKUP_SEMANTICS | _OPEN_REPARSE_POINT, None)
        if handle is None or handle == ctypes.c_void_p(-1).value:
            raise ctypes.WinError(ctypes.get_last_error())
        return handle

    def _information(self, handle):
        information = _HandleInformation()
        if not self._api.GetFileInformationByHandle(handle, ctypes.byref(information)):
            raise ctypes.WinError(ctypes.get_last_error())
        if bool(information.attributes & _DIRECTORY) != self._directory \
                or information.attributes & _REPARSE_POINT:
            kind = "directory" if self._directory else "file"
            raise OSError(f"Expected an ordinary {kind}, not another kind or reparse point.")
        identity = _FileIdentity()
        if not self._api.GetFileInformationByHandleEx(
                handle, _FILE_ID_INFO, ctypes.byref(identity), ctypes.sizeof(identity)):
            raise ctypes.WinError(ctypes.get_last_error())
        return identity.volume, bytes(identity.file_id).hex()

    def assert_unchanged(self):
        if self._handle is None:
            raise OSError("The path guard is closed.")
        if self._information(self._handle) != self.identity:
            raise OSError("The guarded identity changed.")
        # The extra metadata handle shares all access so it remains compatible
        # with our staging DELETE handle. It neither weakens nor replaces it.
        current = self._open(self.path, _READ_ATTRIBUTES, 0x00000007)
        try:
            if self._information(current) != self.identity:
                raise OSError("The guarded path no longer names the owned object.")
        finally:
            self._api.CloseHandle(current)

    def close(self):
        if self._handle is not None:
            handle, self._handle = self._handle, None
            if not self._api.CloseHandle(handle):
                raise ctypes.WinError(ctypes.get_last_error())

    def delete(self):
        """Remove only the held object; a nonempty directory fails in Windows."""
        if not self.delete_access:
            raise OSError("Deletion requires the owned object DELETE handle.")
        self.assert_unchanged()
        disposition = ctypes.c_ubyte(1)
        if not self._api.SetFileInformationByHandle(self._handle, 4,
                                                   ctypes.byref(disposition), 1):
            raise ctypes.WinError(ctypes.get_last_error())
        self.close()

    def __enter__(self):
        self.assert_unchanged()
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        self.close()


class DirectoryGuard(_PathGuard):
    """Guard a directory; caller also holds its ancestors from root to leaf."""

    def allow_publication(self, stage_guard):
        """Allow the target-parent opens required by native publication.

        The held stage keeps this base nonempty and blocks its replacement while
        the base handle is reacquired. Ancestor and staging guards stay strict.
        Call after reserving/guarding staging, before final membership checks.
        """
        self._change_sharing(stage_guard, 0x00000003)

    def restrict_writes(self, stage_guard):
        """Restore strict base protection before deleting its last owned child."""
        self._change_sharing(stage_guard, _SHARE_READ)

    def _change_sharing(self, stage_guard, share_mode):
        if not isinstance(stage_guard, DirectoryGuard) or not stage_guard.delete_access \
                or self.delete_access \
                or os.path.normcase(os.path.dirname(stage_guard.path)) != os.path.normcase(self.path):
            raise OSError("Publication sharing requires the held owned staging child.")
        self.assert_unchanged()
        stage_guard.assert_unchanged()
        if self._share_mode == share_mode:
            return
        identity = self.identity
        self.close()
        self._share_mode = share_mode  # Neither mode permits DELETE sharing.
        try:
            self._handle = self._open(self.path, _GENERIC_READ, self._share_mode)
            if self._information(self._handle) != identity:
                raise OSError("The output base changed while reacquiring its guard.")
            self.assert_unchanged()
            stage_guard.assert_unchanged()
        except BaseException:
            self.close()
            raise

    def rename_to(self, base_guard, basename):
        """Publish this sibling through its retained handle, without replacement."""
        if not self.delete_access:
            raise OSError("Publication requires the owned staging DELETE handle.")
        if not isinstance(base_guard, DirectoryGuard):
            raise TypeError("Publication requires a guarded output base.")
        if base_guard._share_mode != 0x00000003:
            raise OSError("Enable publication sharing under the held staging guard first.")
        if not isinstance(basename, str) or not basename or basename in {".", ".."} \
                or any(character in '\\/:<>"|?*' or ord(character) < 32 for character in basename) \
                or basename.rstrip(" .") != basename:
            raise ValueError("Publication requires one ordinary directory basename.")
        if os.path.normcase(os.path.dirname(self.path)) != os.path.normcase(base_guard.path):
            raise OSError("Staging and publication must be siblings in the guarded base.")
        self.assert_unchanged()
        base_guard.assert_unchanged()
        if self.identity[0] != base_guard.identity[0]:
            raise OSError("Publication must remain on the staging volume.")
        target = os.path.join(base_guard.path, basename)
        # RootDirectory-relative names returned WinError 87 on the tested host.
        # An absolute target works with the same retained staging handle. The
        # caller's root-to-base locks keep its parent traversal unchanged.
        encoded = target.encode("utf-16-le")
        length = ctypes.sizeof(_RenameInformation) + len(encoded)
        buffer = ctypes.create_string_buffer(length)
        rename = _RenameInformation.from_buffer(buffer)
        rename.replace_if_exists = 0
        rename.root_directory = None
        rename.name_length = len(encoded)
        ctypes.memmove(ctypes.addressof(buffer) + _RenameInformation.name.offset,
                       encoded, len(encoded))
        if not self._api.SetFileInformationByHandle(self._handle, _FILE_RENAME_INFO,
                                                   buffer, length):
            raise ctypes.WinError(ctypes.get_last_error())
        self.path = target
        return self.path


class FileGuard(_PathGuard):
    """Guard one known file; match stat_identity before a delete-access cleanup."""

    _directory = False
