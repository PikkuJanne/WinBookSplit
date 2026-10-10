"""Bounded dependency guards and actual BAT shell-shadowing regression."""
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest

from pypdf import apply_configuration, get_configuration
from pypdf.errors import LimitReachedError
from pypdf.filters import decode_stream_data
from pypdf.generic import ArrayObject, EncodedStreamObject, NameObject

ROOT = Path(__file__).resolve().parents[2]


class StreamResourceGuardTests(unittest.TestCase):
    def test_many_filters_reject_without_large_allocation(self):
        stream = EncodedStreamObject()
        stream._data = b"x"
        stream[NameObject('/Filter')] = ArrayObject([NameObject('/DCTDecode')] * 17)
        with self.assertRaisesRegex(LimitReachedError, 'Maximum filter count'):
            decode_stream_data(stream)

    def test_accumulated_decoding_work_limit_stays_enabled(self):
        configuration = get_configuration()
        self.assertEqual(getattr(configuration, 'stream_filters_maximum_length', None), 16)
        self.assertEqual(getattr(configuration, 'stream_decoding_work_maximum_length', None), 300_000_000)
        stream = EncodedStreamObject()
        stream._data = b"xxx"
        stream[NameObject('/Filter')] = NameObject('/DCTDecode')
        with apply_configuration(stream_decoding_work_maximum_length=5):
            with self.assertRaisesRegex(LimitReachedError, 'accumulated'):
                decode_stream_data(stream)


@unittest.skipUnless(sys.platform == 'win32', 'Actual Windows BAT execution required')
class BatchShellTrustTests(unittest.TestCase):
    def setUp(self):
        self.owned = tempfile.TemporaryDirectory(prefix='wbs-shell-trust-')
        self.addCleanup(self.owned.cleanup)
        self.work = Path(self.owned.name)
        self.app = self.work / 'app'
        self.app.mkdir()
        shutil.copyfile(ROOT / 'WinBookSplit.bat', self.app / 'WinBookSplit.bat')
        self.marker = self.work / 'called.txt'
        self.decoy = self.work / 'decoy.txt'
        (self.app / 'WinBookSplit.ps1').write_text(
            "param([string]$InputFile)\n"
            "[IO.File]::WriteAllText($env:WBS_SHELL_TRUST_MARKER, 'inbox:' + $InputFile)\n"
            "exit 37\n", encoding='utf-8')
        self.system = Path(os.environ['SystemRoot']) / 'System32'
        self.wrapper = self.work / 'run.cmd'
        self.wrapper.write_bytes(b'@echo off\r\nsetlocal DisableDelayedExpansion\r\n'
            b'"%WBS_SHELL_TRUST_BAT%" "%WBS_SHELL_TRUST_INPUT%"\r\n')
        self.environment = os.environ.copy()
        self.environment.update(WBS_SHELL_TRUST_BAT=str(self.app / 'WinBookSplit.bat'),
            WBS_SHELL_TRUST_MARKER=str(self.marker), WBS_SHELL_TRUST_DECOY=str(self.decoy),
            WBS_SHELL_TRUST_INPUT="Literal [%UNUSED%!] O'Brien & (book) 日本語.pdf",
            PATHEXT='.COM;.EXE;.BAT;.CMD')

    def invoke(self, cwd):
        return subprocess.run([str(self.system / 'cmd.exe'), '/d', '/c', str(self.wrapper)],
            cwd=cwd, env=self.environment, stdin=subprocess.DEVNULL, capture_output=True,
            timeout=20, check=False)

    def plant_decoy(self, directory):
        directory.mkdir()
        for extension in ('cmd', 'bat'):
            (directory / ('powershell.' + extension)).write_bytes(
                b'@echo off\r\necho decoy>"%WBS_SHELL_TRUST_DECOY%"\r\nexit /b 73\r\n')

    def test_cwd_shell_decoys_do_not_execute_and_literal_argv_exit_survive(self):
        cwd = self.work / 'cwd'
        self.plant_decoy(cwd)
        self.environment['PATH'] = str(self.system)
        result = self.invoke(cwd)
        self.assertEqual(result.returncode, 37, result.stderr.decode(errors='replace'))
        self.assertFalse(self.decoy.exists())
        self.assertEqual(self.marker.read_text(encoding='utf-8'), 'inbox:' + self.environment['WBS_SHELL_TRUST_INPUT'])

    def test_path_shell_decoys_do_not_execute_for_no_input(self):
        directory = self.work / 'path'
        self.plant_decoy(directory)
        self.environment['PATH'] = str(directory) + os.pathsep + str(self.system)
        self.environment['WBS_SHELL_TRUST_INPUT'] = ''
        result = self.invoke(self.work)
        self.assertEqual(result.returncode, 37, result.stderr.decode(errors='replace'))
        self.assertFalse(self.decoy.exists())
        self.assertEqual(self.marker.read_text(encoding='utf-8'), 'inbox:')

    def test_missing_system_shell_fails_without_path_fallback(self):
        directory = self.work / 'path'
        self.plant_decoy(directory)
        self.environment['PATH'] = str(directory)
        self.environment['SystemRoot'] = str(self.work / 'absent-windows')
        result = self.invoke(self.work)
        self.assertEqual(result.returncode, 3)
        self.assertIn(b'Windows PowerShell', result.stdout)
        self.assertFalse(self.decoy.exists())
        self.assertFalse(self.marker.exists())
