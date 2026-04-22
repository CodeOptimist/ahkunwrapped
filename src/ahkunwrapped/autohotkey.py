# SPDX-License-Identifier: AGPL-3.0-or-later
# Copyright (c) 2019-2026 Christopher S. Galpin

import array
import atexit
import io
import math
import os
import shutil
import string
import struct
import subprocess
import sys
import threading
import time
# import traceback
from collections.abc import Iterator, Mapping, Sequence
from contextlib import contextmanager, suppress
from dataclasses import dataclass
from enum import IntEnum
from pathlib import Path
from subprocess import TimeoutExpired
from typing import ClassVar, Final, IO
from warnings import warn

import win32api
import win32con
import win32job
import winerror
# noinspection PyUnresolvedReferences
from win32api import OutputDebugString

_IN_PYINSTALLER: Final = getattr(sys, 'frozen', False)
# noinspection PyProtectedMember,PyUnresolvedReferences
_PACKAGE_PATH: Final = Path(sys._MEIPASS) if _IN_PYINSTALLER else Path(__file__).parent
_SINGLE_JOB_ASSIGNMENTS: Final = sys.getwindowsversion().major < 8  # https://stackoverflow.com/q/13449531/
assert not _SINGLE_JOB_ASSIGNMENTS

type Primitive = bool | float | int | str


# @formatter:off
class AhkException(Exception): pass                         # noqa: E701
class AhkExitException(AhkException): pass                  # noqa: E701
class AhkError(AhkException): pass                          # noqa: E701
class AhkFuncNotFoundError(AhkError): pass                  # noqa: E701
class AhkUnsupportedValueError(AhkError): pass              # noqa: E701
class AhkCantCallOutInInputSyncCallError(AhkError): pass    # noqa: E701
class AhkWarning(UserWarning): pass                         # noqa: E701
# @formatter:on


class AhkLossOfPrecisionWarning(AhkWarning):
    def __init__(self, val: float, val_str: str):
        super().__init__(f'loss of precision from {val} to {val_str}')


@dataclass
class AhkUserException(AhkException):
    from_exception_obj: bool
    message: str
    what: str
    extra: str
    file: str
    line: int

    def __post_init__(self):
        self.from_exception_obj = self.from_exception_obj == "1"

        self.args = (self.message, self.what, self.extra, self.file, self.line)


class AhkCaughtNonExceptionWarning(AhkWarning):
    def __init__(self, exception: AhkUserException):
        message = f"""got "throw X"; recommend "throw Exception(X)" to preserve line info.
\tAlternatively, catch within script https://www.autohotkey.com/docs/commands/Throw.htm#Exception .
\tException message: '{exception.message}'"""
        if not exception.message:
            message += "\n\tMay have been an AutoHotkey object e.g. {abc: 123} intended for use within 'catch'."
        super().__init__(message)


def _comment_debug() -> str:
    return "" if "pytest" in sys.modules else ";"


class Script:
    class _Msg(IntEnum):
        # @formatter:off
        GET    = 0x8001
        SET    = 0x8002
        F      = 0x8003
        F_MAIN = 0x8004
        MORE   = 0x8005
        EXIT   = 0x8006
        # @formatter:off

    SEPARATOR: ClassVar[str] = '\3'  # :Separator

    _BUFFER_SIZE: Final[int] = 4096
    # we read one line at a time, but need a unique reserved character
    #  to signify the end of message, and not just a newline within it
    _EOM: Final = SEPARATOR.encode('utf-16-le') + b'\n'  # :Eom :OneByteNewline
    _EOM_SIZE: Final = 1 + len(_EOM)  # include :IsFinal bool
    _TEXT_SIZE: Final[int] = _BUFFER_SIZE - _EOM_SIZE

    _python_pid: Final = os.getpid()
    _python_job_obj: ClassVar[int]

    _CORE: Final[str] = '''
    _pyUserBatchLines := A_BatchLines
    SetBatchLines, -1
    Process, Exist, ''' + str(_python_pid) + '''
    if (ErrorLevel = 0) ; not found
        ExitApp ; https://stackoverflow.com/q/73506891/#comment129808240_73506891 :AvoidJobRace
    #NoEnv
    #NoTrayIcon
    #Persistent
    SetWorkingDir, ''' + os.getcwd() + '''
    _PY_SEPARATOR := Chr(''' + str(ord(SEPARATOR)) + ''')
    _PY_EOM_BYTES := Chr(1) "`n"  ; internally as 01 00 10 00  :Utf16Internals
    ; Let's write variables as they're stored, avoiding `StrPut()`.
    ; https://www.autohotkey.com/docs/v1/Concepts.htm#string-encoding
    _pyStdOut := FileOpen("*", "w", "utf-16-raw")
    _pyStdErr := FileOpen("**", "w", "utf-16-raw")

    _Py_Response(ByRef pipe, ByRef text, ByRef offset, ByRef onMain) {
        global _PY_SEPARATOR, _PY_EOM_BYTES
        textSize := Max(StrLen(text) * 2 + StrLen(Chr(0)) * 2 - offset, 0)
        isFinal := onMain or textSize <= ''' + str(_TEXT_SIZE) + '''
        ;MsgBox % "offset: " offset " textSize: " textSize " isFinal: " isFinal

        pipe.RawWrite(&text + offset, isFinal ? textSize : ''' + str(_TEXT_SIZE) + ''')
        pipe.RawWrite(&_PY_EOM_BYTES + (isFinal ? +0 : +1), 1)  ; :Utf16Internals :IsFinal
        pipe.Write(_PY_SEPARATOR)  ; :Eom
        pipe.RawWrite(&_PY_EOM_BYTES +2, 1)  ; :OneByteNewline

        pipe.Read(0)
    }

    _Py_MsgMore(ByRef wParam, ByRef lParam, ByRef msg, ByRef hwnd) {
        global _pyStdOut, _pyOutText, _pyOutOffset, _pyStdErr, _pyErrText, _pyErrOffset
        SetBatchLines, -1
        ''' + _comment_debug() + '''DebugMsg(wParam, msg)

        numRead := ''' + str(_TEXT_SIZE) + '''
        _Py_Response(_pyStdOut, _pyOutText, _pyOutOffset += numRead, False)
        _Py_Response(_pyStdErr, _pyErrText, _pyErrOffset += numRead, False)
        return 1  ; :MsgReturn
    }

    ; we can't peek() stdout/stderr, so always write to both or we will over-read and hang waiting
    _Py_StdOut(ByRef outText, ByRef onMain := False) {
        global _pyStdOut, _pyOutText, _pyOutOffset, _pyStdErr, _pyErrText, _pyErrOffset
        _Py_Response(_pyStdOut, _pyOutText := outText, _pyOutOffset := 0, onMain)
        _Py_Response(_pyStdErr, _pyErrText := "", _pyErrOffset := 0, onMain)
    }

    _Py_StdErr(ByRef name, ByRef errText, onMain := False) {
        global _pyStdOut, _pyOutText, _pyOutOffset, _pyStdErr, _pyErrText, _pyErrOffset, _PY_SEPARATOR
        _Py_Response(_pyStdOut, _pyOutText := "", _pyOutOffset := 0, onMain)
        _Py_Response(_pyStdErr, _pyErrText := name _PY_SEPARATOR errText, _pyErrOffset := 0, onMain)
    }

    _Py_MsgCopyData(ByRef wParam, ByRef lParam, ByRef msg, ByRef hwnd) {
        global _pyThreadMsgData, _PY_SEPARATOR
        SetBatchLines, -1
        ''' + _comment_debug() + '''DebugMsg(wParam, msg)

        ;dataTypeId := NumGet(lParam + 0*A_PtrSize) ; unneeded atm
        dataSize := NumGet(lParam + 1*A_PtrSize)
        strAddr := NumGet(lParam + 2*A_PtrSize)
        ; limitation of StrGet(): data is truncated after \\0 :NullTerminator
        data := StrGet(strAddr, dataSize, "utf-8")
        ; OutputDebug, Received: '%data%'

        ; Since messages can arrive from multiple threads (e.g. clicking 'Reload' within OBS Studio 'Scripts' window,
        ;  while a timer is also running within said script) we need to keep their input data separate.
        _pyThreadMsgData[wParam] := []
        ; limitation of Parse and StrSplit(): separator must be a single character :Separator
        Loop, Parse, data, % _PY_SEPARATOR
        {
            ; see Python function _to_ahk_str()
            type := RTrim(SubStr(A_LoopField, 1, 5))  ; :TypePrefix
            val := SubStr(A_LoopField, 7)
            ; others are automatic
            if (type = "bool")
                val := val == "True" ? 1 : 0    ; same as True/False
            _pyThreadMsgData[wParam].Push(val)
        }
        return 1  ; :MsgReturn
    }

    ; call on main thread, much worse latency but may be necessary for DllCall() to avoid:
    ;   Error 0x8001010d An outgoing call cannot be made since the application is dispatching an input-synchronous call.
    _Py_MsgFMain(ByRef wParam, ByRef lParam, ByRef msg, ByRef hwnd) {
        global _pyMsgFMainData
        SetBatchLines, -1
        ''' + _comment_debug() + '''DebugMsg(wParam, msg)

        _pyMsgFMainData.Push(hwnd)
        _pyMsgFMainData.Push(msg)
        _pyMsgFMainData.Push(lParam)
        _pyMsgFMainData.Push(wParam)
        ''' + _comment_debug() + '''OutputDebug, SENDING TO MAIN THREAD
        ; continue on main thread at below label
        ;  ordinarily a new message can interrupt this, but none will be sent because of our lock
        SetTimer, _Py_MsgFMain, -0 ; negative for one-time, and 0 is indeed quicker than 1
        return 1  ; :MsgReturn
    }

    _Py_MsgF(ByRef wParam, ByRef lParam, ByRef msg, ByRef hwnd, ByRef onMain := False) {
        global _pyThreadMsgData, _pyUserBatchLines, _PY_SEPARATOR
        SetBatchLines, -1
        ''' + _comment_debug() + '''if (not onMain)
        ''' + _comment_debug() + '''    DebugMsg(wParam, msg)

        func := _pyThreadMsgData[wParam].RemoveAt(1)
        if (not IsFunc(func)) {
            _pyThreadMsgData.Delete(wParam)
            _Py_StdErr("''' + AhkFuncNotFoundError.__name__ + '''", func, onMain)
            return 1  ; :MsgReturn
        }
        needResult := _pyThreadMsgData[wParam].RemoveAt(1)

        SetBatchLines, % _pyUserBatchLines
        try result := %func%(_pyThreadMsgData[wParam]*)
        catch e {
            SetBatchLines, -1
            _pyThreadMsgData.Delete(wParam)

            ; Exception() just results in a normal object; no easy way to distinguish
            ; https://www.autohotkey.com/docs/commands/Throw.htm
            ; https://web.archive.org/web/20201202074148/https://www.autohotkey.com/boards/viewtopic.php?t=44081
            isExceptionObj := IsObject(e) and (e.Count() == 4 or e.Count() == 5) and e.HasKey("Message") and e.HasKey("What") and e.HasKey("File") and e.HasKey("Line")
            if (isExceptionObj and e.Count() == 5 and !e.HasKey("Extra"))
                isExceptionObj := False

            if (!isExceptionObj)
                e := {Message: e}

            ;MsgBox, % "Message`n" e.Message "`n`nWhat`n" e.What "`n`nExtra`n" e.Extra "`n`nFile`n" e.File "`n`nLine`n" e.Line
            _Py_StdErr("''' + AhkUserException.__name__ + '''"
                , isExceptionObj _PY_SEPARATOR e.Message _PY_SEPARATOR e.What _PY_SEPARATOR e.Extra _PY_SEPARATOR e.File _PY_SEPARATOR e.Line
                , onMain)
            return 1  ; :MsgReturn
        }

        SetBatchLines, -1
        _pyThreadMsgData.Delete(wParam)
        _Py_StdOut(needResult ? result : "", onMain)
        return 1  ; :MsgReturn
    }

    _Py_MsgGet(ByRef wParam, ByRef lParam, ByRef msg, ByRef hwnd) {
        local name, val
        SetBatchLines, -1
        ''' + _comment_debug() + '''DebugMsg(wParam, msg)
        name := _pyThreadMsgData[wParam].RemoveAt(1)
        val := %name%
        _pyThreadMsgData.Delete(wParam)
        _Py_StdOut(val)
        return 1  ; :MsgReturn
    }

    _Py_MsgSet(ByRef wParam, ByRef lParam, ByRef msg, ByRef hwnd) {
        local name
        SetBatchLines, -1
        ''' + _comment_debug() + '''DebugMsg(wParam, msg)
        name := _pyThreadMsgData[wParam].RemoveAt(1)
        %name% := _pyThreadMsgData[wParam].RemoveAt(1)
        _pyThreadMsgData.Delete(wParam)
        return 1  ; :MsgReturn
    }

    _Py_MsgExit() {
        ExitApp
        return 1  ; required even after ExitApp :MsgReturn
    }

    DebugMsg(wParam, msg) {
        OutputDebug, % Format("msg {:#06x}\t\tthread {:#05} -> {:#05}\t\tprocess {:#05} -> {:#05}"
            , msg, wParam, DllCall("GetCurrentThreadId"), ''' + str(_python_pid) + ''', DllCall("GetCurrentProcessId"))
    }

    _pyThreadMsgData := {}
    _pyMsgFMainData := []

    ; these all must return non-zero to signal completion
    ; https://www.autohotkey.com/docs/v1/lib/OnMessage.htm#What_the_Callback_Should_Return  :MsgReturn
    OnMessage(''' + str(win32con.WM_COPYDATA) + ''', Func("_Py_MsgCopyData"))
    OnMessage(''' + str(_Msg.GET) + ''', Func("_Py_MsgGet"))
    OnMessage(''' + str(_Msg.SET) + ''', Func("_Py_MsgSet"))
    OnMessage(''' + str(_Msg.F) + ''', Func("_Py_MsgF"))
    OnMessage(''' + str(_Msg.F_MAIN) + ''', Func("_Py_MsgFMain"))
    OnMessage(''' + str(_Msg.MORE) + ''', Func("_Py_MsgMore"))
    OnMessage(''' + str(_Msg.EXIT) + ''', Func("_Py_MsgExit"))

    _Py_StdOut(A_ScriptHwnd)

    SetBatchLines, % _pyUserBatchLines
    Func("AutoExec").Call() ; call if exists
    _pyUserBatchLines := A_BatchLines

    _Py_StdOut("Initialized")
    return

    ; from _Py_MsgFMain()
    _Py_MsgFMain:
        SetBatchLines, -1
        ''' + _comment_debug() + '''OutputDebug, RECEIVED IN MAIN THREAD
        _Py_MsgF(_pyMsgFMainData.Pop(), _pyMsgFMainData.Pop(), _pyMsgFMainData.Pop(), _pyMsgFMainData.Pop(), True)
    return

    ; an unused label so #Warn won't complain that the script's auto-execute section is unreachable
    ; it is intentionally unreachable (we use `AutoExec()` instead) so scripts can run exclusive standalone code
    _Py_SuppressUnreachableWarning:
    AutoTrim, % A_AutoTrim          ; does nothing and never called, but makes label happy
    '''

    def __init__(self, script: str = "", ahk_path: Path = None, execute_from: Path = None, halt_process_tree_on_exit: bool = False) -> None:
        """Launch an AutoHotkey process.

        :param script: Actual AutoHotkey script. Optional if you only need built-in functions and variables.
        :param ahk_path: Path to an alternative AutoHotkey executable than the one included (`ahk.get('A_AhkVersion')`).
        :param execute_from: Path AutoHotkey executable will be hard-linked/copied to, for the benefit of remembered show/hide status in system tray.
        :param halt_process_tree_on_exit: Descendants of AutoHotkey process will inherit its win32 job object and terminate with it.
            `Script.exit()` (an intentional exit) can override this.
            *Caution*: Universal Windows Platform (UWP) apps (e.g. Windows 10+'s notepad.exe and calc.exe) discard our job object;
            suggest using AutoHotkey's `OnExit()` in those cases: https://github.com/CodeOptimist/ahkunwrapped/issues/1
        """
        self._file = None
        self._script = script
        self._halt_process_tree_on_exit = halt_process_tree_on_exit

        if ahk_path is None:
            ahk_path = _PACKAGE_PATH / r'lib\AutoHotkey\AutoHotkey.exe'
            if _IN_PYINSTALLER and not ahk_path.is_file():
                raise FileNotFoundError(f"""Couldn't find AutoHotkey at '{ahk_path}'.
\tEdit your `.spec` file (may have been auto-generated) to contain:
\t    from pathlib import Path                                      # add these to the top
\t    import ahkunwrapped                                           # add these to the top
\t
\t    a = Analysis(
\t        datas=[                                                   # find this
\t            (Path(ahkunwrapped.__file__).parent / 'lib', 'lib'),  # add this
\t            ('your_script.ahk', '.'),                             # add this if using `Script.from_file()`
\t        ],
\tAnd run it directly: `pyinstaller example.spec` or it will be overwritten (undoing your changes).""")

        if not ahk_path.is_file():
            raise FileNotFoundError(f"Couldn't find file '{ahk_path}' for `ahk_path`.")

        if execute_from is not None:
            execute_from_dir = Path(execute_from)
            if not execute_from_dir.is_dir():
                raise NotADirectoryError(f"Couldn't find folder '{execute_from_dir}' for `execute_from`.")
            ahk_into_folder = execute_from_dir / ahk_path.name

            if ahk_into_folder.exists() and ahk_into_folder.stat().st_mtime != ahk_path.stat().st_mtime:
                ahk_into_folder.unlink(missing_ok=True)

            try:
                ahk_into_folder.hardlink_to(ahk_path)
            except FileExistsError:
                pass
            except OSError as ex:
                if ex.winerror in (winerror.ERROR_ACCESS_DENIED, winerror.ERROR_NOT_SAME_DEVICE):
                    shutil.copyfile(ahk_path, ahk_into_folder)
            ahk_path = ahk_into_folder

        # An asterisk to read the script from standard input. (It will safely terminate if an exception is raised beforehand.)
        # (User script exceptions are already caught and sent to stderr, so `/ErrorStdOut` would only affect debugging CORE.)
        #  cmd = [str(ahk_path), "/ErrorStdOut=utf-16-raw", "*"]
        cmd = [str(ahk_path), "/CP65001", "*"]  # utf-8

        self._popen = subprocess.Popen(cmd, bufsize=Script._BUFFER_SIZE, executable=str(ahk_path),
                                       stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        assert self._popen.stdin and self._popen.stdout and self._popen.stderr
        self._ahk_stdin = io.TextIOWrapper(self._popen.stdin, encoding='utf-8', write_through=True)
        self._ahk_stdout: IO[bytes] = self._popen.stdout
        self._ahk_stderr: IO[bytes] = self._popen.stderr

        # NOTE: PROCESS EXPLORER WILL MISLEAD BY SHOWING ONE OR THE OTHER JOB BUT NOT BOTH @CodeOptimist 2022-10
        # https://learn.microsoft.com/en-gb/windows/win32/api/winbase/nf-winbase-createjobobjecta
        # job containing all AutoHotkey processes to terminate with Python
        Script._python_job_obj = win32job.CreateJobObject(None, f"ahkUnwrapped:python.exe:{Script._python_pid}")  # will find existing or create
        extended_info = win32job.QueryInformationJobObject(Script._python_job_obj, win32job.JobObjectExtendedLimitInformation)
        # silent breakaway so child processes won't inherit job object
        extended_info['BasicLimitInformation']['LimitFlags'] = win32job.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE | win32job.JOB_OBJECT_LIMIT_SILENT_BREAKAWAY_OK
        win32job.SetInformationJobObject(Script._python_job_obj, win32job.JobObjectExtendedLimitInformation, extended_info)

        # Both job objects terminate when their last handle closes (Python exits), but here KILL_ON_JOB_CLOSE (for descendants) is optional.
        # Separately, we can force terminate at any time. :TerminateJob
        self._tree_job_obj = win32job.CreateJobObject(None, f"ahkUnwrapped:AutoHotkey.exe:{self._popen.pid}")  # new job for descendants (and ourself)
        extended_info = win32job.QueryInformationJobObject(self._tree_job_obj, win32job.JobObjectExtendedLimitInformation)
        # no breakaway; this job object will be inherited
        if self._halt_process_tree_on_exit:
            extended_info['BasicLimitInformation']['LimitFlags'] |= win32job.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
        win32job.SetInformationJobObject(self._tree_job_obj, win32job.JobObjectExtendedLimitInformation, extended_info)

        @contextmanager
        def get_handle(handle: int) -> Iterator[int]:
            try:
                yield handle
            finally:
                win32api.CloseHandle(handle)

        # both flags required
        with get_handle(win32api.OpenProcess(win32con.PROCESS_TERMINATE | win32con.PROCESS_SET_QUOTA, False, self._popen.pid)) as ahk_handle:
            win32job.AssignProcessToJobObject(Script._python_job_obj, ahk_handle)  # this one needs to be first to avoid 'Access denied', also see :AvoidJobRace
            win32job.AssignProcessToJobObject(self._tree_job_obj, ahk_handle)  # no race here, AutoHotkey won't `Run` a child process before "Initialized"

        self._ahk_stdin.write(Script._CORE)
        self._ahk_stdin.write(self._script)
        self._ahk_stdin.close()

        self._lock = None
        self._hwnd = int(self._read_response(), 16)
        assert self._read_response() == "Initialized"
        self._lock = threading.Lock()

        # last to make sure things went okay since it runs on its own thread
        atexit.register(self._on_python_exit)  # if we exit, exit AutoHotkey

    @staticmethod
    def from_file(path: Path, format_dict: Mapping[str, str] = None, ahk_path: Path = None, execute_from: Path = None,
                  halt_process_tree_on_exit: bool = False) -> 'Script':  # :FromFile
        """Launch an AutoHotkey process from a script file.

        :param path: Path to file.
        :param format_dict: `.format()` dict to use {{variable}} within script. `globals()` is a common choice.
        :param ahk_path: See `Script()`.
        :param execute_from: See `Script()`.
        :param halt_process_tree_on_exit: See `Script()`.
        """
        script = path.read_text(encoding='utf-8')
        if format_dict is not None:
            # `format()` will mistake function braces as placeholders, so escape those first
            script = script.replace(r'{', r'{{').replace(r'}', r'}}')  # {foo} -> {{foo}}
            # we instead support double braces, so make them single for `format()`
            script = script.replace(r'{{{', r'').replace(r'}}}', r'')  # {{bar}} -> {{{{bar}}}} -> {bar}
            script = script.format(**format_dict)
        script = Script(script, ahk_path, execute_from, halt_process_tree_on_exit)
        script._file = path  # for exceptions
        return script

    def _read_pipes(self) -> tuple[str, str]:
        err, out = bytearray(), bytearray()
        while True:
            def end_of_message(bytearray_: bytearray) -> bool:
                self.poll()
                return bytearray_.endswith(Script._EOM)  # :Eom

            # we're careful not to over-read into the next response,
            # but we can at least go line by line since we always end with \n
            err_buffer, out_buffer = bytearray(), bytearray()
            while not end_of_message(out_buffer):
                out_buffer += self._ahk_stdout.readline()  # :OneByteNewline
            while not end_of_message(err_buffer):
                err_buffer += self._ahk_stderr.readline()

            err += err_buffer[:-Script._EOM_SIZE]  # : Eom
            out += out_buffer[:-Script._EOM_SIZE]

            bool_pos = -Script._EOM_SIZE
            is_final = out_buffer[bool_pos] and err_buffer[bool_pos]  # :IsFinal
            if is_final:
                break
            self._send_message(Script._Msg.MORE)
        if self._lock is not None:
            self._lock.release()
        return err.decode('utf-16-le'), out.decode('utf-16-le')

    def _read_response(self) -> str:
        err, out = self._read_pipes()
        if err:
            name, args = err.split(Script.SEPARATOR, 1)

            exception_class = next((ex for ex in (*AhkError.__subclasses__(), *AhkException.__subclasses__(), AhkException) if ex.__name__ == name), None)
            if exception_class:
                exception = exception_class(*args.split(Script.SEPARATOR))
                if isinstance(exception, AhkUserException):
                    if exception.from_exception_obj and Script._is_num(exception.line):
                        exception.file = self._file or exception.file
                        exception.line = int(exception.line) - Script._CORE.count('\n')

                        if exception.message == '2147549453':
                            exception.message = '0x8001010D - An outgoing call cannot be made since the application is dispatching an input-synchronous call.'
                        if exception.message.startswith('0x8001010D - '):
                            outer_msg = 'Failed a remote procedure call from OnMessage() thread. Solve this with f_main(), call_main() or f_raw_main().'
                            raise AhkCantCallOutInInputSyncCallError(outer_msg) from exception
                    else:
                        warn(AhkCaughtNonExceptionWarning(exception), stacklevel=4)
                raise exception

            warning_class = next((w for w in (*AhkWarning.__subclasses__(), AhkWarning) if w.__name__ == name), None)
            if warning_class:
                warning = warning_class(*args.split(Script.SEPARATOR))
                warn(warning, stacklevel=4)

        return out

    # https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-sendmessage
    def _send_message(self, msg: int, lparam: bytes = None) -> None:
        # this is essential because messages are ignored if we're uninterruptible (e.g. in a menu)
        # wparam is normally source window handle, but in our case source thread id
        # noinspection PyTypeChecker
        while not win32api.SendMessage(self._hwnd, msg, threading.get_ident(), lparam):
            self.poll()
            time.sleep(0.01)

    def _send(self, msg: int, data: Sequence[Primitive]) -> None:
        data_str = Script.SEPARATOR.join(Script._to_ahk_str(v) for v in data)
        # OutputDebugString(f"Sent: {data}")
        # https://learn.microsoft.com/en-us/windows/win32/dataxchg/wm-copydata
        char_buffer = array.array('b', bytes(data_str, 'utf-8'))
        addr, size = char_buffer.buffer_info()
        data_type_id = msg  # anything; unneeded atm
        struct_ = struct.pack('PLP', data_type_id, size, addr)
        self._send_message(win32con.WM_COPYDATA, struct_)
        self._lock.acquire(blocking=True)  # set `False` to witness threads test failure :TestThreads
        self._send_message(msg)

    @staticmethod
    def _to_ahk_str(val: Primitive) -> str:
        if isinstance(val, float):
            if math.isnan(val) or math.isinf(val):
                raise AhkUnsupportedValueError(val)
            val_str = f'{val:.6f}'  # 6 decimal precision to match AutoHotkey
            if float(val_str) != val:
                warn(AhkLossOfPrecisionWarning(val, val_str), stacklevel=6)
            val_str = val_str.rstrip('0').rstrip('.')  # less text to send the better
        else:
            if isinstance(val, str):
                if '\x00' in val:  # :NullTerminator
                    raise AhkUnsupportedValueError(r"string contains null terminator '\x00' which AutoHotkey ignores characters beyond")
                if Script.SEPARATOR in val:  # :Separator
                    raise AhkUnsupportedValueError(f'string contains {repr(Script.SEPARATOR)} which is reserved for messages to AutoHotkey')
            val_str = str(val)
        return f"{type(val).__name__[:5]:<5} {val_str}"  # padded len(5) :TypePrefix

    def _f(self, msg: int, name: str, *args: Primitive, need_result: bool, coerce_result: bool = False) -> Primitive:
        self._send(msg, [name, need_result] + list(args))
        response = self._read_response()
        return self._from_ahk_str(response) if coerce_result else response

    def call(self, name: str, *args: Primitive) -> None:
        """Call a script function without receiving the result, if any. Least latency."""
        self._f(Script._Msg.F, name, *args, need_result=False)

    def call_main(self, name: str, *args: Primitive) -> None:
        """Same as `call()` but executed on AutoHotkey's main thread.
        Worse latency, but solution to `AhkCantCallOutInInputSyncCallError`."""
        self._f(Script._Msg.F_MAIN, name, *args, need_result=False)

    def f_raw(self, name: str, *args: Primitive) -> str:
        """Call a script function and return the result as its raw string (don't mimic AutoHotkey's type inference)."""
        return self._f(Script._Msg.F, name, *args, need_result=True)

    def f_raw_main(self, name: str, *args: Primitive) -> str:
        """Same as `f_raw()` but executed on AutoHotkey's main thread.
        Worse latency, but solution to `AhkCantCallOutInInputSyncCallError`."""
        return self._f(Script._Msg.F_MAIN, name, *args, need_result=True)

    def f(self, name: str, *args: Primitive) -> Primitive:
        """Call a script function and return the result."""
        return self._f(Script._Msg.F, name, *args, need_result=True, coerce_result=True)

    def f_main(self, name: str, *args: Primitive) -> Primitive:
        """Same as `f()` but executed on AutoHotkey's main thread.
        Worse latency, but solution to `AhkCantCallOutInInputSyncCallError`."""
        return self._f(Script._Msg.F_MAIN, name, *args, need_result=True, coerce_result=True)

    @staticmethod
    def _is_num(str_: str) -> bool:
        return str_.isdigit() or (str_.startswith('-') and str_[1:].isdigit())

    @staticmethod
    def _from_ahk_str(str_: str) -> Primitive:
        is_hex = str_.startswith('0x') and all(c in string.hexdigits for c in str_[2:])
        if is_hex:
            return int(str_, 16)

        if Script._is_num(str_):
            return int(str_.lstrip('0') or '0', 0)
        if Script._is_num(str_.replace('.', '', 1)):
            return float(str_)
        return str_

    def get_raw(self, name: str) -> str:
        """Get a global script variable or built-in as its raw string (don't mimic AutoHotkey's type inference)."""
        self._send(Script._Msg.GET, [name])
        return self._read_response()

    def get(self, name: str) -> Primitive:
        """Get a global script variable or built-in like `A_TimeIdle`."""
        self._send(Script._Msg.GET, [name])
        return Script._from_ahk_str(self._read_response())

    def set(self, name: str, val: Primitive) -> None:
        """Set a global script variable."""
        # Every _send() will lock, so others are finished before we set().
        #  We don't need a confirmation response, just the ensurance that it finishes before others begin.
        self._send(Script._Msg.SET, [name, val])
        self._lock.release()  # normally done within `_read_response()`

    # if AutoHotkey is terminated, get error code
    def poll(self) -> None:
        """Detect when AutoHotkey process exits, typically within a loop, by raising `AhkExitException`.
        (Only needed in contexts without other Script functions, as they all run this internally.)"""
        exit_code = self._popen.poll()
        if exit_code is not None:
            # OutputDebugString(f"Exit code: {exit_code}; call stack: {traceback.format_stack()}")
            atexit.unregister(self._on_python_exit)
            raise AhkExitException(exit_code)

    def _on_python_exit(self) -> None:
        with suppress(AhkExitException):  # Expected and not exceptional.
            self.exit()

    def exit(self, timeout: float = 5.0, halt_descendants: bool = None) -> None:
        """Ask AutoHotkey to exit cleanly (remove system tray icon, etc.).
        To my knowledge only an `OnExit()` callback could delay this.

        :param timeout: Seconds to wait before terminating. `None` for infinity.
        :param halt_descendants: Uses `Script()`'s `halt_process_tree_on_exit` (default `False`) unless overridden here.
        """

        if halt_descendants is None:
            halt_descendants = self._halt_process_tree_on_exit

        # No need to `&= ~KILL_ON_JOB_CLOSE` if `halt_descendants` is `False` and `self.halt_process_tree_on_exit` is `True`
        #  because jobs only *automatically* terminate when *Python* exits (job handle closes), not AutoHotkey by itself.

        atexit.unregister(self._on_python_exit)

        exit_code = None
        try:
            try:
                # clean; removes tray icons etc.
                # OutputDebugString(f"Sending ExitApp from thread {threading.get_ident()}")
                self._send_message(Script._Msg.EXIT)
            except AhkExitException as ex:  # exited immediately
                exit_code = ex.args[0]  # for 'finally'
                raise

            exit_code = self._popen.wait(timeout)
            raise AhkExitException(exit_code)  # exited after a delay, before timeout
        except TimeoutExpired as ex:  # never exited before timeout
            self._popen.terminate()
            exit_code = 1
            raise AhkExitException(exit_code) from ex
        finally:
            if halt_descendants:
                win32job.TerminateJobObject(self._tree_job_obj, exit_code)  # :TerminateJob
