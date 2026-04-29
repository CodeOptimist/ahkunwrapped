# SPDX-License-Identifier: AGPL-3.0-or-later
# Copyright (c) 2019-2026 Christopher S. Galpin

import atexit
import ctypes
import io
import math
import os
import shutil
import subprocess
import sys
import threading
import time
# import traceback
from collections.abc import Iterator, Mapping, Sequence
from contextlib import contextmanager, suppress
from ctypes import wintypes
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
_INT64_MIN: Final = -(1 << 63)
_INT64_MAX: Final = (1 << 63) - 1

type Primitive = float | int | bool | str


# @formatter:off
class AhkException(Exception): pass                         # noqa: E701
class AhkExitException(AhkException): pass                  # noqa: E701
class AhkError(AhkException): pass                          # noqa: E701
class AhkFuncNotFoundError(AhkError): pass                  # noqa: E701
class AhkUnsupportedValueError(AhkError): pass              # noqa: E701
class AhkCantCallOutInInputSyncCallError(AhkError): pass    # noqa: E701
class AhkWarning(UserWarning): pass                         # noqa: E701
# @formatter:on


@dataclass
class AhkUserException(AhkException):
    from_error_obj: bool
    message: str
    what: str = None
    extra: str = None
    file: str = None
    line: int = None

    def __post_init__(self):
        self.from_error_obj = self.from_error_obj == "1"
        if self.line is not None:
            self.line = int(self.line)

        self.args = (self.message, self.what, self.extra, self.file, self.line)


class AhkCaughtNonErrorWarning(AhkWarning):
    """Warning issued when AutoHotkey throws a primitive value instead of an Error object."""

    def __init__(self, exception: AhkUserException):
        message = f"""got "throw X"; recommend "throw Error(X)" to preserve line info.
\tAlternatively, catch within script https://www.autohotkey.com/docs/v2/lib/Error.htm .
\tException message: '{exception.message}'"""
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
    _EOM_SIZE: Final = 1 + len(_EOM)  # include :IsFinal bool, before EOM
    _IS_FINAL__POS: Final = 0 - _EOM_SIZE
    _TEXT_SIZE: Final[int] = _BUFFER_SIZE - _EOM_SIZE

    _python_pid: Final = os.getpid()
    _python_job_obj: ClassVar[int]

    _CORE: Final[str] = '''
    if (ProcessExist(''' + str(_python_pid) + ''') = 0) ; not found
        ExitApp() ; https://stackoverflow.com/q/73506891/#comment129808240_73506891 :AvoidJobRace
    Persistent()
    A_WorkingDir := "''' + os.getcwd() + '''"
    _PY_SEPARATOR := Chr(''' + str(ord(SEPARATOR)) + ''')

    _PY_NONFINAL_EOM := Buffer(''' + str(_EOM_SIZE) + ''')
    _PY_FINAL_EOM    := Buffer(''' + str(_EOM_SIZE) + ''')
    ; :IsFinal :Eom :OneByteNewline
    NumPut("UChar", 00, "UShort", Ord(_PY_SEPARATOR), "UChar", Ord("`n"), _PY_NONFINAL_EOM)
    NumPut("UChar", 01, "UShort", Ord(_PY_SEPARATOR), "UChar", Ord("`n"), _PY_FINAL_EOM)

    ; Let's write strings as they're stored, avoiding `StrPut()`.
    ; https://www.autohotkey.com/docs/v2/Concepts.htm#string-encoding
    _pyStdOut := FileOpen("*", "w", "utf-16-raw")
    _pyStdErr := FileOpen("**", "w", "utf-16-raw")

    _Py_Response(pipe, text, offset, onMain) {
        static textSize, isFinal
        textSize := Max(StrLen(text) * 2 - offset, 0)
        isFinal := onMain or textSize <= ''' + str(_TEXT_SIZE) + '''
        ;MsgBox("offset: " offset " textSize: " textSize " isFinal: " isFinal)

        pipe.RawWrite(StrPtr(text) + offset, isFinal ? textSize : ''' + str(_TEXT_SIZE) + ''')
        pipe.RawWrite(isFinal ? _PY_FINAL_EOM : _PY_NONFINAL_EOM, ''' + str(_EOM_SIZE) + ''')

        pipe.Read(0)
    }

    _Py_MsgMore(wParam, lParam, msg, hwnd) {
        ''' + _comment_debug() + '''_Py_DebugMsg(wParam, msg)

        static numRead
        numRead := ''' + str(_TEXT_SIZE) + '''
        global _pyOutOffset, _pyErrOffset
        _Py_Response(_pyStdOut, _pyOutText, _pyOutOffset += numRead, False)
        _Py_Response(_pyStdErr, _pyErrText, _pyErrOffset += numRead, False)
        return 1  ; :MsgReturn
    }

    ; we can't peek() stdout/stderr, so always write to both or we will over-read and hang waiting
    _Py_StdOut(outText, onMain := False) {
        global _pyOutText, _pyOutOffset, _pyErrText, _pyErrOffset
        _Py_Response(_pyStdOut, _pyOutText := outText, _pyOutOffset := 0, onMain)
        _Py_Response(_pyStdErr, _pyErrText := "", _pyErrOffset := 0, onMain)
    }

    _Py_StdErr(name, errText, onMain := False) {
        global _pyOutText, _pyOutOffset, _pyErrText, _pyErrOffset
        _Py_Response(_pyStdOut, _pyOutText := "", _pyOutOffset := 0, onMain)
        _Py_Response(_pyStdErr, _pyErrText := name _PY_SEPARATOR errText, _pyErrOffset := 0, onMain)
    }

    _Py_MsgCopyData(wParam, lParam, msg, hwnd) {
        ''' + _comment_debug() + '''_Py_DebugMsg(wParam, msg)

        static size, addr, copyData
        ;extra := NumGet(lParam, 0*A_PtrSize, "Int64") ; unneeded atm
        size := NumGet(lParam, 1*A_PtrSize, "UInt")
        addr := NumGet(lParam, 2*A_PtrSize, "Ptr")
        copyData := StrGet(addr, size, "utf-8")
        ;OutputDebug("Received: '" copyData "'")

        ; Since messages can arrive from multiple threads (e.g. clicking 'Reload' within OBS Studio 'Scripts' window,
        ;  while a timer is also running within said script) we need to keep their input data separate.
        _pyThreadMsgData[wParam] := []
        ; limitation of Parse and StrSplit(): separator must be a single character :Separator
        Loop Parse, copyData, _PY_SEPARATOR
        {
            static type, val
            type := SubStr(A_LoopField, 1, 1)  ; :TypePrefix
            val := SubStr(A_LoopField, 2)

            if (type = "f")
                val := Float(val)
            else if (type = "i")
                val := Integer(val)
            else if (type = "b")
                val := (val == "1")

            _pyThreadMsgData[wParam].Push(val)
        }
        return 1  ; :MsgReturn
    }

    ; Run on main thread, higher latency but may be necessary for `DllCall()` to avoid:
    ;   '(0x8001010D) An outgoing call cannot be made since the application is dispatching an input-synchronous call.'
    _Py_MsgFMain(wParam, lParam, msg, hwnd) {
        ''' + _comment_debug() + '''_Py_DebugMsg(wParam, msg)

        ''' + _comment_debug() + '''OutputDebug("SENDING TO MAIN THREAD")
        ; Ordinarily a new message can interrupt this, but none will be sent because of our lock.
        RunOnMain() {
            ''' + _comment_debug() + '''OutputDebug("RECEIVED IN MAIN THREAD")
            _Py_MsgF(wParam, lParam, msg, hwnd, True)
        }

        SetTimer(RunOnMain, -1)
        return 1  ; :MsgReturn
    }

    _Py_MsgF(wParam, lParam, msg, hwnd, onMain := False) {
        ''' + _comment_debug() + '''if (not onMain)
        ''' + _comment_debug() + '''    _Py_DebugMsg(wParam, msg)

        static funcName, func
        funcName := _pyThreadMsgData[wParam].RemoveAt(1)
        func := unset
        try func := %funcName%

        if not (IsSet(func) and HasMethod(func)) {
            _pyThreadMsgData.Delete(wParam)
            _Py_StdErr("''' + AhkFuncNotFoundError.__name__ + '''", funcName, onMain)
            return 1  ; :MsgReturn
        }
        needResult := _pyThreadMsgData[wParam].RemoveAt(1)

        static ex
        try result := func(_pyThreadMsgData[wParam]*)
        catch Any as ex {
            _pyThreadMsgData.Delete(wParam)
            if (ex is Error)
                _Py_StdErr("''' + AhkUserException.__name__ + '''", True _PY_SEPARATOR ex.Message _PY_SEPARATOR ex.What _PY_SEPARATOR ex.Extra _PY_SEPARATOR ex.File _PY_SEPARATOR ex.Line, onMain)
            else
                _Py_StdErr("''' + AhkUserException.__name__ + '''", False _PY_SEPARATOR (HasMethod(ex, "ToString") ? String(ex) : Type(ex)), onMain)

            return 1  ; :MsgReturn
        }

        _pyThreadMsgData.Delete(wParam)
        _Py_StdOut(needResult ? SubStr(Type(result), 1, 1) . String(result) : "", onMain)
        return 1  ; :MsgReturn
    }

    _Py_MsgGet(wParam, lParam, msg, hwnd) {
        ''' + _comment_debug() + '''_Py_DebugMsg(wParam, msg)

        static name, val
        name := _pyThreadMsgData[wParam].RemoveAt(1)
        val := %name%
        _pyThreadMsgData.Delete(wParam)
        _Py_StdOut(SubStr(Type(val), 1, 1) . String(val))
        return 1  ; :MsgReturn
    }

    _Py_MsgSet(wParam, lParam, msg, hwnd) {
        global
        ''' + _comment_debug() + '''_Py_DebugMsg(wParam, msg)

        static name
        name := _pyThreadMsgData[wParam].RemoveAt(1)
        %name% := _pyThreadMsgData[wParam].RemoveAt(1)
        _pyThreadMsgData.Delete(wParam)
        return 1  ; :MsgReturn
    }

    _Py_MsgExit(wParam, lParam, msg, hwnd) {
        ExitApp()
        return 1  ; required even after `ExitApp()` :MsgReturn
    }

    _Py_DebugMsg(wParam, msg) {
        OutputDebug(Format("msg {:#06x}\t\tPython thread {:#05} -> AHK thread {:#05}\t\tPython process {:#05} -> AHK {:#05}"
            , msg, wParam, DllCall("GetCurrentThreadId"), ''' + str(_python_pid) + ''', DllCall("GetCurrentProcessId")))
    }

    _pyThreadMsgData := Map()

    ; these all must return non-zero to signal completion
    ; https://www.autohotkey.com/docs/v2/lib/OnMessage.htm#What_the_Callback_Should_Return  :MsgReturn
    OnMessage(''' + str(win32con.WM_COPYDATA) + ''', _Py_MsgCopyData)
    OnMessage(''' + str(_Msg.GET) + '''             , _Py_MsgGet)
    OnMessage(''' + str(_Msg.SET) + '''             , _Py_MsgSet)
    OnMessage(''' + str(_Msg.F) + '''               , _Py_MsgF)
    OnMessage(''' + str(_Msg.F_MAIN) + '''          , _Py_MsgFMain)
    OnMessage(''' + str(_Msg.MORE) + '''            , _Py_MsgMore)
    OnMessage(''' + str(_Msg.EXIT) + '''            , _Py_MsgExit)

    _Py_StdOut(String(A_ScriptHwnd))
    try %"Startup"%() ; call if exists
    _Py_StdOut("Initialized")
    return

    ; an unused label so `#Warn` won't complain that the script's auto-execute section is unreachable
    ; it is intentionally unreachable (we use `AutoExec()` instead) so scripts can run exclusive standalone code
    _Py_SuppressUnreachableWarning:
'''

    # [[[cog
    # SHARED_PARAMS = """
    #         :param ahk_path: Path to an alternative AutoHotkey executable than the one included.
    #         :param execute_from: Path AutoHotkey executable will be hard-linked/copied to, for the benefit of remembered show/hide status in the system tray.
    #         :param halt_process_tree_on_exit: Descendants of the AutoHotkey process will inherit its win32 job object and terminate with it.
    #             `Script.exit()` (an intentional exit) can override this.
    #             *Caution*: Universal Windows Platform (UWP) apps (e.g., Windows 10+'s notepad.exe and calc.exe) discard our job object;
    #             suggest using AutoHotkey's `OnExit()` in those cases.
    # """
    # ]]]
    # [[[end]]]
    def __init__(self, script: str = "", ahk_path: Path = None, execute_from: Path = None, halt_process_tree_on_exit: bool = False, *,
                 _file_path: Path = None) -> None:
        """Launch an AutoHotkey process.

        :param script: AutoHotkey script providing user functions and globals. Optional if you only need built-in functions and `A_` variables.

        # [[[cog
        # cog.outl(SHARED_PARAMS)
        # ]]]

        :param ahk_path: Path to an alternative AutoHotkey executable than the one included.
        :param execute_from: Path AutoHotkey executable will be hard-linked/copied to, for the benefit of remembered show/hide status in the system tray.
        :param halt_process_tree_on_exit: Descendants of the AutoHotkey process will inherit its win32 job object and terminate with it.
            `Script.exit()` (an intentional exit) can override this.
            *Caution*: Universal Windows Platform (UWP) apps (e.g., Windows 10+'s notepad.exe and calc.exe) discard our job object;
            suggest using AutoHotkey's `OnExit()` in those cases.

        # [[[end]]]
        """

        self._script = script
        self._halt_process_tree_on_exit = halt_process_tree_on_exit
        self._file_path = _file_path

        self._err, self._out = bytearray(), bytearray()
        self._err_buffer, self._out_buffer = bytearray(), bytearray()

        if ahk_path is None:
            ahk_path = _PACKAGE_PATH / r'lib\AutoHotkey\AutoHotkey64.exe'
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
        cmd = [str(ahk_path), "*"]  # Default is utf-8 https://www.autohotkey.com/docs/v2/Scripts.htm#cp

        self._popen = subprocess.Popen(cmd, bufsize=Script._BUFFER_SIZE, executable=str(ahk_path),
                                       stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        assert self._popen.stdin and self._popen.stdout and self._popen.stderr
        self._ahk_stdin = io.TextIOWrapper(self._popen.stdin, encoding='utf-8', write_through=True)
        self._ahk_stdout: IO[bytes] = self._popen.stdout
        self._ahk_stderr: IO[bytes] = self._popen.stderr

        # NOTE: PROCESS EXPLORER WILL MISLEAD BY SHOWING ONE OR THE OTHER JOB BUT NOT BOTH @CodeOptimist 2022-10
        # https://learn.microsoft.com/en-gb/windows/win32/api/winbase/nf-winbase-createjobobjecta
        # job containing all AutoHotkey processes to terminate with Python
        Script._python_job_obj = win32job.CreateJobObject(None, f"ahkUnwrapped:{Path(sys.executable).name}:{Script._python_pid}")  # will find existing or create
        extended_info = win32job.QueryInformationJobObject(Script._python_job_obj, win32job.JobObjectExtendedLimitInformation)
        # silent breakaway so child processes won't inherit job object
        extended_info['BasicLimitInformation']['LimitFlags'] = win32job.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE | win32job.JOB_OBJECT_LIMIT_SILENT_BREAKAWAY_OK
        win32job.SetInformationJobObject(Script._python_job_obj, win32job.JobObjectExtendedLimitInformation, extended_info)

        # Both job objects terminate when their last handle closes (Python exits), but here KILL_ON_JOB_CLOSE (for descendants) is optional.
        # Separately, we can force terminate at any time. :TerminateJob
        self._tree_job_obj = win32job.CreateJobObject(None, f"ahkUnwrapped:{ahk_path.name}:{self._popen.pid}")  # new job for descendants (and ourselves)
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

        if self._file_path is None:
            preamble = f"""
#NoTrayIcon
"""
        else:
            preamble = f"""
A_IconTip := "{self._file_path.name}"
A_ScriptName := "{self._file_path.name}"
A_WorkingDir := "{self._file_path.parent}"
#Include "{self._file_path.parent}"
"""

        assert Script._CORE[-1] == '\n'
        injected = f"{preamble}{Script._CORE}"
        self._line_offset = injected.count('\n')

        self._ahk_stdin.write(injected)
        self._ahk_stdin.write(self._script)
        self._ahk_stdin.close()

        self._lock = None
        self._hwnd = int(self._read_response())
        assert self._read_response() == "Initialized"
        self._lock = threading.Lock()

        # last to make sure things went okay since it runs on its own thread
        atexit.register(self._on_python_exit)  # if we exit, exit AutoHotkey

    @staticmethod
    def from_file(path: Path, format_dict: Mapping[str, str] = None, ahk_path: Path = None, execute_from: Path = None,
                  halt_process_tree_on_exit: bool = False) -> 'Script':  # :FromFile
        """Launch an AutoHotkey process from a script file.

        Limitations:
        `A_ScriptFullPath` will permanently return `*`.
        `A_ScriptDir` will be the Python host's initial working directory.
        `A_LineFile` will return `*` (except within an `#Include`d file).

        As expected:
        `A_WorkingDir` and `#Include` will be set to the file's parent folder.
        `A_ScriptName` and `A_IconTip` will be set to the file's name.

        :param path: Path to file.
        :param format_dict: `.format()` dict to use '{{variable}}' within the script. `globals()` is a common choice.

        # [[[cog
        # cog.outl(SHARED_PARAMS)
        # ]]]

        :param ahk_path: Path to an alternative AutoHotkey executable than the one included.
        :param execute_from: Path AutoHotkey executable will be hard-linked/copied to, for the benefit of remembered show/hide status in the system tray.
        :param halt_process_tree_on_exit: Descendants of the AutoHotkey process will inherit its win32 job object and terminate with it.
            `Script.exit()` (an intentional exit) can override this.
            *Caution*: Universal Windows Platform (UWP) apps (e.g., Windows 10+'s notepad.exe and calc.exe) discard our job object;
            suggest using AutoHotkey's `OnExit()` in those cases.

        # [[[end]]]
        """
        script = path.read_text(encoding='utf-8')
        if format_dict is not None:
            # `format()` will mistake function braces as placeholders, so escape those first
            script = script.replace(r'{', r'{{').replace(r'}', r'}}')  # {foo} -> {{foo}}
            # we instead support double braces, so make them single for `format()`
            script = script.replace(r'{{{', r'').replace(r'}}}', r'')  # {{bar}} -> {{{{bar}}}} -> {bar}
            script = script.format(**format_dict)
        script = Script(script, ahk_path, execute_from, halt_process_tree_on_exit, _file_path=path)
        return script

    def _read_pipes(self) -> tuple[str, str]:
        def end_of_message(bytearray_: bytearray) -> bool:
            self.poll()
            return bytearray_.endswith(Script._EOM)  # :Eom

        self._err.clear()
        self._out.clear()

        while True:
            # we're careful not to over-read into the next response,
            # but we can at least go line by line since we always end with `\n`
            self._err_buffer.clear()
            self._out_buffer.clear()

            while not end_of_message(self._out_buffer):
                self._out_buffer += self._ahk_stdout.readline()  # :OneByteNewline
            while not end_of_message(self._err_buffer):
                self._err_buffer += self._ahk_stderr.readline()

            self._err += self._err_buffer[:-Script._EOM_SIZE]  # : Eom
            self._out += self._out_buffer[:-Script._EOM_SIZE]

            if self._out_buffer[Script._IS_FINAL__POS] and self._err_buffer[Script._IS_FINAL__POS]:  # :IsFinal
                break
            self._send_message(Script._Msg.MORE)
        if self._lock is not None:
            self._lock.release()
        return self._err.decode('utf-16-le'), self._out.decode('utf-16-le')

    def _read_response(self) -> str:
        err, out = self._read_pipes()
        if err:
            name, args = err.split(Script.SEPARATOR, 1)

            exception_class = next((ex for ex in (*AhkError.__subclasses__(), *AhkException.__subclasses__(), AhkException) if ex.__name__ == name), None)
            if exception_class:
                exception = exception_class(*args.split(Script.SEPARATOR))
                if isinstance(exception, AhkUserException):
                    if exception.from_error_obj:
                        if exception.file == '*':
                            if self._file_path is not None:
                                exception.file = str(self._file_path.resolve())
                            exception.line -= self._line_offset

                        if exception.message == "2147549453":
                            exception.message = "(0x8001010D) An outgoing call cannot be made since the application is dispatching an input-synchronous call."
                        if exception.message.startswith("(0x8001010D)"):
                            outer_msg = "Failed a remote procedure call from `OnMessage()` thread. Solve this with `f_main()`, or `call_main()`."
                            raise AhkCantCallOutInInputSyncCallError(outer_msg) from exception
                    else:
                        warn(AhkCaughtNonErrorWarning(exception), stacklevel=5)
                raise exception

            warning_class = next((w for w in (*AhkWarning.__subclasses__(), AhkWarning) if w.__name__ == name), None)
            if warning_class:
                warning = warning_class(*args.split(Script.SEPARATOR))
                warn(warning, stacklevel=5)

        return out

    # https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-sendmessage
    def _send_message(self, msg: int, lparam: int = None) -> None:
        # This is essential because messages are ignored if we're uninterruptible (e.g., in a menu).
        # wparam is normally the source window handle, but in our case source thread id.
        #  (We can't put it in the first member of COPYDATASTRUCT because ALL messages need it.)
        # noinspection PyTypeChecker
        while not win32api.SendMessage(self._hwnd, msg, threading.get_ident(), lparam):
            self.poll()
            time.sleep(0.01)

    def _send(self, msg: int, data: Sequence[Primitive]) -> None:
        # OutputDebugString(f"Sending: {data}")
        # https://learn.microsoft.com/en-us/windows/win32/api/winuser/ns-winuser-copydatastruct
        class COPYDATASTRUCT(ctypes.Structure):
            _fields_ = [
                ('dwData', wintypes.WPARAM),
                ('cbData', wintypes.DWORD),
                ('lpData', ctypes.c_char_p),
            ]

        data_str = Script.SEPARATOR.join(Script._to_ahk_str(v) for v in data)
        data_bytes = data_str.encode('utf-8')
        cds = COPYDATASTRUCT(0, len(data_bytes), ctypes.c_char_p(data_bytes))
        self._send_message(win32con.WM_COPYDATA, ctypes.addressof(cds))
        assert self._lock is not None
        self._lock.acquire(blocking=True)  # set `False` to witness threads test failure :TestThreads
        self._send_message(msg)

    @staticmethod
    def _to_ahk_str(val: Primitive) -> str:
        if isinstance(val, float):
            if math.isnan(val) or math.isinf(val):
                raise AhkUnsupportedValueError(val)
        elif isinstance(val, int) and not isinstance(val, bool):
            if val > _INT64_MAX or val < _INT64_MIN:
                raise AhkUnsupportedValueError(f"integer {val} exceeds AutoHotkey's signed 64-bit limits ({_INT64_MIN} to {_INT64_MAX})")
        elif isinstance(val, bool):
            return f"b{int(val)}"
        elif isinstance(val, str):
            if '\x00' in val:  # :NullTerminator
                raise AhkUnsupportedValueError(r"string contains null terminator '\x00' which AutoHotkey ignores characters beyond")
            if Script.SEPARATOR in val:  # :Separator
                raise AhkUnsupportedValueError(f"string contains {repr(Script.SEPARATOR)} which is reserved for messages to AutoHotkey")
        return f"{type(val).__name__[:1]}{val}"  # :TypePrefix

    def _f(self, msg: int, name: str, *args: Primitive, need_result: bool) -> Primitive:
        self._send(msg, [name, need_result] + list(args))
        return self._from_ahk_str()

    def call(self, name: str, *args: Primitive) -> None:
        """Call a script function without receiving the result, if any. Lowest latency."""
        self._f(Script._Msg.F, name, *args, need_result=False)

    def call_main(self, name: str, *args: Primitive) -> None:
        """Same as `call()` but executed on AutoHotkey's main thread.
        Higher latency, but solution to `AhkCantCallOutInInputSyncCallError`."""
        self._f(Script._Msg.F_MAIN, name, *args, need_result=False)

    def f(self, name: str, *args: Primitive) -> Primitive:
        """Call a script function and return the result."""
        return self._f(Script._Msg.F, name, *args, need_result=True)

    def f_main(self, name: str, *args: Primitive) -> Primitive:
        """Same as `f()` but executed on AutoHotkey's main thread.
        Higher latency, but solution to `AhkCantCallOutInInputSyncCallError`."""
        return self._f(Script._Msg.F_MAIN, name, *args, need_result=True)

    def _from_ahk_str(self) -> Primitive:
        str_ = self._read_response()
        if str_:
            match str_[0]:
                case 'F':
                    return float(str_[1:])
                case 'I':
                    return int(str_[1:])
                case 'B':
                    return str_[1:] == 1
                case 'S':
                    return str_[1:]
        return str_

    def get(self, name: str) -> Primitive:
        """Get a global script variable or built-in like `A_TimeIdle`."""
        self._send(Script._Msg.GET, [name])
        return self._from_ahk_str()

    def set(self, name: str, val: Primitive) -> None:
        """Set a global script variable, or some built-ins like `A_Clipboard`."""
        # Every `_send()` will lock, so others are finished before we `set()`.
        #  We don't need a confirmation response, just the ensurance that it finishes before others begin.
        self._send(Script._Msg.SET, [name, val])
        assert self._lock is not None
        self._lock.release()  # normally done within `_read_response()`

    # if AutoHotkey is terminated, get error code
    def poll(self) -> None:
        """Detect when the AutoHotkey process exits, typically within a loop, by raising `AhkExitException`.
        (Only needed in contexts without other `Script` functions, as they all run this internally.)"""
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
        To my knowledge, only an `OnExit()` callback could delay this.

        :param timeout: Seconds to wait before terminating. `None` for infinity.
        :param halt_descendants: Uses `Script()`'s `halt_process_tree_on_exit` (default `False`) unless overridden here.
        """

        if halt_descendants is None:
            halt_descendants = self._halt_process_tree_on_exit

        # No need to `&= ~KILL_ON_JOB_CLOSE` if `halt_descendants` is `False` and `self.halt_process_tree_on_exit` is `True`
        #  because jobs only *automatically* terminate when *Python* exits (the job handle closes), not AutoHotkey by itself.

        atexit.unregister(self._on_python_exit)

        exit_code = None
        try:
            try:
                # clean; removes tray icons etc.
                # OutputDebugString(f"Sending `ExitApp()` from thread {threading.get_ident()}")
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
