# Copyright (C) 2019-2022  Christopher S. Galpin.  Licensed under AGPL-3.0-or-later.  See /NOTICE.
import array
import atexit
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
from contextlib import suppress, contextmanager
from itertools import chain
from pathlib import Path
from subprocess import TimeoutExpired
from typing import ClassVar, Mapping, Optional, Sequence, Tuple, Union, ContextManager
from warnings import warn

import pywintypes
import win32api
import win32con
import win32job
# noinspection PyUnresolvedReferences
from win32api import OutputDebugString

IN_PYINSTALLER = getattr(sys, 'frozen', False)
# noinspection PyProtectedMember,PyUnresolvedReferences
PACKAGE_PATH = Path(sys._MEIPASS) if IN_PYINSTALLER else Path(__file__).parent
SINGLE_JOB_ASSIGNMENTS = sys.getwindowsversion().major < 8  # https://stackoverflow.com/q/13449531/
if SINGLE_JOB_ASSIGNMENTS:
    import inspect


class AhkException(Exception): pass                         # noqa: E701
class AhkExitException(AhkException): pass                  # noqa: E701
class AhkError(AhkException): pass                          # noqa: E701
class AhkFuncNotFoundError(AhkError): pass                  # noqa: E701
class AhkUnsupportedValueError(AhkError): pass              # noqa: E701
class AhkCantCallOutInInputSyncCallError(AhkError): pass    # noqa: E701
class AhkWarning(UserWarning): pass                         # noqa: E701


class AhkLossOfPrecisionWarning(AhkWarning):
    def __init__(self, val: float, val_str: str):
        super().__init__(f'loss of precision from {val} to {val_str}')


# Python 3.7 would use @dataclass
class AhkUserException(AhkException):
    def __init__(self, from_exception_obj: str, message: str, what: str, extra: str, file: str, line: str):
        self.from_exception_obj: bool = from_exception_obj == "1"
        self.message: str = message
        self.what: str = what
        self.extra: str = extra
        self.file: str = file
        self.line: str = line

    def __str__(self) -> str:
        # Python 3.8 would use return f"{message=}, {what=}, {extra=}, {file=}, {line=}"
        return f"(message={repr(self.message)}, what={repr(self.what)}, extra={repr(self.extra)}, file={repr(self.file)}, line={repr(self.line)})"

    def __repr__(self) -> str:
        return f"{self.__class__.__name__}{self}"


class AhkCaughtNonExceptionWarning(AhkWarning):
    def __init__(self, exception: AhkUserException):
        message = f"""got "throw X"; recommend "throw Exception(X)" to preserve line info.
\tAlternatively, catch within script https://www.autohotkey.com/docs/commands/Throw.htm#Exception .
\tException message: '{exception.message}'"""
        if not exception.message:
            message += "\n\tMay have been an AutoHotkey object e.g. {abc: 123} intended for use within 'catch'."
        super().__init__(message)


class WinXPJobObjectWarning(UserWarning):  # for Vista and Windows 7
    def __init__(self, message: str):
        message += f"""
\tRecommend polling within AutoHotkey: https://github.com/CodeOptimist/ahkunwrapped/issues/1
\tAlternatives: https://stackoverflow.com/q/13471611
\tThis isn't an issue on Windows 8+ due to nestable jobs."""
        super().__init__(message)

class ExistingWinXPJobObjectWarning(WinXPJobObjectWarning): pass
class SingleWinXPJobObjectWarning(WinXPJobObjectWarning): pass


def comment_debug() -> str:
    return ";" if "pytest" not in sys.modules else ""


Primitive = Union[bool, float, int, str]


class Script:
    # Python 3.8 would use Final instead of ClassVar https://www.python.org/dev/peps/pep-0591/#id14
    MSG_GET: ClassVar[int] = 0x8001
    MSG_SET: ClassVar[int] = 0x8002
    MSG_F: ClassVar[int] = 0x8003
    MSG_F_MAIN: ClassVar[int] = 0x8004
    MSG_MORE: ClassVar[int] = 0x8005
    MSG_EXIT: ClassVar[int] = 0x8006

    SEPARATOR: ClassVar[str] = '\3'
    EOM_MORE: ClassVar[str] = SEPARATOR * 2
    EOM_END: ClassVar[str] = SEPARATOR * 3

    BUFFER_SIZE: ClassVar[int] = 4096
    BUFFER_W_MORE_SIZE: ClassVar[int] = BUFFER_SIZE - len(EOM_MORE) * 2 - len('\n') - 1  # :SingleByteNewline
    assert BUFFER_W_MORE_SIZE % 2 == 0  # utf-16 is 2 bytes
    BUFFER_W_END_SIZE: ClassVar[int] = BUFFER_SIZE - len(EOM_END) * 2 - len('\n') - 1
    assert BUFFER_W_END_SIZE % 2 == 0

    python_pid: ClassVar = os.getpid()
    python_job: ClassVar = None

    CORE: ClassVar[str] = '''
    _pyUserBatchLines := A_BatchLines
    SetBatchLines, -1
    Process, Exist, ''' + str(python_pid) + '''
    if (ErrorLevel = 0)
        ExitApp ; https://stackoverflow.com/q/73506891/#comment129808240_73506891 :AvoidJobRace
    #NoEnv
    #NoTrayIcon
    #Persistent
    SetWorkingDir, ''' + os.getcwd() + '''
    _PY_SEPARATOR := ''' + f'Chr({ord(SEPARATOR)})' + '''
    _pyStdOut := FileOpen("*", "w", "utf-16-raw")
    _pyStdErr := FileOpen("**", "w", "utf-16-raw")
    
    _Py_Response(ByRef pipe, ByRef text, ByRef offset, ByRef onMain) {
        textSize := Max(StrLen(text) * 2 + StrLen(Chr(0)) * 2 - offset, 0)
        isEnd := onMain or textSize <= ''' + str(BUFFER_W_END_SIZE) + '''
        ;MsgBox % "offset: " offset " textSize: " textSize " isEnd: " isEnd
        
        pipe.RawWrite(&text + offset, isEnd ? textSize : ''' + str(BUFFER_W_MORE_SIZE) + ''')
        if (isEnd)
            pipe.Write(''' + ' '.join(f'Chr({ord(c)})' for c in EOM_END) + ''')
        else
            pipe.Write(''' + ' '.join(f'Chr({ord(c)})' for c in EOM_MORE) + ''')
        newLine := "`n"
        pipe.RawWrite(newLine, 1)  ; :SingleByteNewline
        pipe.Read(0)
    }
  
    _Py_MsgMore(ByRef wParam, ByRef lParam, ByRef msg, ByRef hwnd) {
        global _pyStdOut, _pyOutText, _pyOutOffset, _pyStdErr, _pyErrText, _pyErrOffset
        SetBatchLines, -1
        ''' + comment_debug() + '''DebugMsg(wParam, msg)
       
        numRead := ''' + str(BUFFER_W_MORE_SIZE) + '''
        _Py_Response(_pyStdOut, _pyOutText, _pyOutOffset += numRead, False)
        _Py_Response(_pyStdErr, _pyErrText, _pyErrOffset += numRead, False)
        return 1
    }
     
    ; we can't peek() stdout/stderr, so always write to both or we will over-read and hang waiting
    _Py_StdOut(ByRef outText, ByRef onMain := False) {
        global _pyStdOut, _pyOutText, _pyOutOffset, _pyStdErr, _pyErrText, _pyErrOffset
        _Py_Response(_pyStdOut, _pyOutText := outText, _pyOutOffset := 0, onMain)
        _Py_Response(_pyStdErr, _pyErrText := "", _pyErrOffset := 0, onMain)
        return 1
    }
    _Py_StdErr(ByRef name, ByRef errText, onMain := False) {
        global _pyStdOut, _pyOutText, _pyOutOffset, _pyStdErr, _pyErrText, _pyErrOffset, _PY_SEPARATOR
        _Py_Response(_pyStdOut, _pyOutText := "", _pyOutOffset := 0, onMain)
        _Py_Response(_pyStdErr, _pyErrText := name _PY_SEPARATOR errText, _pyErrOffset := 0, onMain)
        return 1
    }
    
    _Py_MsgCopyData(ByRef wParam, ByRef lParam, ByRef msg, ByRef hwnd) {
        global _pyThreadMsgData, _PY_SEPARATOR
        SetBatchLines, -1
        ''' + comment_debug() + '''DebugMsg(wParam, msg)
        
        ;dataTypeId := NumGet(lParam + 0*A_PtrSize) ; unneeded atm
        dataSize := NumGet(lParam + 1*A_PtrSize)
        strAddr := NumGet(lParam + 2*A_PtrSize)
        ; limitation of StrGet(): data is truncated after \\0
        data := StrGet(strAddr, dataSize, "utf-8")
        ; OutputDebug, Received: '%data%'
        
        ; Since messages can arrive from multiple threads—e.g. clicking 'Reload' within OBS Studio 'Scripts' window,
        ;  while a timer is also running within said script—we need to keep their input data separate.
        _pyThreadMsgData[wParam] := []
        ; limitation of Parse and StrSplit(): separator must be a single character
        Loop, Parse, data, % _PY_SEPARATOR
        {
            ; see Python function _to_ahk_str()
            type := RTrim(SubStr(A_LoopField, 1, 5))
            val := SubStr(A_LoopField, 7)
            ; others are automatic
            if (type = "bool")
                val := val == "True" ? 1 : 0    ; same as True/False
            _pyThreadMsgData[wParam].Push(val)
        }
        return 1
    }
    
    ; call on main thread, much slower but may be necessary for DllCall() to avoid:
    ;   Error 0x8001010d An outgoing call cannot be made since the application is dispatching an input-synchronous call.
    _Py_MsgFMain(ByRef wParam, ByRef lParam, ByRef msg, ByRef hwnd) {
        global _pyMsgFMainData
        SetBatchLines, -1
        ''' + comment_debug() + '''DebugMsg(wParam, msg)
            
        _pyMsgFMainData.Push(hwnd)
        _pyMsgFMainData.Push(msg)
        _pyMsgFMainData.Push(lParam)
        _pyMsgFMainData.Push(wParam)
        ''' + comment_debug() + '''OutputDebug, SENDING TO MAIN THREAD
        ; continue on main thread at below label
        ;  ordinarily a new message can interrupt this, but none will be sent because of our lock
        SetTimer, _Py_MsgFMain, -0 ; negative for one-time, and 0 is indeed quicker than 1
        return 1
    }
    
    _Py_MsgF(ByRef wParam, ByRef lParam, ByRef msg, ByRef hwnd, ByRef onMain := False) {
        global _pyThreadMsgData, _pyUserBatchLines, _PY_SEPARATOR
        SetBatchLines, -1
        ''' + comment_debug() + '''if (not onMain)
        ''' + comment_debug() + '''    DebugMsg(wParam, msg)
       
        func := _pyThreadMsgData[wParam].RemoveAt(1)
        if (not IsFunc(func)) {
            _pyThreadMsgData.Delete(wParam)
            return _Py_StdErr("''' + AhkFuncNotFoundError.__name__ + '''", func, onMain)
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
            
            return _Py_StdErr("''' + AhkUserException.__name__ + '''"
                , isExceptionObj _PY_SEPARATOR e.Message _PY_SEPARATOR e.What _PY_SEPARATOR e.Extra _PY_SEPARATOR e.File _PY_SEPARATOR e.Line
                , onMain)
        }
        
        SetBatchLines, -1
        _pyThreadMsgData.Delete(wParam)
        return _Py_StdOut(needResult ? result : "", onMain)
    }
    
    _Py_MsgGet(ByRef wParam, ByRef lParam, ByRef msg, ByRef hwnd) {
        local name, val
        SetBatchLines, -1
        ''' + comment_debug() + '''DebugMsg(wParam, msg)
        name := _pyThreadMsgData[wParam].RemoveAt(1)
        val := %name%
        _pyThreadMsgData.Delete(wParam)
        return _Py_StdOut(val)
    }
    
    _Py_MsgSet(ByRef wParam, ByRef lParam, ByRef msg, ByRef hwnd) {
        local name
        SetBatchLines, -1
        ''' + comment_debug() + '''DebugMsg(wParam, msg)
        name := _pyThreadMsgData[wParam].RemoveAt(1)
        %name% := _pyThreadMsgData[wParam].RemoveAt(1)
        _pyThreadMsgData.Delete(wParam)
        return 1
    }
    
    _Py_MsgExit() {
        ExitApp
        return 1 ; required even after ExitApp
    }
    
    DebugMsg(wParam, msg) {
        OutputDebug, % Format("msg {:#06x}\t\tthread {:#05} -> {:#05}\t\tprocess {:#05} -> {:#05}"
            , msg, wParam, DllCall("GetCurrentThreadId"), ''' + str(python_pid) + ''', DllCall("GetCurrentProcessId"))
    }
    
    _pyThreadMsgData := {}
    _pyMsgFMainData := []
    
    ; these all must return non-zero to signal completion
    OnMessage(''' + str(win32con.WM_COPYDATA) + ''', Func("_Py_MsgCopyData"))
    OnMessage(''' + str(MSG_GET) + ''', Func("_Py_MsgGet"))
    OnMessage(''' + str(MSG_SET) + ''', Func("_Py_MsgSet"))
    OnMessage(''' + str(MSG_F) + ''', Func("_Py_MsgF"))
    OnMessage(''' + str(MSG_F_MAIN) + ''', Func("_Py_MsgFMain"))
    OnMessage(''' + str(MSG_MORE) + ''', Func("_Py_MsgMore"))
    OnMessage(''' + str(MSG_EXIT) + ''', Func("_Py_MsgExit"))
    
    _Py_StdOut(A_ScriptHwnd)
    
    SetBatchLines, % _pyUserBatchLines
    Func("AutoExec").Call() ; call if exists
    _pyUserBatchLines := A_BatchLines
    
    _Py_StdOut("Initialized")
    return
    
    ; from _Py_MsgFMain()
    _Py_MsgFMain:
        SetBatchLines, -1
        ''' + comment_debug() + '''OutputDebug, RECEIVED IN MAIN THREAD
        _Py_MsgF(_pyMsgFMainData.Pop(), _pyMsgFMainData.Pop(), _pyMsgFMainData.Pop(), _pyMsgFMainData.Pop(), True)
    return
    
    ; an unused label so #Warn won't complain that the user script's auto-execute section is unreachable
    ; it is intentionally unreachable (we use AutoExec() instead) so scripts can run exclusive standalone code
    _Py_SuppressUnreachableWarning:
    AutoTrim, % A_AutoTrim          ; does nothing and never called, but makes label happy
    '''

    def __init__(self, script: str = "", ahk_path: Path = None, execute_from: Path = None, kill_process_tree_on_exit: bool = False) -> None:
        """Launch an AutoHotkey process.

        :param script: Actual AutoHotkey script. Optional if you only need built-in functions and variables.
        :param ahk_path: Path to an alternative AutoHotkey executable than the one included (`ahk.get('A_AhkVersion')`).
        :param execute_from: Path AutoHotkey executable will be hard-linked/copied to, for the benefit of individual show/hide status in system tray.
        :param kill_process_tree_on_exit: Descendants of AutoHotkey process will inherit its win32 job object and terminate with it.
            `Script.exit()` (an intentional exit) can override this.
            *Caution*: Universal Windows Platform (UWP) apps (e.g. Windows 10's notepad.exe and calc.exe) discard our job object;
            suggest using AutoHotkey's `OnExit()` in those cases: https://github.com/CodeOptimist/ahkunwrapped/issues/1
        """
        self.file = None
        self.script = script
        self.kill_process_tree_on_exit = kill_process_tree_on_exit

        if ahk_path is None:
            ahk_path = PACKAGE_PATH / r'lib\AutoHotkey\AutoHotkey.exe'
            if IN_PYINSTALLER and not ahk_path.is_file():
                raise FileNotFoundError(f"""Couldn't find AutoHotkey at '{ahk_path}'.
\tEdit your `.spec` file (may have been auto-generated) to contain:
\t    from pathlib import Path
\t    import ahkunwrapped
\t
\t    a = Analysis(...
\t        datas=[
\t            (Path(ahkunwrapped.__file__).parent / 'lib', 'lib'),
\t            ('your_script.ahk', '.'),  # if using `Script.from_file()`
\t        ],
\tAnd pass it to PyInstaller (or it will be overwritten), e.g. `pyinstaller example.spec`.""")

        if not ahk_path.is_file():
            raise FileNotFoundError(f"Couldn't find file '{ahk_path}' for `ahk_path`.")

        # Windows notification area relies on consistent exe path
        if execute_from is not None:
            execute_from_dir = Path(execute_from)
            if not execute_from_dir.is_dir():
                raise NotADirectoryError(f"Couldn't find folder '{execute_from_dir}' for `execute_from`.")
            ahk_into_folder = execute_from_dir / ahk_path.name

            try:
                if os.path.getmtime(ahk_into_folder) != os.path.getmtime(ahk_path):
                    os.remove(ahk_into_folder)
            except FileNotFoundError:
                pass

            try:
                os.link(ahk_path, ahk_into_folder)
            except FileExistsError:
                pass
            except OSError as ex:
                # 5: "Access is denied"
                # 17: "The system cannot move the file to a different disk drive"
                if ex.winerror in (5, 17):
                    shutil.copyfile(ahk_path, ahk_into_folder)
            ahk_path = ahk_into_folder

        # user script exceptions are already caught and sent to stderr, so /ErrorStdOut would only affect debugging CORE
        # self.cmd = [str(ahk_path), "/ErrorStdOut=utf-16-raw", "/CP65001", "*"]
        self.cmd = [str(ahk_path), "/CP65001", "*"]

        self.popen = subprocess.Popen(self.cmd, bufsize=Script.BUFFER_SIZE, executable=str(ahk_path),
                                      # must pipe all three for PyInstaller onefile exe
                                      stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

        # NOTE: PROCESS EXPLORER WILL SHOW ONE OR THE OTHER JOB BUT NOT BOTH @Chris 2022-10
        # https://learn.microsoft.com/en-gb/windows/win32/api/winbase/nf-winbase-createjobobjecta
        # job containing all AutoHotkey processes to terminate with Python
        Script.python_job = win32job.CreateJobObject(None, f"ahkUnwrapped:python.exe:{Script.python_pid}")  # will find existing or create
        extended_info = win32job.QueryInformationJobObject(Script.python_job, win32job.JobObjectExtendedLimitInformation)
        # silent breakaway so child processes won't inherit job
        extended_info['BasicLimitInformation']['LimitFlags'] = win32job.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE | win32job.JOB_OBJECT_LIMIT_SILENT_BREAKAWAY_OK
        win32job.SetInformationJobObject(Script.python_job, win32job.JobObjectExtendedLimitInformation, extended_info)

        # Both job objects "execute" when their last handle closes (Python exits), but here KILL_ON_JOB_CLOSE (for descendants) is optional.
        # Separately, we can force terminate at any time. :TerminateJob
        self.tree_job = win32job.CreateJobObject(None, f"ahkUnwrapped:AutoHotkey.exe:{self.popen.pid}")  # new job for descendants (and ourself)
        extended_info = win32job.QueryInformationJobObject(self.tree_job, win32job.JobObjectExtendedLimitInformation)
        # no breakaway; this job object will be inherited
        if self.kill_process_tree_on_exit:
            extended_info['BasicLimitInformation']['LimitFlags'] |= win32job.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
        win32job.SetInformationJobObject(self.tree_job, win32job.JobObjectExtendedLimitInformation, extended_info)

        @contextmanager
        def get_handle(handle: object) -> ContextManager[object]:
            try: yield handle
            finally: win32api.CloseHandle(handle)

        with get_handle(win32api.OpenProcess(win32con.PROCESS_TERMINATE | win32con.PROCESS_SET_QUOTA, False, self.popen.pid)) as ahk_handle:  # both flags required
            if SINGLE_JOB_ASSIGNMENTS:
                try:
                    has_tree_job = win32job.AssignProcessToJobObject(self.tree_job, ahk_handle)  # the better choice when we can only have 1
                    win32job.AssignProcessToJobObject(Script.python_job, ahk_handle)  # this will fail
                except pywintypes.error as ex:
                    if ex.winerror != 5:
                        raise
                    stacklevel = 3 if inspect.currentframe().f_back.f_code.co_name == 'from_file' else 2  # :FromFile
                    if 'has_tree_job' in locals():
                        message = f"""Could only assign AutoHotkey (PID {self.popen.pid}) to a single job object: its process tree.
\tAs such, AutoHotkey will only automatically terminate when Python exits unexpectedly if `kill_process_tree_on_exit` is set `True`."""
                        warn(SingleWinXPJobObjectWarning(message), stacklevel=stacklevel)
                    else:
                        message = f"""Couldn't assign AutoHotkey (PID {self.popen.pid}) to a job object because one was already inherited (breakaway is unlikely to succeed).
\tAs such, `Script.exit(kill_descendants=True)` and `Script(kill_process_tree_on_exit=True)` have no effect, nor will AutoHotkey terminate if Python exits unexpectedly."""
                        warn(ExistingWinXPJobObjectWarning(message), stacklevel=stacklevel)
            else:
                win32job.AssignProcessToJobObject(Script.python_job, ahk_handle)  # this one needs to be first to avoid 'Access denied', also see :AvoidJobRace
                win32job.AssignProcessToJobObject(self.tree_job, ahk_handle)  # no race here, AutoHotkey won't `Run` a child process before "Initialized"

        self.popen.stdin.write(Script.CORE.encode('utf-8'))
        self.popen.stdin.write(self.script.encode('utf-8'))
        self.popen.stdin.close()

        self.lock = None
        self.hwnd = int(self._read_response(), 16)
        assert self._read_response() == "Initialized"
        self.lock = threading.Lock()

        # last to make sure things went okay since it runs on its own thread
        atexit.register(self._on_python_exit)  # if we exit, exit AutoHotkey

    @staticmethod
    def from_file(path: Path, format_dict: Mapping[str, str] = None, ahk_path: Path = None, execute_from: Path = None, kill_process_tree_on_exit: bool = None) -> 'Script':  # :FromFile
        """Launch an AutoHotkey process from a script file.

        :param path: Path to file.
        :param format_dict: `.format()` dict to use {{variable}} within script. `globals()` is a common choice.
        :param ahk_path: See `Script()`.
        :param execute_from: See `Script()`.
        :param kill_process_tree_on_exit: See `Script()`.
        """
        with path.open(encoding='utf-8') as f:
            script = f.read()
        if format_dict is not None:
            script = script.replace(r'{', r'{{').replace(r'}', r'}}').replace(r'{{{', r'').replace(r'}}}', r'')
            script = script.format(**format_dict)
        script = Script(script, ahk_path, execute_from, kill_process_tree_on_exit)
        script.file = path  # for exceptions
        return script

    def _read_pipes(self) -> Tuple[str, str]:
        more = bytes(Script.EOM_MORE, 'utf-16-le') + b'\n'  # :SingleByteNewline
        end = bytes(Script.EOM_END, 'utf-16-le') + b'\n'

        err, out = bytearray(), bytearray()
        while True:
            def has_all(bytearray_: bytearray) -> bool:
                self.poll()
                return bytearray_.endswith(end) or bytearray_.endswith(more)

            # we're careful not to over-read into the next response,
            # but we can at least go line by line since we end with \n
            err_buffer, out_buffer = bytearray(), bytearray()
            while not has_all(out_buffer):
                out_buffer += self.popen.stdout.readline()  # :SingleByteNewline
            while not has_all(err_buffer):
                err_buffer += self.popen.stderr.readline()

            is_end = out_buffer.endswith(end) and err_buffer.endswith(end)

            def strip_eom(buffer) -> str:
                head, sep, tail = buffer.rpartition(end)
                return head if sep else buffer.rpartition(more)[0]

            err += strip_eom(err_buffer)
            out += strip_eom(out_buffer)

            if is_end:
                break
            self._send_message(Script.MSG_MORE)
        if self.lock is not None:
            self.lock.release()
        return (err.decode('utf-16-le')), (out.decode('utf-16-le'))

    def _read_response(self) -> str:
        err, out = self._read_pipes()
        if err:
            name, args = err.split(Script.SEPARATOR, 1)

            exception_class = next((ex for ex in chain(AhkError.__subclasses__(), AhkException.__subclasses__(), (AhkException,)) if ex.__name__ == name), None)
            if exception_class:
                exception = exception_class(*args.split(Script.SEPARATOR))
                if isinstance(exception, AhkUserException):
                    if exception.from_exception_obj and Script._is_num(exception.line):
                        exception.file = self.file or exception.file
                        exception.line = int(exception.line) - Script.CORE.count('\n')

                        if exception.message == '2147549453':
                            exception.message = '0x8001010D - An outgoing call cannot be made since the application is dispatching an input-synchronous call.'
                        if exception.message.startswith('0x8001010D - '):
                            outer_msg = 'Failed a remote procedure call from OnMessage() thread. Solve this with f_main(), call_main() or f_raw_main().'
                            raise AhkCantCallOutInInputSyncCallError(outer_msg) from exception
                    else:
                        warn(AhkCaughtNonExceptionWarning(exception), stacklevel=4)
                raise exception

            warning_class = next((w for w in chain(AhkWarning.__subclasses__(), (AhkWarning,)) if w.__name__ == name), None)
            if warning_class:
                warning = warning_class(*args.split(Script.SEPARATOR))
                warn(warning, stacklevel=4)

        return out

    # https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-sendmessage
    def _send_message(self, msg: int, lparam: bytes = None) -> None:
        # this is essential because messages are ignored if uninterruptible (e.g. in menu)
        # wparam is normally source window handle, but in our case source thread id
        while not win32api.SendMessage(self.hwnd, msg, threading.get_ident(), lparam):
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
        self.lock.acquire(blocking=True)  # False to witness test failure
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
                if '\x00' in val:
                    raise AhkUnsupportedValueError(r"string contains null terminator '\x00' which AutoHotkey ignores characters beyond")
                if Script.SEPARATOR in val:
                    raise AhkUnsupportedValueError(f'string contains {repr(Script.SEPARATOR)} which is reserved for messages to AutoHotkey')
            val_str = str(val)
        return f"{type(val).__name__[:5]:<5} {val_str}"

    def _f(self, msg: int, name: str, *args: Primitive, need_result: bool, coerce_result: bool = False) -> Optional[str]:
        self._send(msg, [name, need_result] + list(args))
        response = self._read_response()
        return self._from_ahk_str(response) if coerce_result else response

    def call(self, name: str, *args: Primitive) -> None:
        """Call a script function without receiving the result, if any. Least latency."""
        self._f(Script.MSG_F, name, *args, need_result=False)

    def call_main(self, name: str, *args: Primitive) -> None:
        """Same as `call()` but executed on AutoHotkey's main thread.
        Worse latency, but solution to `AhkCantCallOutInInputSyncCallError`."""
        self._f(Script.MSG_F_MAIN, name, *args, need_result=False)

    def f_raw(self, name: str, *args: Primitive) -> str:
        """Call a script function and return the result as its raw string (don't mimic AutoHotkey's type inference)."""
        return self._f(Script.MSG_F, name, *args, need_result=True)

    def f_raw_main(self, name: str, *args: Primitive) -> str:
        """Same as `f_raw()` but executed on AutoHotkey's main thread.
        Worse latency, but solution to `AhkCantCallOutInInputSyncCallError`."""
        return self._f(Script.MSG_F_MAIN, name, *args, need_result=True)

    def f(self, name: str, *args: Primitive) -> Primitive:
        """Call a script function and return the result."""
        return self._f(Script.MSG_F, name, *args, need_result=True, coerce_result=True)

    def f_main(self, name: str, *args: Primitive) -> Primitive:
        """Same as `f()` but executed on AutoHotkey's main thread.
        Worse latency, but solution to `AhkCantCallOutInInputSyncCallError`."""
        return self._f(Script.MSG_F_MAIN, name, *args, need_result=True, coerce_result=True)

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
        self._send(Script.MSG_GET, [name])
        return self._read_response()

    def get(self, name: str) -> Primitive:
        """Get a global script variable or built-in like `A_TimeIdle`."""
        self._send(Script.MSG_GET, [name])
        return Script._from_ahk_str(self._read_response())

    def set(self, name: str, val: Primitive) -> None:
        """Set a global script variable."""
        # Every _send() will lock, so others are finished before we set().
        #  We don't need a confirmation response, just the ensurance that it finishes before others begin.
        self._send(Script.MSG_SET, [name, val])
        self.lock.release()

    # if AutoHotkey is terminated, get error code
    def poll(self) -> None:
        """Detect when AutoHotkey process exits, typically within a loop, by raising `AhkExitException`.
        (Only needed in contexts without other Script functions, as they all run this internally.)"""
        exit_code = self.popen.poll()
        if exit_code is not None:
            # OutputDebugString(f"Exit code: {exit_code}; call stack: {traceback.format_stack()}")
            atexit.unregister(self._on_python_exit)
            raise AhkExitException(exit_code)

    def _on_python_exit(self) -> None:
        with suppress(AhkExitException):  # Expected and not exceptional.
            self.exit()

    def exit(self, timeout: float = 5.0, kill_descendants: Optional[bool] = None) -> None:
        """Ask AutoHotkey to exit cleanly (remove system tray icon, etc.).
        To my knowledge only an `OnExit()` callback could delay this.

        :param timeout: Seconds to wait before terminating. `None` for infinity.
        :param kill_descendants: Uses `Script()`'s `kill_process_tree_on_exit` (default `False`) unless overriden here.
        """

        if kill_descendants is None:
            kill_descendants = self.kill_process_tree_on_exit

        # No need to &= ~KILL_ON_JOB_CLOSE if `kill_descendants` is `False` and `self.kill_process_tree_on_exit` is `True`
        #  because jobs only *automatically* execute when *Python* exits (job handle closes), not AutoHotkey by itself.

        atexit.unregister(self._on_python_exit)

        exit_code = None
        try:
            try:
                # clean; removes tray icons etc.
                # OutputDebugString(f"Sending ExitApp from thread {threading.get_ident()}")
                self._send_message(Script.MSG_EXIT)
            except AhkExitException as ex:  # exited immediately
                exit_code = ex.args[0]  # for 'finally'
                raise

            exit_code = self.popen.wait(timeout)  # exited after a delay, before timeout
            raise AhkExitException(exit_code)
        except TimeoutExpired as ex:  # never exited before timeout
            self.popen.terminate()
            exit_code = 1
            raise AhkExitException(exit_code) from ex
        finally:
            if kill_descendants:
                win32job.TerminateJobObject(self.tree_job, exit_code)  # :TerminateJob
