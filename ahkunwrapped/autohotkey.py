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
from contextlib import suppress
from itertools import chain
from pathlib import Path
from subprocess import TimeoutExpired
from typing import ClassVar, Mapping, Optional, Sequence, Tuple, Union
from warnings import warn

import win32api
import win32con
import win32job
# noinspection PyUnresolvedReferences
from win32api import OutputDebugString

# support for PyInstaller
# noinspection PyProtectedMember,PyUnresolvedReferences
PACKAGE_PATH = Path(sys._MEIPASS) if getattr(sys, 'frozen', False) else Path(__file__).parent


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


Primitive = Union[bool, float, int, str]


def comment_debug():
    return ";" if "pytest" not in sys.modules else ""


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
    BUFFER_W_MORE_SIZE: ClassVar[int] = BUFFER_SIZE - len(EOM_MORE) * 2 - len('\n')
    BUFFER_W_END_SIZE: ClassVar[int] = BUFFER_SIZE - len(EOM_END) * 2 - len('\n')

    CORE: ClassVar[str] = '''
    _pyUserBatchLines := A_BatchLines
    SetBatchLines, -1
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
        pipe.RawWrite(newLine, 1)  ; must be a single byte for Python's readline()
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
        ; limitation of StrGet(): data is truncated after \0
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
            , msg, wParam, DllCall("GetCurrentThreadId"), ''' + str(os.getpid()) + ''', DllCall("GetCurrentProcessId"))
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

    def __init__(self, script: str = "", ahk_path: Path = None, execute_from: Path = None) -> None:
        self.file = None
        self.script = script
        self.lock = threading.Lock()

        if ahk_path is None:
            ahk_path = PACKAGE_PATH / r'lib\AutoHotkey\AutoHotkey.exe'
        assert ahk_path and ahk_path.is_file()

        # Windows notification area relies on consistent exe path
        if execute_from is not None:
            execute_from_dir = Path(execute_from)
            assert execute_from_dir.is_dir()
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

        # if we exit, exit AutoHotkey
        atexit.register(self.on_python_exit)

        python_pid = os.getpid()
        #  https://learn.microsoft.com/en-gb/windows/win32/api/winbase/nf-winbase-createjobobjecta
        # Supplying a name here prevents the following after creating 100 Scripts:
        #   pywintypes.error: (50, 'AssignProcessToJobObject', 'The request is not supported.')
        # self.job = win32job.CreateJobObject(None, "")
        self.job = win32job.CreateJobObject(None, f"ahkUnwrapped:{python_pid}")  # just to be descriptive, pid doesn't change
        extended_info = win32job.QueryInformationJobObject(self.job, win32job.JobObjectExtendedLimitInformation)
        # if we terminate, terminate AutoHotkey
        extended_info['BasicLimitInformation']['LimitFlags'] = win32job.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
        win32job.SetInformationJobObject(self.job, win32job.JobObjectExtendedLimitInformation, extended_info)
        # add ourselves and subprocess will inherit job membership
        handle = win32api.OpenProcess(win32con.PROCESS_TERMINATE | win32con.PROCESS_SET_QUOTA, False, python_pid)
        win32job.AssignProcessToJobObject(self.job, handle)
        win32api.CloseHandle(handle)

        # user script exceptions are already caught and sent to stderr, so /ErrorStdOut would only affect debugging CORE
        # self.cmd = [str(ahk_path), "/ErrorStdOut=utf-16-raw", "/CP65001", "*"]
        self.cmd = [str(ahk_path), "/CP65001", "*"]
        # must pipe all three within a PyInstaller bundled exe
        self.popen = subprocess.Popen(self.cmd, bufsize=Script.BUFFER_SIZE, executable=str(ahk_path), stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

        # keep grandchild processes from inheriting job membership above
        extended_info['BasicLimitInformation']['LimitFlags'] |= win32job.JOB_OBJECT_LIMIT_SILENT_BREAKAWAY_OK
        win32job.SetInformationJobObject(self.job, win32job.JobObjectExtendedLimitInformation, extended_info)

        self.popen.stdin.write(Script.CORE.encode('utf-8'))
        self.popen.stdin.write(self.script.encode('utf-8'))
        self.popen.stdin.close()

        self.hwnd = int(self._read_response(has_lock=False), 16)
        assert self._read_response(has_lock=False) == "Initialized"

    @staticmethod
    def from_file(path: Path, format_dict: Mapping[str, str] = None, ahk_path: Path = None, execute_from: Path = None) -> 'Script':
        with path.open(encoding='utf-8') as f:
            script = f.read()
        if format_dict is not None:
            script = script.replace(r'{', r'{{').replace(r'}', r'}}').replace(r'{{{', r'').replace(r'}}}', r'')
            script = script.format(**format_dict)
        script = Script(script, ahk_path, execute_from)
        script.file = path  # for exceptions
        return script

    def _read_pipes(self, has_lock: bool) -> Tuple[str, str]:
        more = bytes(Script.EOM_MORE, 'utf-16-le') + b'\n'
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
                out_buffer += self.popen.stdout.readline()  # reads to a *single* '\0d' byte
            while not has_all(err_buffer):                  # a utf-16 newline '\0d\00' would be split
                err_buffer += self.popen.stderr.readline()    # with '\00' starting the next response

            is_end = out_buffer.endswith(end) and err_buffer.endswith(end)

            def strip_eom(buffer) -> str:
                head, sep, tail = buffer.rpartition(end)
                return head if sep else buffer.rpartition(more)[0]

            err += strip_eom(err_buffer)
            out += strip_eom(out_buffer)

            if is_end:
                break
            self._send_message(Script.MSG_MORE)
        if has_lock:
            self.lock.release()
        return (err.decode('utf-16-le')), (out.decode('utf-16-le'))

    def _read_response(self, has_lock: bool = True) -> str:
        err, out = self._read_pipes(has_lock)
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
        self._f(Script.MSG_F, name, *args, need_result=False)

    def call_main(self, name: str, *args: Primitive) -> None:
        self._f(Script.MSG_F_MAIN, name, *args, need_result=False)

    def f_raw(self, name: str, *args: Primitive) -> str:
        return self._f(Script.MSG_F, name, *args, need_result=True)

    def f_raw_main(self, name: str, *args: Primitive) -> str:
        return self._f(Script.MSG_F_MAIN, name, *args, need_result=True)

    def f(self, name: str, *args: Primitive) -> Primitive:
        return self._f(Script.MSG_F, name, *args, need_result=True, coerce_result=True)

    def f_main(self, name: str, *args: Primitive) -> Primitive:
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
        self._send(Script.MSG_GET, [name])
        return self._read_response()

    def get(self, name: str) -> Primitive:
        self._send(Script.MSG_GET, [name])
        return Script._from_ahk_str(self._read_response())

    def set(self, name: str, val: Primitive) -> None:
        # Every _send() will lock, so others are finished before we set().
        #  We don't need a confirmation response, just the ensurance that it finishes before others begin.
        self._send(Script.MSG_SET, [name, val])
        self.lock.release()

    # if AutoHotkey is terminated, get error code
    def poll(self) -> None:
        exit_code = self.popen.poll()
        if exit_code is not None:
            # OutputDebugString(f"Exit code: {exit_code}; call stack: {traceback.format_stack()}")
            raise AhkExitException(exit_code)

    def on_python_exit(self) -> None:
        with suppress(AhkExitException):  # Expected and not exceptional.
            self.exit()

    def exit(self, timeout=5.0) -> None:
        atexit.unregister(self.on_python_exit)

        try:
            # clean; removes tray icons etc.
            # OutputDebugString(f"Sending ExitApp from thread {threading.get_ident()}")
            self._send_message(Script.MSG_EXIT)
            exit_code = self.popen.wait(timeout)
            raise AhkExitException(exit_code)
        except TimeoutExpired as ex:
            self.popen.terminate()
            raise AhkExitException(1) from ex
        finally:
            with suppress(RuntimeError):
                self.lock.release()
