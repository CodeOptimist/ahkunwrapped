# Copyright (C) 2019, 2020  Christopher Galpin.  Licensed under AGPL-3.0-or-later.  See /NOTICE.
import array
import atexit
import math
import os
import shutil
import string
import struct
import subprocess
import sys
import time
from itertools import chain
from pathlib import Path
from subprocess import TimeoutExpired
from typing import ClassVar, Mapping, Optional, Sequence, Union
from warnings import warn
from win32api import OutputDebugString

import win32api
import win32con
import win32job

# noinspection PyProtectedMember
DIR_PATH = Path(sys._MEIPASS) if getattr(sys, 'frozen', False) else Path(__file__).parent
Primitive = Union[bool, float, int, str]


class AhkException(Exception): pass
class AhkExitException(AhkException): pass
class AhkError(AhkException): pass
class AhkFuncNotFoundError(AhkError): pass
class AhkUnexpectedPidError(AhkError): pass
class AhkUnsupportedValueError(AhkError): pass
class AhkWarning(UserWarning): pass
class AhkLossOfPrecisionWarning(AhkWarning): pass
class AhkNewlineReplacementWarning(AhkWarning): pass


class AhkUserException(AhkException):
    def __init__(self, message: str, what: str, extra: str, file: str, line: Union[str, int]):
        self.message: str = message
        self.what: str = what
        self.extra: str = extra
        self.file: str = file
        self.line: int = int(line)

    def __str__(self) -> str:
        # Python 3.8: return f"{message=}, {what=}, {extra=}, {file=}, {line=}"
        return f"(message={repr(self.message)}, what={repr(self.what)}, extra={repr(self.extra)}, file={repr(self.file)}, line={repr(self.line)})"

    def __repr__(self) -> str:
        return f"{self.__class__.__name__}{self}"


class Script:
    # Python 3.8: use Final instead of ClassVar https://www.python.org/dev/peps/pep-0591/#id14
    MSG_GET: ClassVar[int] = 0x8001
    MSG_SET: ClassVar[int] = 0x8002
    MSG_F: ClassVar[int] = 0x8003
    MSG_F_MAIN: ClassVar[int] = 0x8004
    SEPARATOR: ClassVar[str] = '\3'
    END: ClassVar[str] = SEPARATOR + SEPARATOR

    CORE: ClassVar[str] = '''
    _pyUserBatchLines := A_BatchLines
    SetBatchLines, -1
    #NoEnv
    #NoTrayIcon
    #Persistent
    SetWorkingDir, ''' + str(DIR_PATH) + '''
    _PY_SEPARATOR := ''' + f'Chr({ord(SEPARATOR)})' + '''
    _PY_END := _PY_SEPARATOR _PY_SEPARATOR
    _pyStdOut := FileOpen("*", "w", "utf-8-raw")
    _pyStdErr := FileOpen("**", "w", "utf-8-raw")
    
    ; we can't peek() stdout/stderr, so always write to both so we don't hang waiting to read nothing
    _Py_Response(ByRef outText, ByRef errText, ByRef onMain := False) {
        global _pyData, _pyStdOut, _pyStdErr, _PY_SEPARATOR, _PY_END
        
        if (not onMain) {
            ; script hangs on WriteLine without this; flushing in batches won't work
            if (StrPut(outText _PY_END "`n", "utf-8") > 4096 or StrPut(errText _PY_END "`n", "utf-8") > 4096) {
                _pyData.Push(errText)
                _pyData.Push(outText)
                SetTimer, _Py_Response_Main, -1
                return 1
            }
        }
        
        _pyStdOut.WriteLine(outText _PY_END)
        _pyStdOut.Read(0)
        
        if (!errText && InStr(outText, Chr(13)))
            _pyStdErr.WriteLine("''' + AhkNewlineReplacementWarning.__name__ + r'''" _PY_SEPARATOR "'\r\n' and '\r' have been replaced with '\n' in result" _PY_END)''' + '''
        else
            _pyStdErr.WriteLine(errText _PY_END)
        _pyStdErr.Read(0)
        return 1
    }
    
    _Py_StdOut(ByRef text, ByRef onMain := False) {
        return _Py_Response(text, "", onMain)
    }
    
    _Py_StdErr(ByRef name, ByRef text, onMain := False) {
        global _pyData, _PY_SEPARATOR
        _pyData := []
        return _Py_Response("", name _PY_SEPARATOR text, onMain)
    }
    
    _Py_UnexpectedPidError(ByRef wParam) {
        global _pyPid
        return _Py_StdErr("''' + AhkUnexpectedPidError.__name__ + '''", "expected " _pyPid " received " wParam)
    }
    
    _Py_MsgCopyData(ByRef wParam, ByRef lParam, ByRef msg, ByRef hwnd) {
        global _pyData, _pyPid, _PY_SEPARATOR
        SetBatchLines, -1
        if (wParam != _pyPid)
            return _Py_UnexpectedPidError(wParam)
        
        ;dataTypeId := NumGet(lParam + 0*A_PtrSize) ; unneeded atm
        dataSize := NumGet(lParam + 1*A_PtrSize)
        strAddr := NumGet(lParam + 2*A_PtrSize)
        ; limitation of StrGet(): data is truncated after \0
        data := StrGet(strAddr, dataSize, "utf-8")
        ; OutputDebug, Received: '%data%'
        
        ; limitation of Parse and StrSplit(): separator must be a single character
        Loop, Parse, data, % _PY_SEPARATOR
        {
            type := RTrim(SubStr(A_LoopField, 1, 5))
            val := SubStr(A_LoopField, 7)
            ; others are automatic
            if (type = "bool")
                val := val == "True" ? 1 : 0    ; same as True/False
            _pyData.Push(val)
        }
        return 1
    }
    
    ; call on main thread, much slower but may be necessary for DllCall() to avoid:
    ;   Error 0x8001010d An outgoing call cannot be made since the application is dispatching an input-synchronous call.
    _Py_MsgF_Main(ByRef wParam, ByRef lParam, ByRef msg, ByRef hwnd) {
        global _pyData, _pyPid
        SetBatchLines, -1
        if (wParam != _pyPid)
            return _Py_UnexpectedPidError(wParam)
        _pyData.Push(hwnd)
        _pyData.Push(msg)
        _pyData.Push(lParam)
        _pyData.Push(wParam)
        SetTimer, _Py_MsgF_Main, -1
        return 1
    }
    
    _Py_MsgF(ByRef wParam, ByRef lParam, ByRef msg, ByRef hwnd, ByRef onMain := False) {
        global _pyData, _pyPid, _pyUserBatchLines, _PY_SEPARATOR
        SetBatchLines, -1
        if (wParam != _pyPid)
            return _Py_UnexpectedPidError(wParam)
        
        func := _pyData.RemoveAt(1)
        if (not IsFunc(func))
            return _Py_StdErr("''' + AhkFuncNotFoundError.__name__ + '''", func, onMain)
        needResult := _pyData.RemoveAt(1)

        SetBatchLines, % _pyUserBatchLines
        try result := %func%(_pyData*)
        catch e {
            SetBatchLines, -1
            _pyData := []
            return _Py_StdErr("''' + AhkUserException.__name__ + '''"
                , e.Message _PY_SEPARATOR e.What _PY_SEPARATOR e.Extra _PY_SEPARATOR e.File _PY_SEPARATOR e.Line
                , onMain)
        }
        SetBatchLines, -1
        _pyData := []
        
        return _Py_StdOut(needResult ? result : "", onMain)
    }
    
    _Py_MsgGet(ByRef wParam, ByRef lParam, ByRef msg, ByRef hwnd) {
        local name, val
        SetBatchLines, -1
        if (wParam != _pyPid)
            return _Py_UnexpectedPidError(wParam)
        name := _pyData.RemoveAt(1)
        val := %name%
        return _Py_StdOut(val)
    }
    
    _Py_MsgSet(ByRef wParam, ByRef lParam, ByRef msg, ByRef hwnd) {
        local name
        SetBatchLines, -1
        if (wParam != _pyPid)
            return _Py_UnexpectedPidError(wParam)
        name := _pyData.RemoveAt(1)
        %name% := _pyData.RemoveAt(1)
        return 1
    }
    
    _Py_ExitApp() {
        ExitApp
    }
    
    _pyData := []
    _pyPid := ''' + str(os.getpid()) + '''
    
    ; must return non-zero to signal completion
    OnMessage(''' + str(win32con.WM_COPYDATA) + ''', Func("_Py_MsgCopyData"))
    OnMessage(''' + str(MSG_GET) + ''', Func("_Py_MsgGet"))
    OnMessage(''' + str(MSG_SET) + ''', Func("_Py_MsgSet"))
    OnMessage(''' + str(MSG_F) + ''', Func("_Py_MsgF"))
    OnMessage(''' + str(MSG_F_MAIN) + ''', Func("_Py_MsgF_Main"))
    
    _Py_StdOut(A_ScriptHwnd)
    
    SetBatchLines, % _pyUserBatchLines
    Func("AutoExec").Call() ; call if exists
    _pyUserBatchLines := A_BatchLines
    
    _Py_StdOut("Initialized")
    return
    
    _Py_MsgF_Main:
        SetBatchLines, -1
        _Py_MsgF(_pyData.Pop(), _pyData.Pop(), _pyData.Pop(), _pyData.Pop(), True)
    return
    
    _Py_Response_Main:
        SetBatchLines, -1
        _Py_Response(_pyData.Pop(), _pyData.Pop(), True)
    return
    '''

    def __init__(self, script: str = "", ahk_path: Path = None, execute_from: Path = None) -> None:
        self.script = script

        if ahk_path is None:
            lib_path = DIR_PATH / r'lib\AutoHotkey\AutoHotkey.exe'
            prog_path = Path(os.environ.get('ProgramW6432', os.environ['ProgramFiles'])) / r'AutoHotkey\AutoHotkey.exe'
            ahk_path = lib_path if lib_path.is_file() else prog_path if prog_path.is_file() else None
        assert ahk_path and ahk_path.is_file()

        # Windows notification area relies on consistent exe path
        if execute_from is not None:
            execute_from_dir = Path(execute_from)
            assert execute_from_dir.is_dir()
            ahk_into_folder = execute_from_dir / ahk_path.name
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

        self.pid = os.getpid()

        # if we exit, exit AutoHotkey
        atexit.register(self.exit)

        # if we're killed, kill AutoHotkey
        self.job = win32job.CreateJobObject(None, "")
        extended_info = win32job.QueryInformationJobObject(self.job, win32job.JobObjectExtendedLimitInformation)
        extended_info['BasicLimitInformation']['LimitFlags'] = win32job.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
        win32job.SetInformationJobObject(self.job, win32job.JobObjectExtendedLimitInformation, extended_info)
        # add ourselves and subprocess will inherit job membership
        handle = win32api.OpenProcess(win32con.PROCESS_TERMINATE | win32con.PROCESS_SET_QUOTA, False, self.pid)
        win32job.AssignProcessToJobObject(self.job, handle)
        win32api.CloseHandle(handle)

        self.cmd = [str(ahk_path), "/CP65001", "*"]
        # must pipe all three within a PyInstaller bundled exe
        # text=True is a better alias for universal_newlines=True but requires newer Python
        self.popen = subprocess.Popen(self.cmd, executable=str(ahk_path), stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE, encoding='utf-8', universal_newlines=True)
        self.popen.stdin.write(Script.CORE)
        self.popen.stdin.write(self.script)
        self.popen.stdin.close()

        self.hwnd = int(self._read_response(), 16)
        assert self._read_response() == "Initialized"

    @staticmethod
    def from_file(path: Path, format_dict: Mapping[str, str] = None, ahk_path: Path = None, execute_from: Path = None) -> 'Script':
        if not path.is_absolute():
            path = DIR_PATH / path
        with path.open(encoding='utf-8') as f:
            script = f.read()
        if format_dict is not None:
            script = script.replace(r'{', r'{{').replace(r'}', r'}}').replace(r'{{{', r'').replace(r'}}}', r'')
            script = script.format(**format_dict)
        return Script(script, ahk_path, execute_from)

    def _read_response(self) -> str:
        end = f"{Script.END}\n"
        out, err = "", ""
        while out == "" or not out.endswith(end):
            out += self.popen.stdout.readline()
        while err == "" or not err.endswith(end):
            err += self.popen.stderr.readline()
        out, err = out[:-len(end)], err[:-len(end)]
        # OutputDebugString(f"Out: '{out}' Err: '{err}"

        if err:
            name, args = err.split(Script.SEPARATOR, 1)
            exception_class = next((ex for ex in chain(AhkError.__subclasses__(), AhkException.__subclasses__(), (AhkException,)) if ex.__name__ == name), None)
            warning_class = next((w for w in chain(AhkWarning.__subclasses__(), (AhkWarning,)) if w.__name__ == name), None)
            if exception_class:
                if exception_class is AhkUserException:
                    raise AhkUserException(*args.split(Script.SEPARATOR))
                raise exception_class(args)
            if warning_class:
                warn(warning_class(args))
        return out

    def _send_message(self, msg: int, lparam: bytes = None) -> None:
        # this is essential because messages are ignored if uninterruptible (e.g. in menu)
        # wparam is normally source window handle, but we don't have a window
        while not win32api.SendMessage(self.hwnd, msg, self.pid, lparam):
            if self.popen.poll() is not None:
                raise AhkExitException()
            time.sleep(0.01)

    def _send(self, msg: int, data: Sequence[Primitive]) -> None:
        data_str = Script.SEPARATOR.join(Script._to_ahk_str(v) for v in data)
        # OutputDebugString(f"Sent: {data}")
        char_buffer = array.array('b', bytes(data_str, 'utf-8'))
        addr, size = char_buffer.buffer_info()
        data_type_id = msg  # anything; unneeded atm
        struct_ = struct.pack('PLP', data_type_id, size, addr)
        self._send_message(win32con.WM_COPYDATA, struct_)
        self._send_message(msg)

    @staticmethod
    def _to_ahk_str(val: Primitive) -> str:
        if isinstance(val, float):
            if math.isnan(val) or math.isinf(val):
                raise AhkUnsupportedValueError(val)
            val_str = f'{val:.6f}'  # 6 decimal precision to match AutoHotkey
            if float(val_str) != val:
                warn(AhkLossOfPrecisionWarning(f'loss of precision from {val} to {val_str}'))
            val_str = val_str.rstrip('0').rstrip('.')  # less text to send the better
        else:
            if isinstance(val, str):
                if '\x00' in val:
                    raise AhkUnsupportedValueError(r"string contains null terminator '\x00' which AutoHotkey ignores characters beyond")
                if Script.SEPARATOR in val:
                    raise AhkUnsupportedValueError(f'string contains {repr(Script.SEPARATOR)} which is reserved for messages to AutoHotkey')
            val_str = str(val)
        return f"{type(val).__name__[:5]:<5} {val_str}"

    def _f(self, msg: int, name: str, *args: Primitive, need_result: bool) -> Optional[str]:
        self._send(msg, [name, need_result] + list(args))
        return self._read_response()

    def call(self, name: str, *args: Primitive) -> None:
        self._f(Script.MSG_F, name, *args, need_result=False)

    def f(self, name: str, *args: Primitive, coerce_type: bool = True) -> Primitive:
        response = self._f(Script.MSG_F, name, *args, need_result=True)
        return self._from_ahk_str(response) if coerce_type else response

    def call_main(self, name: str, *args: Primitive) -> None:
        self._f(Script.MSG_F_MAIN, name, *args, need_result=False)

    def f_main(self, name: str, *args: Primitive, coerce_type: bool = True) -> Primitive:
        response = self._f(Script.MSG_F_MAIN, name, *args, need_result=True)
        return self._from_ahk_str(response) if coerce_type else response

    @staticmethod
    def _from_ahk_str(str_: str) -> Primitive:
        is_hex = str_.startswith('0x') and all(c in string.hexdigits for c in str_[2:])
        if is_hex:
            return int(str_, 16)

        # noinspection PyShadowingNames
        def is_num(str_: str) -> bool:
            return str_.isdigit() or (str_.startswith('-') and str_[1:].isdigit())

        if is_num(str_):
            return int(str_.lstrip('0') or '0', 0)
        if is_num(str_.replace('.', '', 1)):
            return float(str_)
        return str_

    def get(self, name: str, coerce_type: bool = True) -> Primitive:
        self._send(Script.MSG_GET, [name])
        response = self._read_response()
        return Script._from_ahk_str(response) if coerce_type else response

    def set(self, name: str, val: Primitive) -> None:
        self._send(Script.MSG_SET, [name, val])

    # if AutoHotkey is killed, get error code
    def poll(self) -> Optional[int]:
        return self.popen.poll()

    def exit(self, timeout=5.0) -> None:
        try:
            self.call("_Py_ExitApp")  # clean, removes tray icons etc.
            return_code = self.popen.wait(timeout)
            if return_code:
                raise subprocess.CalledProcessError(return_code, self.cmd)
        except AhkExitException:
            pass
        except TimeoutExpired:
            self.popen.terminate()
        except Exception:
            self.popen.terminate()
            raise
