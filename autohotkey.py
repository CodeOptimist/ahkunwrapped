# Copyright (C) 2019  Christopher Galpin.  See /NOTICE.
import array
import atexit
import os
import shutil
import string
import struct
import subprocess
import sys
import time
from itertools import chain
from subprocess import TimeoutExpired
from typing import ClassVar
from typing import Mapping
from typing import Optional
from typing import Sequence
from typing import TypeVar

import win32api
import win32con

# noinspection PyProtectedMember
DIR_PATH = sys._MEIPASS if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
Primitive = TypeVar("Primitive", bool, float, int, str)


class AhkException(Exception): pass
class AhkExitException(AhkException): pass
class AhkError(AhkException): pass
class AhkFuncNotFoundError(AhkError): pass
class AhkUnexpectedPidError(AhkError): pass


class Script:
    GET: ClassVar = 0x8001
    SET: ClassVar = 0x8002
    F: ClassVar = 0x8003
    F_MAIN: ClassVar = 0x8004
    SEPARATOR: ClassVar = '\3'
    END: ClassVar = SEPARATOR + SEPARATOR

    CORE: ClassVar = '''
    _pyUserBatchLines := A_BatchLines
    SetBatchLines, -1
    #NoEnv
    #NoTrayIcon
    #Persistent
    FileEncoding, utf-8-raw
    SetWorkingDir, ''' + DIR_PATH + '''
    ; _PY_SEPARATOR assignment is prepended to core below
    _PY_END := _PY_SEPARATOR _PY_SEPARATOR "`n"
    
    ; we can't peek() stdout/stderr, so write to both
    _Py_Response(ByRef text) {
        global _PY_END
        FileAppend, % text _PY_END, *                       ; stdout
        FileAppend, % "" _PY_END, **                        ; stderr
        return 1
    }
    
    _Py_Exception(ByRef name, ByRef text) {
        global _pyData, _PY_SEPARATOR, _PY_END
        _pyData := []
        FileAppend, % "" _PY_END, *                         ; stdout
        FileAppend, % name _PY_SEPARATOR text _PY_END, **   ; stderr
        return 1
    }
    
    _Py_UnexpectedPidError(ByRef wParam) {
        global _pyPid
        return _Py_Exception("''' + AhkUnexpectedPidError.__name__ + '''", "expected " _pyPid " received " wParam)
    }
    
    _Py_CopyData(ByRef wParam, ByRef lParam, ByRef msg, ByRef hwnd) {
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
            _pyData.InsertAt(1, val)
        }
        return 1
    }
    
    ; call on main thread, much slower but may be necessary for DllCall() to avoid:
    ;   Error 0x8001010d An outgoing call cannot be made since the application is dispatching an input-synchronous call.
    _Py_F_Main(ByRef wParam, ByRef lParam, ByRef msg, ByRef hwnd) {
        global _pyData, _pyPid
        SetBatchLines, -1
        if (wParam != _pyPid)
            return _Py_UnexpectedPidError(wParam)
        a := _pyData
        a.Push(hwnd)
        a.Push(msg)
        a.Push(lParam)
        a.Push(wParam)
        SetTimer, _Py_F_Main, -1
        return 1
    }
    
    _Py_F(ByRef wParam, ByRef lParam, ByRef msg, ByRef hwnd) {
        global _pyData, _pyPid, _pyUserBatchLines
        SetBatchLines, -1
        if (wParam != _pyPid)
            return _Py_UnexpectedPidError(wParam)
        a := _pyData
        
        name := a.Pop()
        if (not IsFunc(name))
            return _Py_Exception("''' + AhkFuncNotFoundError.__name__ + '''", name)
        
        needResult := a.Pop()
        f := name
        len := a.Length()
        
        SetBatchLines, % _pyUserBatchLines
        if (len = 0)
            result := %f%()
        else if (len = 1)
            result := %f%(a.Pop())
        else if (len = 2)
            result := %f%(a.Pop(), a.Pop())
        else if (len = 3)
            result := %f%(a.Pop(), a.Pop(), a.Pop())
        else if (len = 4)
            result := %f%(a.Pop(), a.Pop(), a.Pop(), a.Pop())
        else if (len = 5)
            result := %f%(a.Pop(), a.Pop(), a.Pop(), a.Pop(), a.Pop())
        else if (len = 6)
            result := %f%(a.Pop(), a.Pop(), a.Pop(), a.Pop(), a.Pop(), a.Pop())
        else if (len = 7)
            result := %f%(a.Pop(), a.Pop(), a.Pop(), a.Pop(), a.Pop(), a.Pop(), a.Pop())
        else if (len = 8)
            result := %f%(a.Pop(), a.Pop(), a.Pop(), a.Pop(), a.Pop(), a.Pop(), a.Pop(), a.Pop())
        else if (len = 9)
            result := %f%(a.Pop(), a.Pop(), a.Pop(), a.Pop(), a.Pop(), a.Pop(), a.Pop(), a.Pop(), a.Pop())
        else if (len = 10)
            result := %f%(a.Pop(), a.Pop(), a.Pop(), a.Pop(), a.Pop(), a.Pop(), a.Pop(), a.Pop(), a.Pop(), a.Pop())
        SetBatchLines, -1
        
        return _Py_Response(needResult ? result : "")
    }
    
    _Py_Get(ByRef wParam, ByRef lParam, ByRef msg, ByRef hwnd) {
        local name, val
        SetBatchLines, -1
        if (wParam != _pyPid)
            return _Py_UnexpectedPidError(wParam)
        name := _pyData.Pop()
        val := %name%
        return _Py_Response(val)
    }
    
    _Py_Set(ByRef wParam, ByRef lParam, ByRef msg, ByRef hwnd) {
        local name
        SetBatchLines, -1
        if (wParam != _pyPid)
            return _Py_UnexpectedPidError(wParam)
        name := _pyData.Pop()
        %name% := _pyData.Pop()
        return 1
    }
    
    _Py_ExitApp() {
        ExitApp
    }
    
    _pyData := []
    _pyPid := ''' + str(os.getpid()) + '''
    
    ; must return non-zero to signal completion
    OnMessage(''' + str(win32con.WM_COPYDATA) + ''', Func("_Py_CopyData"))
    OnMessage(''' + str(GET) + ''', Func("_Py_Get"))
    OnMessage(''' + str(SET) + ''', Func("_Py_Set"))
    OnMessage(''' + str(F) + ''', Func("_Py_F"))
    OnMessage(''' + str(F_MAIN) + ''', Func("_Py_F_Main"))
    
    _Py_Response(A_ScriptHwnd)
    
    SetBatchLines, % _pyUserBatchLines
    Func("AutoExec").Call() ; call if exists
    _pyUserBatchLines := A_BatchLines
    
    _Py_Response("Initialized")
    return
    
    _Py_F_Main:
        SetBatchLines, -1
        _Py_F(_pyData.Pop(), _pyData.Pop(), _pyData.Pop(), _pyData.Pop())
    return
    '''

    def __init__(self, script: str = "", ahk_path: str = None, execute_from: str = None) -> None:
        self.pid = os.getpid()

        self.script = f"_PY_SEPARATOR := Chr({ord(Script.SEPARATOR)})"
        self.script += Script.CORE
        self.script += script

        if ahk_path is None:
            lib_path = os.path.join(DIR_PATH, r'lib\AutoHotkey\AutoHotkey.exe')
            prog_path = os.path.join(os.environ.get('ProgramW6432', os.environ['ProgramFiles']), r'AutoHotkey\AutoHotkey.exe')
            ahk_path = lib_path if os.path.exists(lib_path) else prog_path if os.path.exists(prog_path) else None
        assert os.path.exists(ahk_path)

        # Windows notification area relies on consistent exe path
        if execute_from is not None:
            assert os.path.isdir(execute_from)
            ahk_into_folder = os.path.join(execute_from, os.path.basename(ahk_path))
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

        self.cmd = [ahk_path, "/CP65001", "*"]
        # must pipe all three within a PyInstaller bundled exe
        # text=True is a better alias for universal_newlines=True but requires newer Python
        self.ahk = subprocess.Popen(self.cmd, executable=ahk_path, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE, encoding='utf-8', universal_newlines=True)
        atexit.register(self.exit)
        self.ahk.stdin.write(self.script)
        self.ahk.stdin.close()

        self.hwnd = int(self._read_response(), 16)
        assert self._read_response() == "Initialized"

    @staticmethod
    def from_file(path: str, format_dict: Mapping[str, str] = None, ahk_path: str = None, execute_from: str = None) -> 'Script':
        with open(os.path.join(DIR_PATH, path), encoding='utf-8') as f:
            script = f.read()
        if format_dict is not None:
            script = script.replace(r'{', r'{{').replace(r'}', r'}}').replace(r'{{{', r'').replace(r'}}}', r'')
            script = script.format(**format_dict)
        return Script(script, ahk_path, execute_from)

    def _read_response(self) -> str:
        end = f"{Script.END}\n"
        out, err = "", ""
        while out == "" or not out.endswith(end):
            out += self.ahk.stdout.readline()
        while err == "" or not err.endswith(end):
            err += self.ahk.stderr.readline()
        out, err = out[:-len(end)], err[:-len(end)]
        # OutputDebugString(f"Out: '{out}' Err: '{err}"

        if err:
            name, text = tuple(map(str, err.split(Script.SEPARATOR)))
            exception = next((ex for ex in chain(AhkError.__subclasses__(), AhkException.__subclasses__()) if ex.__name__ == name), None)
            if exception:
                raise exception(text)
        return out

    def _send_message(self, msg: int, lparam: bytes = None) -> None:
        # this is essential because messages are ignored if uninterruptible (e.g. in menu)
        # wparam is normally source window handle, but we don't have a window
        while not win32api.SendMessage(self.hwnd, msg, self.pid, lparam):
            if self.ahk.poll() is not None:
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
        str_ = f"{val:f}" if isinstance(val, float) else str(val)
        return f"{type(val).__name__[:5]:<5} {str_}"

    def _f(self, msg: int, name: str, *args: Primitive, need_result: bool) -> Optional[str]:
        self._send(msg, [name, need_result] + list(args))
        return Script._from_ahk_str(self._read_response())

    def call(self, name: str, *args: Primitive) -> None:
        self._f(Script.F, name, *args, need_result=False)

    def f(self, name: str, *args: Primitive) -> Primitive:
        return self._f(Script.F, name, *args, need_result=True)

    def call_main(self, name: str, *args: Primitive) -> None:
        self._f(Script.F_MAIN, name, *args, need_result=False)

    def f_main(self, name: str, *args: Primitive) -> Primitive:
        result = self._f(Script.F_MAIN, name, *args, need_result=True)
        return result

    @staticmethod
    def _from_ahk_str(str_: str) -> Primitive:
        is_hex = str_.startswith('0x') and all(c in string.hexdigits for c in str_[2:])
        if is_hex:
            return int(str_, 16)

        # noinspection PyShadowingNames
        def is_num(str_):
            return str_.isdigit() or (str_.startswith('-') and str_[1:].isdigit())

        if is_num(str_):
            return int(str_.lstrip('0') or '0', 0)
        if is_num(str_.replace('.', '', 1)):
            return float(str_)
        return str_

    def get(self, name: str) -> Primitive:
        self._send(Script.GET, [name])
        return Script._from_ahk_str(self._read_response())

    def set(self, name: str, val: Primitive) -> None:
        self._send(Script.SET, [name, val])

    def exit(self, timeout=5.0) -> None:
        try:
            self.call("_Py_ExitApp")  # clean, removes tray icons etc.
            return_code = self.ahk.wait(timeout)
            if return_code:
                raise subprocess.CalledProcessError(return_code, self.cmd)
        except AhkExitException:
            pass
        except TimeoutExpired:
            self.ahk.terminate()
        except Exception:
            self.ahk.terminate()
            raise
