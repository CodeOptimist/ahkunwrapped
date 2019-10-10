# Copyright (C) 2019  Christopher Galpin.  See /NOTICE.
import array
import atexit
import os
import struct
import subprocess
import sys
import time
from typing import ClassVar
from typing import Mapping
from typing import TypeVar

import win32api
import win32con

# noinspection PyProtectedMember
DIR_PATH = sys._MEIPASS if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
Primitive = TypeVar("Primitive", bool, float, int, str)


class AhkExitException(Exception): pass


class Script:
    GET: ClassVar = 0x8001
    SET: ClassVar = 0x8002
    F: ClassVar = 0x8003
    F_MAIN: ClassVar = 0x8004

    CORE: ClassVar = '''
    #NoEnv
    #NoTrayIcon
    #Persistent
    FileEncoding, utf-8-raw
    SetWorkingDir, ''' + DIR_PATH + '''
    
    _Py_CopyData(wParam, lParam, msg, hwnd) {
        global _pyData
        strAddr := NumGet(lParam + 2*A_PtrSize)
        val := StrGet(strAddr, "utf-8")
        _pyData.InsertAt(1, val)
        return 1
    }
    
    ; call on main thread, much slower but may be necessary for DllCall() to avoid:
    ;   Error 0x8001010d An outgoing call cannot be made since the application is dispatching an input-synchronous call.
    _Py_F_Main(wParam, lParam, msg, hwnd) {
        global _pyData
        a := _pyData
        a.Push(hwnd)
        a.Push(msg)
        a.Push(lParam)
        a.Push(wParam)
        SetTimer, _Py_F_Main, -1
        return 1
    }
    
    _Py_F(wParam, lParam, msg, hwnd) {
        global _pyData
        a := _pyData
        
        name := a.Pop()
        
        f := name
        len := a.Length()
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
        
        FileAppend, % result Chr(3) "`n", *
        return 1
    }
    
    _Py_Get(wParam, lParam, msg, hwnd) {
        local name, val
        name := _pyData.Pop()
        val := %name%
        FileAppend, % val Chr(3) "`n", *
        return 1
    }
    
    _Py_Set(wParam, lParam, msg, hwnd) {
        local name
        name := _pyData.Pop()
        %name% := _pyData.Pop()
        return 1
    }
    
    _pyData := []
    
    OnMessage(''' + str(win32con.WM_COPYDATA) + ''', Func("_Py_CopyData"))
    OnMessage(''' + str(GET) + ''', Func("_Py_Get"))
    OnMessage(''' + str(SET) + ''', Func("_Py_Set"))
    OnMessage(''' + str(F) + ''', Func("_Py_F"))
    OnMessage(''' + str(F_MAIN) + ''', Func("_Py_F_Main"))
    
    FileAppend, % A_ScriptHwnd Chr(3) "`n", *
    Func("AutoExec").Call() ; call if exists
    FileAppend, % "Initialized" Chr(3) "`n", *
    
    return
    
    _Py_F_Main:
        _Py_F(_pyData.Pop(), _pyData.Pop(), _pyData.Pop(), _pyData.Pop())
        FileAppend, % "F_Main" Chr(3) "`n", *
    return
    '''

    def __init__(self, script: str = "", ahk_path: str = None) -> None:
        self.pid = os.getpid()

        self.script = Script.CORE
        self.script += script

        if ahk_path is None:
            lib_path = os.path.join(DIR_PATH, r'lib\AutoHotkey\AutoHotkey.exe')
            prog_path = os.path.join(os.environ.get('ProgramW6432', os.environ['ProgramFiles']), r'AutoHotkey\AutoHotkey.exe')
            ahk_path = lib_path if os.path.exists(lib_path) else prog_path if os.path.exists(prog_path) else None
        assert os.path.exists(ahk_path)

        self.cmd = [ahk_path, "*"]
        # must pipe all three within a PyInstaller bundled exe
        # text=True is a better alias for universal_newlines=True but requires newer Python
        self.ahk = subprocess.Popen(self.cmd, executable=ahk_path, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE, encoding='utf-8', universal_newlines=True)
        atexit.register(self.ahk.terminate)
        self.ahk.stdin.write(self.script)
        self.ahk.stdin.close()

        self.hwnd = int(self._read_text(), 16)
        assert self._read_text() == "Initialized"

    @staticmethod
    def from_file(path: str, format_dict: Mapping[str, str] = None, ahk_path: str = None) -> 'Script':
        with open(os.path.join(DIR_PATH, path)) as f:
            script = f.read()
        if format_dict is not None:
            script = script.replace(r'{', r'{{').replace(r'}', r'}}').replace(r'{{{', r'').replace(r'}}}', r'')
            script = script.format(**format_dict)
        return Script(script, ahk_path)

    def _read_text(self) -> str:
        end = '\3\n'
        result = ""
        while result == "" or not result.endswith(end):
            result += self.ahk.stdout.readline()
        return result[:-len(end)]

    def _send_message(self, msg: int, lparam: bytes = None) -> None:
        # this is essential because messages are ignored if uninterruptible (e.g. in menu)
        # wparam is normally source window handle, but we don't have a window
        while not win32api.SendMessage(self.hwnd, msg, self.pid, lparam):
            if self.ahk.poll() is not None:
                raise AhkExitException()
            time.sleep(0.01)

    def _send(self, val: Primitive) -> None:
        char_buffer = array.array('b', bytes(Script._to_ahk_str(val), 'utf-8'))
        addr, size = char_buffer.buffer_info()
        struct_ = struct.pack('PLP', 12345, size, addr)
        self._send_message(win32con.WM_COPYDATA, struct_)

    @staticmethod
    def _to_ahk_str(val: Primitive) -> str:
        return f"{val}\0"

    def _f(self, msg: int, name: str, *args: Primitive) -> None:
        self._send(name)
        for arg in args:
            self._send(arg)
        self._send_message(msg)

    def call(self, name: str, *args: Primitive) -> None:
        self._f(Script.F, name, *args)
        self._read_text()

    def f(self, name: str, *args: Primitive) -> Primitive:
        self._f(Script.F, name, *args)
        return Script._from_ahk_str(self._read_text())

    # call from main AHK thread
    def call_main(self, name: str, *args: Primitive) -> None:
        self._f(Script.F_MAIN, name, *args)
        self._read_text()
        assert self._read_text() == "F_Main"

    def f_main(self, name: str, *args: Primitive) -> Primitive:
        self._f(Script.F_MAIN, name, *args)
        result = Script._from_ahk_str(self._read_text())
        assert self._read_text() == "F_Main"
        return result

    @staticmethod
    def _from_ahk_str(str_: str) -> Primitive:
        is_hex = str_.startswith('0x') and str_[2:].isdigit()
        is_negative = str_.startswith('-') and str_[1:].isdigit()
        if str_.isdigit() or is_hex or is_negative:
            return int(str_, 0)
        return str_

    def get(self, name: str) -> Primitive:
        self._send(name)
        self._send_message(Script.GET)
        return Script._from_ahk_str(self._read_text())

    def set(self, name: str, val: Primitive) -> None:
        self._send(name)
        self._send(val)
        self._send_message(Script.SET)

    def close(self) -> None:
        self.ahk.stdout.close()
        return_code = self.ahk.wait()
        if return_code:
            raise subprocess.CalledProcessError(return_code, self.cmd)
