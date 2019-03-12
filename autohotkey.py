# Copyright (C) 2019  Christopher Galpin.  See /NOTICE.
import array
import atexit
import os
import struct
import subprocess
import sys
from typing import ClassVar
from typing import TypeVar

import win32api
import win32con

Primitive = TypeVar("Primitive", bool, float, int, str)


class Script:
    GET: ClassVar = 0x8001
    SET: ClassVar = 0x8002
    F: ClassVar = 0x8003

    CORE: ClassVar = '''
    #NoEnv
    #NoTrayIcon
    #Persistent
    FileEncoding, utf-8-raw
    SetWorkingDir, % A_ScriptDir
    
    _Py_CopyData(wParam, lParam, msg, hwnd) {
        global _pyData
        strAddr := NumGet(lParam + 2*A_PtrSize)
        val := StrGet(strAddr, "utf-8")
        _pyData.InsertAt(1, val)
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
        
        FileAppend, %result%`n, *
    }
    
    _Py_Get(wParam, lParam, msg, hwnd) {
        local name, val
        name := _pyData.Pop()
        val := %name%
        FileAppend, %val%`n, *
    }
    
    _Py_Set(wParam, lParam, msg, hwnd) {
        local name
        name := _pyData.Pop()
        %name% := _pyData.Pop()
    }
    
    _pyData := []
    
    OnMessage(''' + str(win32con.WM_COPYDATA) + ''', Func("_Py_CopyData"))
    OnMessage(''' + str(GET) + ''', Func("_Py_Get"))
    OnMessage(''' + str(SET) + ''', Func("_Py_Set"))
    OnMessage(''' + str(F) + ''', Func("_Py_F"))
    
    FileAppend, %A_ScriptHwnd%`n, *
    '''

    def __init__(self, script: str = "") -> None:
        self.script = Script.CORE
        self.script += script

        # noinspection PyProtectedMember
        dir_path = sys._MEIPASS if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
        ahk_path = os.path.join(dir_path, r'lib\AutoHotkey\AutoHotkey.exe')
        assert os.path.exists(ahk_path)

        self.cmd = [ahk_path, "*"]
        # must pipe all three within a PyInstaller bundled exe
        # text=True is a better alias for universal_newlines=True but requires newer Python
        self.ahk = subprocess.Popen(self.cmd, executable=ahk_path, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE, encoding='utf-8', universal_newlines=True)
        atexit.register(self.ahk.terminate)
        self.ahk.stdin.write(self.script)
        self.ahk.stdin.close()

        self.hwnd = int(self._line(), 16)

    def _line(self) -> str:
        return self.ahk.stdout.readline()[:-1]

    def _send(self, val: Primitive) -> None:
        char_buffer = array.array('b', bytes(Script._to_ahk_str(val), 'utf-8'))
        addr, size = char_buffer.buffer_info()
        struct_ = struct.pack('PLP', 12345, size, addr)
        win32api.SendMessage(self.hwnd, win32con.WM_COPYDATA, None, struct_)

    @staticmethod
    def _to_ahk_str(val: Primitive) -> str:
        return f"{val}\0"

    def _f(self, msg: int, name: str, *args: Primitive) -> None:
        self._send(name)
        for arg in args:
            self._send(arg)
        win32api.SendMessage(self.hwnd, msg, 0, 0)

    def call(self, name: str, *args: Primitive) -> None:
        self._f(Script.F, name, *args)
        self._line()

    def f(self, name: str, *args: Primitive) -> Primitive:
        self._f(Script.F, name, *args)
        return Script._from_ahk_str(self._line())

    @staticmethod
    def _from_ahk_str(str_: str) -> Primitive:
        is_hex = str_.startswith('0x') and str_[2:].isdigit()
        is_negative = str_.startswith('-') and str_[1:].isdigit()
        if str_.isdigit() or is_hex or is_negative:
            return int(str_, 0)
        return str_

    def get(self, name: str) -> Primitive:
        self._send(name)
        win32api.SendMessage(self.hwnd, Script.GET, 0, 0)
        return Script._from_ahk_str(self._line())

    def set(self, name: str, val: Primitive) -> None:
        self._send(name)
        self._send(val)
        win32api.SendMessage(self.hwnd, Script.SET, 0, 0)

    def close(self) -> None:
        self.ahk.stdout.close()
        return_code = self.ahk.wait()
        if return_code:
            raise subprocess.CalledProcessError(return_code, self.cmd)
