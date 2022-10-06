# ahkUnwrapped
Program in Python; use [AutoHotkey](https://www.autohotkey.com/) for its simplification and mastery of the Windows API.

## Features
* Lightweight single module:
  * 170 lines of AutoHotkey.
* Performance is a feature.
* Zero wrappers or abstractions.
* GPL licensed to bake-in AutoHotkey.
* Test suite for all types (primitives).
* Supports exceptions and hard exits.
* Execute arbitrary AHK or load from files.
* Errors for unsupported values (`NaN` `Inf` `\0`).
* Warnings for loss of data (maximum 6 decimal places).
* Supports [PyInstaller](https://www.pyinstaller.org/) for single _.exe_ bundles.


* **Full language support; all of AutoHotkey:**
  * Easy mouse and keyboard events.
  * Detect and manipulate windows.
  * WinAPI calls with minimal code.
  * Rapid GUI layouts.

## How it Works
Each `Script` launches an _AutoHotkey.exe_ process with framework and user code passed [via stdin](https://www.autohotkey.com/docs/AHKL_ChangeLog.htm#v1.1.17.00). The framework listens for [windows messages](https://www.autohotkey.com/docs/commands/OnMessage.htm) from Python and responds [via stdout](https://docs.python.org/3.7/library/subprocess.html).

## Usage
`call(proc, ...)` `f(func, ...)` `get(var)` `set(var, val)`  

```python
from ahkunwrapped import Script

ahk = Script()
# built-in functions are directly callable
isNotepadActive = ahk.f('WinActive', "ahk_class Notepad")
# built-in variables (and user globals) can be set directly
ahk.set('Clipboard', "Copied text!")
print(isNotepadActive)
```
---
```python
from ahkunwrapped import Script

ahk = Script('''
LuckyMinimize(winTitle) {
  global myVar
  myVar := 7
  
  Clipboard := "You minimized: " winTitle
  WinMinimize, % winTitle
}
''')

# built-in commands can be used within functions
ahk.call('LuckyMinimize', "ahk_class Notepad")
print("Lucky number", ahk.get('myVar'))
```
---
```python
from pathlib import Path
from ahkunwrapped import Script

# load from a file
ahk = Script.from_file(Path('my_msg.ahk'))
ahk.call('MyMsg', "Wooo!")
```

_my_msg.ahk_:
```autohotkey
; auto-execute section when ran standalone
#SingleInstance force
#Warn
AutoExec()                   ; we can call this if we want
MyMsg("test our function")
return

; auto-execute section when ran from Python
AutoExec() {
  SetBatchLines, 100ms       ; slow our code to reduce CPU
}

MyMsg(text) {
  MsgBox, % text
}
```

Settings from [AutoExec()](https://www.autohotkey.com/docs/Scripts.htm#auto) will [still apply](https://www.autohotkey.com/docs/commands/OnMessage.htm#Remarks) even though we execute from `OnMessage()` for speed.  
AutoHotkey's [#Warn](https://www.autohotkey.com/docs/commands/_Warn.htm#Remarks) is special and will apply to both standalone and from-Python execution, unless you add/remove it dynamically.

`call(proc, ...)` is for performance, to avoid receiving a large unneeded result.  
`get(var)` `set(var, val)` are shorthand for accessing global variables and [built-ins](https://www.autohotkey.com/docs/Variables.htm#BuiltIn) like `A_TimeIdle`.  
`f(func, ...)` `get(var)` will infer `float` and `int` (base-16 beginning with `0x`) like AutoHotkey.  

`f_raw(func, ...)` `get_raw(var)` will return the raw string as-stored.  

`call_main(proc, ...)` `f_main(func, ...)` `f_raw_main(func, ...)` will execute on AutoHotkey's main thread, instead of [OnMessage()](https://www.autohotkey.com/docs/commands/OnMessage.htm#Remarks).  
This is necessary if `AhkCantCallOutInInputSyncCallError` is thrown, generally from some uses of [ComObjCreate()](https://www.autohotkey.com/docs/commands/ComObjCreate.htm).  
This is slower (except with very large data), but still fast and unlikely to bottleneck.

## Example event loop
### See bottom of AHK script for hotkeys

```python
import sys
import time
from datetime import datetime
from enum import Enum
from pathlib import Path

from ahkunwrapped import Script, AhkExitException

choice = None
HOTKEY_SEND_CHOICE = 'F2'


class Event(Enum):
    QUIT, SEND_CHOICE, CLEAR_CHOICE, CHOOSE_MONTH, CHOOSE_DAY = range(5)


ahk = Script.from_file(Path('example.ahk'), format_dict=globals())


def main() -> None:
    print("Scroll your mousewheel in Notepad.")

    ts = 0
    while True:
        try:
            ahk.poll()
            s_elapsed = time.time() - ts
            if s_elapsed >= 60:
                ts = time.time()
                print_minute()

            event = ahk.get('event')
            if event:
                ahk.set('event', '')
                on_event(event)
        except AhkExitException as e:
            sys.exit(e.args[0])
        time.sleep(0.01)


def print_minute() -> None:
    print(f'It is now {datetime.now().time()}')


def on_event(event: str) -> None:
    global choice

    def get_choice() -> str:
        return choice or datetime.now().strftime('%#I:%M %p')

    if event == str(Event.QUIT):
        ahk.exit()
    if event == str(Event.CLEAR_CHOICE):
        choice = None
    if event == str(Event.SEND_CHOICE):
        ahk.call('Send', f'{get_choice()} ')
    if event == str(Event.CHOOSE_MONTH):
        choice = datetime.now().strftime('%b')
        ahk.call('ToolTip', f'Month is {get_choice()}, {HOTKEY_SEND_CHOICE} to insert.')
    if event == str(Event.CHOOSE_DAY):
        choice = datetime.now().strftime('%#d')
        ahk.call('ToolTip', f'Day is {get_choice()}, {HOTKEY_SEND_CHOICE} to insert.')


if __name__ == '__main__':
    main()
```

_example.ahk_:
```autohotkey
#SingleInstance, force
#Warn
ToolTip("Standalone script test!")
return

AutoExec() {
    global event
    event := ""
    SendMode, input
}

Send(text) {
    Send, % text
}

ToolTip(text, s := 2) {
    ToolTip, % text
    SetTimer, RemoveToolTip, % s * 1000
}

RemoveToolTip:
    SetTimer, RemoveToolTip, off
    ToolTip,
    event = {{Event.CLEAR_CHOICE}}
return

MouseIsOver(winTitle) {
    MouseGetPos,,, winId
    result := WinExist(winTitle " ahk_id " winId)
    return result
}

#If WinActive("ahk_class Notepad")
{{HOTKEY_SEND_CHOICE}}::event = {{Event.SEND_CHOICE}}
^Q::event = {{Event.QUIT}}
#If MouseIsOver("ahk_class Notepad")
WheelUp::event = {{Event.CHOOSE_MONTH}}
WheelDown::event = {{Event.CHOOSE_DAY}}
```

## PyInstaller Example (single _.exe_ or folder)
_example.[spec](https://pyinstaller.readthedocs.io/en/stable/spec-files.html)_:

```python
# -*- mode: python -*-
from pathlib import Path

import ahkunwrapped

a = Analysis(
  ['example.py'],
  datas=[
    (Path(ahkunwrapped.__file__).parent / 'lib', 'lib'),
    ('example.ahk', '.'),
  ]
)
pyz = PYZ(a.pure)

# for onefile (slow to launch as it decompresses, but convenient)
exe = EXE(pyz, a.scripts, a.binaries, a.datas, name='my-example', upx=True, console=False)
# for onedir
# exe = EXE(pyz, a.scripts, exclude_binaries=True, name='my-example', upx=True, console=False)
# dir = COLLECT(exe, a.binaries, a.datas, name='my-example-folder')
```

_example.py_:
```python
from pathlib import Path

from ahkunwrapped import Script

# tray icon visibility settings rely on consistent exe paths
LOCALAPP_DIR = Path(os.getenv('LOCALAPPDATA') / 'pyinstaller-example')

# because working directory could be somewhere else
# https://pyinstaller.readthedocs.io/en/stable/runtime-information.html
CUR_DIR = Path(getattr(sys, '_MEIPASS', Path(__file__).parent))

ahk = Script.from_file(CUR_DIR / 'example.ahk', format_dict=globals(), execute_from=LOCALAPP_DIR)

# ...
```

_example.ahk_:
```autohotkey
AutoExec() {
  Menu, Tray, Icon, {{CUR_DIR}}\black.ico
  Menu, Tray, Icon  ; unhide
}

; ...
```