# ahkUnwrapped
I wanted to automate Windows with the coverage and simplicity of the _complete_ [AutoHotkey API](https://www.autohotkey.com/), yet code in Python, so I created `ahkUnwrapped`.

AutoHotkey already abstracts the Windows API, so another layer to introduce complexity and slowdowns is undesirable. 

Instead, we bundle and bridge *AutoHotkey.exe*, sending your initial script [via stdin](https://www.autohotkey.com/docs/AHKL_ChangeLog.htm#v1.1.17.00) with minimal boilerplate to listen for [window messages](https://www.autohotkey.com/docs/commands/OnMessage.htm) from Python and respond [via stdout](https://docs.python.org/3.7/library/subprocess.html).

## Features
* **All** of AutoHotkey!
* Execute arbitrary AHK code or load scripts.
* [Hypothesis](https://hypothesis.readthedocs.io/en/latest/) powered testing of convoluted unicode, et al.
* Warnings for loss of precision (maximum 6 decimal places).
* Errors for unsupported values (`NaN` `Inf` `\0`).
* Unhandled AHK exceptions carry over to Python.
* Won't explode when used from multiple threads.
* Separate auto-execute sections to ease scripting.
* Supports [PyInstaller](https://www.pyinstaller.org/) for _onefile/onedir_ installations.
* Special care for:
  * Descriptive errors with accurate line numbers.
  * Persistent _Windows notification area_ settings.
  * Unexpected exit handling.
  * Minimal latency.

## Get started
`> pip install ahkunwrapped`

`call(proc, ...)` `f(func, ...)` `get(var)` `set(var, val)`  

```python
from ahkunwrapped import Script

ahk = Script()
# built-in functions are directly callable
isNotepadActive = ahk.f('WinActive', 'ahk_class Notepad')
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
  
  WinMinimize, % winTitle
  Clipboard := "You minimized: " winTitle
}
''')

ahk.call('LuckyMinimize', 'ahk_class Notepad')
print("Lucky number", ahk.get('myVar'))
```
---
```python
from pathlib import Path
from ahkunwrapped import Script

ahk = Script.from_file(Path('my_msg.ahk'))
ahk.call('MyMsg', "Wooo!")
```

_my_msg.ahk:_
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

Settings from [AutoExec()](https://www.autohotkey.com/docs/Scripts.htm#auto) will [still apply](https://www.autohotkey.com/docs/commands/OnMessage.htm#Remarks) even though we execute directly from the message listening thread for speed.  
<sub>(AutoHotkey's [#Warn](https://www.autohotkey.com/docs/commands/_Warn.htm#Remarks) is special and will apply to both standalone and from-Python execution, unless you add/remove it dynamically.)</sub>

## Usage
`call(proc, ...)` is for performance, to avoid receiving a large unneeded result.  

`get(var)` `set(var, val)` are shorthand for accessing global variables and [built-ins](https://www.autohotkey.com/docs/Variables.htm#BuiltIn) like `A_TimeIdle`.  

`f(func, ...)` `get(var)` will infer `float` and `int` (base-16 beginning with `0x`) like AutoHotkey.  

`f_raw(func, ...)` `get_raw(var)` will return the raw string as-stored.

`call_main(proc, ...)` `f_main(func, ...)` `f_raw_main(func, ...)` will execute on AutoHotkey's main thread instead of the [OnMessage()](https://www.autohotkey.com/docs/commands/OnMessage.htm#Remarks) listener.
This avoids `AhkCantCallOutInInputSyncCallError`, e.g. from some uses of [ComObjCreate()](https://www.autohotkey.com/docs/commands/ComObjCreate.htm).
This is slower (except with very large data), but still fast and unlikely to bottleneck.

## Event loop with hotkeys

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


# format_dict= so we can use {{VARIABLE}} within example.ahk
ahk = Script.from_file(Path('example.ahk'), format_dict=globals())


def main() -> None:
  print("Scroll your mousewheel in Notepad.")

  ts = 0
  while True:
    try:
      # ahk.poll()  # detect exit, but all ahk functions include this

      s_elapsed = time.time() - ts
      if s_elapsed >= 60:
        ts = time.time()
        print_minute()

      event = ahk.get('event')  # contains ahk.poll()
      if event:
        ahk.set('event', '')
        on_event(event)
    except AhkExitException as e:
      sys.exit(e.args[0])
    time.sleep(0.01)


def print_minute() -> None:
  print(f"It is now {datetime.now().time()}")


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
    ahk.call('ToolTip', f"Month is {get_choice()}, {HOTKEY_SEND_CHOICE} to insert.")
  if event == str(Event.CHOOSE_DAY):
    choice = datetime.now().strftime('%#d')
    ahk.call('ToolTip', f"Day is {get_choice()}, {HOTKEY_SEND_CHOICE} to insert.")


if __name__ == '__main__':
  main()
```

_example.ahk:_
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
    ; negative for non-repeating
    SetTimer, RemoveToolTip, % s * -1000
}

RemoveToolTip:
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

## PyInstaller (single _.exe_ or folder)
_example.[spec](https://pyinstaller.readthedocs.io/en/stable/spec-files.html):_

```python
# -*- mode: python -*-
from pathlib import Path

import ahkunwrapped

a = Analysis(
  ['example.py'],
  datas=[
    (Path(ahkunwrapped.__file__).parent / 'lib', 'lib'),  # required
    ('example.ahk', '.'),
  ]
)
pyz = PYZ(a.pure)

# for onefile
exe = EXE(pyz, a.scripts, a.binaries, a.datas, name='my-example', upx=True, console=False)

# for onedir
# exe = EXE(pyz, a.scripts, exclude_binaries=True, name='my-example', upx=True, console=False)
# dir = COLLECT(exe, a.binaries, a.datas, name='my-example-folder')
```

### Folder considerations

_example.py:_
```python
from pathlib import Path

from ahkunwrapped import Script

# tray icon visibility settings rely on consistent exe paths
LOCALAPP_DIR = Path(os.getenv('LOCALAPPDATA') / 'pyinstaller-example')

# because working directory could be somewhere else
#  https://pyinstaller.readthedocs.io/en/stable/runtime-information.html
CUR_DIR = Path(getattr(sys, '_MEIPASS', Path(__file__).parent))

ahk = Script.from_file(CUR_DIR / 'example.ahk', format_dict=globals(), execute_from=LOCALAPP_DIR)

# ...
```

_example.ahk:_
```autohotkey
AutoExec() {
  Menu, Tray, Icon, {{CUR_DIR}}\black.ico
  Menu, Tray, Icon  ; unhide
}

; ...
```