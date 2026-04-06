# ahkUnwrapped
I wanted to automate Windows with the coverage and simplicity of the _complete_ [AutoHotkey API](https://www.autohotkey.com/docs/v2/), yet code in Python, so I created `ahkUnwrapped`.

AutoHotkey already abstracts the Windows API, so another layer to introduce complexity and slowdowns is undesirable. 

Instead, we bundle and bridge *AutoHotkey64.exe [v2.0](https://www.autohotkey.com/docs/v2/v2-changes.htm)*, sending your initial script [via stdin](https://www.autohotkey.com/docs/v2/lib/FileOpen.htm) with a minimal framework to listen for [window messages](https://www.autohotkey.com/docs/v2/lib/OnMessage.htm) from Python and respond [via stdout](https://docs.python.org/3.13/library/subprocess.html).

## Features
* **All** of AutoHotkey!
* Minimal framework code and AHK glue.
* Execute arbitrary AHK code or load scripts.
* Won't explode when used from multiple threads.
* [Hypothesis](https://hypothesis.readthedocs.io/en/latest/) powered testing of convoluted Unicode, etc.
* Separate startup sections to test AHK scripts standalone.
* Supports [PyInstaller](https://www.pyinstaller.org/) for _onedir/onefile_ installations.
* Complete exception handling:
  * Errors for unsupported values (`NaN` `Inf` `\0`).
  * Unhandled AHK exceptions carry over to Python.
* Special care for:
  * Exceptions with accurate line numbers.
  * Persistent tray icon visibility settings.
  * Descendant process handling.
  * Unexpected exit handling.
  * Minimal latency.
### New with 3.0
* Transitioned from AutoHotkey v1 to v2.
* Type support:
    * Persistent types instead of coercion.
  * `(..., t=type)` for type assertion.
* Dot syntax for object properties and methods.
  * Direct access to UI object members.
  * [Arrays](https://www.autohotkey.com/docs/v2/lib/Array.htm) with `.Push`, `.Get`, etc.

## Get started
`> uv add ahkunwrapped`

`call(name, ...)` `f(name, ..., t=type)`  
`get(var, t=type)` `set(var, val)`  

You can immediately interact with standard AutoHotkey functions and variables:
```python
from ahkunwrapped import Script

ahk = Script()
ahk.set('A_Clipboard', "Hello from Python!")

is_notepad_active = ahk.f('WinActive', "ahk_class Notepad", t=bool)
if not is_notepad_active:
    ahk.call('Run', "notepad.exe")
```
---
You can write a script inline:
```python
from ahkunwrapped import Script

ahk = Script('''
Startup() {
  global myVar := 0
}

LuckyMinimize(winTitle) {
  global myVar := 7
  WinMinimize(winTitle)
}
''')

print(ahk.get('myVar'))
ahk.call('LuckyMinimize', "ahk_class Notepad")
lucky_num = ahk.get('myVar', t=int)
```
---
Or load it from a file:
```python
from pathlib import Path
from ahkunwrapped import Script

ahk = Script.from_file(Path('hello.ahk'))
ahk.call('Hello', "World!")
```

_hello.ahk:_
```autohotkey
; global directives
#Warn
#SingleInstance

; AHK-only startup section
A_ScriptName := "AutoHotkey"
Hello("test")
return

; Python-only startup section
Startup() {
  A_ScriptName := "Python"
}

Hello(text) {
  MsgBox("Hello " text)
}
```

## Usage
`call(name, ...)` `f(name, ..., t=type)`  
Execute a standalone function or a dot-notated object method (e.g., `myObj.MyMethod`).
- `call` is for performance, to avoid receiving a large unneeded result.
- `f` returns the result, optionally type-asserted with `t=` (otherwise the union `float | int | bool | str`).
  
`get(var, t=type)` `set(var, val)`  
Shorthand for accessing [built-ins](https://www.autohotkey.com/docs/v2/Variables.htm#BuiltIn) like `A_Clipboard`, or global variables and dot-notated properties (e.g., `myObj.prop.subProp`).

`call_main(...)` `f_main(...)`  
Execute on AutoHotkey's main thread instead of the [OnMessage()](https://www.autohotkey.com/docs/v2/lib/OnMessage.htm#Remarks) listener.  
This avoids `AhkCantCallOutInInputSyncCallError` in constrained threading contexts, but has higher latency.

## Event loop with hotkeys
_example.py:_
<!--[[[cog
from pathlib import Path
cog.outl("```python")
cog.outl(Path("examples/example.py").read_text(encoding='utf-8').split('\n', 2)[-1].strip())
cog.outl("```")
]]]-->
```python
import sys
import time
from datetime import datetime
from enum import IntEnum
from pathlib import Path

import schedule

from ahkunwrapped import Script, AhkExitException

choice = None
HOTKEY_SEND_CHOICE = 'F2'


class Event(IntEnum):
    QUIT, SEND_CHOICE, CLEAR_CHOICE, CHOOSE_MONTH, CHOOSE_DAY = range(5)


# `format_dict=` so we can use `{{VARIABLE}}` within example.ahk
ahk = Script.from_file(Path(__file__).parent / 'example.ahk', format_dict=globals())


def main() -> None:
    print("Scroll your mousewheel up and down in Notepad.")
    schedule.every(10).seconds.do(print_time)

    try:
        while True:
            # ahk.poll()  # detect exit, but all `ahk.` functions include this

            event = ahk.get('event', t=int)  # contains `ahk.poll()`
            if event >= 0:
                ahk.set('event', -1)
                on_event(event)

            schedule.run_pending()
            time.sleep(0.1)
    except AhkExitException as e:
        sys.exit(e.args[0])


def print_time() -> None:
    print(f"It is now {datetime.now().time()}")


def on_event(event: int) -> None:
    global choice

    def get_choice() -> str:
        return choice or datetime.now().strftime('%#I:%M %p')

    match event:
        case Event.QUIT:
            ahk.exit()
        case Event.CLEAR_CHOICE:
            choice = None
        case Event.SEND_CHOICE:
            ahk.call('Send', f"{get_choice()} ")
        case Event.CHOOSE_MONTH:
            choice = datetime.now().strftime('%b')
            ahk.call('Notify', f"Month is {get_choice()}, {HOTKEY_SEND_CHOICE} to insert.")
        case Event.CHOOSE_DAY:
            choice = datetime.now().strftime('%#d')
            ahk.call('Notify', f"Day is {get_choice()}, {HOTKEY_SEND_CHOICE} to insert.")


if __name__ == '__main__':
    main()
```
<!--[[[end]]]-->

_example.ahk:_
<!--[[[cog
from pathlib import Path
cog.outl("```autohotkey")
cog.outl(Path("examples/example.ahk").read_text(encoding='utf-8').split('\n', 2)[-1].strip())
cog.outl("```")
]]]-->
```autohotkey
#Warn
#SingleInstance

Notify("Standalone script test!")
return

Startup() {
    global event := -1
}

Notify(text, duration := 2000) {
    ToolTip(text)

    static RemoveToolTip() {
        ToolTip()
        global event := {{Event.CLEAR_CHOICE}}
    }
    SetTimer(RemoveToolTip, -duration)  ; negative for non-repeating
}

MouseIsOver(winTitle) {
    MouseGetPos(unset, unset, &winId)
    result := WinExist(winTitle " ahk_id " winId)
    return result
}

#HotIf WinActive("ahk_class Notepad")
{{HOTKEY_SEND_CHOICE}}::global event := {{Event.SEND_CHOICE}}
^Q::global event := {{Event.QUIT}}
#HotIf MouseIsOver("ahk_class Notepad")
WheelUp::global event := {{Event.CHOOSE_MONTH}}
WheelDown::global event := {{Event.CHOOSE_DAY}}
```
<!--[[[end]]]-->

## PyInstaller (single _.exe_ or folder)
_example.[spec](https://pyinstaller.readthedocs.io/en/stable/spec-files.html):_

<!--[[[cog
from pathlib import Path
cog.outl("```python")
cog.outl(Path("examples/example.spec").read_text(encoding='utf-8').split('\n', 2)[-1].strip())
cog.outl("```")
]]]-->
```python
from PyInstaller.utils.hooks import collect_data_files

import ahkunwrapped

a = Analysis(
    ['example.py'],                               # Python file
    datas=collect_data_files('ahkunwrapped') + [
        ('example.ahk', '.'),                     # AutoHotkey script (if using `Script.from_file()`)
    ]
)
pyz = PYZ(a.pure)

name = 'my-example'                               # used below

# for onefile
#exe = EXE(pyz, a.scripts, a.binaries, a.datas, name=name, upx=True, console=False)
# for onedir
exe = EXE(pyz, a.scripts, exclude_binaries=True, name=name, upx=True, console=False)
dir = COLLECT(exe, a.binaries, a.datas, name=name)
```
<!--[[[end]]]-->

### PyInstaller folder considerations
```python
import os
from pathlib import Path

from ahkunwrapped import Script

# Works both in and out of PyInstaller
CUR_DIR = Path(__file__).parent

# Windows needs a consistent exe path to remember tray icon visibility
LOCALAPP_DIR = Path(os.environ['LOCALAPPDATA'] / 'pyinstaller-example')

ahk = Script.from_file(CUR_DIR / 'example.ahk', format_dict=globals(), execute_from=LOCALAPP_DIR)

# ...
```
