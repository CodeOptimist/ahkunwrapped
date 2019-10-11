import time
from datetime import datetime
from autohotkey import Script


ahk = Script("""
Send(text) {
    Send, % text
}

AutoExec() {
    SetTitleMatchMode, 2
    MsgBox, Initialized.
}

#IfWinActive, Notepad
CapsLock::hotkey = insert time
""")


def poll():
    hotkey = ahk.get('hotkey')
    if not hotkey:
        return
    ahk.set('hotkey', '')
    if hotkey == "insert time":
        ahk.call('Send', datetime.now().strftime('%#I:%M %p'))


if __name__ == '__main__':
    # ahk2 = Script.from_file(r'example.ahk')
    # ahk2.call('MsgBox', "Hello World!")

    while True:
        poll()
        time.sleep(0.01)
