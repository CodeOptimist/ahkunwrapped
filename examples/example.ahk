; SPDX-License-Identifier: 0BSD

#Warn
#SingleInstance

Notify("Standalone script test!")
return

Startup() {
    global event := ""
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
