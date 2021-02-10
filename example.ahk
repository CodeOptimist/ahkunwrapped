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
