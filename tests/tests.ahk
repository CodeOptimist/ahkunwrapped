; SPDX-License-Identifier: AGPL-3.0-or-later
; Copyright (c) 2019-2026 Christopher S. Galpin

#SingleInstance, force
#Warn
return

HasUtf16Internals() {
    str := "0.333333"
    float := 1 / 3  ; stored identically to above

    loop, % (StrLen(str) + 1) * 2 { ; include null-terminator, and 2 bytes each
;        MsgBox % A_Index - 1 " str: " NumGet(str, A_Index - 1, "UChar") " float: " NumGet(float, A_Index - 1, "UChar")
        if (NumGet(str, A_Index - 1, "UChar") != NumGet(float, A_Index - 1, "UChar"))
            return False
    }
    return True
}

GetSmile() {
    return "🙂"
}

ComWmiRpcCallout() {
    ComObjCreate("WbemScripting.SWbemLocator").ConnectServer()
}

ComFsoTempName() {
    comFso := ComObjCreate("Scripting.FileSystemObject")
    return comFso.GetTempName()
}

UserException() {
    throw Exception("UserException", "example what", "example extra")
}

NonException1() {
    throw 12345
}

NonException2() {
    throw "hello"
}

NonException3() {
    throw {abc: 123, def: "hi"}
}

NonException4() {
    throw {Message: "example message", What: "example what", File: "some file", Line: "not a number"}
}

ContrivedException() {
    throw {Message: "ContrivedException", What: "example what", File: "some file", Line: 9999999999}
}

Echo(val) {
    return val
}

Copy(val) {
    Clipboard := val
}
