; SPDX-License-Identifier: AGPL-3.0-or-later
; Copyright (c) 2019-2026 Christopher S. Galpin

#Warn
#SingleInstance

return

Startup() {
    global myVar := ""
}

IsUtf16Ieee754() {
    str := "0.33333333333333331"
    floatStr := String(1 / 3)

    Loop (StrLen(str) + 1) * 2 { ; include null-terminator, and 2 bytes each
        byteA := NumGet(StrPtr(str), A_Index - 1, "UChar")
        byteB := NumGet(StrPtr(floatStr), A_Index - 1, "UChar")

        if (byteA != byteB)
            return False
    }
    return True
}

GetSmile() {
    return "🙂"
}

ComWmiRpcCallout() {
    ComObject("WbemScripting.SWbemLocator").ConnectServer()
}

ComFsoTempName() {
    comFso := ComObject("Scripting.FileSystemObject")
    return comFso.GetTempName()
}

UserException() {
    throw Error("UserException", "example what", "example extra")
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

Echo(val) {
    return val
}
