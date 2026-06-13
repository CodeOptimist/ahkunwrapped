; SPDX-License-Identifier: AGPL-3.0-or-later
; Copyright (c) 2019-2026 Christopher S. Galpin

#Warn
#SingleInstance

class MyClass {
    static MyMethod(val) {
        return "Static" val
    }
    class MyNestedClass {
        static MyNestedMethod(val) {
            return "Nested" val
        }
    }
}

return

Startup() {
    global myVar := ""

    global myObj := {
        myProp: {
            myMethod: (this, val) => "Instance" val,
            str1: "Hello",
            str2: " World",
            str3: "!",
        }
    }

    global myArray := ['A',]
    global myMap := Map("abc", 123)
    global result := Map()
}

#Include "included.ahk"

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

HasInt64Limit() {
    return (9223372036854775807 + 1 == -9223372036854775808)
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

NonError1() {
    throw 12345
}

NonError2() {
    throw "hello"
}

NonError3() {
    throw {abc: 123, def: "hi"}
}

Echo(val) {
    return val
}
