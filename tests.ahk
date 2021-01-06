#SingleInstance, force
#Warn
NoWarn:
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

Echo(val) {
    return val
}

GetSmile() {
    return "🙂"
}

Copy(val) {
    Clipboard := val
}
