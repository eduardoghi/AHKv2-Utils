AddTab(TabText, ControlHwnd, Pos := GetTabCount(ControlHwnd) + 1) {
    TCIF_TEXT := 0x0001
    OffTxP := (3 * 4) + (A_PtrSize - 4)
    TCM_INSERTITEMW := 0x133E

    Size := (5 * 4) + (2 * A_PtrSize) + (A_PtrSize - 4)
    TCITEM := Buffer(Size, 0)

    NumPut("UInt", TCIF_TEXT, TCITEM, 0)
    NumPut("Ptr", StrPtr(TabText), TCITEM, OffTxP)

    SendMessage(TCM_INSERTITEMW, Pos - 1, TCITEM, , "ahk_id " ControlHwnd)
}

RemoveTab(hwnd, Item) {
    TCM_DELETEITEM := 0x1308
    SendMessage(TCM_DELETEITEM, Item -1, 0, , "ahk_id " hwnd)
}

GetTabCount(hwnd) {
    TCM_GETITEMCOUNT := 0x1304 
    return SendMessage(TCM_GETITEMCOUNT, 0, 0, , "ahk_id " hwnd)
}

GetSelTabIndex(hwnd) {
    TCM_GETCURSEL := 0x130B
    return SendMessage(TCM_GETCURSEL, 0, 0, , "ahk_id " hwnd) + 1
}

SetTabSelection(ControlHwnd, index := 0) {
    TCM_SETCURFOCUS := 0x1330
    TCM_SETCURSEL := 0x130C

    SendMessage(TCM_SETCURFOCUS, index, 0, , ControlHwnd)
    SendMessage(TCM_SETCURSEL, index, 0, , ControlHwnd)
}

GetTabCaption(ControlHwnd, index) {
    static TCM_GETITEM := 0x133C
        , TCIF_TEXT := 1
        , MAX_TEXT_LENGTH := 260
        , MAX_TEXT_SIZE := MAX_TEXT_LENGTH * 2

    pid := WinGetPID("ahk_id " ControlHwnd)

    static PROCESS_VM_OPERATION := 0x8
        , PROCESS_VM_READ := 0x10
        , PROCESS_VM_WRITE := 0x20
        , READ_WRITE_ACCESS := PROCESS_VM_READ | PROCESS_VM_WRITE | PROCESS_VM_OPERATION
        , PROCESS_QUERY_INFORMATION := 0x400
        , MEM_COMMIT := 0x1000
        , MEM_RELEASE := 0x8000
        , PAGE_READWRITE := 4

    hproc := DllCall("OpenProcess", "uint", READ_WRITE_ACCESS | PROCESS_QUERY_INFORMATION, "int", false, "uint", pid, "ptr")
    if !hproc
        return ""

    if A_Is64bitOS {
        try DllCall("IsWow64Process", "ptr", hproc, "int*", &is32bit := true)
    } else {
        is32bit := true
    }

    RPtrSize := is32bit ? 4 : 8
    TCITEM_SIZE := 16 + RPtrSize * 3

    remote_item := DllCall("VirtualAllocEx", "ptr", hproc, "ptr", 0, "uptr", TCITEM_SIZE + MAX_TEXT_SIZE, "uint", MEM_COMMIT, "uint", PAGE_READWRITE, "ptr")
    remote_text := remote_item + TCITEM_SIZE

    local_item := Buffer(TCITEM_SIZE, 0)
    NumPut("uint", TCIF_TEXT, local_item, 0)
    NumPut("UPtr", remote_text, local_item, 8 + RPtrSize)
    NumPut("int", MAX_TEXT_LENGTH, local_item, 8 + RPtrSize * 2)

    VarSetStrCapacity(&local_text, MAX_TEXT_SIZE)

    DllCall("WriteProcessMemory", "ptr", hproc, "ptr", remote_item, "ptr", local_item, "uptr", TCITEM_SIZE, "ptr", 0)

    SendMessage(TCM_GETITEM, index, remote_item,, "ahk_id " ControlHwnd)

    DllCall("ReadProcessMemory", "ptr", hproc, "ptr", remote_text, "ptr", StrPtr(local_text), "uptr", MAX_TEXT_SIZE, "ptr", 0)
    VarSetStrCapacity(&local_text, -1)

    DllCall("VirtualFreeEx", "ptr", hproc, "ptr", remote_item, "uptr", 0, "uint", MEM_RELEASE)
    DllCall("CloseHandle", "ptr", hproc)

    return local_text
}
