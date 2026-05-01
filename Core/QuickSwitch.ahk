#Requires AutoHotkey v2.0
#SingleInstance Force
SetWorkingDir(A_ScriptDir)
SendMode("Input")

/*
QuickSwitch - Use open file manager folders in Windows file dialogs
Original (AHK v1) by NotNull: https://www.voidtools.com/forum/viewtopic.php?f=2&t=9881
Rewritten for AHK v2

Supported file managers: Windows Explorer, Total Commander, XYPlorer, Directory Opus
Hotkey: Ctrl+Q (when a file dialog is active)
*/

global cm_CopySrcPathToClip := 2029
global cm_CopyTrgPathToClip := 2030

global gINI     := StrReplace(A_ScriptFullPath, ".ahk", ".ini")
global _tempfile := EnvGet("TEMP") . "\dopusinfo.xml"
try FileDelete(_tempfile)

global gWinID       := 0
global gDialogType  := ""
global gDialogAction := ""
global gFingerPrint := ""

Hotkey("^Q", ShowMenu, "Off")

loop {
    WinWaitActive("ahk_class #32770")

    gWinID      := WinExist("A")
    gDialogType := SmellsLikeAFileDialog(gWinID)

    if gDialogType {
        ahkExe       := WinGetProcessName("ahk_id " gWinID)
        windowTitle  := WinGetTitle("ahk_id " gWinID)
        gFingerPrint := ahkExe "___" windowTitle
        gDialogAction := IniRead(gINI, "Dialogs", gFingerPrint, "")

        if (gDialogAction = "1") {
            folderPath := Get_Zfolder(gWinID)
            if ValidFolder(folderPath)
                FeedDialog(gWinID, folderPath, gDialogType)
        } else if (gDialogAction = "0") {
            ; "Never here" — do nothing
        } else {
            ShowMenu()
        }

        Hotkey("^Q", ShowMenu, "On")
    }

    WinWaitNotActive("ahk_id " gWinID)
    Hotkey("^Q", ShowMenu, "Off")

    gWinID := 0, gDialogType := "", gDialogAction := "", gFingerPrint := ""
}


; ── Dialog detection ──────────────────────────────────────────────────────────

SmellsLikeAFileDialog(_thisID) {
    try controls := WinGetControls("ahk_id " _thisID)
    catch
        return false

    hasSysListView321 := false, hasToolbarWindow321 := false
    hasDirectUIHWND1  := false, hasEdit1 := false

    for ctrl in controls {
        switch ctrl {
            case "SysListView321":   hasSysListView321  := true
            case "ToolbarWindow321": hasToolbarWindow321 := true
            case "DirectUIHWND1":   hasDirectUIHWND1   := true
            case "Edit1":           hasEdit1           := true
        }
    }

    if (hasDirectUIHWND1 && hasToolbarWindow321 && hasEdit1)
        return "GENERAL"
    if (hasSysListView321 && hasToolbarWindow321 && hasEdit1)
        return "SYSLISTVIEW"
    return false
}


; ── Feed folder path into dialog ──────────────────────────────────────────────

FeedDialog(_thisID, _thisFOLDER, _type) {
    if _type = "GENERAL"
        FeedDialogGENERAL(_thisID, _thisFOLDER)
    else if _type = "SYSLISTVIEW"
        FeedDialogSYSLISTVIEW(_thisID, _thisFOLDER)
}

FeedDialogGENERAL(_thisID, _thisFOLDER) {
    WinActivate("ahk_id " _thisID)
    Sleep(50)
    ControlFocus("Edit1", "ahk_id " _thisID)

    useToolbar := "", enterToolbar := ""
    for ctrl in WinGetControls("ahk_id " _thisID) {
        if !InStr(ctrl, "ToolbarWindow32")
            continue
        ctrlHandle   := ControlGetHwnd(ctrl, "ahk_id " _thisID)
        parentHandle := DllCall("GetParent", "Ptr", ctrlHandle, "Ptr")
        parentClass  := WinGetClass("ahk_id " parentHandle)
        if InStr(parentClass, "Breadcrumb Parent")
            useToolbar := ctrl
        if InStr(parentClass, "msctls_progress32")
            enterToolbar := ctrl
    }

    if !(useToolbar && enterToolbar) {
        MsgBox("This type of dialog cannot be handled (yet).`nPlease report it!")
        return
    }

    folderSet := false
    loop 5 {
        SendInput("^l")
        Sleep(100)
        ctrlFocus := ControlGetFocus("ahk_id " _thisID)
        if (InStr(ctrlFocus, "Edit") && ctrlFocus != "Edit1") {
            ControlSetText(_thisFOLDER, ctrlFocus, "ahk_id " _thisID)
            if (ControlGetText(ctrlFocus, "ahk_id " _thisID) = _thisFOLDER) {
                folderSet := true
                break
            }
        }
    }

    if folderSet {
        ControlClick(enterToolbar, "ahk_id " _thisID)
        Sleep(15)
        ControlFocus("Edit1", "ahk_id " _thisID)
    }
}

FeedDialogSYSLISTVIEW(_thisID, _thisFOLDER) {
    WinActivate("ahk_id " _thisID)
    oldText := ControlGetText("Edit1", "ahk_id " _thisID)
    Sleep(20)

    _thisFOLDER := RTrim(_thisFOLDER, Chr(92)) . Chr(92)

    folderSet := false
    loop 20 {
        Sleep(10)
        ControlSetText(_thisFOLDER, "Edit1", "ahk_id " _thisID)
        if (ControlGetText("Edit1", "ahk_id " _thisID) = _thisFOLDER) {
            folderSet := true
            break
        }
    }

    if !folderSet
        return

    Sleep(20)
    ControlFocus("Edit1", "ahk_id " _thisID)
    Sleep(20)
    ControlSend("{Enter}", "Edit1", "ahk_id " _thisID)
    Sleep(15)
    ControlFocus("Edit1", "ahk_id " _thisID)
    Sleep(20)

    loop 5 {
        ControlSetText(oldText, "Edit1", "ahk_id " _thisID)
        Sleep(15)
        if (ControlGetText("Edit1", "ahk_id " _thisID) = oldText)
            break
    }
}


; ── Context menu ──────────────────────────────────────────────────────────────

ShowMenu(hk := "") {
    global gDialogType, gDialogAction, gWinID, gFingerPrint, gINI, _tempfile
    global cm_CopySrcPathToClip, cm_CopyTrgPathToClip

    contextMenu := Menu()
    contextMenu.Add("QuickSwitch Menu", (*) => 0)
    contextMenu.Default := "QuickSwitch Menu"
    contextMenu.Disable("QuickSwitch Menu")

    showMenu := false
    OpusInfo := ""

    for winID in WinGetList() {
        thisClass := WinGetClass("ahk_id " winID)

        ; ── Total Commander ──────────────────────────────────
        if thisClass = "TTOTAL_CMD" {
            tcExe     := GetModuleFileNameEx(WinGetPID("ahk_id " winID))
            clipSaved := ClipboardAll()

            A_Clipboard := ""
            try SendMessage(1075, cm_CopySrcPathToClip, 0, , "ahk_id " winID)
            if ValidFolder(A_Clipboard) {
                f := A_Clipboard
                contextMenu.Add(f, FolderChoiceCB)
                contextMenu.SetIcon(f, tcExe, 0)
                showMenu := true
            }

            A_Clipboard := ""
            try SendMessage(1075, cm_CopyTrgPathToClip, 0, , "ahk_id " winID)
            if ValidFolder(A_Clipboard) {
                f := A_Clipboard
                contextMenu.Add(f, FolderChoiceCB)
                contextMenu.SetIcon(f, tcExe, 0)
                showMenu := true
            }

            A_Clipboard := clipSaved
        }

        ; ── XYPlorer ─────────────────────────────────────────
        if thisClass = "ThunderRT6FormDC" {
            xyExe     := GetModuleFileNameEx(WinGetPID("ahk_id " winID))
            clipSaved := ClipboardAll()

            for xyCmd in ["::copytext get('path', a);", "::copytext get('path', i);"] {
                A_Clipboard := ""
                Send_XYPlorer_Message(winID, xyCmd)
                if ValidFolder(A_Clipboard) {
                    f := A_Clipboard
                    contextMenu.Add(f, FolderChoiceCB)
                    contextMenu.SetIcon(f, xyExe, 0)
                    showMenu := true
                }
            }

            A_Clipboard := clipSaved
        }

        ; ── Directory Opus ───────────────────────────────────
        if thisClass = "dopus.lister" {
            dopusExe := GetModuleFileNameEx(WinGetPID("ahk_id " winID))

            if !OpusInfo {
                try Run('"' dopusExe '\..\dopusrt.exe" /info "' _tempfile '"')
                Sleep(100)
                try { OpusInfo := FileRead(_tempfile), FileDelete(_tempfile) }
            }

            for tabState in [1, 2] {
                if RegExMatch(OpusInfo, 'mO)^.*lister="' winID '".*tab_state="' tabState '".*>(.*)<\/path>$', &out) {
                    if ValidFolder(out[1]) {
                        contextMenu.Add(out[1], FolderChoiceCB)
                        contextMenu.SetIcon(out[1], dopusExe, 0)
                        showMenu := true
                    }
                }
            }
        }

        ; ── File Explorer ────────────────────────────────────
        if thisClass = "CabinetWClass" {
            for expWin in ComObject("Shell.Application").Windows {
                try {
                    if winID = expWin.hwnd {
                        p := expWin.Document.Folder.Self.Path
                        if ValidFolder(p) {
                            contextMenu.Add(p, FolderChoiceCB)
                            contextMenu.SetIcon(p, "shell32.dll", 5)
                            showMenu := true
                        }
                    }
                }
            }
        }
    }

    if !showMenu
        return

    contextMenu.Add()
    contextMenu.Add("Settings for this dialog", (*) => 0)
    contextMenu.Disable("Settings for this dialog")
    contextMenu.Add("Allow AutoSwitch", AutoSwitchCB)
    contextMenu.Add("Never here",       NeverCB)
    contextMenu.Add("Not now",          NotNowCB)

    if (gDialogAction = "1") {
        contextMenu.Check("Allow AutoSwitch")
        contextMenu.Add("AutoSwitch exception", AutoSwitchExceptionCB)
    } else if (gDialogAction = "0") {
        contextMenu.Check("Never here")
    } else {
        contextMenu.Check("Not now")
    }

    contextMenu.Add("Debug this dialog", DebugCB)
    contextMenu.Show(100, 100)
}


; ── Menu callbacks ────────────────────────────────────────────────────────────

FolderChoiceCB(ItemName, ItemPos, MyMenu) {
    global gDialogType, gWinID
    if ValidFolder(ItemName)
        FeedDialog(gWinID, ItemName, gDialogType)
}

AutoSwitchCB(ItemName, ItemPos, MyMenu) {
    global gDialogType, gWinID, gFingerPrint, gINI, gDialogAction
    IniWrite("1", gINI, "Dialogs", gFingerPrint)
    gDialogAction := "1"
    folderPath := Get_Zfolder(gWinID)
    if ValidFolder(folderPath)
        FeedDialog(gWinID, folderPath, gDialogType)
}

NeverCB(ItemName, ItemPos, MyMenu) {
    global gFingerPrint, gINI, gDialogAction
    IniWrite("0", gINI, "Dialogs", gFingerPrint)
    gDialogAction := "0"
}

NotNowCB(ItemName, ItemPos, MyMenu) {
    global gFingerPrint, gINI, gDialogAction
    IniDelete(gINI, "Dialogs", gFingerPrint)
    gDialogAction := ""
}

AutoSwitchExceptionCB(ItemName, ItemPos, MyMenu) {
    global gDialogType, gWinID, gFingerPrint, gINI

    result := MsgBox(
        "For AutoSwitch to work, a file manager is typically '2 windows away':`n" .
        "File manager → Application → Dialog`n`n" .
        "If AutoSwitch isn't working, the app may have extra hidden windows.`n`n" .
        "To fix:`n" .
        "1. Cancel this dialog`n" .
        "2. Alt-Tab to your file manager`n" .
        "3. Alt-Tab back to the file dialog`n" .
        "4. Press Ctrl+Q → AutoSwitch Exception → OK`n`n" .
        "The correct window depth will be detected and saved.",
        "AutoSwitch Exceptions", 1)

    if result != "OK"
        return

    debugGui := Gui(, "Window Z-Order")
    lv := debugGui.Add("ListView", "r20 w900", ["Nr", "ID", "Title", "Program", "Class"])

    allWindows := WinGetList()
    level1 := 0, level2 := 0

    for idx, wid in allWindows {
        thisClass := WinGetClass("ahk_id " wid)
        selected  := ""
        if wid = gWinID {
            selected := "Select"
            level1 := idx
        }
        if (!level2 && (thisClass = "TTOTAL_CMD" || thisClass = "CabinetWClass" || thisClass = "ThunderRT6FormDC")) {
            selected := "Select"
            level2 := idx
        }
        lv.Add(selected, idx, wid, WinGetTitle("ahk_id " wid), WinGetProcessName("ahk_id " wid), thisClass)
    }
    lv.ModifyCol()
    debugGui.Show()

    delta := level2 - level1
    result2 := MsgBox(
        "The file manager appears to be " delta " levels away (default = 2).`n`nSave this as the default for this dialog?",
        "File manager found", 1)

    if result2 = "OK" {
        if delta = 2
            IniDelete(gINI, "AutoSwitchException", gFingerPrint)
        else
            IniWrite(delta, gINI, "AutoSwitchException", gFingerPrint)

        folderPath := Get_Zfolder(gWinID)
        if ValidFolder(folderPath)
            FeedDialog(gWinID, folderPath, gDialogType)
    }
    debugGui.Destroy()
}

DebugCB(ItemName, ItemPos, MyMenu) {
    global gFingerPrint
    debugGui := Gui(, "Control Debug")
    lv := debugGui.Add("ListView", "r25 w1024", ["Control", "ID", "Parent", "Text", "X", "Y", "W", "H"])

    for ctrl in WinGetControls("A") {
        try {
            handle := ControlGetHwnd(ctrl, "A")
            text   := ControlGetText(ctrl, "A")
            ControlGetPos(&x, &y, &w, &h, ctrl, "A")
            parent := DllCall("GetParent", "Ptr", handle, "Ptr")
            lv.Add(, ctrl, handle, parent, text, x, y, w, h)
        }
    }
    lv.ModifyCol()

    btnExport := debugGui.Add("Button", "y+10 w100 h30", "Export")
    btnExport.OnEvent("Click", (*) => ExportDebugData(lv, gFingerPrint))
    btnCancel := debugGui.Add("Button", "x+10 w100 h30", "Cancel")
    btnCancel.OnEvent("Click", (*) => debugGui.Destroy())
    debugGui.Show()
}

ExportDebugData(lv, fingerPrint) {
    fileName := A_ScriptDir . Chr(92) . fingerPrint . ".csv"
    try {
        f := FileOpen(fileName, "w")
        f.WriteLine("Control;ID;Parent;Text;X;Y;W;H")
        loop lv.GetCount() {
            rowNum := A_Index
            line := ""
            loop 8 {
                line .= (A_Index > 1 ? ";" : "") . lv.GetText(rowNum, A_Index)
            }
            f.WriteLine(line)
        }
        f.Close()
        MsgBox('Exported to:`n"' fileName '"')
    } catch as e {
        MsgBox("Export failed: " e.Message)
    }
}


; ── Core logic ────────────────────────────────────────────────────────────────

Get_Zfolder(_thisID) {
    global gFingerPrint, gINI, _tempfile, cm_CopySrcPathToClip

    zDelta     := Integer(IniRead(gINI, "AutoSwitchException", gFingerPrint, "2"))
    allWindows := WinGetList()
    thisZ      := 0

    for idx, wid in allWindows {
        if wid = _thisID {
            thisZ := idx
            break
        }
    }

    if (!thisZ || thisZ + zDelta > allWindows.Length)
        return ""

    nextID    := allWindows[thisZ + zDelta]
    nextClass := WinGetClass("ahk_id " nextID)
    ZFolder   := ""

    if nextClass = "TTOTAL_CMD" {
        clipSaved := ClipboardAll()
        A_Clipboard := ""
        try SendMessage(1075, cm_CopySrcPathToClip, 0, , "ahk_id " nextID)
        ZFolder := A_Clipboard
        A_Clipboard := clipSaved
    }

    else if nextClass = "ThunderRT6FormDC" {
        clipSaved := ClipboardAll()
        A_Clipboard := ""
        Send_XYPlorer_Message(nextID, "::copytext get('path', a);")
        ClipWait(0)
        ZFolder := A_Clipboard
        A_Clipboard := clipSaved
    }

    else if nextClass = "CabinetWClass" {
        for expWin in ComObject("Shell.Application").Windows {
            try {
                if nextID = expWin.hwnd {
                    ZFolder := expWin.Document.Folder.Self.Path
                    break
                }
            }
        }
    }

    else if nextClass = "dopus.lister" {
        dopusExe := GetModuleFileNameEx(WinGetPID("ahk_id " nextID))
        try Run('"' dopusExe '\..\dopusrt.exe" /info "' _tempfile '"')
        Sleep(100)
        try {
            OpusInfo := FileRead(_tempfile)
            FileDelete(_tempfile)
            if RegExMatch(OpusInfo, 'mO)^.*lister="' nextID '".*tab_state="1".*>(.*)<\/path>$', &out)
                ZFolder := out[1]
        }
    }

    return ZFolder
}

ValidFolder(_thisPath) {
    if (_thisPath = "" || StrLen(_thisPath) >= 259)
        return false
    return InStr(FileExist(_thisPath), "D") ? true : false
}


; ── DLL helpers ───────────────────────────────────────────────────────────────

GetModuleFileNameEx(p_pid) {
    hProcess := DllCall("OpenProcess", "UInt", 0x0410, "Int", false, "UInt", p_pid, "Ptr")
    if !hProcess
        return ""
    nameBuf := Buffer(520, 0)
    DllCall("psapi.dll\GetModuleFileNameExW", "Ptr", hProcess, "Ptr", 0, "Ptr", nameBuf, "UInt", 260)
    DllCall("CloseHandle", "Ptr", hProcess)
    return StrGet(nameBuf, "UTF-16")
}

Send_XYPlorer_Message(xyHwnd, message) {
    size    := StrLen(message)
    dataBuf := Buffer(size * 2 + 2, 0)
    StrPut(message, dataBuf, "UTF-16")

    COPYDATA := Buffer(A_PtrSize * 3, 0)
    NumPut("Ptr",  4194305,     COPYDATA, 0)
    NumPut("UInt", size * 2,    COPYDATA, A_PtrSize)
    NumPut("Ptr",  dataBuf.Ptr, COPYDATA, A_PtrSize * 2)

    DllCall("User32.dll\SendMessageW", "Ptr", xyHwnd, "UInt", 74, "Ptr", 0, "Ptr", COPYDATA)
}
