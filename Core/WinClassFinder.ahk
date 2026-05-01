#Requires AutoHotkey v2.0
#SingleInstance Force

; WinClassFinder - Press Ctrl+Shift+W while any window is focused
; to see its class name and process path. Use this to identify
; file managers for QuickSwitch integration.

Hotkey("^+w", ShowClassInfo)

ShowClassInfo(*) {
    activeWin := WinExist("A")
    cls  := WinGetClass("A")
    proc := WinGetProcessPath("A")
    pid  := WinGetPID("A")
    titl := WinGetTitle("A")
    MsgBox(
        "Title: " titl "`n" .
        "Class: " cls "`n" .
        "PID:   " pid "`n" .
        "Exe:   " proc,
        "Window Info (Ctrl+Shift+W)", 0)
}
