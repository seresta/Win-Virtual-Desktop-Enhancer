#SingleInstance, force
#WinActivateForce
#HotkeyInterval 20
#MaxHotkeysPerInterval 20000
#MenuMaskKey vk07
#UseHook
; Credits to Ciantic: https://github.com/Ciantic/VirtualDesktopAccessor

#Include, %A_ScriptDir%\libraries\read-ini.ahk

; ======================================================================
; Set Up Library Hooks
; ======================================================================

DetectHiddenWindows, On
hwnd := WinExist("ahk_pid " . DllCall("GetCurrentProcessId","Uint"))
hwnd += 0x1000 << 32
hVirtualDesktopAccessor := DllCall("LoadLibrary", "Str", A_ScriptDir . "\libraries\virtual-desktop-accessor.dll", "Ptr")

global GoToDesktopNumberProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "GoToDesktopNumber", "Ptr")
global RegisterPostMessageHookProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "RegisterPostMessageHook", "Ptr")
global UnregisterPostMessageHookProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "UnregisterPostMessageHook", "Ptr")
global GetCurrentDesktopNumberProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "GetCurrentDesktopNumber", "Ptr")
global GetDesktopCountProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "GetDesktopCount", "Ptr")
global IsWindowOnDesktopNumberProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "IsWindowOnDesktopNumber", "Ptr")
global MoveWindowToDesktopNumberProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "MoveWindowToDesktopNumber", "Ptr")
global IsPinnedWindowProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "IsPinnedWindow", "Ptr")
global PinWindowProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "PinWindow", "Ptr")
global UnPinWindowProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "UnPinWindow", "Ptr")
global IsPinnedAppProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "IsPinnedApp", "Ptr")
global PinAppProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "PinApp", "Ptr")
global UnPinAppProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "UnPinApp", "Ptr")

DllCall(RegisterPostMessageHookProc, Int, hwnd, Int, 0x1400 + 30)
OnMessage(0x1400 + 30, "VWMess")
VWMess(wParam, lParam, msg, hwnd) {
    OnDesktopSwitch(lParam + 1)
}

; ======================================================================
; Auto Execute
; ======================================================================

; Set up tray and tray menu
Menu, Tray, NoStandard
Menu, Tray, Add, &Manage Desktops, OpenDesktopManager
Menu, Tray, Default, &Manage Desktops
Menu, Tray, Add, Reload Settings, Reload
Menu, Tray, Add, Exit, Exit
Menu, Tray, Click, 1

; Read and groom settings
ReadIni("settings.ini")

global GeneralWorkspaceSize := (GeneralWorkspaceSize != "" and GeneralWorkspaceSize ~= "[1-3]") ? GeneralWorkspaceSize : 1
global GeneralWorkspaceNum := GeneralWorkspaceSize * GeneralWorkspaceSize

; Initialize
global taskbarPrimaryID=0
global taskbarSecondaryID=0
global previousDesktopNo=0
global doFocusAfterNextSwitch=0
global hasSwitchedDesktopsBefore=1

initialDesktopNo := _GetCurrentDesktopNumber()

SwitchToDesktop(GeneralDefaultDesktop)
; Call "OnDesktopSwitch" since it wouldn't be called otherwise, if the default desktop matches the current one
if (GeneralDefaultDesktop == initialDesktopNo) {
    OnDesktopSwitch(GeneralDefaultDesktop)
}

; ======================================================================
; Set Up Key Bindings
; ======================================================================

; Translate the modifier keys strings

hkModifiersSwitch          := KeyboardShortcutsModifiersSwitchDesktop
hkModifiersMove            := KeyboardShortcutsModifiersMoveWindowToDesktop
hkModifiersMoveAndSwitch   := KeyboardShortcutsModifiersMoveWindowAndSwitchToDesktop
hkIdentifierPrevious       := KeyboardShortcutsIdentifiersPreviousDesktop
hkIdentifierNext           := KeyboardShortcutsIdentifiersNextDesktop
hkIdentifierTop            := KeyboardShortcutsIdentifiersTopDesktop
hkIdentifierBottom         := KeyboardShortcutsIdentifiersBottomDesktop
hkComboPinWin              := KeyboardShortcutsCombinationsPinWindow
hkComboUnpinWin            := KeyboardShortcutsCombinationsUnpinWindow
hkComboTogglePinWin        := KeyboardShortcutsCombinationsTogglePinWindow
hkComboPinApp              := KeyboardShortcutsCombinationsPinApp
hkComboUnpinApp            := KeyboardShortcutsCombinationsUnpinApp
hkComboTogglePinApp        := KeyboardShortcutsCombinationsTogglePinApp
hkComboOpenDesktopManager  := KeyboardShortcutsCombinationsOpenDesktopManager
hkComboQuickLaunchProgram  := KeyboardShortcutsCombinationsQuickLaunchProgram

arrayS := Object(),                     arrayR := Object()
arrayS.Insert("\s*|,"),                 arrayR.Insert("")
arrayS.Insert("L(Ctrl|Shift|Alt|Win)"), arrayR.Insert("<$1")
arrayS.Insert("R(Ctrl|Shift|Alt|Win)"), arrayR.Insert(">$1")
arrayS.Insert("Ctrl"),                  arrayR.Insert("^")
arrayS.Insert("Shift"),                 arrayR.Insert("+")
arrayS.Insert("Alt"),                   arrayR.Insert("!")
arrayS.Insert("Win"),                   arrayR.Insert("#")

for index in arrayS {
    hkModifiersSwitch         := RegExReplace(hkModifiersSwitch, arrayS[index], arrayR[index])
    hkModifiersMove           := RegExReplace(hkModifiersMove, arrayS[index], arrayR[index])
    hkModifiersMoveAndSwitch  := RegExReplace(hkModifiersMoveAndSwitch, arrayS[index], arrayR[index])
    hkModifiersPlusTen        := RegExReplace(hkModifiersPlusTen, arrayS[index], arrayR[index])
    hkComboPinWin             := RegExReplace(hkComboPinWin, arrayS[index], arrayR[index])
    hkComboUnpinWin           := RegExReplace(hkComboUnpinWin, arrayS[index], arrayR[index])
    hkComboTogglePinWin       := RegExReplace(hkComboTogglePinWin, arrayS[index], arrayR[index])
    hkComboPinApp             := RegExReplace(hkComboPinApp, arrayS[index], arrayR[index])
    hkComboUnpinApp           := RegExReplace(hkComboUnpinApp, arrayS[index], arrayR[index])
    hkComboTogglePinApp       := RegExReplace(hkComboTogglePinApp, arrayS[index], arrayR[index])
    hkComboOpenDesktopManager := RegExReplace(hkComboOpenDesktopManager, arrayS[index], arrayR[index])  
    hkComboQuickLaunchProgram := RegExReplace(hkComboQuickLaunchProgram, arrayS[index], arrayR[index])
}

; Setup key bindings dynamically
;  If they are set incorrectly in the settings, an error will be thrown.

setUpHotkey(hk, handler, settingPaths) {
    Hotkey, %hk%, %handler%, UseErrorLevel
    if (ErrorLevel <> 0) {
        MsgBox, 16, Error, One or more keyboard shortcut settings have been defined incorrectly in the settings file: `n%settingPaths%. `n`nPlease read the README for instructions.
        Exit
    }
}

setUpHotkeyWithOneSetOfModifiersAndIdentifier(modifiers, identifier, handler, settingPaths) {
    modifiers <> "" && identifier <> "" ? setUpHotkey(modifiers . identifier, handler, settingPaths) :
}

setUpHotkeyWithTwoSetOfModifiersAndIdentifier(modifiersA, modifiersB, identifier, handler, settingPaths) {
    modifiersA <> "" && modifiersB <> "" && identifier <> "" ? setUpHotkey(modifiersA . modifiersB . identifier, handler, settingPaths) :
}

setUpHotkeyWithCombo(combo, handler, settingPaths) {
    combo <> "" ? setUpHotkey(combo, handler, settingPaths) :
}

i := 0
while (i < 9) {
    hkDesktopI0 := KeyboardShortcutsIdentifiersDesktop%i%
    hkDesktopI1 := KeyboardShortcutsIdentifiersDesktopAlt%i%
    j := 0
    while (j < 2) {
        hkDesktopI := hkDesktopI%j%
        setUpHotkeyWithOneSetOfModifiersAndIdentifier(hkModifiersSwitch, hkDesktopI, "OnShiftNumberedPress", "[KeyboardShortcutsModifiers] SwitchDesktop")
        setUpHotkeyWithOneSetOfModifiersAndIdentifier(hkModifiersMove, hkDesktopI, "OnMoveNumberedPress", "[KeyboardShortcutsModifiers] MoveWindowToDesktop")
        setUpHotkeyWithOneSetOfModifiersAndIdentifier(hkModifiersMoveAndSwitch, hkDesktopI, "OnMoveAndShiftNumberedPress", "[KeyboardShortcutsModifiers] MoveWindowAndSwitchToDesktop")
        j := j + 1
    }
    i := i + 1
}

if (!(GeneralUseNativePrevNextDesktopSwitchingIfConflicting && _IsPrevNextDesktopSwitchingKeyboardShortcutConflicting(hkModifiersSwitch, hkIdentifierPrevious))) {
    setUpHotkeyWithOneSetOfModifiersAndIdentifier(hkModifiersSwitch, hkIdentifierPrevious, "OnShiftLeftPress", "[KeyboardShortcutsModifiers] SwitchDesktop, [KeyboardShortcutsIdentifiers] PreviousDesktop")
}
if (!(GeneralUseNativePrevNextDesktopSwitchingIfConflicting && _IsPrevNextDesktopSwitchingKeyboardShortcutConflicting(hkModifiersSwitch, hkIdentifierNext))) {
    setUpHotkeyWithOneSetOfModifiersAndIdentifier(hkModifiersSwitch, hkIdentifierNext, "OnShiftRightPress", "[KeyboardShortcutsModifiers] SwitchDesktop, [KeyboardShortcutsIdentifiers] NextDesktop")
}
if (!(GeneralUseNativePrevNextDesktopSwitchingIfConflicting && _IsPrevNextDesktopSwitchingKeyboardShortcutConflicting(hkModifiersSwitch, hkIdentifierTop))) {
    setUpHotkeyWithOneSetOfModifiersAndIdentifier(hkModifiersSwitch, hkIdentifierTop, "OnShiftUpPress", "[KeyboardShortcutsModifiers] SwitchDesktop, [KeyboardShortcutsIdentifiers] TopDesktop")
}
if (!(GeneralUseNativePrevNextDesktopSwitchingIfConflicting && _IsPrevNextDesktopSwitchingKeyboardShortcutConflicting(hkModifiersSwitch, hkIdentifierBottom))) {
    setUpHotkeyWithOneSetOfModifiersAndIdentifier(hkModifiersSwitch, hkIdentifierBottom, "OnShiftDownPress", "[KeyboardShortcutsModifiers] SwitchDesktop, [KeyboardShortcutsIdentifiers] BottomDesktop")
}

setUpHotkeyWithOneSetOfModifiersAndIdentifier(hkModifiersMove, hkIdentifierPrevious, "OnMoveLeftPress", "[KeyboardShortcutsModifiers] MoveWindowToDesktop, [KeyboardShortcutsIdentifiers] PreviousDesktop")
setUpHotkeyWithOneSetOfModifiersAndIdentifier(hkModifiersMove, hkIdentifierNext, "OnMoveRightPress", "[KeyboardShortcutsModifiers] MoveWindowToDesktop, [KeyboardShortcutsIdentifiers] NextDesktop")

setUpHotkeyWithOneSetOfModifiersAndIdentifier(hkModifiersMoveAndSwitch, hkIdentifierPrevious, "OnMoveAndShiftLeftPress", "[KeyboardShortcutsModifiers] MoveWindowAndSwitchToDesktop, [KeyboardShortcutsIdentifiers] PreviousDesktop")
setUpHotkeyWithOneSetOfModifiersAndIdentifier(hkModifiersMoveAndSwitch, hkIdentifierNext, "OnMoveAndShiftRightPress", "[KeyboardShortcutsModifiers] MoveWindowAndSwitchToDesktop, [KeyboardShortcutsIdentifiers] NextDesktop")

setUpHotkeyWithCombo(hkComboPinWin, "OnPinWindowPress", "[KeyboardShortcutsCombinations] PinWindow")
setUpHotkeyWithCombo(hkComboUnpinWin, "OnUnpinWindowPress", "[KeyboardShortcutsCombinations] UnpinWindow")
setUpHotkeyWithCombo(hkComboTogglePinWin, "OnTogglePinWindowPress", "[KeyboardShortcutsCombinations] TogglePinWindow")

setUpHotkeyWithCombo(hkComboPinApp, "OnPinAppPress", "[KeyboardShortcutsCombinations] PinApp")
setUpHotkeyWithCombo(hkComboUnpinApp, "OnUnpinAppPress", "[KeyboardShortcutsCombinations] UnpinApp")
setUpHotkeyWithCombo(hkComboTogglePinApp, "OnTogglePinAppPress", "[KeyboardShortcutsCombinations] TogglePinApp")

setUpHotkeyWithCombo(hkComboOpenDesktopManager, "OpenDesktopManager", "[KeyboardShortcutsCombinations] OpenDesktopManager")

setUpHotkeyWithCombo(hkComboQuickLaunchProgram, "OnQuickLaunchProgramPress", "[KeyboardShortcutsCombinations] QuickLaunchProgram")

; ======================================================================
; Event Handlers
; ======================================================================

OnShiftNumberedPress() {
    SwitchToDesktop(substr(A_ThisHotkey, 0, 1))
}

OnMoveNumberedPress() {
    MoveToDesktop(substr(A_ThisHotkey, 0, 1))
}

OnMoveAndShiftNumberedPress() {
    MoveAndSwitchToDesktop(substr(A_ThisHotkey, 0, 1))
}

OnShiftLeftPress() {
    SwitchToDesktop(_GetPreviousDesktopNumberInRow())
}

OnShiftRightPress() {
    SwitchToDesktop(_GetNextDesktopNumberInRow())
}

OnShiftUpPress() {
    SwitchToDesktop(_GetPreviousDesktopNumberInColumn())
}

OnShiftDownPress() {
    SwitchToDesktop(_GetNextDesktopNumberInColumn())
}

OnMoveLeftPress() {
    MoveToDesktop(_GetPreviousDesktopNumberInRow())
}

OnMoveRightPress() {
    MoveToDesktop(_GetNextDesktopNumberInRow())
}

OnMoveUpPress() {
    MoveToDesktop(_GetPreviousDesktopNumberInColumn())
}

OnMoveDownPress() {
    MoveToDesktop(_GetNextDesktopNumberInColumn())
}

OnMoveAndShiftLeftPress() {
    MoveAndSwitchToDesktop(_GetPreviousDesktopNumberInRow())
}

OnMoveAndShiftRightPress() {
    MoveAndSwitchToDesktop(_GetNextDesktopNumberInRow())
}

OnMoveAndShiftUpPress() {
    SwitchToDesktop(_GetPreviousDesktopNumberInColumn())
}

OnMoveAndShiftDownPress() {
    SwitchToDesktop(_GetNextDesktopNumberInColumn())
}

OnPinWindowPress() {
    windowID := _GetCurrentWindowID()
    windowTitle := _GetCurrentWindowTitle()
    _PinWindow(windowID)
}

OnUnpinWindowPress() {
    windowID := _GetCurrentWindowID()
    windowTitle := _GetCurrentWindowTitle()
    _UnpinWindow(windowID)
}

OnTogglePinWindowPress() {
    windowID := _GetCurrentWindowID()
    windowTitle := _GetCurrentWindowTitle()
    if (_GetIsWindowPinned(windowID)) {
        _UnpinWindow(windowID)
    }
    else {
        _PinWindow(windowID)
    }
}

OnPinAppPress() {
    windowID := _GetCurrentWindowID()
    windowTitle := _GetCurrentWindowTitle()
    _PinApp()
}

OnUnpinAppPress() {
    windowID := _GetCurrentWindowID()
    windowTitle := _GetCurrentWindowTitle()
    _UnpinApp()
}

OnTogglePinAppPress() {
    windowID := _GetCurrentWindowID()
    windowTitle := _GetCurrentWindowTitle()
    if (_GetIsAppPinned(windowID)) {
        _UnpinApp(windowID)
    }
    else {
        _PinApp(windowID)
    }
}

OnDesktopSwitch(n:=1) {
    ; Give focus first, then display the popup, otherwise the popup could
    ; steal the focus from the legitimate window until it disappears.
    _FocusIfRequested()
    _ChangeIcon(n)
    _ChangeBackground(n)

    previousDesktopNo := n
}

OnQuickLaunchProgramPress(n:=1) {
    n := _GetCurrentDesktopNumber()
   _RunProgram(QuickLaunchProgram%n%, "[QuickLaunchProgram] " . n)
}

; ======================================================================
; Functions
; ======================================================================

SwitchToDesktop(n:=1) {
    doFocusAfterNextSwitch=1
    _ChangeDesktop(n)
}

MoveToDesktop(n:=1) {
    _MoveCurrentWindowToDesktop(n)
    _Focus()
}

MoveAndSwitchToDesktop(n:=1) {
    doFocusAfterNextSwitch=1
    _MoveCurrentWindowToDesktop(n)
    _ChangeDesktop(n)
}

OpenDesktopManager() {
    Send #{Tab}
}

; Let the user change desktop names with a prompt, without having to edit the 'settings.ini'
; file and reload the program.
; The changes are temprorary (names will be overwritten by the default values of
; 'settings.ini' when the program will be restarted.

Reload() {
    Reload
}

Exit() {
    ExitApp
}

_IsPrevNextDesktopSwitchingKeyboardShortcutConflicting(hkModifiersSwitch, hkIdentifierNextOrPrevious) {
    return ((hkModifiersSwitch == "<#<^" || hkModifiersSwitch == ">#<^" || hkModifiersSwitch == "#<^" || hkModifiersSwitch == "<#>^" || hkModifiersSwitch == ">#>^" || hkModifiersSwitch == "#>^" || hkModifiersSwitch == "<#^" || hkModifiersSwitch == ">#^" || hkModifiersSwitch == "#^") && (hkIdentifierNextOrPrevious == "Left" || hkIdentifierNextOrPrevious == "Right"))
}

_IsCursorHoveringTaskbar() {
    MouseGetPos,,, mouseHoveringID
    if (!taskbarPrimaryID) {
        WinGet, taskbarPrimaryID, ID, ahk_class Shell_TrayWnd
    }
    if (!taskbarSecondaryID) {
        WinGet, taskbarSecondaryID, ID, ahk_class Shell_SecondaryTrayWnd
    }
    return (mouseHoveringID == taskbarPrimaryID || mouseHoveringID == taskbarSecondaryID)
}

_GetCurrentWindowID() {
    WinGet, activeHwnd, ID, A
    return activeHwnd
}

_GetCurrentWindowTitle() {
    WinGetTitle, activeHwnd, A
    return activeHwnd
}

_TruncateString(string:="", n:=10) {
    return (StrLen(string) > n ? SubStr(string, 1, n-3) . "..." : string)
}

_GetDesktopName(n:=1) {
    if (n == 0) {
        n := 9
    }
    name := DesktopNames%n%
    if (!name) {
        name := "Desktop " . n
    }
    return name
}

; Set the name of the nth desktop to the value of a given string.
_SetDesktopName(n:=1, name:=0) {
    if (n == 0) {
        n := 9
    }
    if (!name) {
        ; Default value: "Desktop N".
        name := "Desktop " %n%
    }
    DesktopNames%n% := name
}

_GetNextDesktopNumberInRow() {
    i := _GetCurrentDesktopNumber()
    i := ((mod(i,GeneralWorkspaceSize) == 0) ? i : i+1)

    return i
}

_GetPreviousDesktopNumberInRow() {
    i := _GetCurrentDesktopNumber()
	i := ((mod(i,GeneralWorkspaceSize) == 1) ? i : i-1)

    return i
}

_GetNextDesktopNumberInColumn() {
    i := _GetCurrentDesktopNumber()
	i := ( ((((i-1)//GeneralWorkspaceSize)) == GeneralWorkspaceSize-1) ? i : i+GeneralWorkspaceSize)

    return i
}

_GetPreviousDesktopNumberInColumn() {
    i := _GetCurrentDesktopNumber()
	i := ( ((((i-1)//GeneralWorkspaceSize)) == 0) ? i : i-GeneralWorkspaceSize)

    return i
}

_GetCurrentDesktopNumber() {
    return DllCall(GetCurrentDesktopNumberProc) + 1
}

_GetNumberOfDesktops() {
    return DllCall(GetDesktopCountProc)
}

_MoveCurrentWindowToDesktop(n:=1) {
    activeHwnd := _GetCurrentWindowID()
    DllCall(MoveWindowToDesktopNumberProc, UInt, activeHwnd, UInt, n-1)
}

_ChangeDesktop(n:=1) {
    if (n == 0) {
        n := 9
    }
    DllCall(GoToDesktopNumberProc, Int, n-1)
}

_CallWindowProc(proc, window:="") {
    if (window == "") {
        window := _GetCurrentWindowID()
    }
    return DllCall(proc, UInt, window)
}

_PinWindow(windowID:="") {
    _CallWindowProc(PinWindowProc, windowID)
}

_UnpinWindow(windowID:="") {
    _CallWindowProc(UnpinWindowProc, windowID)
}

_GetIsWindowPinned(windowID:="") {
    return _CallWindowProc(IsPinnedWindowProc, windowID)
}

_PinApp(windowID:="") {
    _CallWindowProc(PinAppProc, windowID)
}

_UnpinApp(windowID:="") {
    _CallWindowProc(UnpinAppProc, windowID)
}

_GetIsAppPinned(windowID:="") {
    return _CallWindowProc(IsPinnedAppProc, windowID)
}

_RunProgram(program:="", settingName:="") {
    if (program <> "") {
        if (FileExist(program)) {
            Run, % program
        }
        else {
            MsgBox, 16, Error, The program "%program%" is not valid. `nPlease reconfigure the "%settingName%" setting. `n`nPlease read the README for instructions.
        }
    }
}

_ChangeBackground(n:=1) {
    line := Wallpapers%n%
    isHex := RegExMatch(line, "^0x([0-9A-Fa-f]{1,6})", hexMatchTotal)
    if (isHex) {
        hexColorReversed := SubStr("00000" . hexMatchTotal1, -5)

        RegExMatch(hexColorReversed, "^([0-9A-Fa-f]{2})([0-9A-Fa-f]{2})([0-9A-Fa-f]{2})", match)
        hexColor := "0x" . match3 . match2 . match1, hexColor += 0

        DllCall("SystemParametersInfo", UInt, 0x14, UInt, 0, Str, "", UInt, 1)
        DllCall("SetSysColors", "Int", 1, "Int*", 1, "UInt*", hexColor)
    }
    else {
        filePath := line

        isRelative := (substr(filePath, 1, 1) == ".")
        if (isRelative) {
            filePath := (A_WorkingDir . substr(filePath, 2))
        }
        if (filePath and FileExist(filePath)) {
            DllCall("SystemParametersInfo", UInt, 0x14, UInt, 0, Str, filePath, UInt, 1)
        }
    }
}

_ChangeIcon(n:=1) {
    Menu, Tray, Tip, % _GetDesktopName(n)
    Try
        Menu, Tray, Icon, icons/%GeneralWorkspaceSize%/%n%.ico
    Catch Exception
        Menu, Tray, Icon, icons/+.ico
}

; Only give focus to the foremost window if it has been requested.
_FocusIfRequested() {
    if (doFocusAfterNextSwitch) {
        _Focus()
        doFocusAfterNextSwitch=0
    }
}

; Give focus to the foremost window on the desktop.
_Focus() {
    foremostWindowId := _GetForemostWindowIdOnDesktop(_GetCurrentDesktopNumber())
    WinActivate, ahk_id %foremostWindowId%
}

; Select the ahk_id of the foremost window in a given virtual desktop.
_GetForemostWindowIdOnDesktop(n) {
    if (n == 0) {
        n := 9
    }
    ; Desktop count starts at 1 for this script, but at 0 for Windows.
    n -= 1

    ; winIDList contains a list of windows IDs ordered from the top to the bottom for each desktop.
    WinGet winIDList, list
    Loop % winIDList {
        windowID := % winIDList%A_Index%
        windowIsOnDesktop := DllCall(IsWindowOnDesktopNumberProc, UInt, WindowID, UInt, n)
        ; Select the first (and foremost) window which is in the specified desktop.
        if (WindowIsOnDesktop == 1) {
            return WindowID
        }
    }
}