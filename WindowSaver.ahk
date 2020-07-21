; Window Saver
; Save and Restore window positions when docking/undocking
;
; Tested and minimum required AHK version: 1.1.32
;

#NoEnv
#SingleInstance, Force
    SendMode Input
SetKeyDelay, 75
SetWorkingDir %A_ScriptDir%
#Include Class_LV_InCellEdit.ahk
FileInstall, Icon.ico, Icon.ico
DetectHiddenWindows, On
SetTitleMatchMode, 2 ; partial match is enough
#KeyHistory, 0
ListLines, Off

global FileName :="window.cfg"
global Author := "David Szilvasi"
global Email := "szilvasi.dave@gmail.com"
global Website := "https://github.com/szilvasidave/WindowSaver/releases"
global AppVersion := "2.0"
global AppTitle := "Window Saver " . AppVersion
global 	DesktopCount = 2 ; Windows starts with 2 desktops at boot
global CurrentDesktop = 0 ; Desktop count is 0-indexed
global VDIDArr := []
global ParmVals := "Title StartIfNotRunning Class ID FullPath X Y W H VDID"

initMenu()
initApp()

WM_DISPLAYCHANGE := 0x7E
OnMessage(WM_DISPLAYCHANGE, "RestoreWindows")

;Restore window positions from file
RestoreWindows() {
    global VDIDArr, ParmVals
    Sleep 700
    WinGetActiveTitle, SavedActiveWindow
    Loop % StrSplit(ParmVals, A_Space).MaxIndex()
    {
        Win_%A_LoopField% := ""
        }
    
    SysGet, MonitorCount, MonitorCount
    SysGet, MonitorPrimary, MonitorPrimary
    WinGetPos, PMx, PMy, PMw, PMh, Program Manager
    IniRead, sectionNameList, %FileName%
    sectionName := "SECTION: Monitors=" . MonitorCount . ",MonitorPrimary=" . MonitorPrimary . "; Desktop size:" . PMx . "," . PMy . "," . PMw . "," . PMh
    
    If !InStr(sectionNameList, sectionName)
        MsgBox 64, %AppTitle%, Current configuration wasnt found.`nPlease make a save first with %SaveCombo%!
    
    IniRead, sectionValues, %FileName%, %sectionName%
    
    Loop, Parse, sectionValues, "`n"
    {
        currentLine := A_LoopField
        Loop, Parse, currentLine, "`,"
        {
            EqualPos := InStr(A_LoopField,"=")
            Var := SubStr(A_LoopField,1,EqualPos-1)
            Val := SubStr(A_LoopField,EqualPos+1)
            If InStr(ParmVals, %Var%)
            {
                ;Remove any surrounding double quotes (")
                If (SubStr(Val,1,1)==Chr(34)) 
                {
                    Val := SubStr(Val, 2, StrLen(Val)-2)
                }
                Win_%Var% := Val
            }
        }
        ; Try to find if window is already open. If it wasnt found, open a new window using it's path
        
        mapDesktopsFromRegistry()
        If WinExist("ahk_id" . Win_ID) {
            WinMove, ahk_id %Win_ID%,,%Win_X%,%Win_Y%,%Win_W%,%Win_H%
        } Else If (WinExist(Win_Title)) {
            WinGet, Win_Title_Count, Count, %Win_Title%
            If (Win_Title_Count > 1) {
                WinGet, windows, List, %Win_Title%
                Loop % windows
                {
                    id := windows%A_Index%
                    Error := getWindowDesktopID(id, desktopID)
                    if !(Error=0) ;S_OK
                        MsgBox an error2: %Error%
                    Else
                    {
                        i := 0
                        while (i < VDIDArr.MaxIndex()) {
                            if (VDIDArr[i] = Guid_ToStr(desktopID)) {
                                Win_VDID := i
                                break
                            }
                            i++
                        }
                    }
                    WinGet, Win_FullPathA, ProcessPath, ahk_id %id%
                    If (Win_VDID == Win_VDIDA AND Win_FullPathA == Win_FullPath) {
                        WinMove, %Win_Title%,,%Win_X%,%Win_Y%,%Win_W%,%Win_H%
                        ;Break
                    }
                }
            }
            Else If (Win_Title_Count == 1) {
                WinMove, %Win_Title%,,%Win_X%,%Win_Y%,%Win_W%,%Win_H%
            }
        } Else If WinExist("ahk_exe" . Win_FullPath) {
            WinGet windows, List
            Loop % windows
            {
                Win_IDA := windows%A_Index%
                WinGetTitle Win_TitleA, ahk_id %Win_IDA%
                If (Win_TitleA = "")
                    continue
                Win_TitleA := StrReplace(Win_TitleA, "`,")
                WinGetClass Win_ClassA, ahk_id %Win_IDA%
                If (Win_ClassA = "ApplicationFrameWindow") 
                {
                    WinGetText, text, ahk_id %Win_IDA%
                    If (text = "")
                        continue
                }
                WinGet, style, style, ahk_id %Win_IDA%
                if !(style & 0xC00000) OR !(style & 0x10000000) ; if the window doesn't have a title bar or is not on any virtual desktop
                {
                    ; If Win_TitleA not contains ...  ; add exceptions
                    continue
                }
                WinGetPos Win_XA, Win_YA, Win_WA, Win_HA, ahk_id %Win_IDA%
                if ((Win_XA = 0 OR Win_YA = 0) AND Win_WA = 0 AND Win_HA = 0)
                {
                    continue
                }
                WinGet, Win_FullPathA, ProcessPath, ahk_id %Win_IDA%
                
                Error := getWindowDesktopID(Win_IDA, desktopID)
                if !(Error=0) ;S_OK
                    MsgBox "an error3:" . Error
                Else
                {
                    i := 0
                    while (i < VDIDArr.MaxIndex()) {
                        if (VDIDArr[i] = Guid_ToStr(desktopID)) {
                            Win_VDID := i
                            break
                        }
                        i++
                    }
                }
                
                If (Win_FullPathA == Win_FullPath && Win_VDID == Win_VDIDA)
                {
                    WinMove, ahk_id %Win_IDA%,,%Win_X%,%Win_Y%,%Win_W%,%Win_H%
                    ;Break
                } Else {
                    OutputDebug % "Target: " . Win_FullPath . "`nCurrently tested: " . Win_FullPathA . "`nWin_VDID Target: " . Win_VDID . "`nWin_VDID Current: " . Win_VDIDA . "`nIndex: " . A_Index . "`nMaxIndex: " . windows.Count()
                }
            }
        } Else If (Win_StartIfNotRunning == 1) {
            switchDesktopByNumber(Win_VDID)
            Run %Win_FullPath%,,,CurrentAppNewPID
            WinWait ahk_pid %CurrentAppNewPID%,,5
            WinGetTitle, Win_TitleA, ahk_pid %CurrentAppNewPID%
            WinMove, %Win_TitleA%,,%Win_X%,%Win_Y%,%Win_W%,%Win_H%
        }
    }
    WinActivate, %SavedActiveWindow% ;Restore window that was active at beginning of script
    
    Return
}

Return ; <- Return of Main

;Save current windows to file
SaveWindows() {
    MsgBox 36, %AppTitle%, Save window positions?
    IfMsgBox, NO, Return
        
    WinGetActiveTitle, SavedActiveWindow
    
    SysGet, MonitorCount, MonitorCount
    SysGet, MonitorPrimary, MonitorPrimary
    WinGetPos, PMx, PMy, PMw, PMh, Program Manager
    IniRead, sectionNameList, %FileName%
    sectionName := "SECTION: Monitors=" . MonitorCount . ",MonitorPrimary=" . MonitorPrimary . "; Desktop size:" . PMx . "," . PMy . "," . PMw . "," . PMh
    
    If InStr(sectionNameList, sectionName)
        MsgBox 52, %AppTitle%, Configuration already exists. Overwrite?
    IfMsgBox, NO, Return
        IfMsgBox, YES
    {
        IniDelete, %FileName%, %sectionName%
    }
    
    ;Get all Virtual Desktops and put it in VDIDArr for access
    mapDesktopsFromRegistry()
    
    ;Get all non-hidden windows
    WinGet windows, List
    Pairs := ""
    Loop % windows
    {
        Win_ID := windows%A_Index%
        WinGetTitle Win_Title, ahk_id %Win_ID%
        If (Win_Title = "")
            continue
        Win_Title := StrReplace(Win_Title, "`,")
        WinGetClass Win_Class, ahk_id %Win_ID%
        If (Win_Class = "ApplicationFrameWindow") 
        {
            WinGetText, text, ahk_id %Win_ID%
            If (text = "")
                continue
        }
        WinGet, style, style, ahk_id %Win_ID%
        if !(style & 0xC00000) OR !(style & 0x10000000) ; if the window doesn't have a title bar or is not on any virtual desktop
        {
            ; If Win_Title not contains ...  ; add exceptions
            continue
        }
        WinGetPos Win_X, Win_Y, Win_W, Win_H, ahk_id %Win_ID%
        if ((Win_X = 0 OR Win_Y = 0) AND Win_W = 0 AND Win_H = 0)
        {
            continue
        }
        WinGet, Win_FullPath, ProcessPath, ahk_id %Win_ID%
        
        Error := getWindowDesktopID(Win_ID, desktopID)
        if !(Error=0) ;S_OK
            MsgBox an error: %Error%
        Else
        {
            i := 0
            while (i < VDIDArr.MaxIndex()) {
                if (VDIDArr[i] = Guid_ToStr(desktopID)) {
                    Win_VDID := i
                    break
                }
                i++
            }
        }
        
        Win_StartIfNotRunning := 1
            
        For index, value in StrSplit(ParmVals, A_Space)
        {
            Pairs .= value . "=" . Win_%value% . "`,"
        }
        Pairs .= "`n"
    }
    IniWrite, %Pairs%, %FileName%, %sectionName% 
    
    WinActivate, %SavedActiveWindow% ;Restore window that was active at beginning of script
    Return
}

showSettings() {
    
    global ParmVals, sectionName, DDL1
    changeSectionButton := false
    Loop % StrSplit(ParmVals, A_Space).MaxIndex()
    {
        Win_%A_LoopField% := ""
        }
    Gui, Destroy
    Gui, Add, DropDownList, w380 vDDL1,
    Gui, Add, Button, w100 x+m gChangeSection, Change Section
    Gui, Add, ListView, xm w1200 r15 -Multi -ReadOnly hwndHLV1, .
    Gui, Add, Button, Default xm y+m w100 gSettingsWindowOK, &OK
    Gui, Add, Button, x+m  w100 gCancel, &Cancel
    Gui, Default
    
    IniRead, sectionNameList, %FileName%
    Loop, Parse, sectionNameList, "`n"
    {
        GuiControl, , DDL1 , %A_LoopField%
        }
    
    ICELV2 := New LV_InCellEdit(HLV1, True, True)
    
    Gui, Show,, %AppTitle%
    
    Loop
    {
        If !(ChangeSectionButton)
            continue
        GuiControlGet, sectionName, , DDL1
        IniRead, sectionValues, %FileName%, %sectionName%
        
        LV_Delete()
        Loop % LV_GetCount("Col")
        {
            LV_DeleteCol(1)
        }
        
        If (sectionName == "Settings") {
            LV_InsertCol(1)
            LV_InsertCol(2,,"Setting Name")
            LV_InsertCol(3,,"Setting Value")
            Loop, Parse, sectionValues, "`n"
            {
                EqualPos := InStr(A_LoopField,"=")
                Var := SubStr(A_LoopField,1,EqualPos-1)
                Val := SubStr(A_LoopField,EqualPos+1)
                ;Remove any surrounding double quotes (")
                If (SubStr(Val,1,1)==Chr(34)) 
                {
                    Val := SubStr(Val, 2, StrLen(Val)-2)
                }
                LV_Add(, "",Var,Val)
            }
        } Else If (sectionName != "") {
            LV_InsertCol(1)
            For index, value in StrSplit(ParmVals, A_Space)
            {
                LV_InsertCol(index+1,,value)
            }
            Loop, Parse, sectionValues, "`n"
            {
                currentLine := A_LoopField
                Loop, Parse, currentLine, "`,"
                {
                    EqualPos := InStr(A_LoopField,"=")
                    Var := SubStr(A_LoopField,1,EqualPos-1)
                    Val := SubStr(A_LoopField,EqualPos+1)
                    If InStr(ParmVals, %Var%)
                    {
                        ;Remove any surrounding double quotes (")
                        If (SubStr(Val,1,1)==Chr(34)) 
                        {
                            Val := SubStr(Val, 2, StrLen(Val)-2)
                        }
                        Win_%Var% := Val
                    }
                }
                LV_Add(, "",Win_Title, Win_StartIfNotRunning, Win_Class, Win_ID, Win_FullPath, Win_X, Win_Y, Win_W, Win_H, Win_VDID)
                }
        }
        
        Loop % LV_GetCount("Col")
        {
            If (A_Index == 1) {
                LV_ModifyCol(1, 0)
                    continue
            }
            LV_ModifyCol(A_Index,"+AutoHdr")	
            }
        
        changeSectionButton := false
    }
    return
    
    ;----------------------
    SettingsWindowOK:
        global sectionName
        IniDelete, %FileName%, %sectionName%
        Pairs := ""
        Loop % LV_GetCount()
        {
            currentRowNr := A_Index
            For index, value in StrSplit(ParmVals, A_Space)
            {
                LV_GetText(Win_%value%, currentRowNr, index+1)
                Pairs .= value . "=" . Win_%value% . "`,"
            }
            Pairs .= "`n"
        }
        IniWrite, %Pairs%, %FileName%, %sectionName% 
        Gui, Cancel
    return
    ;----------------------
    ChangeSection:
        changeSectionButton := true
    return
}

Exit:
    ExitApp
    
About:
    MsgBox 64, %AppTitle%,
    (
    App created by %Author%
    For feedback and questions contact %Email%
        
    This app was created to be able to save and restore window positions and sizes when docking/undocking your notebook.
    
How to set up:
    1. Save your current window positions and sizes by pressing %SaveCombo%
    2. Undock/Dock your notebook - to have a different resolution/no. of monitors than before - and arrange the windows as you like
        3. Save this configuration also by pressing %SaveCombo%
    (You can add as many monitor setups as you like. If a monitor setup was saved before, it can be overwritten)
        
    When you Dock/Undock again, just press %LoadCombo% and it will restore your previously saved configuration.
    More info on the hotkeys can be found in the %FileName% file
    )
Return

Reload:
    Reload
Return

CheckForUpdate:
    MsgBox 64, %AppTitle%, Redirecting to GitHub page.., 4
    Run, %Website%
Return

mapDesktopsFromRegistry() {
    global VDIDArr, CurrentDesktop, DesktopCount
    
    ; Get the current desktop UUID. Length should be 32 always, but there's no guarantee this couldn't change in a later Windows release so we check.
    IdLength := 32
    SessionId := getSessionId()
    if (SessionId) {
        RegRead, CurrentDesktopId, HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\SessionInfo\%SessionId%\VirtualDesktops, CurrentVirtualDesktop
        if (CurrentDesktopId) {
            IdLength := StrLen(CurrentDesktopId)
        }
    }
    ; Get a list of the UUIDs for all virtual desktops on the system
    RegRead, DesktopList, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VirtualDesktops, VirtualDesktopIDs
    if (DesktopList) {
        DesktopListLength := StrLen(DesktopList)
        ; Figure out how many virtual desktops there are
        DesktopCount := DesktopListLength / IdLength
    }
    else {
        DesktopCount := 1
    }
    ; Parse the REG_DATA string that stores the array of UUID's for virtual desktops in the registry.
    i := 0
    while (CurrentDesktopId and i < DesktopCount) {
        StartPos := (i * IdLength) + 1
        DesktopIter := SubStr(DesktopList, StartPos, IdLength)
        If (DesktopIter == CurrentDesktopId) {
            CurrentDesktop := i
        }
        VDIDArr[i] := "{" . SubStr(DesktopIter,7,2) . SubStr(DesktopIter,5,2) . SubStr(DesktopIter,3,2) . SubStr(DesktopIter,1,2) . "-" . SubStr(DesktopIter,11,2) . SubStr(DesktopIter,9,2) . "-" . SubStr(DesktopIter,15,2) . SubStr(DesktopIter,13,2) . "-" . SubStr(DesktopIter,17,4) . "-" . SubStr(DesktopIter,21,12) . "}"
        
        i++
    }
return
}

getSessionId() {
    ProcessId := DllCall("GetCurrentProcessId", "UInt")
    if ErrorLevel {
        OutputDebug, Error getting current process id: %ErrorLevel%
        return
    }
    OutputDebug, Current Process Id: %ProcessId%
    DllCall("ProcessIdToSessionId", "UInt", ProcessId, "UInt*", SessionId)
    if ErrorLevel {
        OutputDebug, Error getting session id: %ErrorLevel%
        return
    }
    OutputDebug, Current Session Id: %SessionId%
return SessionId
}

getWindowDesktopID(hWnd, ByRef desktopID) {
    ;returns the GUID of the virtual desktop that contains the given hwnd (you might have to own the process to get the guid)
    ;IVirtualDesktopManager interface
    ;Exposes methods that enable an application to interact with groups of windows that form virtual workspaces.
    ;https://msdn.microsoft.com/en-us/library/windows/desktop/mt186440(v=vs.85).aspx
    VarSetCapacity(desktopID, 16, 0) ; a GUID structure occupies 16 bytes in memory; allocate the space for one for ::GetWindowDesktopId to fill in
    
    CLSID := "{aa509086-5ca9-4c25-8f95-589d3c07b48a}" ;search VirtualDesktopManager clsid
    IID := "{a5cd92ff-29be-454c-8d04-d82879fb3f1b}" ;search IID_IVirtualDesktopManager
    IVirtualDesktopManager := ComObjCreate(CLSID, IID)
    
    ;IVirtualDesktopManager::GetWindowDesktopId  method
    ;https://msdn.microsoft.com/en-us/library/windows/desktop/mt186441(v=vs.85).aspx
    Error := DllCall(NumGet(NumGet(IVirtualDesktopManager+0), 4*A_PtrSize), "Ptr", IVirtualDesktopManager, "Ptr", hWnd, "Ptr", &desktopID)
    
    ;free IVirtualDesktopManager
    ObjRelease(IVirtualDesktopManager)
    
return Error
}

Guid_ToStr(ByRef VarOrAddress) {
    pGuid := IsByRef(VarOrAddress) ? &VarOrAddress : VarOrAddress
    VarSetCapacity(sGuid, 78) ; (38 + 1) * 2
    if !DllCall("ole32\StringFromGUID2", "Ptr", pGuid, "Ptr", &sGuid, "Int", 39)
        throw Exception("Invalid GUID", -1, Format("<at {1:p}>", pGuid))
return StrGet(&sGuid, "UTF-16")
}

switchDesktopByNumber(targetDesktop) {
    global CurrentDesktop, DesktopCount
    ; Re-generate the list of desktops and where we fit in that.
    mapDesktopsFromRegistry()
    ; Don't attempt to switch to an invalid desktop
    if (targetDesktop > DesktopCount || targetDesktop < 0) {
        OutputDebug, [invalid] target: %targetDesktop% current: %CurrentDesktop%
        return
    }
    ; Go right until we reach the desktop we want
    while(CurrentDesktop < targetDesktop) {
        Send ^#{Right}
        CurrentDesktop++
        OutputDebug, [right] target: %targetDesktop% current: %CurrentDesktop%
    }
    ; Go left until we reach the desktop we want
    while(CurrentDesktop > targetDesktop) {
        Send ^#{Left}
        CurrentDesktop--
        OutputDebug, [left] target: %targetDesktop% current: %CurrentDesktop%
    }
return
}

initMenu() {
    Menu, Tray, DeleteAll
    Menu, Tray, Icon, Icon.ico
    Menu, Tray, Tip, %AppTitle%
    Menu, Tray, NoStandard
    Menu, Tray, Add, Settings, showSettings
    Menu, Tray, Add, About, About
    Menu, Tray, Add, Reload, Reload
    Menu, Tray, Add, Check for update, CheckForUpdate
        Menu, Tray, Add
    Menu, Tray, Add, Exit, Exit
    Menu, Tray, Default, About
return
}

initApp() {
    global FileName
    FileRead, FileText, %FileName%
    IniRead, SettingsSection, %Filename%, Settings
    If (FileExist(FileName) == "" OR FileText == "" OR SettingsSection == "") {
        FileAppend , , %FileName%
        IniWrite, ^F12, %FileName%, Settings, SaveCombo
        IniWrite, ^F1, %FileName%, Settings, LoadCombo
        IniWrite, For a list of special key symbols go to https://www.autohotkey.com/docs/Hotkeys.htm, %FileName%, Settings, Info
        }
    IniRead, SaveCombo, %FileName%, Settings, SaveCombo
    IniRead, LoadCombo, %FileName%, Settings, LoadCombo
    
    Hotkey, %SaveCombo%, SaveWindows
    Hotkey, %LoadCombo%, RestoreWindows
    
    MsgBox  64, %AppTitle%,
    (
    Welcome to %AppTitle%
    
    To save window positions press %SaveCombo%
    To Load: %LoadCombo%
    More info on the hotkeys can be found in the %FileName% file
    )
    
return
}