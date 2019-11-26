; Window Saver
; Save and Restore window positions when docking/undocking
;
; Author: David Szilvasi
; Email: David_Szilvasi@Dell.com
; Version : v0.8
;
; Tested and minimum required AHK version: 1.1.32
;
;To-do: start with windows, when opening apps open on proper VD

#NoEnv
SendMode Input
SetWorkingDir %A_ScriptDir%
#SingleInstance, Force
DetectHiddenWindows, On
SetTitleMatchMode, 2

AppTitle := "Window Saver"
SaveCombo := "Ctrl+F12"
LoadCombo := "Ctrl+F1"
FileName :="window.cfg"
Menu, Tray, Icon, WindowSaver.png

MsgBox  0, %AppTitle%, Welcome to %AppTitle% `n`nTo save window positions press %SaveCombo%`nTo Load: %LoadCombo%

SysGet, OriginalMonCount, MonitorCount
SysGet, OriginalMonitorPrimary, MonitorPrimary
WinGetPos, Originalx, Originaly, OriginalWidth, OriginalHeight, Program Manager

;SetTimer, GetMonCount, 10000

;Save current windows to file
^F12::
SaveWindows:
	MsgBox  4, %AppTitle%, Save window positions?
		IfMsgBox, NO, Return

	WinGetActiveTitle, SavedActiveWindow

	SysGet, MonitorCount, MonitorCount
	SysGet, MonitorPrimary, MonitorPrimary
    WinGetPos, PMx, PMy, PMw, PMh, Program Manager
	IniRead, sectionNameList, %Filename%
	sectionName := "SECTION: Monitors=" . MonitorCount . ",MonitorPrimary=" . MonitorPrimary . "; Desktop size:" . PMx . "," . PMy . "," . PMw . "," . PMh

	If InStr(sectionNameList, sectionName)
		MsgBox 4, %AppTitle%, Configuration already exists. Overwrite?
			IfMsgBox, NO, Return
			IfMsgBox, YES
			{
				IniDelete, %FileName%, %sectionName%
			}

	;Get all non-hidden windows
	WinGet windows, List
	Pairs := ""
	Loop % windows
	{
		id := windows%A_Index%
		WinGetTitle Win_Title, ahk_id %id%
		If (Win_Title = "")
			continue
		WinGetClass class, ahk_id %id%
		If (class = "ApplicationFrameWindow") 
		{
			WinGetText, text, ahk_id %id%
			If (text = "")
				continue
		}
		WinGet, style, style, ahk_id %id%
		if !(style & 0xC00000) OR !(style & 0x10000000) ; if the window doesn't have a title bar or is not on any virtual desktop
		{
			; If Win_Title not contains ...  ; add exceptions
				continue
		}
		WinGetPos Win_X, Win_Y, Win_W, Win_H, ahk_id %id%
		WinGet, PID, PID , ahk_id %id%
		Win_FullPath := GetModuleExeName(PID)
		if ((Win_X = 0 OR Win_Y = 0) AND Win_W = 0 AND Win_H = 0)
		{
			continue
		}

		Pairs .= "Title=" . Win_Title . "`,Class=" . class . "`,ID=" . id . "`,FullPath=" . Win_FullPath . "`,X=" . Win_X . "`,Y=" . Win_Y . "`,W=" . Win_W . "`,H=" . Win_H . "`n"
	}
	IniWrite, %Pairs%, %Filename%, %sectionName% 

  	WinActivate, %SavedActiveWindow% ;Restore window that was active at beginning of script
	Return
Return

;Restore window positions from file
^F1::
RestoreWindows:
	WinGetActiveTitle, SavedActiveWindow
  	ParmVals := "Title Class ID FullPath X Y W H"
	Win_Title:="", Win_Class:="", Win_ID:="", Win_FullPath:="", Win_X:=0, Win_Y:=0, Win_W:=0, Win_H:=0
	
	SysGet, MonitorCount, MonitorCount
	SysGet, MonitorPrimary, MonitorPrimary
    WinGetPos, PMx, PMy, PMw, PMh, Program Manager
	IniRead, sectionNameList, %Filename%
	sectionName := "SECTION: Monitors=" . MonitorCount . ",MonitorPrimary=" . MonitorPrimary . "; Desktop size:" . PMx . "," . PMy . "," . PMw . "," . PMh

	If !InStr(sectionNameList, sectionName)
		MsgBox 4, %AppTitle%, Current configuration wasnt found.`nWould you like to save first?
			IfMsgBox, Yes, Goto SaveWindows

	IniRead, sectionValues, %FileName%, %sectionName%
	sectionValuesArr := StrSplit(sectionValues , "`n")

	Loop % sectionValuesArr.MaxIndex()
	{
		currentLine := sectionValuesArr[A_Index]
		Loop, Parse, currentLine, "`,"
		{
			EqualPos := InStr(A_LoopField,"=")
			Var := SubStr(A_LoopField,1,EqualPos-1)
			Val := SubStr(A_LoopField,EqualPos+1)
			If InStr(ParmVals, %Var%)
			{
				;Remove any surrounding double quotes (")
				If (SubStr(Val,1,1)=Chr(34)) 
				{
					Val := SubStr(Val, 2, StrLen(Val)-2)
				}
				Win_%Var% := Val
			}
		}
		; Try to find if window is already open. If it wasnt found, open a new window using it's path
		If WinExist("ahk_id" . Win_ID) {
			WinMove, ahk_id %Win_ID%,,%Win_X%,%Win_Y%,%Win_W%,%Win_H%
			;sleep 100
		} Else If WinExist(Win_Title) {
			WinMove, %Win_Title%,,%Win_X%,%Win_Y%,%Win_W%,%Win_H%
			;sleep 100
		} Else If WinExist("ahk_class" . Win_Class) {
			WinMove, ahk_class %Win_Class%,,%Win_X%,%Win_Y%,%Win_W%,%Win_H% 
		} Else If WinExist("ahk_exe" . Win_FullPath) {
			WinGet, Win_Title , , ahk_exe %Win_FullPath%
			WinMove, ahk_id %Win_ID%,,%Win_X%,%Win_Y%,%Win_W%,%Win_H%
		} Else {
			Run %Win_FullPath%,,,CurrentAppNewPID
			sleep 1000
			WinMove, ahk_pid %CurrentAppNewPID%,,%Win_X%,%Win_Y%,%Win_W%,%Win_H% ; This line isnt working
		}
	}

  	WinActivate, %SavedActiveWindow% ;Restore window that was active at beginning of script
	Return
Return

GetModuleExeName(PID) 
{
	for process in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process where ProcessId=" PID)
		return process.ExecutablePath
}

; GetMonCount:
; 	SysGet, MonCount, MonitorCount
; 	SysGet, MonitorPrimary, MonitorPrimary
; 	WinGetPos, tmp_x, tmp_y, tmp_Width, tmp_Height, Program Manager

; 	if MonCount != OriginalMonCount OR MonitorPrimary != OriginalMonitorPrimary OR tmp_Width != OriginalWidth OR tmp_Height != OriginalHeight
; 	{
; 		OriginalMonCount := MonCount
; 		OriginalMonitorPrimary := MonitorPrimary
; 		OriginalWidth := tmp_Width
; 		OriginalHeight := tmp_Height
; 		GoSub RestoreWindows
;  	}
;  	return