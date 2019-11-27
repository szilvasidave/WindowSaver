; Window Saver
; Save and Restore window positions when docking/undocking
;
; Tested and minimum required AHK version: 1.1.32
;
;To-do: start with windows, when opening apps open on proper VD

#NoEnv
#SingleInstance, Force
SendMode Input
SetWorkingDir %A_ScriptDir%
DetectHiddenWindows, On
SetTitleMatchMode, 2
FileEncoding , UTF-16

AppTitle := "Window Saver" . AppVersion
FileName :="window.cfg"
Author := David Szilvasi
Email := David_Szilvasi@Dell.com
AppVersion := " 0.9"


Menu, Tray, Icon, Icon.ico
Menu, Tray, Tip, %AppTitle%
Menu, Tray, NoStandard
Menu, Tray, Add, About, About
Menu, Tray, Add, Reload, Reload
Menu, Tray, Add
Menu, Tray, Add, Exit, Exit
Menu, Tray, Default, About

If FileExist(FileName) == ""
	{
	FileAppend , , %FileName%
	IniWrite, "^F12", %FileName%, Settings, SaveCombo
	IniWrite, "^F1", %FileName%, Settings, LoadCombo
	IniWrite , "For a list of special keys' symbols go to https://www.autohotkey.com/docs/Hotkeys.htm", %FileName%, Settings, Info

	}
IniRead, SaveCombo, %FileName%, Settings, SaveCombo
IniRead, LoadCombo, %FileName%, Settings, LoadCombo

Hotkey, %SaveCombo%, SaveWindows
Hotkey, %LoadCombo%, RestoreWindows

MsgBox  0, %AppTitle%,
(
Welcome to %AppTitle%

To save window positions press %SaveCombo%
To Load: %LoadCombo%
More info on the hotkeys can be found in the %FileName% file
)

SysGet, OriginalMonCount, MonitorCount
SysGet, OriginalMonitorPrimary, MonitorPrimary
WinGetPos, Originalx, Originaly, OriginalWidth, OriginalHeight, Program Manager
Return
;SetTimer, GetMonCount, 10000

;Save current windows to file
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
		Win_Title := StrReplace(Win_Title, "`,")
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
		if ((Win_X = 0 OR Win_Y = 0) AND Win_W = 0 AND Win_H = 0)
		{
			continue
		}
		WinGet, Win_FullPath, ProcessPath, ahk_id %id%
		Pairs .= "Title=" . Win_Title . "`,Class=" . class . "`,ID=" . id . "`,FullPath=" . Win_FullPath . "`,X=" . Win_X . "`,Y=" . Win_Y . "`,W=" . Win_W . "`,H=" . Win_H . "`n"
	}
	IniWrite, %Pairs%, %Filename%, %sectionName% 

  	WinActivate, %SavedActiveWindow% ;Restore window that was active at beginning of script
	Return

;Restore window positions from file
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
		} Else If WinExist(Win_Title) {
			WinMove, %Win_Title%,,%Win_X%,%Win_Y%,%Win_W%,%Win_H%

		} Else If WinExist("ahk_exe" . Win_FullPath) {
			;WinGet, Win_Title, "ahk_exe" . Win_FullPath
			WinGet, Win_ID, ID , ahk_exe %Win_FullPath%
			WinMove, ahk_id %Win_ID%,,%Win_X%,%Win_Y%,%Win_W%,%Win_H%
		} Else {
			Run %Win_FullPath%,,,CurrentAppNewPID
			sleep 1000
			WinMove, ahk_pid %CurrentAppNewPID%,,%Win_X%,%Win_Y%,%Win_W%,%Win_H% ; This line isnt working
		}
	}

  	WinActivate, %SavedActiveWindow% ;Restore window that was active at beginning of script
	Return

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

Exit:
	ExitApp

About:
	MsgBox 0, %AppTitle%,
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