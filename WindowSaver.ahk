﻿; Window Saver
; Save and Restore window positions when docking/undocking
;
; Tested and minimum required AHK version: 1.1.32
;

#NoEnv
#SingleInstance, Force
SendMode Input
SetWorkingDir %A_ScriptDir%
FileInstall, version.data, version.data
FileInstall, Icon.ico, Icon.ico
DetectHiddenWindows, On
SetTitleMatchMode, 2
#KeyHistory, 0
ListLines, Off

FileName :="window.cfg"
Author := "David Szilvasi"
Email := "szilvasi.dave@gmail.com"
FileRead, AppVersion, version.data
AppTitle := "Window Saver " . AppVersion
debug := 0

Menu, Tray, Icon, Icon.ico
Menu, Tray, Tip, %AppTitle%
Menu, Tray, NoStandard
Menu, Tray, Add, About, About
Menu, Tray, Add, Reload, Reload
Menu, Tray, Add, Check for update, CheckForUpdate
Menu, Tray, Add, Debug Mode, DebugMode
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

;Check every Monday for updates
If A_WDay == 2 ;If it's Monday, check for updates
	GoSub CheckForUpdate

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
		;WinGet, Win_Controls, ControlList , ahk_id %id%  "Win_Controls=" . Win_Controls .
		Pairs .= "Title=" . Win_Title . "`,Class=" . class . "`,ID=" . id . "`,FullPath=" . Win_FullPath . "`,X=" . Win_X . "`,Y=" . Win_Y . "`,W=" . Win_W . "`,H=" . Win_H . "`n"
	}
	If debug == 1
		MsgBox % Pairs
	IniWrite, %Pairs%, %Filename%, %sectionName% 

  	WinActivate, %SavedActiveWindow% ;Restore window that was active at beginning of script
	Return

;Restore window positions from file
RestoreWindows:
	;DetectHiddenWindows, Off
	WinGetActiveTitle, SavedActiveWindow
  	ParmVals := "Title Class ID FullPath X Y W H Controls"
	Win_Title:="", Win_Class:="", Win_ID:="", Win_FullPath:="", Win_X:=0, Win_Y:=0, Win_W:=0, Win_H:=0, Win_Controls:=""
	
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
				If (SubStr(Val,1,1)==Chr(34)) 
				{
					Val := SubStr(Val, 2, StrLen(Val)-2)
				}
				Win_%Var% := Val
			}
		}
		; Try to find if window is already open. If it wasnt found, open a new window using it's path
		WinGet, Win_Class_Count, Count, ahk_class %Win_Class%
		WinGet, Win_Title_Count, Count, %Win_Title%
		If WinExist("ahk_id" . Win_ID) {
			;MsgBox Using ID - %Win_Title%
			WinMove, ahk_id %Win_ID%,,%Win_X%,%Win_Y%,%Win_W%,%Win_H%
		} Else If (WinExist("ahk_class" . Win_Class)) { ; AND (Win_Class_Count == 1)
			;MsgBox Using class- %Win_Title%
			WinMove, ahk_class %Win_Class%,,%Win_X%,%Win_Y%,%Win_W%,%Win_H%
		} Else If (WinExist(Win_Title)){ ; AND (Win_Title_Count == 1)
			;MsgBox Using Title - %Win_Title%
			WinMove, %Win_Title%,,%Win_X%,%Win_Y%,%Win_W%,%Win_H%
		} Else If WinExist("ahk_exe" . Win_FullPath) {
			;MsgBox Using EXE - %Win_Title%
			WinMove, ahk_exe %Win_FullPath%,,%Win_X%,%Win_Y%,%Win_W%,%Win_H%
		} Else {
			;MsgBox Starting program - %Win_Title%
			Run %Win_FullPath%,,,CurrentAppNewPID
			WinWait, ahk_pid %CurrentAppNewPID%
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

CheckForUpdate:
; 	UrlDownloadToFile, https://david.szilvasi.family/WindowSaver/version.data, version_new.data
; 	FileRead, AppVersion_new, version_new.data
; 	If (AppVersion != AppVersion_new)
; 	{
; 		MsgBox 4, %AppTitle%,
; 		(
; A newer, better version of %AppTitle% is available!
; Current version: %AppVersion%
; New version: %AppVersion_new%

; Would you like to update?
; 		)
; 		IfMsgBox, Yes
; 		{
; 			UrlDownloadToFile, https://dell.box.com/shared/static/yrsal6y2pg3xentufcxb46zn9nft8hry.ex_e, WindowSaver.exe
; 			FileCopy, version_new.data, version.data, 1
; 			FileDelete, version_new.data
; 			Goto Reload
; 		}
; 	} Else {
		MsgBox 0, %AppTitle%, You already have the latest version of %AppTitle%!
	; }
	Return

DebugMode:
	If (debug == 0) {
		MsgBox 4, %AppTitle%, This mode is for debugging errors in the app. Please use it ONLY if you know what you're doing!`nAre you sure you want to continue?
		IfMsgBox, YES
		{
			debug := 1
			#KeyHistory, 10
			ListLines, On
			Menu, Tray, ToggleCheck, Debug Mode
			;SplashTextOn, 500, 500, Debug Mode, %KeyHistory%
		}
	} Else {
		debug := 0
		Menu, Tray, ToggleCheck, Debug Mode
		;SplashTextOff
	}
	Return