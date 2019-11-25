; Window Saver
; Save and Restore window positions when docking/undocking
;
; Author: David Szilvasi
; Email: David_Szilvasi@Dell.com
; Version : v0.6
;
; Tested and minimum required AHK version: 1.1.32+
;
;To-do: start with windows, when opening apps open on proper VD

#NoEnv
SendMode Input
SetWorkingDir %A_ScriptDir%
#SingleInstance, Force
DetectHiddenWindows, On
SetTitleMatchMode, 2

AppTitle := "Window Saver"
SaveCombo := "Ctrl+0"
LoadCombo := "Ctrl+1"
FileName :="window.cfg"
CrLf=`r`n
Menu, Tray, Icon, WindowSaver.png

MsgBox  , %AppTitle%, Welcome to %AppTitle% `n`nTo save window positions press %SaveCombo%`nLoading window positions is automatic when you plug in or disconnect a monitor

SysGet, OriginalMonCount, MonitorCount
SysGet, OriginalMonitorPrimary, MonitorPrimary
WinGetPos, Originalx, Originaly, OriginalWidth, OriginalHeight, Program Manager

;SetTimer, GetMonCount, 10000

;Save current windows to file
^1::
	MsgBox  4, %AppTitle%, Save window positions?
		IfMsgBox, NO, Return

	WinGetActiveTitle, SavedActiveWindow

	file := FileOpen(FileName, "a")
	if !IsObject(file)
	{
		MsgBox, Can't open "%FileName%" for writing.
		Return
	}
  	Loop, Read, %FileName%
  	{
		If (SubStr(A_LoopReadLine,1) == SectionHeader())
		{
			MsgBox  , %AppTitle%, Data for your current monitor/virtual desktop/resolution setup already exists! Please delete section from line %A_Index% in the configuration file.`nYour current data is now also saved
			Break
		}
	}
	line:= SectionHeader() . CrLf
	file.Write(line)

	;Get all non-hidden windows
	WinGet windows, List
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
		if !(style & 0xC00000) OR !(style & 0x10000000) ; if the window doesn't have a wintitle bar or is not on any virtual desktop
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
		line := "Title=" . Win_Title . "`,Class=" . class . "`,ID=" . id . "`,FullPath=" . Win_FullPath . "`,X=" . Win_X . "`,Y=" . Win_Y . "`,W=" . Win_W . "`,H=" . Win_H . "`r`n"
		file.Write(line)
	}
	file.write(CrLf)  ;Add blank line after section
  	file.Close()
  	WinActivate, %SavedActiveWindow% ;Restore window that was active at beginning of script
	Return


;Restore window positions from file
^1::
RestoreWindows:
	WinGetActiveTitle, SavedActiveWindow
  	ParmVals := "Title Class ID FullPath X Y W H"
  	SectionToFind := SectionHeader()
  	SectionFound := 0

  	Loop, Read, %FileName%
  	{
		;If no section was found and this line is not the current config section line->jump to next iteration
    	if !SectionFound AND (A_LoopReadLine!=SectionToFind) 
			Continue
		
		;Exit if another section was already found and iteration reached end of recorded apps
		If (SectionFound AND SubStr(A_LoopReadLine,1,8) = "SECTION:")
			Break
		
        SectionFound:=1
		Win_Title:="", Win_Class:="", Win_ID:="", Win_FullPath:="", Win_X:=0, Win_Y:=0, Win_W:=0, Win_H:=0
		
		If (A_LoopReadLine == SectionToFind) OR (A_LoopReadLine == "")
			Continue ; Jump to next iteration to start reading data if current config section found or if line empty
		Loop, Parse, A_LoopReadLine, "`,"
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
		WinRestore
		WinActivate
		; Try to find if window is already open. If it wasnt found, open a new window using it's path
		If WinExist("ahk_id" . Win_ID) {
			WinMove, ahk_id %Win_ID%,,%Win_X%,%Win_Y%,%Win_W%,%Win_H%
		} Else If WinExist(Win_Title) {
		 	WinMove, %Win_Title%,,%Win_X%,%Win_Y%,%Win_W%,%Win_H%
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

  	if !SectionFound
  	{
    	MsgBox,,%AppTitle%, Section does not exist in %FileName% `nLooking for: %SectionToFind%`n`nTo address this issue, you can press %SaveCombo% to save your current setup!
  	}
	
  	WinActivate, %SavedActiveWindow% ;Restore window that was active at beginning of script
	Return
return

;Create standardized section header for later retrieval
SectionHeader()
{
	SysGet, MonitorCount, MonitorCount
	SysGet, MonitorPrimary, MonitorPrimary
	line := "SECTION: Monitors=" . MonitorCount . ",MonitorPrimary=" . MonitorPrimary

    WinGetPos, x, y, Width, Height, Program Manager
	line:= line . "; Desktop size:" . x . "," . y . "," . width . "," . height

	Return %line%
}

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