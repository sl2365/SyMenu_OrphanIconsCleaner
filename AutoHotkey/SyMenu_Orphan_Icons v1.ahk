;DESCRIPTION:
;	Script for AutoHotKey for Move/Delete the SyMenu Orphan Icons
;CHANGELOG:
; Requires v1.1 autohotkey to compile.
;		V.1.0.-2017.01.27: First version
;
;NOTES:
;
;------------------------------------------------------------
;Compilation orders
	#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
	#Warn  ; Enable warnings to assist with detecting common errors.
	#SingleInstance Off  ;The word OFF allows multiple instances of the script to run concurrently
	;#NoTrayIcon ;Disables the showing of a tray icon.
	SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
	SplitPath, A_ScriptName,,,, ScriptName ;Get ScriptName for windows title.
	;SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

;Program behavior and Global variables
	global s_LogFile:= A_ScriptDir . "\" . ScriptName . ".log"
	global b_DeleteAction:= False		; Delete the Orphan icon: True= Delete, False=Move to %SyMenuPath%\ProgramFiles\SPSSuite\SyMenuSuite\_Trash\_OrphanIcons
	global b_NewLogFile:= True			; Delete the present Log File and start one new: False= Append, True= Star new one
	global b_VerboseLog:= False			; Log (show) all messages with script action over files: False= No, True= Yes
	global s_TextLog:= ""				; Program Log
	global s_SyMenuPath:=""				; Working SyMenu path so icon folder= %SyMenuPath%\Icons , config folder=%SyMenuPath%\Config, Trash folder %SyMenuPath%\ProgramFiles\SPSSuite\SyMenuSuite\_Trash
	global b_Debug:= False				; Full Debug mode. Normaly used for extended verbose messages of large loops for calculate (without action) over file.

;Make the GUI for get the script input parameters and action variables
	Gui 1: -Disabled +SysMenu -Owner  ; -AlwaysOnTop ; +Owner No tray button but provides a normal title bar. +ToolWindow No tray button but provides a narrower title bar.
	Gui 1:Add, Button, x10 y8 w105 Default gGetFolder, Get SyMenu Folder
	Gui 1:Font, Bold
	Gui 1:Color, cWhite
	Gui 1:Add, Text, x120 y10 w880 cBlue Border, % "  " . s_SyMenuPath
	Gui 1:Font, Norm
	Gui 1:Color, Default
	
	Gui 1:Add, CheckBox, x10 vb_DeleteAction, ¿Delete? (or only move to     'SyMenu\ProgramFiles\SPSSuite\SyMenuSuite\_Trash\_OrphanIcons'     folder)

	Gui 1:Add, Text, cBlue, Log Window:
	Gui 1:Add, Edit, x10 vs_TextLog w990 r20 ,
	Gui 1:Add, CheckBox, x10 y380 vb_NewLogFile Checked, ¿Delete the present Log File and start one new?
	Gui 1:Add, CheckBox, x800 y380 vb_VerboseLog , ¿Full information log (verbose mode)?
	Gui 1:Add, Button, x10 y400 w75 gScriptProcess, ¡Go!
	Gui 1:Show
	Return  ;The script wait in the Gui window showed. It continue with the Gui events.

GetFolder:
	;Get folder for recursive search
	FileSelectFolder, s_NewWorkPath, , 2, ¿May you select the folder of the working SyMenu?
	if (s_NewWorkPath = "")
		Return
	s_SyMenuPath:= RegExReplace(s_NewWorkPath, "\\$")  ; Removes the trailing backslash, if present.
	GuiControl,, Static1, % "  " . s_SyMenuPath  ;Search the control name with AutoHotKey Window Spy
	Return
GuiClose:
GuiEscape:
	if b_VerboseLog ScriptLog(A_YYYY . "." . A_MM . "." . A_DD . "-" . A_Hour . ":" . A_Min . ":" . A_Sec . "-" . ScriptName "-(" . b_DeleteAction . "|" . b_NewLogFile . "|" . b_VerboseLog . ") canceled. ¡Good bye, " . A_UserName . "!" . "|`n", 0)
	Gui 1:Destroy	
	ExitApp
ScriptProcess:
	Gui 1:Submit  ; Save the input from the user to each control's associated variable.
	Gui 1:Show
	;Create vars for easy comprehension
	s_SyMenuPathIcons:= s_SyMenuPath . "\Icons"
	s_SyMenuPathConfig:= s_SyMenuPath . "\Config"
	s_SyMenuPath_Trash:= s_SyMenuPath . "\ProgramFiles\SPSSuite\SyMenuSuite\_Trash"
	s_OrphanIconsPath:= s_SyMenuPath_Trash . "\_OrphanIcons"
	;Cleaning Logfile
	if (b_NewLogFile and FileExist(s_LogFile))
		FileDelete, %s_LogFile%
	;Get the SyMenuItem.xlm file in the array o_SyMenuItemConfig
	try{  ;Create the s_OrphanIconsPath folder and unzip the 'SyMenuItem.xml' file to it.
		if !FileExist(s_OrphanIconsPath)
			FileCreateDir, % s_OrphanIconsPath
		if FileExist(s_OrphanIconsPath . "\_TmpZip") ;Unz needs a clean folder to comprobate that all files are unzipped.
			FileDelete, % s_OrphanIconsPath . "\_TmpZip\*.*"
		Unz(s_SyMenuPathConfig . "\SyMenuItem.zip", s_OrphanIconsPath . "\_TmpZip\")
		FileMove, % s_OrphanIconsPath . "\_TmpZip\SyMenuItem.xml" , % s_OrphanIconsPath , 1
		FileRemoveDir, % s_OrphanIconsPath . "\_TmpZip", 1
	}	
	catch excep {
		MsgBox 16, % A_ThisFunc . "-(" . A_LineNumber . ")", % "Error catch " excep.What ", which was called at line " excep.Line ", in:`n|" s_OrphanIconsPath "|"
		ExitApp            
	}	
	;Get SyMenuItem.xml text.
	s_SyMenuItemConfig:=""
	FileRead, s_SyMenuItemConfig, % s_OrphanIconsPath . "\SyMenuItem.xml"
	;Scan the icons files
	Loop, %s_SyMenuPathIcons%\*.ico{
		if (0 < InStr(s_SyMenuItemConfig , A_LoopFileName)){
			if b_VerboseLog
				ScriptLog(A_YYYY . "." . A_MM . "." . A_DD . "-" . A_Hour . ":" . A_Min . ":" . A_Sec . "-" . ScriptName "-(" . b_DeleteAction . "|" . b_NewLogFile . "|" . b_VerboseLog . ") Active Icon detected:" . A_LoopFileName . "|`n",1)
		}else{
			ScriptLog(A_YYYY . "." . A_MM . "." . A_DD . "-" . A_Hour . ":" . A_Min . ":" . A_Sec . "-" . ScriptName "-(" . b_DeleteAction . "|" . b_NewLogFile . "|" . b_VerboseLog . ") Orphan Icon moved: " . A_LoopFileName . "|`n", 2)
			FileMove, %A_LoopFileFullPath%, %s_OrphanIconsPath%, 1
		}
	}
	;Finish
	if b_DeleteAction { ;Delete the Orphan icon: True= Delete, False=Move to %SyMenuPath%\ProgramFiles\SPSSuite\SyMenuSuite\_Trash\_OrphanIcons
		if b_VerboseLog 
			ScriptLog(A_YYYY . "." . A_MM . "." . A_DD . "-" . A_Hour . ":" . A_Min . ":" . A_Sec . "-" . ScriptName "-(" . b_DeleteAction . "|" . b_NewLogFile . "|" . b_VerboseLog . ") Delete moved Icons" . "|`n",0)
		FileRemoveDir, % s_OrphanIconsPath, 1
	}
	ScriptLog(A_YYYY . "." . A_MM . "." . A_DD . "-" . A_Hour . ":" . A_Min . ":" . A_Sec . "-" . ScriptName "-(" . b_DeleteAction . "|" . b_NewLogFile . "|" . b_VerboseLog . ") has finish." . A_UserName . "!------------------------------------------------------------------------------------------`n", 0)
	MsgBox, 64, % ScriptName, % ScriptName "-(" . b_DeleteAction . "|" . b_NewLogFile . "|" . b_VerboseLog . ") has finish.`n" . A_UserName . "!" 
	Return
;==========================================================================================
;SUBRUTINES--------------------------------------------------------------------------------
ScriptLog(_s_Log, _i_TabDeep=0){ ;Log in screen (GUI 1:) and file the messages of script. Needs global s_LogFile and s_TextLog. No Return value.
	try{  ;Asegurate global s_LogFile access
		While (A_Index <= _i_TabDeep) ; Index with tab and | the message
			_s_Log:= % "`t|" . _s_Log
		s_TextLog:= % s_TextLog . _s_Log
	
		;With user validation screen output
			;MsgBox, 64, % A_ScriptName . "-" . A_ThisFunc . "-(" . A_LineNumber . ")" ,%s_TextLog%   ;Debug
		;Fast screen output (no scroll)
			;ToolTip, %_s_Log% ;, 5, 300
		;In GUI scroll display
			GuiControl,, Edit1, %s_TextLog% ;Search the control name with AutoHotKey Window Spy
			if !b_Debug {	; Not Debug mode. Too slow and keyboard use problems
				ControlSend, Edit1, ^{End}, %A_ScriptName%   ;Assuming your edit control is Writtable and your window is %A_ScriptName%
			}
		;Log to file
			FileAppend %_s_Log%, %s_LogFile%
	}
	catch e {
        MsgBox 16, % A_ThisFunc . "-(" . A_LineNumber . ")", % "Error catch " e.What ", which was called at line " e.Line ", in:`n|" s_LogFile "|"
        ExitApp            
    }		
	Return
}
;==========================================================================================
;AutoHotKey Snipets------------------------------------------------------------------------ 
Zip(FilesToZip,sZip){ ; Zip/Unzip natively on any Windows > XP. https://autohotkey.com/board/topic/60706-native-zip-and-unzip-xpvista7-ahk-l/
/*
Unz needs a clean folder to comprobate that all files are unzipped.
Options for zipping, unzipping:
 4 Do not display a progress dialog box. 
 8 Give the file being operated on a new name in a move, copy, or rename operation if a file with the target name already exists. 
 16 Respond with "Yes to All" for any dialog box that is displayed. 
 64 Preserve undo information, if possible. 
 128 Perform the operation on files only if a wildcard file name (*.*) is specified. 
 256 Display a progress dialog box but do not show the file names. 
 512 Do not confirm the creation of a new directory if the operation requires one to be created. 
 1024 Do not display a user interface if an error occurs. 
 2048 Version 4.71. Do not copy the security attributes of the file. 
 4096 Only operate in the local directory. Don't operate recursively into subdirectories. 
 9182 Version 5.0. Do not move connected files as a group. Only move the specified files. 
 --------- 	EXAMPLE CODE	-------------------------------------
 FilesToZip = D:\Projects\AHK\_Temp\Test\  ;Example of folder to compress
 FilesToZip = D:\Projects\AHK\_Temp\Test\*.ahk  ;Example of wildcards to compress
 FilesToZip := A_ScriptFullPath   ;Example of file to compress
 sZip := A_ScriptDir . "\Test.zip"  ;Zip file to be created
 sUnz := A_ScriptDir . "\ext\"      ;Directory to unzip files

 Zip(FilesToZip,sZip)
 Sleep, 500
 Unz(sZip,sUnz)
 --------- 	END EXAMPLE 	-------------------------------------
*/
If Not FileExist(sZip)
	CreateZipFile(sZip)
psh := ComObjCreate( "Shell.Application" )
pzip := psh.Namespace( sZip )
if InStr(FileExist(FilesToZip), "D")
	FilesToZip .= SubStr(FilesToZip,0)="\" ? "*.*" : "\*.*"
loop,%FilesToZip%,1
{
	zipped++
	ToolTip Zipping %A_LoopFileName% ..
	pzip.CopyHere( A_LoopFileLongPath, 4|16 )
	Loop
	{
		done := pzip.items().count
		if done = %zipped%
			break
	}
	done := -1
}
ToolTip
}
CreateZipFile(sZip){
	Header1 := "PK" . Chr(5) . Chr(6)
	VarSetCapacity(Header2, 18, 0)
	file := FileOpen(sZip,"w")
	file.Write(Header1)
	file.RawWrite(Header2,18)
	file.close()
}
Unz(sZip, sUnz){
    fso := ComObjCreate("Scripting.FileSystemObject")
    If Not fso.FolderExists(sUnz)  ;http://www.autohotkey.com/forum/viewtopic.php?p=402574
       fso.CreateFolder(sUnz)
    psh  := ComObjCreate("Shell.Application")
    zippedItems := psh.Namespace( sZip ).items().count
    psh.Namespace( sUnz ).CopyHere( psh.Namespace( sZip ).items, 4|16 )
    Loop {
        sleep 50
        unzippedItems := psh.Namespace( sUnz ).items().count
        ToolTip Unzipping in progress..
        IfEqual,zippedItems,%unzippedItems%
            break
    }
    ToolTip
}
;==========================================================================================
;General Documentation:
;--------------------------------------------------------------------------------
;--------------------------------------------------------------------------------
;==========================================================================================
;Program Documentation:
;==========================================================================================

