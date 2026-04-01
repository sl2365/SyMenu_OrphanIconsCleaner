;DESCRIPTION:
;	Script for AutoHotKey for Move/Delete the SyMenu Orphan Icons
;CHANGELOG:
; Requires v2.0 AutoHotkey to compile.
;		V.1.0.-2017.01.27: First version
;		V.2.0.-2026.04.01: Converted to AHK v2
;		V.2.1.-2026.04.01: Added exclusion list support via Exclusions.txt
;		V.2.2.-2026.04.01: Inline edit/save for exclusions, path fixes
;		V.2.3.-2026.04.01: Settings persistence via Settings.ini
;		V.2.4.-2026.04.01: Side-by-side layout (exclusions left, log right)
;		V.2.5.-2026.04.01: Resizable window, draggable splitter, saved layout
;		V.2.6.-2026.04.01: Fixed splitter drag via WM messages, anchored buttons
;		V.2.7.-2026.04.01: Config read from SyMenuItem.xml, exclusion list support
;
;------------------------------------------------------------
	#Requires AutoHotkey v2.0
	#Warn
	#SingleInstance Off
	SendMode("Input")
	SplitPath(A_ScriptName, , , , &ScriptName)

;Program behavior and Global variables
	global s_LogFile       := A_ScriptDir . "\" . ScriptName . ".log"
	global s_ExcludeFile   := A_ScriptDir . "\Exclusions.txt"
	global s_SettingsFile  := A_ScriptDir . "\Settings.ini"
	global b_DeleteAction  := False
	global b_NewLogFile    := True
	global b_VerboseLog    := False
	global s_TextLog       := ""
	global s_SyMenuPath    := ""
	global b_Debug         := False
	global a_ExcludeList   := Map()

;Layout defaults
	global i_WinW          := 1010
	global i_WinH          := 460
	global i_SplitterX     := 495
	global i_TopAreaH      := 60
	global i_Margin        := 10
	global i_SplitterW     := 6
	global b_Dragging      := false
	global i_LastWinW      := 0
	global i_LastWinH      := 0

;Load saved settings and the exclusion list
	LoadSettings()
	LoadExcludeList()

;Make the GUI
	global MainGui := Gui("+SysMenu +Resize +MinSize640x300", ScriptName)
	MainGui.OnEvent("Close", GuiCloseHandler)
	MainGui.OnEvent("Escape", GuiCloseHandler)
	MainGui.OnEvent("Size", GuiSizeHandler)

	; --- Top row: folder selector ---
	global btnFolder := MainGui.Add("Button", "x10 y8 w70 Default vbtnFolder", "Browse...")
	btnFolder.OnEvent("Click", GetFolder)
	MainGui.SetFont("Bold")
	MainGui.BackColor := "White"
	global txtPath := MainGui.Add("Text", "x90 y12 w870 cBlue Border vtxtPath", "  " . s_SyMenuPath)
	MainGui.SetFont("Norm")
	MainGui.BackColor := "Default"

	; --- Second row: delete checkbox ---
	global chkDelete := MainGui.Add("CheckBox", "x10 y35 vchkDelete", "Delete Icons")
	chkDelete.Value := b_DeleteAction
	chkDelete.OnEvent("Click", SaveSettingsEvent)

	; === LEFT COLUMN: Exclusion list ===
	global txtExcLabel := MainGui.Add("Text", "x10 y60 cPurple vtxtExcLabel", "Excluded icons (one 'filename.ico' per line):")
	global edtExclude := MainGui.Add("Edit", "x10 y78 w480 r15 ReadOnly vedtExclude", GetExcludeDisplay())
	global btnGo := MainGui.Add("Button", "x10 w75 vbtnGo", "Go!")
	btnGo.OnEvent("Click", ScriptProcess)
	global btnEdit := MainGui.Add("Button", "x90 yp w75 vbtnEdit", "Edit")
	btnEdit.OnEvent("Click", EditExcludeList)
	global btnSave := MainGui.Add("Button", "x170 yp w75 Disabled vbtnSave", "Save")
	btnSave.OnEvent("Click", SaveExcludeList)

	; === SPLITTER BAR ===
	global ctrlSplitter := MainGui.Add("Text", "x495 y60 w6 h300 BackgroundSilver vctrlSplitter Border")

	; === RIGHT COLUMN: Log window ===
	global txtLogLabel := MainGui.Add("Text", "x510 y60 cBlue vtxtLogLabel", "Log Window:")
	global edtLog := MainGui.Add("Edit", "x510 y78 w490 r15 vedtLog", "")

	; --- Bottom row: options ---
	global chkNewLog := MainGui.Add("CheckBox", "x510 vchkNewLog Checked", "Delete Log")
	chkNewLog.Value := b_NewLogFile
	chkNewLog.OnEvent("Click", SaveSettingsEvent)
	global chkVerbose := MainGui.Add("CheckBox", "x510 vchkVerbose", "Log All Events")
	chkVerbose.Value := b_VerboseLog
	chkVerbose.OnEvent("Click", SaveSettingsEvent)

	; Register mouse messages for splitter dragging
	OnMessage(0x0201, WM_LBUTTONDOWN)
	OnMessage(0x0200, WM_MOUSEMOVE)
	OnMessage(0x0202, WM_LBUTTONUP)

	MainGui.Show("w" . i_WinW . " h" . i_WinH)

	; Create tooltip control
	global hToolTip := DllCall("CreateWindowEx", "UInt", 0x8, "Str", "tooltips_class32", "Str", ""
		, "UInt", 0x80000002, "Int", 0x80000000, "Int", 0x80000000, "Int", 0x80000000, "Int", 0x80000000
		, "Ptr", MainGui.Hwnd, "Ptr", 0, "Ptr", 0, "Ptr", 0, "Ptr")
	DllCall("SendMessage", "Ptr", hToolTip, "UInt", 0x0418, "Ptr", 0, "Ptr", 32767)  ; TTM_SETMAXTIPWIDTH for multiline

	AddToolTip(btnFolder, "Select your SyMenu.exe folder")
	AddToolTip(chkDelete, "If checked, orphan icons are permanently deleted.`nUnchecked items are moved to _Trash\_OrphanIcons")
	AddToolTip(btnGo, "Start scanning for orphan icons")
	AddToolTip(btnEdit, "Edit exclusion list")
	AddToolTip(btnSave, "Save exclusion list")
; 	AddToolTip(edtExclude, "One icon filename per line.`nComment/blank lines are ignored")
	AddToolTip(chkNewLog, "Start a fresh log file each run")
	AddToolTip(chkVerbose, "Log every icon checked, not just orphans")

	ApplyLayout(i_WinW, i_WinH)
	Return

;==========================================================================================
; LAYOUT / RESIZE / SPLITTER
;==========================================================================================
ApplyLayout(winW, winH) {
	global i_SplitterX, i_Margin, i_SplitterW
	global txtPath, edtExclude, ctrlSplitter, edtLog
	global txtExcLabel, txtLogLabel, btnEdit, btnSave, btnGo
	global chkNewLog, chkVerbose

	global i_LastWinW := winW
	global i_LastWinH := winH

	; Clamp splitter position
	minLeft := 260
	minRight := 200
	maxSplitX := winW - i_Margin - minRight - i_SplitterW
	if (i_SplitterX < minLeft)
		i_SplitterX := minLeft
	if (i_SplitterX > maxSplitX)
		i_SplitterX := maxSplitX

	editTop := 78
	btnRowY := winH - 45
	editH := btnRowY - editTop - 8
	if (editH < 50)
		editH := 50

	leftW := i_SplitterX - i_Margin
	rightX := i_SplitterX + i_SplitterW + 4
	rightW := winW - rightX - i_Margin

	txtPath.Move(90, 12, winW - 100)

	; Left column
	txtExcLabel.Move(i_Margin, 60)
	edtExclude.Move(i_Margin, editTop, leftW, editH)
	btnGo.Move(i_Margin, btnRowY)
	btnEdit.Move(i_Margin + 80, btnRowY)
	btnSave.Move(i_Margin + 160, btnRowY)

	; Splitter bar
	ctrlSplitter.Move(i_SplitterX, 60, i_SplitterW, editH + 18)

	; Right column
	txtLogLabel.Move(rightX, 60)
	edtLog.Move(rightX, editTop, rightW, editH)
	chkNewLog.Move(rightX, btnRowY)
	chkVerbose.Move(rightX, btnRowY + 20)
}

GuiSizeHandler(thisGui, minMax, w, h) {
	if (minMax = -1)
		return
	ApplyLayout(w, h)
}

WM_LBUTTONDOWN(wParam, lParam, msg, hwnd) {
	global b_Dragging, i_SplitterX, i_SplitterW, MainGui, ctrlSplitter
	mx := lParam & 0xFFFF
	my := (lParam >> 16) & 0xFFFF
	if (mx >= i_SplitterX - 6 && mx <= i_SplitterX + i_SplitterW + 6 && my >= 60) {
		if (hwnd = MainGui.Hwnd || hwnd = ctrlSplitter.Hwnd) {
			b_Dragging := true
			DllCall("SetCapture", "Ptr", MainGui.Hwnd)
		}
	}
}

WM_MOUSEMOVE(wParam, lParam, msg, hwnd) {
	global b_Dragging, i_SplitterX, i_LastWinW, i_LastWinH, MainGui, ctrlSplitter, i_SplitterW
	mx := lParam & 0xFFFF
	my := (lParam >> 16) & 0xFFFF
	if (mx > 0x7FFF)
		mx := mx - 0x10000
	if (my > 0x7FFF)
		my := my - 0x10000
	if (b_Dragging) {
		i_SplitterX := mx
		ApplyLayout(i_LastWinW, i_LastWinH)
		hCursor := DllCall("LoadCursor", "Ptr", 0, "Ptr", 32644, "Ptr")
		DllCall("SetCursor", "Ptr", hCursor)
		return 0
	}
	if (mx >= i_SplitterX - 6 && mx <= i_SplitterX + i_SplitterW + 6 && my >= 60) {
		hCursor := DllCall("LoadCursor", "Ptr", 0, "Ptr", 32644, "Ptr")
		DllCall("SetCursor", "Ptr", hCursor)
		return 0
	}
}

WM_LBUTTONUP(wParam, lParam, msg, hwnd) {
	global b_Dragging
	if (b_Dragging) {
		b_Dragging := false
		DllCall("ReleaseCapture")
	}
}

;==========================================================================================
; SETTINGS PERSISTENCE
;==========================================================================================
LoadSettings() {
	global s_SettingsFile, s_SyMenuPath, b_DeleteAction, b_NewLogFile, b_VerboseLog
	global i_WinW, i_WinH, i_SplitterX

	if !FileExist(s_SettingsFile)
		return

	s_SyMenuPath    := IniRead(s_SettingsFile, "General", "SyMenuPath", "")
	b_DeleteAction  := IniRead(s_SettingsFile, "General", "DeleteAction", "0") = "1"
	b_NewLogFile    := IniRead(s_SettingsFile, "General", "NewLogFile", "1") = "1"
	b_VerboseLog    := IniRead(s_SettingsFile, "General", "VerboseLog", "0") = "1"

	savedW := IniRead(s_SettingsFile, "Window", "Width", "0")
	savedH := IniRead(s_SettingsFile, "Window", "Height", "0")
	savedS := IniRead(s_SettingsFile, "Window", "SplitterX", "0")
	if (Integer(savedW) > 0)
		i_WinW := Integer(savedW)
	if (Integer(savedH) > 0)
		i_WinH := Integer(savedH)
	if (Integer(savedS) > 0)
		i_SplitterX := Integer(savedS)
}

SaveSettings() {
	global s_SettingsFile, s_SyMenuPath, b_DeleteAction, b_NewLogFile, b_VerboseLog
	global chkDelete, chkNewLog, chkVerbose

	b_DeleteAction := chkDelete.Value
	b_NewLogFile   := chkNewLog.Value
	b_VerboseLog   := chkVerbose.Value

	IniWrite(s_SyMenuPath, s_SettingsFile, "General", "SyMenuPath")
	IniWrite(b_DeleteAction ? "1" : "0", s_SettingsFile, "General", "DeleteAction")
	IniWrite(b_NewLogFile ? "1" : "0", s_SettingsFile, "General", "NewLogFile")
	IniWrite(b_VerboseLog ? "1" : "0", s_SettingsFile, "General", "VerboseLog")
}

SaveWindowLayout() {
	global s_SettingsFile, i_SplitterX, MainGui
	try {
		MainGui.GetClientPos(, , &w, &h)
		IniWrite(w, s_SettingsFile, "Window", "Width")
		IniWrite(h, s_SettingsFile, "Window", "Height")
		IniWrite(i_SplitterX, s_SettingsFile, "Window", "SplitterX")
	}
}

SaveSettingsEvent(ctrl, info) {
	SaveSettings()
}

;==========================================================================================
; EXCLUSION LIST FUNCTIONS
;==========================================================================================
LoadExcludeList() {
	global a_ExcludeList, s_ExcludeFile
	a_ExcludeList := Map()
	a_ExcludeList.CaseSense := "Off"

	if !FileExist(s_ExcludeFile) {
		defaultContent := "`; SyMenu Orphan Icons - Exclusion List`n"
		defaultContent .= "`; ------------------------------------------------------------`n"
		defaultContent .= "`; Add one icon filename per line to prevent it from being moved or deleted.`n"
		defaultContent .= "`; Lines starting with `; are comments and will be ignored.`n"
		defaultContent .= "`; Blank lines are also ignored.`n"
		defaultContent .= "`; Comments can be added/removed and will be saved.`n"
		defaultContent .= "`; Examples:`n"
		defaultContent .= "`;  MyCustomIcon.ico`n"
		defaultContent .= "`;  AnotherIcon.ico`n"
		FileAppend(defaultContent, s_ExcludeFile)
		return
	}

	fileContent := FileRead(s_ExcludeFile)
	Loop Parse, fileContent, "`n", "`r" {
		line := Trim(A_LoopField)
		if (line = "" || SubStr(line, 1, 1) = "`;")
			continue
		a_ExcludeList[line] := true
	}
}

GetExcludeDisplay() {
	global s_ExcludeFile
	if !FileExist(s_ExcludeFile)
		return "(No Exclusions.txt found. Click 'Edit' to add icon filenames.)"
	return FileRead(s_ExcludeFile)
}

IsExcluded(filename) {
	global a_ExcludeList
	return a_ExcludeList.Has(filename)
}

EditExcludeList(ctrl, info) {
	global edtExclude, btnEdit, btnSave
	edtExclude.Opt("-ReadOnly")
	btnEdit.Enabled := false
	btnSave.Enabled := true
}

SaveExcludeList(ctrl, info) {
	global edtExclude, btnEdit, btnSave, s_ExcludeFile, a_ExcludeList

	newContent := edtExclude.Value
	if FileExist(s_ExcludeFile)
		FileDelete(s_ExcludeFile)
	FileAppend(newContent, s_ExcludeFile)

	a_ExcludeList := Map()
	a_ExcludeList.CaseSense := "Off"
	Loop Parse, newContent, "`n", "`r" {
		line := Trim(A_LoopField)
		if (line = "" || SubStr(line, 1, 1) = "`;")
			continue
		a_ExcludeList[line] := true
	}

	edtExclude.Opt("+ReadOnly")
	btnEdit.Enabled := true
	btnSave.Enabled := false

	MsgBox("Exclusion list saved: " . a_ExcludeList.Count . " icon(s) excluded.", "Saved", 64)
}

;==========================================================================================
; GUI EVENT HANDLERS
;==========================================================================================
GetFolder(ctrl, info) {
	global s_SyMenuPath, txtPath
	s_NewWorkPath := DirSelect("", 2, "May you select the folder of the working SyMenu?")
	if (s_NewWorkPath = "")
		Return
	s_SyMenuPath := RegExReplace(s_NewWorkPath, "\\$")
	txtPath.Value := "  " . s_SyMenuPath
	SaveSettings()
}

GuiCloseHandler(thisGui) {
	global b_VerboseLog, b_DeleteAction, b_NewLogFile, ScriptName
	SaveSettings()
	SaveWindowLayout()
	if b_VerboseLog
		ScriptLog(A_YYYY . "." . A_MM . "." . A_DD . "-" . A_Hour . ":" . A_Min . ":" . A_Sec . "-" . ScriptName . "-(" . b_DeleteAction . "|" . b_NewLogFile . "|" . b_VerboseLog . ") canceled. Good bye, " . A_UserName . "!" . "|`n", 0)
	thisGui.Destroy()
	ExitApp()
}

;==========================================================================================
; MAIN PROCESS
;==========================================================================================
ScriptProcess(ctrl, info) {
	global s_SyMenuPath, b_DeleteAction, b_NewLogFile, b_VerboseLog, b_Debug
	global s_LogFile, s_TextLog, ScriptName, edtLog, chkDelete, chkNewLog, chkVerbose

	b_DeleteAction := chkDelete.Value
	b_NewLogFile   := chkNewLog.Value
	b_VerboseLog   := chkVerbose.Value
	SaveSettings()

	if (s_SyMenuPath = "") {
		MsgBox("Please select a SyMenu folder first using the 'Browse...' button.", ScriptName, 48)
		return
	}

	s_SyMenuPathIcons  := s_SyMenuPath . "\Icons"
	s_SyMenuPathConfig := s_SyMenuPath . "\Config"
	s_SyMenuPath_Trash := s_SyMenuPath . "\ProgramFiles\SPSSuite\SyMenuSuite\_Trash"
	s_OrphanIconsPath  := s_SyMenuPath_Trash . "\_OrphanIcons"

	if !FileExist(s_SyMenuPathIcons) {
		MsgBox("Icons folder not found at:`n" . s_SyMenuPathIcons, "Folder Not Found", 16)
		return
	}
	if !FileExist(s_SyMenuPathConfig . "\SyMenuItem.zip") {
		MsgBox("SyMenuItem.zip not found at:`n" . s_SyMenuPathConfig, "Config Not Found", 16)
		return
	}

	if (b_NewLogFile and FileExist(s_LogFile))
		FileDelete(s_LogFile)

	ScriptLog(A_YYYY . "." . A_MM . "." . A_DD . "-" . A_Hour . ":" . A_Min . ":" . A_Sec . "- " . ScriptName . " STARTED`n", 0)

	; ============================================================
	; PHASE 1: Extract SyMenuItem.xml from Config\SyMenuItem.zip
	; ============================================================
	try {
		if !FileExist(s_OrphanIconsPath)
			DirCreate(s_OrphanIconsPath)
		if FileExist(s_OrphanIconsPath . "\_TmpZip")
			DirDelete(s_OrphanIconsPath . "\_TmpZip", true)
		DirCreate(s_OrphanIconsPath . "\_TmpZip")

		Unz(s_SyMenuPathConfig . "\SyMenuItem.zip", s_OrphanIconsPath . "\_TmpZip\")

		FileMove(s_OrphanIconsPath . "\_TmpZip\SyMenuItem.xml", s_OrphanIconsPath, 1)
		DirDelete(s_OrphanIconsPath . "\_TmpZip", true)
	}
	catch as e {
		MsgBox("Error extracting SyMenuItem.zip:`n" . e.Message, ScriptName, 16)
		return
	}

	; Read the config XML
	s_SyMenuItemConfig := FileRead(s_OrphanIconsPath . "\SyMenuItem.xml")

	ScriptLog(A_YYYY . "." . A_MM . "." . A_DD . "-" . A_Hour . ":" . A_Min . ":" . A_Sec . "- Config loaded from: Config\SyMenuItem.zip`n", 1)

	; ============================================================
	; PHASE 2: Scan icons against config
	; ============================================================
	excludedCount := 0
	orphanCount := 0
	activeCount := 0

	Loop Files, s_SyMenuPathIcons . "\*.ico" {
		if IsExcluded(A_LoopFileName) {
			excludedCount++
			if b_VerboseLog
				ScriptLog(A_YYYY . "." . A_MM . "." . A_DD . "-" . A_Hour . ":" . A_Min . ":" . A_Sec . "- EXCLUDED (skipped): " . A_LoopFileName . "`n", 2)
			continue
		}

		if (0 < InStr(s_SyMenuItemConfig, A_LoopFileName)) {
			activeCount++
			if b_VerboseLog
				ScriptLog(A_YYYY . "." . A_MM . "." . A_DD . "-" . A_Hour . ":" . A_Min . ":" . A_Sec . "- Active Icon: " . A_LoopFileName . "`n", 2)
		} else {
			orphanCount++
			ScriptLog(A_YYYY . "." . A_MM . "." . A_DD . "-" . A_Hour . ":" . A_Min . ":" . A_Sec . "- Orphan Icon moved: " . A_LoopFileName . "`n", 2)
			FileMove(A_LoopFileFullPath, s_OrphanIconsPath, 1)
		}
	}

	; ============================================================
	; PHASE 3: Cleanup
	; ============================================================
	; Delete the extracted SyMenuItem.xml from orphan folder
	if FileExist(s_OrphanIconsPath . "\SyMenuItem.xml")
		FileDelete(s_OrphanIconsPath . "\SyMenuItem.xml")

	if b_DeleteAction {
		if b_VerboseLog
			ScriptLog(A_YYYY . "." . A_MM . "." . A_DD . "-" . A_Hour . ":" . A_Min . ":" . A_Sec . "- Deleting orphan icons folder`n", 1)
		if FileExist(s_OrphanIconsPath)
			DirDelete(s_OrphanIconsPath, true)
	}

	summaryMsg := ScriptName . " finished.`n"
	summaryMsg .= "Active: " . activeCount . " | Orphans: " . orphanCount . " | Excluded: " . excludedCount
	ScriptLog(A_YYYY . "." . A_MM . "." . A_DD . "-" . A_Hour . ":" . A_Min . ":" . A_Sec . "- " . summaryMsg . "`n", 0)
	MsgBox(summaryMsg . "`n" . A_UserName . "!", ScriptName, 64)
}

;==========================================================================================
;SUBROUTINES-------------------------------------------------------------------------------
ScriptLog(_s_Log, _i_TabDeep := 0) {
	global s_LogFile, s_TextLog, b_Debug, edtLog, ScriptName
	try {
		i := 0
		While (i < _i_TabDeep) {
			_s_Log := "`t|" . _s_Log
			i++
		}
		s_TextLog := s_TextLog . _s_Log
		edtLog.Value := s_TextLog
		if !b_Debug {
			ControlSend("^{End}", edtLog)
		}
		FileAppend(_s_Log, s_LogFile)
	}
	catch as e {
		MsgBox("Error catch " . e.Message . ", which was called at line " . e.Line . ", in:`n|" . s_LogFile . "|", A_ThisFunc . "-(" . A_LineNumber . ")", 16)
		ExitApp()
	}
}

AddToolTip(ctrl, text) {
	global hToolTip
	static TOOLINFO_size := 24 + A_PtrSize * 6

	ti := Buffer(TOOLINFO_size, 0)
	NumPut("UInt", TOOLINFO_size, ti, 0)          ; cbSize
	NumPut("UInt", 0x11, ti, 4)                     ; uFlags: TTF_SUBCLASS | TTF_IDISHWND
	NumPut("Ptr", ctrl.Gui.Hwnd, ti, 8)            ; hwnd
	NumPut("Ptr", ctrl.Hwnd, ti, 8 + A_PtrSize)    ; uId
	NumPut("Ptr", StrPtr(text), ti, 24 + A_PtrSize * 3)  ; lpszText

	DllCall("SendMessage", "Ptr", hToolTip, "UInt", 0x0432, "Ptr", 0, "Ptr", ti)  ; TTM_ADDTOOLW
}

;==========================================================================================
;AutoHotKey Snippets-----------------------------------------------------------------------
Zip(FilesToZip, sZip) {
	If Not FileExist(sZip)
		CreateZipFile(sZip)
	psh := ComObject("Shell.Application")
	pzip := psh.Namespace(sZip)
	if !IsObject(pzip) {
		MsgBox("Failed to open zip file:`n" . sZip, "Zip Error", 16)
		return
	}
	if InStr(FileExist(FilesToZip), "D")
		FilesToZip .= SubStr(FilesToZip, -1) = "\" ? "*.*" : "\*.*"
	zipped := 0
	Loop Files, FilesToZip, "FD" {
		zipped++
		ToolTip("Zipping " . A_LoopFileName . " ..")
		pzip.CopyHere(A_LoopFileFullPath, 4 | 16)
		Loop {
			done := pzip.items().count
			if (done = zipped)
				break
		}
	}
	ToolTip()
}

CreateZipFile(sZip) {
	Header1 := "PK" . Chr(5) . Chr(6)
	file := FileOpen(sZip, "w")
	file.Write(Header1)
	Loop 18
		file.Write(Chr(0))
	file.close()
}

Unz(sZip, sUnz) {
	fso := ComObject("Scripting.FileSystemObject")
	If Not fso.FolderExists(sUnz)
		fso.CreateFolder(sUnz)
	psh := ComObject("Shell.Application")
	nsZip := psh.Namespace(sZip)
	if !IsObject(nsZip) {
		MsgBox("Failed to open zip file:`n" . sZip . "`n`nEnsure the path is a full absolute path and the file exists.", "Unz Error", 16)
		return
	}
	nsUnz := psh.Namespace(sUnz)
	if !IsObject(nsUnz) {
		MsgBox("Failed to open output folder:`n" . sUnz . "`n`nEnsure the path is a full absolute path and the folder exists.", "Unz Error", 16)
		return
	}
	zippedItems := nsZip.items().count
	nsUnz.CopyHere(nsZip.items, 4 | 16)
	Loop {
		Sleep(50)
		unzippedItems := nsUnz.items().count
		ToolTip("Unzipping in progress..")
		if (zippedItems = unzippedItems)
			break
	}
	ToolTip()
}