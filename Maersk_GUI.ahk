; ============================================
; Maersk GUI – Macro Control HQ (All-in-One)
; v3.0  — AutoHotkey v2 ONLY
; ============================================

#Requires AutoHotkey v2
#SingleInstance Force
#Warn
Persistent
SetTitleMatchMode 2
CoordMode("Mouse", "Screen")   ; absolute coords for multi-monitor

; ============================================
; PERSISTENCE PATHS (dual-write; AppData first)
; ============================================
global APPDATA_DIR := A_AppData "\Maersk_GUI"
global USER_INI    := APPDATA_DIR "\GUI_Settings.ini"
global LOCAL_INI   := A_ScriptDir "\GUI_Settings.ini"
if !DirExist(APPDATA_DIR)
    DirCreate(APPDATA_DIR)

; ============================================
; GLOBALS / SETTINGS
; ============================================
; user settings (persisted)
global settings := Map(
    "SnapToWindow", false,
    "ScreenshotFolder", A_ScriptDir "\Pictures"
)

; analytics (persisted)
global analytics := Map(
    "LastFile","",
    "RunCount",0,
    "TotalMs",0,
    "AvgMs",0,
    "LastSuccess","",
    "Errors",0
)

; hotkeys (persisted) — Option 2: require a modifier
global HOTKEYS := Map(
    "RunMacro","Ctrl+F9",
    "KillAll","Ctrl+F10",
    "StartRec","Ctrl+F11",
    "StopRec","Ctrl+F12",
    "Screenshot","Ctrl+Shift+S"
)

; runtime state
global currentMacroFile := ""
global currentMacroFolder := A_ScriptDir
global recordedActions := []         ; textual macro lines (runner consumes)
global recording := false
global recordStartTime := 0
global stopAll := false
global totalSteps := 0
global doneSteps := 0
global coordArmed := false

; recorder (rich) state for listview inspector
global gMacro := []                  ; [{type, key/btn, x, y, delay}]
global _lastMouse := Map("x",0,"y",0,"t",A_TickCount)

; gui refs
global main, tabs
global layout := Map()
global tabNames := ["Manager","Recorder","Tools","Editor","Hotkeys","Analytics"]

; controls (declared for clarity)
global folderPathTxt, macroList, currentMacroTxt, speedBox, speedSpin, repeatBox, repeatSpin
global outputBox
global snapChk, coordLiveBox, snapFolderTxt
global editorBox, editorStatusTxt, editorPath := ""
global hkEdits := Map(), hkButtons := Map()
global statusText, progBar, progPct
global listMenu, ctxSelectedRow := 0

; Recorder/Player widgets (inspector LV)
global MacroLV

; ============================================
; UTIL HELPERS
; ============================================
EnsureDir(path) {
    dir := RegExReplace(path, "(?<!\\)\\[^\\]+$")    ; parent dir
    if dir != "" && !DirExist(dir)
        DirCreate(dir)
}
WriteAll(path, data) {
    EnsureDir(path)
    try FileDelete(path)
    f := FileOpen(path, "w")
    f.RawWrite(data)
    f.Close()
}
WriteText(path, text) {
    EnsureDir(path)
    try FileDelete(path)
    FileAppend(text, path, "UTF-8")
}
ReadIniSectionAll(iniPath, section, mapObj) {
    for k, v in mapObj
        mapObj[k] := IniRead(iniPath, section, k, v)
}
WriteIniSectionAll(iniPath, section, mapObj) {
    EnsureDir(iniPath)
    for k, v in mapObj
        IniWrite(v, iniPath, section, k)
}
; Dual save (AppData + Local)
SaveAllIni() {
    global settings, HOTKEYS, analytics, USER_INI, LOCAL_INI
    try WriteIniSectionAll(USER_INI,  "settings", settings)
    try WriteIniSectionAll(USER_INI,  "hotkeys",  HOTKEYS)
    try WriteIniSectionAll(USER_INI,  "analytics",analytics)
    ; also write alongside script for portability/backups
    try WriteIniSectionAll(LOCAL_INI, "settings", settings)
    try WriteIniSectionAll(LOCAL_INI, "hotkeys",  HOTKEYS)
    try WriteIniSectionAll(LOCAL_INI, "analytics",analytics)
}
; Load precedence: AppData first, fallback to Local, else defaults already in maps
LoadAllIni() {
    global settings, HOTKEYS, analytics, USER_INI, LOCAL_INI
    ini := FileExist(USER_INI) ? USER_INI : (FileExist(LOCAL_INI) ? LOCAL_INI : "")
    if ini != "" {
        ReadIniSectionAll(ini, "settings", settings)
        ReadIniSectionAll(ini, "hotkeys",  HOTKEYS)
        ReadIniSectionAll(ini, "analytics",analytics)
        ; normalize ints
        analytics["RunCount"] := Integer(analytics["RunCount"])
        analytics["TotalMs"]  := Integer(analytics["TotalMs"])
        analytics["AvgMs"]    := Integer(analytics["AvgMs"])
        analytics["Errors"]   := Integer(analytics["Errors"])
    }
    if !DirExist(settings["ScreenshotFolder"])
        DirCreate(settings["ScreenshotFolder"])
}

; ============================================
; INIT / EXIT
; ============================================
OnExit(SaveSettings)
InitSettings()
InitAnalytics()
InitHotkeys()   ; bind defaults

InitSettings() {
    LoadAllIni()
}
InitAnalytics() {
    ; already loaded via LoadAllIni()
}
SaveSettings(*) {
    SaveAllIni()
}

; ============================================
; LAYOUT ENGINE (responsive grid)
; ============================================
LayoutInit(tab, gapX := 10, gapY := 10, margin := 20) {
    global layout
    layout[tab] := { rows: [], gapX: gapX, gapY: gapY, margin: margin }
}
LayoutRow(tab) {
    global layout
    layout[tab].rows.Push([])
}
LayoutAdd(tab, ctrl, wantW := 0, wantH := 0) {
    global layout
    if layout[tab].rows.Length = 0
        LayoutRow(tab)
    layout[tab].rows[-1].Push([ctrl, wantW, wantH])
}
LayoutApply(tab, availW) {
    global layout
    cfg := layout[tab]
    if cfg.rows.Length = 0
        return
    x := cfg.margin
    y := cfg.margin + 10  ; 10px below tab header
    maxRowH := 0
    for _, row in cfg.rows {
        x := cfg.margin
        maxRowH := 0
        for _, item in row {
            ctrl := item[1], wantW := item[2], wantH := item[3]
            ctrl.GetPos(&cx,&cy,&cw,&ch)
            w := wantW>0 ? wantW : cw
            h := wantH>0 ? wantH : ch
            if (x + w + cfg.margin > availW) {
                x := cfg.margin
                y += maxRowH + cfg.gapY
                maxRowH := 0
            }
            ctrl.Move(x, y, w, h)
            x += w + cfg.gapX
            if (h > maxRowH)
                maxRowH := h
        }
        y += maxRowH + cfg.gapY
    }
}

; ============================================
; GUI + TABS (fully dynamic)
; ============================================
main := Gui("+Resize", "Macro Control HQ v3.0 - Maersk GUI")
tabs := main.Add("Tab3", "x10 y10 w-20 h-60", tabNames)

; --------- MANAGER TAB ---------
tabs.UseTab("Manager"), LayoutInit("Manager")

folderLbl := main.Add("Text",, "Folder:")
folderPathTxt := main.Add("Edit", "w520 vFolderPath", currentMacroFolder)
browseFolderBtn := main.Add("Button",, "Browse Folder"), browseFolderBtn.OnEvent("Click", BrowseFolder)
refreshBtn := main.Add("Button",, "↻ Refresh"), refreshBtn.OnEvent("Click", RefreshFolderList)
LayoutRow("Manager"), LayoutAdd("Manager", folderLbl), LayoutAdd("Manager", folderPathTxt, 520), LayoutAdd("Manager", browseFolderBtn), LayoutAdd("Manager", refreshBtn)

macroListLbl := main.Add("Text",, "Macros in Folder:")
LayoutRow("Manager"), LayoutAdd("Manager", macroListLbl)

macroList := main.Add("ListView", "w820 h220 vMacroList", ["File Name"])
macroList.OnEvent("DoubleClick", LoadSelectedMacro)
macroList.OnEvent("ContextMenu", MacroListContext)
LayoutRow("Manager"), LayoutAdd("Manager", macroList, 820, 220)

oneOffLbl := main.Add("Text",, "Or select a one-off macro file:")
browseFileBtn := main.Add("Button",, "Browse File"), browseFileBtn.OnEvent("Click", BrowseFile)
LayoutRow("Manager"), LayoutAdd("Manager", oneOffLbl), LayoutAdd("Manager", browseFileBtn)

curLbl := main.Add("Text",, "Current Macro:")
currentMacroTxt := main.Add("Text", "w600 vCurrentMacro +0x200", "None loaded")
LayoutRow("Manager"), LayoutAdd("Manager", curLbl), LayoutAdd("Manager", currentMacroTxt, 600)

speedLbl := main.Add("Text",, "Speed Multiplier (0.1 - 10):")
speedBox := main.Add("Edit", "w80 vSpeed Number", "1.0")
speedSpin := main.Add("UpDown", "Range1-100 vSpeedSpin", 10)  ; 10 → 1.0, 5 → 0.5
LayoutRow("Manager"), LayoutAdd("Manager", speedLbl), LayoutAdd("Manager", speedBox, 80), LayoutAdd("Manager", speedSpin)

repeatLbl := main.Add("Text",, "Repeat Count (1 - 100):")
repeatBox := main.Add("Edit", "w80 vRepeat Number", "1")
repeatSpin := main.Add("UpDown", "Range1-100 vRepeatSpin", 1)
LayoutRow("Manager"), LayoutAdd("Manager", repeatLbl), LayoutAdd("Manager", repeatBox, 80), LayoutAdd("Manager", repeatSpin)

runBtn := main.Add("Button",, "▶️ Run"), runBtn.OnEvent("Click", RunMacro)
killBtn := main.Add("Button",, "⛔ Kill All"), killBtn.OnEvent("Click", KillAllMacros)
LayoutRow("Manager"), LayoutAdd("Manager", runBtn), LayoutAdd("Manager", killBtn)

tipLbl := main.Add("Text", "w820", "💡 Tip: End stops playback. ESC or Ctrl+Alt+K stops everything immediately.")
LayoutRow("Manager"), LayoutAdd("Manager", tipLbl, 820)

; --------- RECORDER TAB ---------
tabs.UseTab("Recorder"), LayoutInit("Recorder")

recBtn := main.Add("Button",, "▶️ Record"), recBtn.OnEvent("Click", StartRecording)
stopBtn := main.Add("Button",, "⏹ Stop"),   stopBtn.OnEvent("Click", StopRecording)
clearBtn := main.Add("Button",, "🧹 Clear"), clearBtn.OnEvent("Click", ClearRecording)
saveBtn := main.Add("Button",, "💾 Save"),   saveBtn.OnEvent("Click", SaveRecording)
LayoutRow("Recorder"), LayoutAdd("Recorder", recBtn), LayoutAdd("Recorder", stopBtn), LayoutAdd("Recorder", clearBtn), LayoutAdd("Recorder", saveBtn)

outputBox := main.Add("Edit", "w820 h150 vOutputBox +Wrap", "")
LayoutRow("Recorder"), LayoutAdd("Recorder", outputBox, 820, 150)

main.Add("Text",, "Recorded Steps (Inspector):")
MacroLV := main.Add("ListView", "w820 h180", ["#", "Type", "Key/Button", "X", "Y", "Delay (ms)"])
MacroLV.ModifyCol(1, 40), MacroLV.ModifyCol(2, 80), MacroLV.ModifyCol(3, 220), MacroLV.ModifyCol(4, 70), MacroLV.ModifyCol(5, 70), MacroLV.ModifyCol(6, 90)
LayoutRow("Recorder"), LayoutAdd("Recorder", MacroLV, 820, 180)

; --------- TOOLS TAB ---------
tabs.UseTab("Tools"), LayoutInit("Tools")

grabBtn := main.Add("Button",, "🎯 Grab Coords"), grabBtn.OnEvent("Click", ArmCoordGrab)
snapChk := main.Add("CheckBox",, "Snap to Window"), snapChk.Value := settings["SnapToWindow"] ? 1 : 0, snapChk.OnEvent("Click", ToggleSnap)
coordLbl := main.Add("Text",, "Coords:")
coordLiveBox := main.Add("Edit", "w160 +ReadOnly", "(idle)")
LayoutRow("Tools"), LayoutAdd("Tools", grabBtn), LayoutAdd("Tools", snapChk), LayoutAdd("Tools", coordLbl), LayoutAdd("Tools", coordLiveBox, 160)

snapBtn := main.Add("Button",, "📸 Screenshot (snip)"), snapBtn.OnEvent("Click", TakeScreenshot)  ; ms-screenclip: user selects area
snapFolderLbl := main.Add("Text",, "Save to:")
snapFolderTxt := main.Add("Edit", "w520", settings["ScreenshotFolder"])
snapFolderBtn := main.Add("Button",, "Change Save Folder"), snapFolderBtn.OnEvent("Click", ChangeSnapFolder)
LayoutRow("Tools"), LayoutAdd("Tools", snapBtn), LayoutAdd("Tools", snapFolderLbl), LayoutAdd("Tools", snapFolderTxt, 520), LayoutAdd("Tools", snapFolderBtn)

; --------- EDITOR TAB ---------
tabs.UseTab("Editor"), LayoutInit("Editor")

edOpen   := main.Add("Button",, "📂 Open"),   edOpen.OnEvent("Click", OpenMacroForEdit)
edSave   := main.Add("Button",, "💾 Save"),   edSave.OnEvent("Click", SaveMacroEdit)
edSaveAs := main.Add("Button",, "💾 Save As"),edSaveAs.OnEvent("Click", SaveMacroEditAs)
edReload := main.Add("Button",, "🔄 Reload"), edReload.OnEvent("Click", ReloadMacroEdit)
edClear  := main.Add("Button",, "🧹 Clear"),  edClear.OnEvent("Click", (*) => (editorBox.Value := ""))
LayoutRow("Editor"), LayoutAdd("Editor", edOpen), LayoutAdd("Editor", edSave), LayoutAdd("Editor", edSaveAs), LayoutAdd("Editor", edReload), LayoutAdd("Editor", edClear)

editorBox := main.Add("Edit", "w820 h360 +Wrap", "")
LayoutRow("Editor"), LayoutAdd("Editor", editorBox, 820, 360)
editorStatusTxt := main.Add("Text", "w820", "Editing: (none)")
LayoutRow("Editor"), LayoutAdd("Editor", editorStatusTxt, 820)

; --------- HOTKEYS TAB ---------
tabs.UseTab("Hotkeys"), LayoutInit("Hotkeys")

AddHotkeyRow(label, action) {
    global main, hkEdits, hkButtons, HOTKEYS
    lbl := main.Add("Text",, label)
    btn := main.Add("Button",, "Set")
    btn.OnEvent("Click", (*) => HotkeyCaptureDialog(action))
    edt := main.Add("Edit", "w160 +ReadOnly", HOTKEYS[action])
    hkButtons[action] := btn
    hkEdits[action] := edt
    LayoutRow("Hotkeys"), LayoutAdd("Hotkeys", lbl), LayoutAdd("Hotkeys", btn), LayoutAdd("Hotkeys", edt, 160)
}
AddHotkeyRow("Run Macro:", "RunMacro")
AddHotkeyRow("Kill All:", "KillAll")
AddHotkeyRow("Start Recording:", "StartRec")
AddHotkeyRow("Stop Recording:", "StopRec")
AddHotkeyRow("Screenshot:", "Screenshot")

; --------- ANALYTICS TAB ---------
tabs.UseTab("Analytics"), LayoutInit("Analytics")

aLast  := main.Add("Text", "w820", "")
aCount := main.Add("Text", "w820", "")
aAvg   := main.Add("Text", "w820", "")
aTotal := main.Add("Text", "w820", "")
aErrs  := main.Add("Text", "w820", "")
aWhen  := main.Add("Text", "w820", "")
btnCSV := main.Add("Button",, "Export CSV"), btnCSV.OnEvent("Click", ExportAnalyticsCSV)
btnTXT := main.Add("Button",, "Export TXT"), btnTXT.OnEvent("Click", ExportAnalyticsTXT)
LayoutRow("Analytics"), LayoutAdd("Analytics", aLast, 820)
LayoutRow("Analytics"), LayoutAdd("Analytics", aCount, 820)
LayoutRow("Analytics"), LayoutAdd("Analytics", aAvg, 820)
LayoutRow("Analytics"), LayoutAdd("Analytics", aTotal, 820)
LayoutRow("Analytics"), LayoutAdd("Analytics", aErrs, 820)
LayoutRow("Analytics"), LayoutAdd("Analytics", aWhen, 820)
LayoutRow("Analytics"), LayoutAdd("Analytics", btnCSV), LayoutAdd("Analytics", btnTXT)

; --------- STATUS / PROGRESS ---------
tabs.UseTab()
statusText := main.Add("Text", "Section w640", "Status: Ready")
progBar := main.Add("Progress", "x+10 w160 h18 vPlayProgress Range0-100", 0)
progPct := main.Add("Text", "x+6 w40", "0%")

; context menu for macro list
listMenu := Menu()
listMenu.Add("Run",    CtxRunMacro)
listMenu.Add("Edit",   CtxEditMacro)
listMenu.Add("Rename", CtxRenameMacro)
listMenu.Add("Delete", CtxDeleteMacro)

; ============================================
; HOTKEYS (binding + capture)
; ============================================
InitHotkeys() {
    global HOTKEYS
    for action, combo in HOTKEYS {
        if (combo != "")
            BindHotkey(action, combo)
    }
}

BindHotkey(action, combo) {
    static lastBound := Map() ; action -> combo
    if (lastBound.Has(action) && lastBound[action] != "")
        try Hotkey(lastBound[action], , "Off")

    ; Require a modifier for dynamic binding (Option 2)
    if !RegExMatch(combo, "i)^(Ctrl|Alt|Shift|Win)\+[\w\+]+$")
    {
        MsgBox "Invalid hotkey: " combo "`nUse a modifier like Ctrl/Alt/Shift/Win."
        return
    }

    fn := ActionToFn(action)
    if (fn = "")
        return

    try {
        Hotkey(combo, Func(fn), "On")
        lastBound[action] := combo
    }
    catch Error as e {
        MsgBox "Failed to bind hotkey: " combo " for " action "`n`nError: " e.Message
    }
}

ActionToFn(action) {
    switch action {
        case "RunMacro":   return "RunMacro"
        case "KillAll":    return "KillAllMacros"
        case "StartRec":   return "StartRecording"
        case "StopRec":    return "StopRecording"
        case "Screenshot": return "TakeScreenshot"
        default:           return ""
    }
}

HotkeyCaptureDialog(actionName) {
    global HOTKEYS, hkEdits
    g := Gui(, "Set Hotkey – " actionName)
    g.Add("Text",, "Press any keys… (Esc to cancel)")
    preview := g.Add("Edit", "w220 +ReadOnly", "")
    g.Show("w260 h120")

    ; InputHook capture: detect modifiers + final key
    ih := InputHook("V", "{All}")
    ih.KeyOpt("{All}", "E")
    ih.Start(), ih.Wait()
    combo := ""
    mods := []
    if GetKeyState("Ctrl")  mods.Push("Ctrl")
    if GetKeyState("Alt")   mods.Push("Alt")
    if GetKeyState("Shift") mods.Push("Shift")
    if GetKeyState("Win")   mods.Push("Win")
    k := ih.EndKey
    if k
        mods.Push(k)
    if mods.Length
        combo := StrJoin("+", mods*)
    g.Destroy()

    if (combo = "")
        return

    if !RegExMatch(combo, "i)^(Ctrl|Alt|Shift|Win)\+[\w\+]+$") {
        MsgBox "Invalid hotkey: " combo "`nUse a modifier like Ctrl/Alt/Shift/Win."
        return
    }

    HOTKEYS[actionName] := combo
    hkEdits[actionName].Value := combo
    SaveAllIni()
    BindHotkey(actionName, combo)
}

; universal kill
Esc::KillAllMacros()
^!k::KillAllMacros()

; ============================================
; SHOW + APPLY LAYOUT
; ============================================
main.OnEvent("Size", OnResize)
main.Show("w1000 h760")
RefreshFolderList()
ApplyAllLayouts()
SetStatus("Ready")
UpdateAnalyticsDisplay()

OnResize(*) => ApplyAllLayouts()
ApplyAllLayouts() {
    global main, tabs, layout, statusText, progBar, progPct, tabNames
    main.GetPos(&x,&y,&w,&h)
    tabs.Move(10, 10, w-20, h-60)
    availW := w - 40
    for _, tab in tabNames
        LayoutApply(tab, availW)
    statusText.Move(10, h-38, w-220, 20)
    progBar.Move(w-200, h-40, 150, 18)
    progPct.Move(w-46, h-40, 36, 18)
}

; ============================================
; RECORDER (multi-monitor safe) + LISTVIEW inspector
; ============================================
OnMessage(0x201, Mouse_Click) ; LButtonDown
OnMessage(0x204, Mouse_Click) ; RButtonDown
OnMessage(0x100, Key_Press)   ; KeyDown

StartRecording(*) {
    global recording, recordedActions, recordStartTime, outputBox, gMacro, _lastMouse
    recordedActions := []
    gMacro := []
    _lastMouse["x"] := 0, _lastMouse["y"] := 0, _lastMouse["t"] := A_TickCount
    recordStartTime := A_TickCount
    recording := true
    outputBox.Value := "Recording..."
    outputBox.Opt("+ReadOnly")
    SetStatus("Recording")
}
StopRecording(*) {
    global recording, recordedActions, outputBox
    recording := false
    script := ""
    for _, line in recordedActions
        script .= line "`n"
    outputBox.Opt("-ReadOnly")
    outputBox.Value := script
    SetStatus("Ready")
    RefreshLV()
}
ClearRecording(*) {
    global recordedActions, outputBox, gMacro
    recordedActions := []
    gMacro := []
    outputBox.Value := ""
    RefreshLV()
}
SaveRecording(*) {
    global outputBox, currentMacroFolder
    script := outputBox.Value
    if script = "" {
        MsgBox "No recording to save!"
        return
    }
    if !DirExist(currentMacroFolder)
        DirCreate(currentMacroFolder)
    timestamp := FormatTime(A_Now, "yyyyMMdd_HHmmss")
    filePath := currentMacroFolder "\Recorded_" timestamp ".ahk"
    WriteText(filePath, script)
    MsgBox "Saved recording as:`n" filePath
    RefreshFolderList()
}

Mouse_Click(wParam, lParam, msg, hwnd) {
    global recording, recordedActions, recordStartTime, coordArmed, coordLiveBox, settings, gMacro, _lastMouse
    if recording {
        MouseGetPos &x, &y   ; absolute, multi-monitor
        delay := A_TickCount - recordStartTime
        recordStartTime := A_TickCount
        recordedActions.Push("Sleep " delay)
        btn := (msg=0x201) ? "Left" : "Right"
        recordedActions.Push("Click " x ", " y)
        ; inspector
        gMacro.Push(Map("type","CLICK","btn",btn,"x",x,"y",y,"delay",delay))
        _lastMouse["t"] := A_TickCount
        RefreshLVRow(gMacro.Length)
    }
    ; single-shot coordinate capture (Tools tab)
    if coordArmed {
        MouseGetPos &cx, &cy
        coordLiveBox.Value := "(" cx ", " cy ")"
        coordArmed := false
        if settings["SnapToWindow"] {
            ToolTip("(" cx ", " cy ")")
            SetTimer(() => ToolTip(), -1500) ; auto-hide
        } else {
            ToolTip()
        }
    }
}
Key_Press(wParam, lParam, msg, hwnd) {
    global recording, recordedActions, recordStartTime, gMacro, _lastMouse
    if !recording
        return
    key := GetKeyName(Format("vk{:x}", wParam))
    ; ignore lone modifiers (we’ll capture in combos when held)
    if key in ["LControl","RControl","LShift","RShift","LAlt","RAlt","LWin","RWin"]
        return
    delay := A_TickCount - recordStartTime
    recordStartTime := A_TickCount
    recordedActions.Push("Sleep " delay)
    recordedActions.Push("Send, {" key "}")
    gMacro.Push(Map("type","KEY","key",key,"delay",delay))
    _lastMouse["t"] := A_TickCount
    RefreshLVRow(gMacro.Length)
}

RefreshLV() {
    global MacroLV, gMacro
    MacroLV.Delete()
    for idx, st in gMacro {
        t := st["type"], k := "", x := "", y := "", d := st.Has("delay")?st["delay"]:""
        if t="KEY" {
            k := st["key"]
        } else if t="CLICK" {
            k := st["btn"], x := st["x"], y := st["y"]
        } else if t="MOVE" {
            x := st["x"], y := st["y"]
        }
        MacroLV.Add("", idx, t, k, x, y, d)
    }
}
RefreshLVRow(i) {
    global MacroLV, gMacro
    st := gMacro[i]
    t := st["type"], k := "", x := "", y := "", d := st.Has("delay")?st["delay"]:""
    if t="KEY" {
        k := st["key"]
    } else if t="CLICK" {
        k := st["btn"], x := st["x"], y := st["y"]
    } else if t="MOVE" {
        x := st["x"], y := st["y"]
    }
    MacroLV.Add("", i, t, k, x, y, d)
}

; ============================================
; MANAGER
; ============================================
BrowseFile(*) {
    global currentMacroFile, main
    selFile := FileSelect(3, A_ScriptDir, "Select a Macro Script", "AHK Scripts (*.ahk)")
    if selFile {
        currentMacroFile := selFile
        SplitPath selFile, &name
        main["CurrentMacro"].Value := name
        HighlightCurrentMacro(name)
    }
}
BrowseFolder(*) {
    global currentMacroFolder, folderPathTxt
    start := folderPathTxt.Value != "" ? folderPathTxt.Value : ""
    selFolder := DirSelect(start, 3, "Select Macro Folder")
    if selFolder {
        currentMacroFolder := selFolder
        folderPathTxt.Value := selFolder
        RefreshFolderList()
    }
}
RefreshFolderList(*) {
    global macroList, currentMacroFolder
    if !DirExist(currentMacroFolder)
        return
    macroList.Delete()
    Loop Files, currentMacroFolder "\*.ahk" {
        macroList.Add("", A_LoopFileName)
    }
}
LoadSelectedMacro(*) {
    global macroList, currentMacroFile, currentMacroFolder, main
    row := macroList.GetNext(0, "F")
    if !row
        return
    fileName := macroList.GetText(row, 1)
    filePath := currentMacroFolder "\" fileName
    if FileExist(filePath) {
        currentMacroFile := filePath
        main["CurrentMacro"].Value := fileName
        HighlightCurrentMacro(fileName)
    }
}
HighlightCurrentMacro(fileName) {
    global macroList
    count := macroList.GetCount()
    if count = 0
        return
    Loop count {
        row := A_Index
        macroList.Modify(row, "", macroList.GetText(row, 1))
    }
    Loop count {
        row := A_Index
        if (macroList.GetText(row, 1) = fileName)
            macroList.Modify(row, "+Select")
    }
}

; context menu
MacroListContext(*) {
    global macroList, listMenu, ctxSelectedRow
    row := macroList.GetNext(0, "F")
    if !row
        return
    ctxSelectedRow := row
    MouseGetPos &mx, &my
    listMenu.Show(mx, my)
}
CtxRunMacro(*) {
    global ctxSelectedRow, macroList, currentMacroFolder, currentMacroFile, main
    if !ctxSelectedRow
        return
    fileName := macroList.GetText(ctxSelectedRow, 1)
    path := currentMacroFolder "\" fileName
    if FileExist(path) {
        currentMacroFile := path
        main["CurrentMacro"].Value := fileName
        RunMacro()
    }
}
CtxEditMacro(*) {
    global ctxSelectedRow, macroList, currentMacroFolder
    if !ctxSelectedRow
        return
    fileName := macroList.GetText(ctxSelectedRow, 1)
    path := currentMacroFolder "\" fileName
    if FileExist(path)
        Run(path)
}
CtxRenameMacro(*) {
    global ctxSelectedRow, macroList, currentMacroFolder
    if !ctxSelectedRow
        return
    old := macroList.GetText(ctxSelectedRow, 1)
    ib := InputBox("Enter new file name (with .ahk):", "Rename Macro", old)
    if ib.Result != "OK"
        return
    new := ib.Value
    if (new = "" || new = old)
        return
    oldPath := currentMacroFolder "\" old
    newPath := currentMacroFolder "\" new
    if FileExist(newPath) {
        MsgBox "A file with that name already exists."
        return
    }
    try FileMove(oldPath, newPath)
    catch {
        MsgBox "Rename failed."
        return
    }
    RefreshFolderList()
}
CtxDeleteMacro(*) {
    global ctxSelectedRow, macroList, currentMacroFolder, currentMacroFile, main
    if !ctxSelectedRow
        return
    name := macroList.GetText(ctxSelectedRow, 1)
    path := currentMacroFolder "\" name
    if MsgBox("Delete '" name "'?", "Confirm Delete", "YesNo Icon!") = "Yes" {
        try FileDelete(path)
        catch {
            MsgBox "Delete failed."
            return
        }
        if (currentMacroFile = path) {
            currentMacroFile := ""
            main["CurrentMacro"].Value := "None loaded"
        }
        RefreshFolderList()
    }
}

; ============================================
; RUNNER (speed + repeat + progress + analytics)
; ============================================
RunMacro(*) {
    global currentMacroFile, main, stopAll, totalSteps, doneSteps, analytics
    if currentMacroFile = "" {
        MsgBox "Please select a macro file first!"
        return
    }
    vals := main.Submit()
    speed := Round(vals.SpeedSpin / 10, 1)   ; 10 -> 1.0, 5 -> 0.5, etc.
    repeat := vals.RepeatSpin
    if speed < 0.1
        speed := 0.1
    if speed > 10
        speed := 10
    stopAll := false

    content := FileRead(currentMacroFile, "UTF-8")
    lines := StrSplit(content, "`n")
    lc := 0
    for line in lines
        if Trim(line) != ""
            lc++
    totalSteps := Max(1, lc * Max(1, repeat))
    doneSteps := 0
    UpdateProgress(0)
    SetStatus("Running")

    analytics["LastFile"] := currentMacroFile
    analytics["RunCount"] := Integer(analytics["RunCount"]) + 1
    startMs := A_TickCount
    SetTimer(() => PlayMacro(currentMacroFile, speed, repeat, startMs), -10)
}
PlayMacro(file, speed, repeat, startMs) {
    global stopAll, totalSteps, doneSteps, analytics
    content := FileRead(file, "UTF-8")
    if content = ""
        return
    errCount := 0

    Loop repeat {
        if stopAll
            break
        for raw in StrSplit(content, "`n") {
            if stopAll
                break
            line := Trim(raw)
            if line = ""
                continue

            if RegExMatch(line, "i)^Sleep\s+(\d+)", &m) {
                Sleep Round(m[1] / speed)   ; speed multiplier applied globally
            } else {
                try Exec(line)
                catch {
                    errCount++
                    ToolTip("Error in macro line:`n" line)
                    SetTimer(() => ToolTip(), -800)
                }
            }
            doneSteps++
            UpdateProgress(Round((doneSteps/totalSteps)*100))
        }
    }
    dur := A_TickCount - startMs
    analytics["TotalMs"] := Integer(analytics["TotalMs"]) + dur
    rc := Integer(analytics["RunCount"])
    analytics["AvgMs"] := Round(analytics["TotalMs"] / Max(1, rc))
    if errCount = 0
        analytics["LastSuccess"] := FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss")
    analytics["Errors"] := Integer(analytics["Errors"]) + errCount
    SaveAllIni()
    UpdateAnalyticsDisplay()

    if !stopAll {
        SetStatus("Ready")
        UpdateProgress(100), Sleep 300, UpdateProgress(0)
    } else {
        SetStatus("Killed")
        UpdateProgress(0)
    }
}
KillAllMacros(*) {
    global stopAll, recording, outputBox
    stopAll := true
    recording := false
    if IsSet(outputBox) {
        outputBox.Opt("-ReadOnly")
        outputBox.Value := "[⛔ All macros and recordings stopped manually]"
    }
    ToolTip("⛔ All macros stopped!")
    SetTimer(() => ToolTip(), -800)
}
Exec(line) {
    try {
        if RegExMatch(line, "i)^Click\s+(-?\d+)\s*,\s*(-?\d+)", &m) {
            Click Integer(m[1]), Integer(m[2])
        } else if RegExMatch(line, "i)^Send,\s*(.+)$", &m2) {
            Send m2[1]
        } else {
            ; allow arbitrary AHK commands in macro lines if desired
            ; you can expand here as needed
        }
    }
}

UpdateProgress(pct) {
    global progBar, progPct
    pct := pct<0 ? 0 : pct>100 ? 100 : pct
    progBar.Value := pct
    progPct.Value := pct "%"
}
SetStatus(text) {
    global statusText
    statusText.Value := "Status: " text
}

; ============================================
; TOOLS
; ============================================
ArmCoordGrab(*) {
    global coordArmed, coordLiveBox, settings
    coordArmed := true
    coordLiveBox.Value := "(click once…)"
    if settings["SnapToWindow"] {
        MouseGetPos &x, &y
        ToolTip("(" x ", " y ")")
    } else ToolTip()
    SetTimer(() => ToolTip(), -2000)  ; safety auto-hide regardless
}
ToggleSnap(*) {
    global settings, snapChk
    settings["SnapToWindow"] := (snapChk.Value = 1)
    SaveAllIni()
    if !settings["SnapToWindow"]
        ToolTip()
}
ChangeSnapFolder(*) {
    global settings, snapFolderTxt
    start := snapFolderTxt.Value != "" ? snapFolderTxt.Value : settings["ScreenshotFolder"]
    sel := DirSelect(start, 3, "Select Screenshot Save Folder")
    if sel {
        settings["ScreenshotFolder"] := sel
        snapFolderTxt.Value := sel
        if !DirExist(sel)
            DirCreate(sel)
        SaveAllIni()
        SetStatus("Screenshot folder set")
    }
}
TakeScreenshot(*) {
    global settings
    if !DirExist(settings["ScreenshotFolder"])
        DirCreate(settings["ScreenshotFolder"])
    ts := FormatTime(A_Now, "yyyyMMdd_HHmmss")
    imgFile := settings["ScreenshotFolder"] "\Screenshot_" ts ".png"
    Run("ms-screenclip:")
    ; wait for user to snip and clipboard receive
    if !ClipWait(10) {
        MsgBox "No image on clipboard. Snip cancelled or timed out."
        return
    }
    img := ClipboardAll()
    if !img {
        MsgBox "No image data on clipboard."
        return
    }
    WriteAll(imgFile, img)
    MsgBox "Screenshot saved:`n" imgFile
    SetStatus("Screenshot saved")
}

; ============================================
; EDITOR
; ============================================
OpenMacroForEdit(*) {
    global editorBox, editorStatusTxt, editorPath
    sel := FileSelect(3, A_ScriptDir, "Open Macro", "AHK Scripts (*.ahk)")
    if !sel
        return
    editorPath := sel
    editorBox.Value := FileRead(sel, "UTF-8")
    Editor_UpdateStatus()
}
SaveMacroEdit(*) {
    global editorBox, editorPath
    if (editorPath = "")
        return SaveMacroEditAs()
    WriteText(editorPath, editorBox.Value)
    MsgBox "Saved: " editorPath
}
SaveMacroEditAs(*) {
    global editorBox, editorPath
    init := editorPath!="" ? editorPath : A_ScriptDir "\*.ahk"
    tgt := FileSelect("S", init, "Save Macro As", "AHK Scripts (*.ahk)")
    if !tgt
        return
    if !InStr(tgt, ".ahk")
        tgt .= ".ahk"
    WriteText(tgt, editorBox.Value)
    editorPath := tgt
    MsgBox "Saved: " tgt
    Editor_UpdateStatus()
}
ReloadMacroEdit(*) {
    global editorBox, editorPath
    if (editorPath = "" || !FileExist(editorPath))
        return
    editorBox.Value := FileRead(editorPath, "UTF-8")
    Editor_UpdateStatus()
}
Editor_UpdateStatus() {
    global editorStatusTxt, editorPath
    editorStatusTxt.Value := "Editing: " (editorPath!="" ? editorPath : "(none)")
}

; ============================================
; ANALYTICS
; ============================================
UpdateAnalyticsDisplay() {
    global analytics, aLast, aCount, aAvg, aTotal, aErrs, aWhen
    aLast.Value  := "Last Macro: " (analytics["LastFile"]!="" ? analytics["LastFile"] : "(none)")
    aCount.Value := "Run Count: " analytics["RunCount"]
    aAvg.Value   := "Average Duration: " analytics["AvgMs"] " ms"
    aTotal.Value := "Total Runtime: " analytics["TotalMs"] " ms"
    aErrs.Value  := "Errors: " analytics["Errors"]
    aWhen.Value  := "Last Success: " (analytics["LastSuccess"]!="" ? analytics["LastSuccess"] : "(n/a)")
}
ExportAnalyticsCSV(*) {
    global analytics
    ts := FormatTime(A_Now, "yyyyMMdd_HHmmss")
    path := A_ScriptDir "\Analytics_" ts ".csv"
    text := "LastFile,RunCount,TotalMs,AvgMs,LastSuccess,Errors`n"
    text .= '"' analytics["LastFile"] '",' analytics["RunCount"] "," analytics["TotalMs"] "," analytics["AvgMs"] ',"' analytics["LastSuccess"] '",' analytics["Errors"] "`n"
    WriteText(path, text)
    MsgBox "Exported CSV:`n" path
}
ExportAnalyticsTXT(*) {
    global analytics
    ts := FormatTime(A_Now, "yyyyMMdd_HHmmss")
    path := A_ScriptDir "\Analytics_" ts ".txt"
    text := "Last File: " analytics["LastFile"] "`n"
    text .= "Run Count: " analytics["RunCount"] "`n"
    text .= "Total Ms: " analytics["TotalMs"] "`n"
    text .= "Avg Ms: " analytics["AvgMs"] "`n"
    text .= "Last Success: " analytics["LastSuccess"] "`n"
    text .= "Errors: " analytics["Errors"] "`n"
    WriteText(path, text)
    MsgBox "Exported TXT:`n" path
}
