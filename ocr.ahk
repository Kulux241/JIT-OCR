#Requires AutoHotkey v2.0
#SingleInstance Force

TraySetIcon("shell32.dll", 14)

global ocrPID := 0
global currentHotkey := ""

RegisterHotkey()
BuildMenu()

RunOCR() {
    global ocrPID

    if ocrPID && ProcessExist(ocrPID)
        return

    ocrPath := A_ScriptDir "\ocr.exe"

    if !FileExist(ocrPath) {
        MsgBox("Cannot find: " ocrPath, "OCR Tool Error")
        return
    }

    Run('"' ocrPath '"', A_ScriptDir, , &ocrPID)
}

; ─── Hotkey Management ───────────────────────────────

ReadHotkeyFromSettings() {
    path := A_ScriptDir "\settings.json"
    if !FileExist(path)
        return "ctrl+shift+p"
    content := FileRead(path)
    if RegExMatch(content, '"hotkey"\s*:\s*"([^"]+)"', &m)
        return m[1]
    return "ctrl+shift+p"
}

ConvertHotkey(human) {
    human := StrLower(Trim(human))
    parts := StrSplit(human, "+")
    if parts.Length < 2
        return "^+p"

    result := ""

    for i, part in parts {
        part := Trim(part)
        if (i = parts.Length)
            continue
        if (part = "ctrl" || part = "control")
            result .= "^"
        else if (part = "alt")
            result .= "!"
        else if (part = "shift")
            result .= "+"
        else if (part = "win" || part = "windows")
            result .= "#"
    }

    result .= parts[parts.Length]
    return result
}

HumanHotkey(human) {
    parts := StrSplit(human, "+")
    result := ""
    for i, part in parts {
        part := Trim(part)
        part := StrUpper(SubStr(part, 1, 1)) SubStr(part, 2)
        if (i > 1)
            result .= "+"
        result .= part
    }
    return result
}

RegisterHotkey() {
    global currentHotkey
    human := ReadHotkeyFromSettings()
    ahkKey := ConvertHotkey(human)

    if (currentHotkey != "" && currentHotkey != ahkKey) {
        try Hotkey(currentHotkey, , "Off")
    }

    try {
        Hotkey(ahkKey, (*) => RunOCR())
        currentHotkey := ahkKey
    } catch {
        Hotkey("^+p", (*) => RunOCR())
        currentHotkey := "^+p"
    }
}

; ─── Model Switching ─────────────────────────────────

SetActiveModel(id, *) {
    path := A_ScriptDir "\settings.json"
    if !FileExist(path)
        return
    content := FileRead(path)
    content := RegExReplace(content, '("active_model"\s*:\s*")[^"]*"', '${1}' id '"')
    FileDelete(path)
    FileAppend(content, path)
    BuildMenu()
    TrayTip("Model: " id, "OCR Tool")
}

; ─── Autostart ───────────────────────────────────────

HasAutostart() {
    return FileExist(A_Startup "\OCR Tool.lnk")
}

ToggleAutostart() {
    shortcutPath := A_Startup "\OCR Tool.lnk"
    exePath := A_ScriptDir "\hotkey.exe"

    if FileExist(shortcutPath) {
        FileDelete(shortcutPath)
        TrayTip("Autostart disabled", "OCR Tool")
    } else {
        ComObj := ComObject("WScript.Shell")
        Shortcut := ComObj.CreateShortcut(shortcutPath)
        Shortcut.TargetPath := exePath
        Shortcut.WorkingDirectory := A_ScriptDir
        Shortcut.Description := "OCR Tool"
        Shortcut.Save()
        TrayTip("Autostart enabled", "OCR Tool")
    }

    BuildMenu()
}

; ─── Build Tray Menu ─────────────────────────────────

BuildMenu() {
    path := A_ScriptDir "\settings.json"
    human := ReadHotkeyFromSettings()
    displayHK := HumanHotkey(human)

    A_TrayMenu.Delete()
    A_TrayMenu.Add("Scan Region`t" displayHK, (*) => RunOCR())
    A_TrayMenu.Add()

    if FileExist(path) {
        content := FileRead(path)

        active := ""
        if RegExMatch(content, '"active_model"\s*:\s*"([^"]+)"', &am)
            active := am[1]

        modelMenu := Menu()
        foundModels := false

        if RegExMatch(content, '"models"\s*:\s*\{', &mBlock) {
            searchStart := mBlock.Pos + mBlock.Len
            pos := searchStart

            Loop {
                if !RegExMatch(content, '"([\w][\w.-]*)"\s*:\s*\{', &idMatch, pos)
                    break
                if (idMatch.Pos - searchStart > 5000)
                    break

                modelId := idMatch[1]
                chunkStart := idMatch.Pos + idMatch.Len
                chunk := SubStr(content, chunkStart, 500)

                if RegExMatch(chunk, '"name"\s*:\s*"([^"]+)"', &nameMatch) {
                    modelName := nameMatch[1]
                    fn := SetActiveModel.Bind(modelId)
                    modelMenu.Add(modelName, fn)
                    if (modelId = active)
                        modelMenu.Check(modelName)
                    foundModels := true
                }

                pos := chunkStart + 1
            }
        }

        if foundModels
            A_TrayMenu.Add("Models", modelMenu)
    }

    A_TrayMenu.Add()
    A_TrayMenu.Add("Settings", (*) => OpenSettings())
    A_TrayMenu.Add("Start with Windows", (*) => ToggleAutostart())
    if HasAutostart()
        A_TrayMenu.Check("Start with Windows")
    A_TrayMenu.Add()
    A_TrayMenu.Add("Exit", (*) => ExitApp())
}

OpenSettings() {
    Run('notepad "' A_ScriptDir '\settings.json"')
    SetTimer(() => (RegisterHotkey(), BuildMenu()), -3000)
}