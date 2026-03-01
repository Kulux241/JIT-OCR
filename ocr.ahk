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
    Run('pythonw "' A_ScriptDir '\ocr.py"', A_ScriptDir, , &ocrPID)
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
    ; "ctrl+shift+p" → "^+p"
    human := StrLower(Trim(human))
    parts := StrSplit(human, "+")
    if parts.Length < 2
        return "^+p"

    result := ""
    lastPart := parts[parts.Length]

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

    result .= lastPart
    return result
}

HumanHotkey(human) {
    ; "ctrl+shift+p" → "Ctrl+Shift+P" (display format)
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
        ; Fallback
        Hotkey("^+p", (*) => RunOCR())
        currentHotkey := "^+p"
    }
}

; ─── Model Switching ─────────────────────────────────

GetActiveModel() {
    path := A_ScriptDir "\settings.json"
    if !FileExist(path)
        return ""
    content := FileRead(path)
    if RegExMatch(content, '"active_model"\s*:\s*"([^"]+)"', &m)
        return m[1]
    return ""
}

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

; ─── Build Tray Menu ─────────────────────────────────

BuildMenu() {
    path := A_ScriptDir "\settings.json"
    human := ReadHotkeyFromSettings()
    displayHK := HumanHotkey(human)

    A_TrayMenu.Delete()
    A_TrayMenu.Add("Scan Region`t" displayHK, (*) => RunOCR())
    A_TrayMenu.Add()

    ; Models submenu
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
    A_TrayMenu.Add()
    A_TrayMenu.Add("Exit", (*) => ExitApp())
}

OpenSettings() {
    Run('notepad "' A_ScriptDir '\settings.json"')
    ; Re-read hotkey after settings close
    SetTimer(() => (RegisterHotkey(), BuildMenu()), -3000)
}