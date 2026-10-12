#Requires AutoHotkey v2.0
#SingleInstance Force
#NoTrayIcon                            ; ingen ikon i aktivitetsfältet; avsluta genom att starta om (Force)
; critic-export från Obsidian: exporterar den anteckning som är öppen.
; Hittar filen via fönstertiteln ("Anteckning - Valv - Obsidian v1.x")
; och valvets sökväg i Obsidians egen obsidian.json.
; Startas genom att dubbelklicka på critic-export.cmd i mappen ovanför.

#HotIf WinActive("ahk_exe Obsidian.exe")
^!w::CriticExport("--to docx --open")                      ; Ctrl+Alt+W  Word
^!p::CriticExport("--to pdf --critic markup --open")       ; Ctrl+Alt+P  PDF, granskning
^!+p::CriticExport("--to pdf --critic accept --open")      ; Ctrl+Alt+Shift+P  ren PDF
^!h::CriticExport("--to html --open")                      ; Ctrl+Alt+H  HTML
#HotIf

CriticExport(args) {
    file := ObsidianActiveFile()
    if !file
        return
    Send "^s"                          ; se till att sista ändringen är sparad
    Sleep 400
    ; Python körs dolt och kan inte fråga, så titeln efterfrågas här.
    ; Författaren tas från document_author i settings.toml.
    args .= " --no-prompt"
    if !NoteHasTitle(file) {
        ib := InputBox("Anteckningen saknar titel. Titel för exporten (tomt = ingen):",
            "critic-export", "w360 h130", SuggestedTitle(file))
        if ib.Result = "Cancel"
            return
        title := StrReplace(Trim(ib.Value), '"', "”")
        if title != "" {
            args .= ' --title "' title '"'
            if MsgBox("Spara titeln i anteckningens YAML-huvud?", "critic-export", 36) = "Yes"
                args .= " --save-metadata"
        }
    }
    ToolTip "critic-export: " args
    log := A_Temp "\critic-export.log"
    ; PYTHONIOENCODING: annars skriver Python åäö i fel teckenkodning till loggen
    startCmd := A_ScriptDir "\..\critic-export.cmd"
    exitCode := RunWait(A_ComSpec ' /c set PYTHONIOENCODING=utf-8&& "' startCmd '" "' file '" '
        . args ' > "' log '" 2>&1', , "Hide")
    ToolTip
    output := FileExist(log) ? FileRead(log, "UTF-8") : ""
    if exitCode
        MsgBox "Exporten misslyckades:`n`n" output, "critic-export", 16
    else if RegExMatch(output, "i)warning|note:")
        MsgBox "Exporten klar, men med varningar:`n`n" output, "critic-export", 48
    else
    {
        ToolTip "critic-export: exporten klar"   ; TrayTip kräver en ikon i aktivitetsfältet
        SetTimer () => ToolTip(), -2500
    }
}

; Förslag: en H1-rubrik som inleder anteckningen, annars filnamnet.
SuggestedTitle(file) {
    text := FileRead(file, "UTF-8")
    text := RegExReplace(text, "s)\A---\R.*?\R(?:---|\.\.\.)\R", "", , 1)
    if RegExMatch(text, "\A\s*#[ \t]+(.+?)(?:[ \t]+#+)?[ \t]*(?:\R|$)", &m)
        return Trim(m[1])
    SplitPath file, , , , &stem
    return stem
}

NoteHasTitle(file) {
    text := FileRead(file, "UTF-8")
    if !RegExMatch(text, "s)\A---\R(.*?)\R(?:---|\.\.\.)\R", &m)
        return false
    return RegExMatch(m[1], "m)^title\s*:[ \t]*[^\s#]")
}

ObsidianActiveFile() {
    title := WinGetTitle("A")
    ; "Anteckning - Valv - Obsidian 1.13.7" (äldre versioner: "Obsidian v1.x")
    if !RegExMatch(title, "^(.*) - (.*?) - Obsidian(?: v?[\d.]+)?$", &m) {
        MsgBox "Ingen anteckning verkar vara öppen.`n`nFönstertitel: " title, "critic-export", 48
        return ""
    }
    note := m[1], vaultName := m[2]

    vault := ObsidianVaultPath(vaultName)
    if !vault {
        MsgBox "Hittade inte valvet " vaultName " i Obsidians inställningar.", "critic-export", 48
        return ""
    }

    matches := []
    Loop Files vault "\*.md", "R" {
        if InStr(A_LoopFileFullPath, "\.obsidian\") || InStr(A_LoopFileFullPath, "\.trash\")
            continue
        if (A_LoopFileName = note ".md")
            matches.Push(A_LoopFileFullPath)
    }
    if matches.Length = 0 {
        MsgBox "Hittade inte " note ".md i " vault, "critic-export", 48
        return ""
    }
    if matches.Length > 1 {
        list := ""
        for p in matches
            list .= "`n" p
        MsgBox "Flera anteckningar heter " note ".md. Byt namn på en av dem:" list, "critic-export", 48
        return ""
    }
    return matches[1]
}

ObsidianVaultPath(vaultName) {
    try json := FileRead(A_AppData "\obsidian\obsidian.json", "UTF-8")
    catch
        return ""
    pos := 1
    while pos := RegExMatch(json, '"path"\s*:\s*"((?:[^"\\]|\\.)*)"', &m, pos) {
        path := StrReplace(StrReplace(m[1], "\\", "\"), "\/", "/")
        SplitPath path, &name
        if (name = vaultName)
            return path
        pos += m.Len
    }
    return ""
}
