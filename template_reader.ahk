#Requires AutoHotkey v2.0
#SingleInstance Force

csvPath := "template.csv"

if !FileExist(csvPath) {
    MsgBox("Error: Can't find " csvPath " in this folder.", "Missing File", 48)
    ExitApp()
}

loadedCount := 0

; Loop through every line of the file safely
Loop Read, csvPath {
    ; Skip empty lines
    if (A_LoopReadLine = "")
        continue
        
    ; Split each line at the first comma
    commaPos := InStr(A_LoopReadLine, ",")
    if (commaPos > 0) {
        ; Extract and trim any accidental spaces around the shortcut
        trigger := Trim(SubStr(A_LoopReadLine, 1, commaPos - 1))
        ; Extract the full text message
        message := SubStr(A_LoopReadLine, commaPos + 1)
        
        ; Remove accidental wrapping quotes if Excel added them
        trigger := Trim(trigger, '"')
        message := Trim(message, '"')
        
        ; Create a native automatic hotstring (Just like your working example)
        ; :* means expand instantly without needing spacebar
        Hotstring(":*:" . trigger, message)
        loadedCount++
    }
}

TrayTip("Success!", "Loaded " loadedCount " shortcuts cleanly.")
