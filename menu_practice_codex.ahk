#Requires AutoHotkey v2.0
#SingleInstance Force

; =========================
; グローバル定義（設定）
; =========================
global InputBuffer := ""
global TargetHwnd := 0
global MenuIsOpen := false

; 定数扱いのパスを先に定義し、利用箇所は既存変数名を維持
global ChatGptLogDir := "C:\Users\mezzs\OneDrive - 株式会社BON\SP_AIログ\ChatGPT"
global TempLogFile := ChatGptLogDir "\temp_chatgpt_log.txt"
global MenuTitle := "数値メニュー"

global MenuGui := Gui("+AlwaysOnTop -MinimizeBox -MaximizeBox", MenuTitle)
MenuGui.SetFont("s10", "Meiryo")

global TxtCode := MenuGui.AddText("w360", "入力：")
global TxtMeaning := MenuGui.AddText("w360", "意味：")
global TxtCandidates := MenuGui.AddText("w360 h220", "候補：")

; =========================
; テンプレ登録
; =========================
global Templates := Map()
Templates["300"] := "チャットが重くなったので、引継ぎのプロンプトをお願いします。"

; =========================
; ホットキー
; =========================
F2::StartMenuMode()

#HotIf MenuIsOpen
Numpad0::HandleDigit("0")
Numpad1::HandleDigit("1")
Numpad2::HandleDigit("2")
Numpad3::HandleDigit("3")
Numpad4::HandleDigit("4")
Numpad5::HandleDigit("5")
Numpad6::HandleDigit("6")
Numpad7::HandleDigit("7")
Numpad8::HandleDigit("8")
Numpad9::HandleDigit("9")
Backspace::HandleBackspace()
Esc::CancelMenuMode()
#HotIf

; =========================
; メニュー開始
; =========================
StartMenuMode() {
    global InputBuffer, TargetHwnd, MenuIsOpen

    InputBuffer := ""
    TargetHwnd := WinExist("A")
    MenuIsOpen := true

    UpdateGui()
    ShowGuiNearMouse()
}

; =========================
; 数字入力処理
; =========================
HandleDigit(digit) {
    global InputBuffer

    if StrLen(InputBuffer) >= 3 {
        return
    }

    InputBuffer .= digit
    UpdateGui()

    if StrLen(InputBuffer) = 3 {
        ExecuteCode(InputBuffer)
    }
}

; =========================
; Backspace処理
; =========================
HandleBackspace() {
    global InputBuffer

    if StrLen(InputBuffer) > 0 {
        InputBuffer := SubStr(InputBuffer, 1, StrLen(InputBuffer) - 1)
        UpdateGui()
    }
}

; =========================
; キャンセル
; =========================
CancelMenuMode() {
    global InputBuffer, MenuIsOpen, MenuGui

    InputBuffer := ""
    MenuIsOpen := false
    MenuGui.Hide()
}

; =========================
; 実行処理
; =========================
ExecuteCode(code) {
    global Templates, TargetHwnd, InputBuffer, MenuIsOpen, MenuGui

    MenuGui.Hide()
    MenuIsOpen := false

    if TargetHwnd {
        WinActivate("ahk_id " TargetHwnd)
        Sleep(200)
    }

    if code = "000" {
        ChatGPTLogStep1(false)
    } else if code = "001" {
        ChatGPTLogStep2()
    } else if code = "002" {
        ChatGPTLogAuto()
    } else if Templates.Has(code) {
        A_Clipboard := Templates[code]
        ClipWait(1)
        Send("^v")
    } else {
        MsgBox("未登録の番号です：" code, "数値メニュー", "Icon!")
    }

    InputBuffer := ""
}

; =========================
; 000：ChatGPTログ取得＋要約依頼送信
; =========================
ChatGPTLogStep1(silent := false) {
    global TempLogFile

    A_Clipboard := ""
    Send("^a")
    Sleep(300)
    Send("^c")

    if !ClipWait(2) {
        MsgBox("コピーに失敗しました。ChatGPT画面をクリックしてから再実行してください。", "ログ保存 STEP1", "Icon!")
        return
    }

    logText := A_Clipboard

    if Trim(logText) = "" {
        MsgBox("コピー内容が空です。ChatGPT画面をクリックしてから再実行してください。", "ログ保存 STEP1", "Icon!")
        return
    }

    if FileExist(TempLogFile) {
        FileDelete(TempLogFile)
    }

    FileAppend(logText, TempLogFile, "UTF-8")

    dq := Chr(34)
    prompt := "この会話の内容を、ファイル名用に20〜35文字で要約してください。"
        . "`n必ず以下の形式で出力してください："
        . "`n`nTITLE_START"
        . "`n要約タイトル"
        . "`nTITLE_END"
        . "`n`n禁止：記号（\/:*?" dq "<>|）"
        . "`n出力はこの形式のみ"

    A_Clipboard := prompt
    ClipWait(1)
    Sleep(200)

    Send("^v")
    Sleep(200)
    Send("{Enter}")

    if !silent {
        MsgBox("STEP1完了：ログ本文を一時保存し、要約依頼を送信しました。", "ログ保存 STEP1")
    }
}

; GUI更新
; =========================
UpdateGui() {
    global InputBuffer, TxtCode, TxtMeaning, TxtCandidates

    TxtCode.Text := "入力：" InputBuffer
    TxtMeaning.Text := "意味：" GetMeaning(InputBuffer)
    TxtCandidates.Text := GetCandidates(InputBuffer)
}

; =========================
; 意味表示
; =========================
GetMeaning(code) {
    if code = "" {
        return "F2入力待機中"
    }

    if code = "0" {
        return "大分類：AIログ保存"
    }

    if code = "00" {
        return "中分類：ChatGPTログ保存"
    }

    if code = "000" {
        return "実行：STEP1 ログ取得＋要約依頼送信"
    }

    if code = "001" {
        return "実行予定：STEP2 タイトル取得＋正式保存"
    }

    if code = "3" {
        return "大分類：ChatGPT操作"
    }

    if code = "30" {
        return "中分類：引継ぎ・会話整理"
    }

    if code = "300" {
        return "実行：引継ぎプロンプト依頼文を貼り付け"
    }

    return "※未登録"
}

; =========================
; 候補一覧
; =========================
GetCandidates(code) {
    if code = "" {
        return "候補：`n0：AIログ保存`n3：ChatGPT操作"
    }

    if code = "0" {
        return "候補：`n00：ChatGPTログ保存"
    }

    if code = "00" {
        return "候補：`n000：STEP1 ログ取得＋要約依頼送信`n001：STEP2 タイトル取得＋正式保存`n002：完全自動保存"
    }

    if code = "000" {
        return "候補：`n000を実行します"
    }

    if code = "001" {
        return "候補：`n001は次ステップで実装します"
    }

    if code = "3" {
        return "候補：`n30：引継ぎ・会話整理"
    }

    if code = "30" {
        return "候補：`n300：引継ぎプロンプト依頼`n301：※未使用`n302：※未使用`n303：※未使用`n304：※未使用`n305：※未使用`n306：※未使用`n307：※未使用`n308：※未使用`n309：※未使用"
    }

    if code = "300" {
        return "候補：`n300を実行します"
    }

    return "候補：`n※未登録"
}

; =========================
; マウス近傍にGUI表示
; 画面端ではみ出さない補正あり
; =========================
ShowGuiNearMouse() {
    global MenuGui

    MouseGetPos(&mx, &my)

    guiW := 400
    guiH := 340
    offsetX := 20
    offsetY := 20

    x := mx + offsetX
    y := my + offsetY

    monitorCount := MonitorGetCount()

    selectedLeft := 0
    selectedTop := 0
    selectedRight := A_ScreenWidth
    selectedBottom := A_ScreenHeight

    Loop monitorCount {
        MonitorGetWorkArea(A_Index, &left, &top, &right, &bottom)

        if (mx >= left && mx <= right && my >= top && my <= bottom) {
            selectedLeft := left
            selectedTop := top
            selectedRight := right
            selectedBottom := bottom
            break
        }
    }

    if (x + guiW > selectedRight) {
        x := mx - guiW - offsetX
    }

    if (y + guiH > selectedBottom) {
        y := my - guiH - offsetY
    }

    if (x < selectedLeft) {
        x := selectedLeft
    }

    if (y < selectedTop) {
        y := selectedTop
    }

    MenuGui.Show("x" x " y" y " w" guiW " h" guiH)
}



; =========================
; ChatGPT URLからCID取得
; =========================
; =========================
; 001：タイトル取得＋正式保存
; =========================
ChatGPTLogStep2() {
    global TempLogFile

    A_Clipboard := ""
    Send("^a")
    Sleep(300)
    Send("^c")

    if !ClipWait(2) {
        MsgBox("タイトル抽出用テキストの取得に失敗しました", "STEP2", "Icon!")
        return
    }

    fullText := A_Clipboard
    fullText := StrReplace(fullText, "`r`n", "`n")
    fullText := StrReplace(fullText, "`r", "`n")

    title := ""

    if RegExMatch(fullText, "s)TITLE_START\s*(.*?)\s*TITLE_END", &m) {
        title := Trim(m[1])
    }

    badChars := ["\", "/", ":", "*", "?", Chr(34), "<", ">", "|", "`r", "`n", "`t"]
    for bad in badChars {
        title := StrReplace(title, bad, "")
    }

    title := Trim(title)

    if (title = "" || title = "～" || title = "無題") {
        MsgBox("タイトル未取得のため保存を中止しました", "ERROR")
        return
    }

    if StrLen(title) > 40 {
        title := SubStr(title, 1, 40)
    }

    if !FileExist(TempLogFile) {
        MsgBox("一時ログが見つかりません", "STEP2", "Icon!")
        return
    }

    logText := FileRead(TempLogFile, "UTF-8")

    saveDir := "C:\Users\mezzs\OneDrive - 株式会社BON\SP_AIログ\ChatGPT"

    if !DirExist(saveDir) {
        MsgBox("保存先フォルダが見つかりません：" saveDir, "STEP2", "Icon!")
        return
    }

    cid := GetChatGptCid()

    if cid != "" {
        dt := FormatTime(, "yyyyMMdd")
        filePath := saveDir "\" dt "_AIログ_" title "_CID" cid ".txt"

        if FileExist(filePath) {
            FileDelete(filePath)
        }
    } else {
        dt := FormatTime(, "yyyyMMdd_HHmmss")
        filePath := saveDir "\" dt "_AIログ_" title "_CID取得失敗.txt"
    }

    FileAppend(logText, filePath, "UTF-8")
    FileDelete(TempLogFile)

    MsgBox("保存完了：" filePath, "STEP2")
}

; =========================
; 002：完全自動保存（検知待機付き）
; =========================
ChatGPTLogAuto() {
    global TempLogFile

    ; STEP1実行
    ChatGPTLogStep1(true)

    ; 待機（最大7秒）
    found := false

    Loop 14 {
        Sleep(500)

        A_Clipboard := ""
        Send("^a")
        Sleep(200)
        Send("^c")

        if !ClipWait(1)
            continue

        txt := A_Clipboard

        if InStr(txt, "TITLE_END") {
            found := true
            break
        }
    }

    if !found {
        MsgBox("要約取得タイムアウト（7秒）", "002", "Icon!")
        return
    }

    ; STEP2実行
    ChatGPTLogStep2()
}

; =========================
; ChatGPT URLからCID取得
; =========================
GetChatGptCid() {
    savedClip := ClipboardAll()

    A_Clipboard := ""
    Send("^l")
    Sleep(200)
    Send("^c")

    if !ClipWait(1) {
        A_Clipboard := savedClip
        return ""
    }

    url := A_Clipboard
    Send("{Esc}")
    Sleep(100)

    A_Clipboard := savedClip

    if RegExMatch(url, "/c/([^/?#]+)", &m) {
        raw := m[1]
        raw := StrReplace(raw, "-", "")
        raw := RegExReplace(raw, "[^A-Za-z0-9]", "")

        if StrLen(raw) >= 12 {
            return SubStr(raw, 1, 12)
        }
    }

    return ""
}

; test




