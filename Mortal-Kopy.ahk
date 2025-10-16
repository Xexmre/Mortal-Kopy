; This code was partially generated with the assistance of GPT-5.
; 该程序由GPT-5协助编写。

; ====== 自动填写问卷脚本 ======
; 版本：AHK v2
; 功能：自动填写单选、多选、填空、判断题、名词解释、简答题
; 注意：JSON 必须为 UTF-8 无 BOM 编码
; 执行：按下 F10 运行

#Requires AutoHotkey v2.0

; --------- 工具函数：JSON 字符串反转义（用于 term_explain / short_answer 文本）---------
JsonUnescape(s) {
    ; 处理常见转义：\"  \\  \n  \r  \t
    s := StrReplace(s, "\" "", "" "")   ; \"  ->  "
    s := StrReplace(s, "\\", "\")     ; \\  ->  \
    s := StrReplace(s, "\n", "`n")    ; \n  ->  换行
    s := StrReplace(s, "\r", "`r")    ; \r  ->  回车
    s := StrReplace(s, "\t", "`t")    ; \t  ->  制表
    return s
}

; --------- 小工具：解析输入中的卷号与 --skip 列表 ----------
ParseRollAndSkips(inputText) {
    ; 返回对象 { roll: "2.2.2", skip: Map() }
    res := { roll: "", skip: Map() }
    inputText := Trim(inputText)

    ; 支持两种写法： --skip.xxx,yyy  或  --skip=xxx,yyy
    if RegExMatch(inputText, 'i)^\s*([^ \t]+)\s*(?:--skip[.=]([A-Za-z_,]+))?\s*$', &m) {
        res.roll := m[1]
        list := (m.Count >= 2) ? m[2] : ""
        if (list != "") {
            for token in StrSplit(list, ",") {
                k := Trim(StrLower(token))
                ; 别名规范化
                if (k = "term")
                    k := "term_explain"
                else if (k = "short")
                    k := "short_answer"
                res.skip[k] := true
            }
        }
    } else {
        ; 不符合格式就把整串当作卷号
        res.roll := inputText
    }
    return res
}

F10::
{
    ; ========== 0. 让用户输入卷号与可选跳过标志 ==========
    msg := "请输入卷号，可选 --skip：`n"
        . "示例1：2.2.2`n"
        . "示例2：2.2.2 --skip.multi_choice,judgment`n"
        . "示例3：2.2.2 --skip=term_explain,short_answer"
    ib := InputBox(msg, "真题快抄 — 选择卷号 & 跳过模块", "w420 h160", "2.2.2")
    if (ib.Result = "Cancel")
        return

    parsed := ParseRollAndSkips(ib.Value)
    roll := Trim(parsed.roll)
    skip := parsed.skip
    if (roll = "") {
        TrayTip("未输入卷号，已取消。", "提示", 1)
        return
    }

    ; 若用户已输入“.json”，也能兼容；否则自动补全
    if !RegExMatch(roll, "\.json$")
        rollFile := roll ".json"
    else
        rollFile := roll

    jsonPath := A_ScriptDir "\paper_db\" rollFile
    if !FileExist(jsonPath) {
        TrayTip("未找到文件：" rollFile "`n请确认它位于 .\paper_db\ 目录。", "错误", 3)
        return
    }

    ; ---------- 1. 读取 JSON ----------
    try jsonText := FileRead(jsonPath, "UTF-8")
    catch {
        TrayTip("无法读取 " rollFile "！请检查编码是否为 UTF-8 无 BOM。", "错误", 3)
        return
    }

    ; ---------- 2. 通用参数 ----------
    y_start := 265
    row_gap := 55
    module_gap := 115

    ; 计算模块起始 Y（把前面已存在的模块行数累计上去；与是否跳过无关）
    CalcStartY(counts*) {
        y := y_start
        for each, c in counts {
            if (c > 0) {
                rows := Ceil(c / 5)
                y += row_gap * (rows - 1) + module_gap
            }
        }
        return y
    }

    ; ======================================================================
    ;                           single_choice
    ; ======================================================================
    single_choice := Map()
    single_section := RegExReplace(jsonText, 's).*?"single_choice"\s*:\s*\{(.*?)\}.*', '$1')
    for , field in StrSplit(single_section, ",") {
        if RegExMatch(field, '"(\d+)"\s*:\s*"([A-D])"', &m)
            single_choice[m[1]] := m[2]
    }
    single_count := single_choice.Count
    TrayTip("已读取 " single_count " 道单选题", "信息", 1)

    ; ---------- 填写 single_choice（可跳过） ----------
    if !skip.Has("single_choice") {
        loop single_count {
            i := A_Index
            row := Ceil(i / 5)
            col := Mod(i - 1, 5) + 1
            x := 995 + 50 * (col - 1)
            y := y_start + row_gap * (row - 1)

            Sleep 100
            MouseMove(x, y, 0)
            Click()
            Sleep 800

            key := String(i)
            if !single_choice.Has(key)
                continue
            choice := single_choice[key]

            if choice = "A"
                MouseMove(95, 265, 0)
            else if choice = "B"
                MouseMove(95, 315, 0)
            else if choice = "C"
                MouseMove(95, 365, 0)
            else if choice = "D"
                MouseMove(95, 415, 0)
            Click()
        }
    }

    ; ======================================================================
    ;                           multi_choice
    ; ======================================================================
    multi_choice := Map()
    multi_section := RegExReplace(jsonText, 's).*?"multi_choice"\s*:\s*\{(.*?)\}.*', '$1')
    for , field in StrSplit(multi_section, ",") {
        if RegExMatch(field, '"(\d+)"\s*:\s*"([A-D]+)"', &m)
            multi_choice[m[1]] := m[2]
    }
    multi_count := multi_choice.Count
    TrayTip("已读取 " multi_count " 道多选题", "信息", 1)

    y_multi := CalcStartY(single_count)

    ; ---------- 填写 multi_choice（可跳过） ----------
    if !skip.Has("multi_choice") {
        loop multi_count {
            i := A_Index
            row := Ceil(i / 5)
            col := Mod(i - 1, 5) + 1
            x := 995 + 50 * (col - 1)
            y := y_multi + row_gap * (row - 1)

            Sleep 100
            MouseMove(x, y, 0)
            Click()
            Sleep 800

            key := String(i + single_count)
            if !multi_choice.Has(key)
                continue
            choiceStr := multi_choice[key]

            for choice in StrSplit(choiceStr) {
                if choice = "A"
                    MouseMove(95, 265, 0)
                else if choice = "B"
                    MouseMove(95, 315, 0)
                else if choice = "C"
                    MouseMove(95, 365, 0)
                else if choice = "D"
                    MouseMove(95, 415, 0)
                Click()
                Sleep 200
            }
        }
    }

    ; ======================================================================
    ;                           fill_blank（多空）
    ; ======================================================================
    fill_blank := Map()
    fill_section := RegExReplace(jsonText, 's).*?"fill_blank"\s*:\s*\{(.*?)\}.*', '$1')
    pos := 1
    while pos := RegExMatch(fill_section, '"(\d+)"\s*:\s*\[(.*?)\]', &m, pos) {
        idx := m[1]
        raw := m[2]
        clean := RegExReplace(raw, '"|\r|\n')
        clean := RegExReplace(clean, '\s+')
        fill_blank[idx] := clean   ; 以逗号串保存，稍后再 split
        pos += StrLen(m[0])
    }
    fill_count := fill_blank.Count
    TrayTip("已读取 " fill_count " 道填空题", "信息", 1)

    y_fill := CalcStartY(single_count, multi_count)

    ; ---------- 填写 fill_blank（可跳过） ----------
    if !skip.Has("fill_blank") {
        loop fill_count {
            i := A_Index
            row := Ceil(i / 5)
            col := Mod(i - 1, 5) + 1
            x := 995 + 50 * (col - 1)
            y := y_fill + row_gap * (row - 1)

            Sleep 100
            MouseMove(x, y, 0)
            Click()
            Sleep 800

            key := String(i + single_count + multi_count)
            if !fill_blank.Has(key)
                continue
            blanks := fill_blank[key]
            arr := StrSplit(blanks, ",")

            for n, fillValue in arr {
                posY := 290 + 80 * (n - 1)
                MouseMove(180, posY, 0)
                Click()
                Sleep 200
                SendText(fillValue)
                Sleep 200
            }
        }
    }

    ; ======================================================================
    ;                           judgment（判断）
    ; ======================================================================
    judgment := Map()
    judge_section := RegExReplace(jsonText, 's).*?"judgment"\s*:\s*\{(.*?)\}.*', '$1')
    for , field in StrSplit(judge_section, ",") {
        if RegExMatch(field, '"(\d+)"\s*:\s*(true|false)', &m)
            judgment[m[1]] := m[2]
    }
    judge_count := judgment.Count
    TrayTip("已读取 " judge_count " 道判断题", "信息", 1)

    y_judge := CalcStartY(single_count, multi_count, fill_count)

    ; ---------- 填写 judgment（可跳过） ----------
    if !skip.Has("judgment") {
        loop judge_count {
            i := A_Index
            row := Ceil(i / 5)
            col := Mod(i - 1, 5) + 1
            x := 995 + 50 * (col - 1)
            y := y_judge + row_gap * (row - 1)

            Sleep 100
            MouseMove(x, y, 0)
            Click()
            Sleep 800

            key := String(i + single_count + multi_count + fill_count)
            if !judgment.Has(key)
                continue
            val := judgment[key]

            if val = "true"
                MouseMove(95, 265, 0)
            else if val = "false"
                MouseMove(95, 315, 0)
            Click()
        }
    }

    ; ======================================================================
    ;                           term_explain
    ; ======================================================================
    term_explain := Map()
    term_section := RegExReplace(jsonText, 's).*?"term_explain"\s*:\s*\{(.*?)\}.*', '$1')
    pos := 1
    while pos := RegExMatch(term_section, '"(\d+)"\s*:\s*"((?:[^"\\]|\\.)*)"', &m, pos) {
        idx := m[1]
        val := JsonUnescape(m[2])
        term_explain[idx] := val
        pos += StrLen(m[0])
    }
    term_count := term_explain.Count
    TrayTip("已读取 " term_count " 道名词解释", "信息", 1)

    y_term := CalcStartY(single_count, multi_count, fill_count, judge_count)

    ; ---------- 填写 term_explain（可跳过） ----------
    if !skip.Has("term_explain") {
        loop term_count {
            i := A_Index
            row := Ceil(i / 5)
            col := Mod(i - 1, 5) + 1
            x := 995 + 50 * (col - 1)
            y := y_term + row_gap * (row - 1)

            Sleep 100
            MouseMove(x, y, 0)
            Click()
            Sleep 800

            key := String(i + single_count + multi_count + fill_count + judge_count)
            if !term_explain.Has(key)
                continue
            ans := term_explain[key]

            MouseMove(180, 350, 0)
            Click()
            Sleep 200
            SendText(ans)
            Sleep 200
        }
    }

    ; ======================================================================
    ;                           short_answer
    ; ======================================================================
    short_answer := Map()
    short_section := RegExReplace(jsonText, 's).*?"short_answer"\s*:\s*\{(.*?)\}.*', '$1')
    pos := 1
    while pos := RegExMatch(short_section, '"(\d+)"\s*:\s*"((?:[^"\\]|\\.)*)"', &m, pos) {
        idx := m[1]
        val := JsonUnescape(m[2])
        short_answer[idx] := val
        pos += StrLen(m[0])
    }
    short_count := short_answer.Count
    TrayTip("已读取 " short_count " 道简答题", "信息", 1)

    y_short := CalcStartY(single_count, multi_count, fill_count, judge_count, term_count)

    ; ---------- 填写 short_answer（可跳过） ----------
    if !skip.Has("short_answer") {
        loop short_count {
            i := A_Index
            row := Ceil(i / 5)
            col := Mod(i - 1, 5) + 1
            x := 995 + 50 * (col - 1)
            y := y_short + row_gap * (row - 1)

            Sleep 100
            MouseMove(x, y, 0)
            Click()
            Sleep 800

            key := String(i + single_count + multi_count + fill_count + judge_count + term_count)
            if !short_answer.Has(key)
                continue
            ans := short_answer[key]

            MouseMove(180, 470, 0)
            Click()
            Sleep 200
            SendText(ans)
            Sleep 200
        }
    }

    TrayTip("所有题型处理完成！（解析完整，动作已按 --skip 执行）", "完成", 1)
}

F12:: ExitApp