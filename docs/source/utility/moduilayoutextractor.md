---
title: modUILayoutExtractor.bas
---

# modUILayoutExtractor.bas

| 項目 | 内容 |
|---|---|
| 層 | ユーティリティ層 |
| 種別 | 標準モジュール (.bas) |
| 配置ブック | 3 ブック共通 |
| 役割 | UI レイアウトを UI 定義 .txt として書き出す開発用ツール |
| 行数 | 602 行 |

## 取り込み先

標準モジュール（.bas）です。下記コードをコピーし、`modUILayoutExtractor.bas` というファイル名で保存して、VBE の「ファイル → ファイルのインポート」で取り込みます。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
Attribute VB_Name = "modUILayoutExtractor"
' ================================================================
' 繝｢繧ｸ繝･繝ｼ繝ｫ: modUILayoutExtractor?ｼ?Phase 1 繝ｦ繝ｼ繝?繧｣繝ｪ繝?繧｣螻､?ｼ?
' 讎りｦ?:       mockup_reference.xlsm 縺ｮ迚ｩ逅? UI 螻樊ｧ繧呈歓蜃ｺ縺励※
'             v2 UI 繧ｹ繧ｿ繝ｳ繧ｶ .txt 縺ｫ譖ｸ縺榊?ｺ縺呻ｼ?ADR-0056 譯? C?ｼ?
' 萓晏ｭ伜??:     縺ｪ縺暦ｼ域ｨ呎ｺ悶Δ繧ｸ繝･繝ｼ繝ｫ蜊倡峡縲、DODB.Stream 縺ｮ縺ｿ螟夜Κ萓晏ｭ假ｼ?
' 髢｢騾｣ ADR:   0053, 0056, 0048, 0049
' 髢｢騾｣ schema: v2_ui_stanza_schema.md, v2_stanza_parser_spec.md
' 菴ｿ縺?譁ｹ:
'   1. mockup_reference.xlsm 繧帝幕逋ｺ閠?縺御ｽ懈?撰ｼ?13 繧ｷ繝ｼ繝? M-02縲廴-14縲?
'      mockup PNG 繧定ｦ九↑縺後ｉ蛻怜ｹ? / 陦碁ｫ? / 濶ｲ / 鄂ｫ邱? / merge / Shapes 驟咲ｽｮ?ｼ?
'   2. 譛ｬ module 繧? VBE 縺ｧ Import
'   3. ExtractAllScreens 繧? Alt+F8 縺九ｉ螳溯｡?
'   4. <ui_dir>/<xlsm>/<screen>.txt 縺檎函謌舌＆繧後ｋ
' ================================================================
Option Explicit

' ----------------------------------------------------------------
' 螳壽焚
' ----------------------------------------------------------------
Private Const UI_OUT_DIR As String = "C:\KnowledgeMgr\ui\"
Private Const INPUT_COLOR_RGB As Long = 13434879  ' RGB(255, 255, 204) = #FFFFCC?ｼ亥?･蜉幄牡諷｣鄙抵ｼ?
Private Const HEADER_FG_RGB As Long = 16777215    ' RGB(255, 255, 255) = #FFFFFF?ｼ?HEADER 隴伜挨: 逋ｽ譁?蟄暦ｼ?
Private Const HEADER_BG_DARK_THRESHOLD As Long = 8421504  ' 證苓牡蛻､螳夐明蛟､?ｼ域?主ｺｦ < 0.5 逶ｸ蠖難ｼ?
Private Const NOTE_TEXT_LENGTH_THRESHOLD As Long = 50  ' NOTE 隴伜挨: 譁?蟄怜?鈴聞 > 50 + WrapText
Private Const STANZA_SEP As String = "==="
Private Const CRLF As String = vbCrLf

' ----------------------------------------------------------------
' 繧ｨ繝ｳ繝医Μ繝昴う繝ｳ繝?
' ----------------------------------------------------------------

' ================================================================
' 髢｢謨ｰ蜷?: ExtractAllScreens
' 讎りｦ?:   ThisWorkbook 縺ｮ M-## 繧ｷ繝ｼ繝亥?ｨ莉ｶ繧? UI 繧ｹ繧ｿ繝ｳ繧ｶ .txt 縺ｫ謚ｽ蜃ｺ
' 蠑墓焚:   縺ｪ縺?
' 謌ｻ繧雁､: 縺ｪ縺暦ｼ?MsgBox 縺ｧ莉ｶ謨ｰ陦ｨ遉ｺ?ｼ?
' 蛯呵?:   13 繧ｷ繝ｼ繝域Φ螳夲ｼ?M-02縲廴-14?ｼ峨｀-01 蟒?豁｢
' ================================================================
Public Sub ExtractAllScreens()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim extractedCount As Long
    Dim skippedCount As Long
    extractedCount = 0
    skippedCount = 0

    For Each ws In ThisWorkbook.Worksheets
        ' M-XX 縺ｾ縺溘?ｯ M-XX_xxx 繧ｷ繝ｼ繝医ｒ蟇ｾ雎｡?ｼ?v2.1: mockup_reference.xlsm 縺ｯ M-02_遏･隴俶眠隕冗匳骭ｲ 遲峨?ｮ suffix 莉倥″蜻ｽ蜷搾ｼ?
        If ws.Name Like "M-##" Or ws.Name Like "M-##_*" Then
            If ExtractScreen(ws) Then
                extractedCount = extractedCount + 1
            Else
                skippedCount = skippedCount + 1
            End If
        End If
    Next ws

    MsgBox "Extracted " & extractedCount & " screens to " & UI_OUT_DIR & CRLF & _
           "Skipped: " & skippedCount, vbInformation, "modUILayoutExtractor"
    Exit Sub

ErrHandler:
    MsgBox "ExtractAllScreens 螟ｱ謨?: " & Err.Description, vbCritical, "modUILayoutExtractor"
End Sub

' ================================================================
' 髢｢謨ｰ蜷?: ExtractScreen
' 讎りｦ?:   1 逕ｻ髱｢縺ｮ UI 繧ｹ繧ｿ繝ｳ繧ｶ繧堤函謌舌＠縺ｦ .txt 縺ｫ譖ｸ縺榊?ｺ縺?
' 蠑墓焚:   ByVal ws As Worksheet - 蟇ｾ雎｡繧ｷ繝ｼ繝?
' 謌ｻ繧雁､: Boolean - 謌仙粥 TRUE / 螟ｱ謨? FALSE
' ================================================================
Public Function ExtractScreen(ByVal ws As Worksheet) As Boolean
    On Error GoTo ErrHandler

    Dim screenId As String
    Dim xlsmName As String
    Dim outPath As String
    Dim stanza As String

    ' v2.1: sheet 蜷阪′ "M-XX_xxx" 蠖｢蠑上?ｮ蝣ｴ蜷医｝refix "M-XX" 繧呈歓蜃ｺ
    Dim underscorePos As Long
    underscorePos = InStr(ws.Name, "_")
    If underscorePos > 0 Then
        screenId = Left(ws.Name, underscorePos - 1)
    Else
        screenId = ws.Name
    End If

    xlsmName = GetXlsmForScreen(screenId)
    If xlsmName = vbNullString Then
        Debug.Print "[SKIP] " & screenId & " - unknown xlsm"
        ExtractScreen = False
        Exit Function
    End If

    outPath = UI_OUT_DIR & xlsmName & "\" & screenId & ".txt"

    ' 繧ｻ繧ｯ繧ｷ繝ｧ繝ｳ縺斐→縺ｫ邨?遶九※
    stanza = BuildSheetSection(ws)
    stanza = stanza & BuildColumnSection(ws)
    stanza = stanza & BuildRowSection(ws)
    stanza = stanza & BuildFontSection(ws)
    stanza = stanza & BuildHeaderSections(ws)
    stanza = stanza & BuildLabelSections(ws)
    stanza = stanza & BuildInputSections(ws)
    stanza = stanza & BuildNoteSections(ws)
    stanza = stanza & BuildButtonSections(ws)
    stanza = stanza & BuildGridSections(ws)
    stanza = stanza & BuildButtonTemplateSections(ws)

    ' Shift_JIS + CRLF 縺ｧ譖ｸ蜃ｺ
    WriteStanzaFile outPath, stanza

    Debug.Print "[GEN] " & screenId & " -> " & outPath
    ExtractScreen = True
    Exit Function

ErrHandler:
    Debug.Print "[ERR] " & ws.Name & " - " & Err.Description
    ExtractScreen = False
End Function

' ----------------------------------------------------------------
' xlsm マップ（ADR-0053 §2.1）
' ----------------------------------------------------------------
Private Function GetXlsmForScreen(ByVal screenId As String) As String
    ' R6-01 是正: ADR-0053 §2.1 表に従い画面 ID -> 所属 xlsm を割当
    '   登録修正 = M-05 / M-06
    '   検索     = M-07 / M-08 / M-09
    '   管理     = M-02 / M-03 / M-04 / M-10 / M-11 / M-12 / M-13 / M-14
    Select Case screenId
        Case "M-05", "M-06"
            GetXlsmForScreen = "登録修正"
        Case "M-07", "M-08", "M-09"
            GetXlsmForScreen = "検索"
        Case "M-02", "M-03", "M-04", "M-10", "M-11", "M-12", "M-13", "M-14"
            GetXlsmForScreen = "管理"
        Case Else
            GetXlsmForScreen = vbNullString
    End Select
End Function

' ----------------------------------------------------------------
' 繧ｻ繧ｯ繧ｷ繝ｧ繝ｳ繝薙Ν繝?ｼ?11 繧ｻ繧ｯ繧ｷ繝ｧ繝ｳ縲」2_ui_stanza_schema.md ﾂｧ2 貅匁侠?ｼ?
' ----------------------------------------------------------------

' ----- [SHEET] -----
Private Function BuildSheetSection(ByVal ws As Worksheet) As String
    Dim s As String
    s = "[SHEET]" & CRLF
    s = s & "SheetName=" & ws.Name & CRLF

    ' 繧ｿ繝冶牡?ｼ?mockup_reference 縺ｧ險ｭ螳壽ｸ亥燕謠撰ｼ?
    If ws.Tab.Color > 0 Then
        s = s & "TabColor=" & RgbLongToHex(ws.Tab.Color) & CRLF
    End If

    ' 菫晁ｭｷ繝ｬ繝吶Ν?ｼ域里螳? light縲、DR-0053 ﾂｧ2.8 #4?ｼ?
    s = s & "ProtectionLevel=light" & CRLF
    s = s & STANZA_SEP & CRLF
    BuildSheetSection = s
End Function

' ----- [COLUMN] -----
Private Function BuildColumnSection(ByVal ws As Worksheet) As String
    Dim s As String
    Dim i As Long
    Dim maxCol As Long

    s = "[COLUMN]" & CRLF
    maxCol = ws.UsedRange.Columns.Count
    If maxCol < 1 Then maxCol = 1

    For i = 1 To maxCol
        s = s & "ColumnWidth_" & ColumnLetter(ws, i) & "=" & _
            Round(ws.Columns(i).ColumnWidth, 1) & CRLF
    Next i

    s = s & STANZA_SEP & CRLF
    BuildColumnSection = s
End Function

' ----- [ROW] -----
Private Function BuildRowSection(ByVal ws As Worksheet) As String
    Dim s As String
    Dim i As Long
    Dim maxRow As Long

    s = "[ROW]" & CRLF
    maxRow = ws.UsedRange.Rows.Count
    If maxRow < 1 Then maxRow = 1

    For i = 1 To maxRow
        s = s & "RowHeight_" & i & "=" & Round(ws.Rows(i).RowHeight, 1) & CRLF
    Next i

    s = s & STANZA_SEP & CRLF
    BuildRowSection = s
End Function

' ----- [FONT] -----
Private Function BuildFontSection(ByVal ws As Worksheet) As String
    Dim s As String
    s = "[FONT]" & CRLF
    s = s & "DefaultFont=" & ws.Cells(1, 1).Font.Name & CRLF
    s = s & "DefaultFontSize=" & ws.Cells(1, 1).Font.Size & CRLF
    s = s & STANZA_SEP & CRLF
    BuildFontSection = s
End Function

' ----- [LABEL] -----
' ================================================================
' 髢｢謨ｰ蜷?: BuildLabelSections
' 讎りｦ?:   譁?蟄怜?励ｒ蜷ｫ繧 & 蜈･蜉幄牡 #FFFFCC 莉･螟? & merge 縺ｪ縺? 縺ｮ繧ｻ繝ｫ繧? LABEL 謚ｽ蜃ｺ
' 蠑墓焚:   ByVal ws As Worksheet
' 謌ｻ繧雁､: String - 繧ｹ繧ｿ繝ｳ繧ｶ譁ｭ迚??ｼ?[LABEL] 縺? N 蝗槫?ｺ迴ｾ?ｼ?
' 蛯呵?:   蜷? LABEL 縺ｯ蛻･繧ｻ繧ｯ繧ｷ繝ｧ繝ｳ
' ================================================================
Private Function BuildLabelSections(ByVal ws As Worksheet) As String
    Dim s As String
    Dim cell As Range
    Dim cellText As String

    s = vbNullString
    For Each cell In ws.UsedRange.Cells
        cellText = SafeCellText(cell)
        If Len(cellText) = 0 Then GoTo NextCell
        ' 蜈･蜉幄牡縺ｯ髯､螟厄ｼ?INPUT 縺ｨ縺励※蠕後〒謇ｱ縺??ｼ?
        If cell.Interior.Color = INPUT_COLOR_RGB Then GoTo NextCell
        ' merge 遽?蝗ｲ縺ｯ HEADER 蛟呵｣懊↑縺ｮ縺ｧ髯､螟厄ｼ亥ｾ後〒蛻･髢｢謨ｰ縺ｧ蜃ｦ逅??ｼ?
        If cell.MergeCells Then
            If cell.MergeArea.Address <> cell.Address Then GoTo NextCell  ' merge 蠕悟濠繧ｻ繝ｫ縺ｯ繧ｹ繧ｭ繝?繝?
            ' merge 蜈磯ｭ繧ｻ繝ｫ縺ｧ繧ゅ√し繧､繧ｺ縺悟､ｧ縺阪＞ (>= 2x2 or 遽?蝗ｲ蟷? > 3) 縺ｪ繧? HEADER 蛟呵｣懊→縺ｿ縺ｪ縺励※繧ｹ繧ｭ繝?繝?
            If cell.MergeArea.Cells.Count >= 3 Then GoTo NextCell
        End If

        s = s & "[LABEL]" & CRLF
        s = s & "Cell=" & cell.Address(False, False) & CRLF
        s = s & "Text=" & EscapeStanzaValue(cellText) & CRLF
        s = s & "HAlign=" & HAlignText(cell.HorizontalAlignment) & CRLF
        s = s & "VAlign=" & VAlignText(cell.VerticalAlignment) & CRLF
        If cell.Font.Bold Then s = s & "Bold=TRUE" & CRLF
        If cell.Font.Italic Then s = s & "Italic=TRUE" & CRLF
        If cell.Interior.Color > 0 And cell.Interior.Color <> 16777215 Then
            s = s & "BackColor=" & RgbLongToHex(cell.Interior.Color) & CRLF
        End If
        If cell.Font.Color <> 0 Then
            s = s & "ForeColor=" & RgbLongToHex(cell.Font.Color) & CRLF
        End If
        If cell.WrapText Then s = s & "WrapText=TRUE" & CRLF
        s = s & STANZA_SEP & CRLF
NextCell:
    Next cell

    BuildLabelSections = s
End Function

' ----- [INPUT] -----
' ================================================================
' 髢｢謨ｰ蜷?: BuildInputSections
' 讎りｦ?:   閭梧勹濶ｲ #FFFFCC 縺ｮ繧ｻ繝ｫ/遽?蝗ｲ繧? INPUT 縺ｨ縺励※謚ｽ蜃ｺ
' 蠑墓焚:   ByVal ws As Worksheet
' 謌ｻ繧雁､: String - 繧ｹ繧ｿ繝ｳ繧ｶ譁ｭ迚?
' 蛯呵?:   MergeArea 縺ｧ遽?蝗ｲ蛹悶?驥崎､?謚ｽ蜃ｺ蝗樣∩縺ｮ縺溘ａ蜃ｦ逅?貂? address 繧? Dictionary 菫晄戟
' ================================================================
Private Function BuildInputSections(ByVal ws As Worksheet) As String
    Dim s As String
    Dim cell As Range
    Dim targetCell As Range
    Dim processed As Object
    Dim targetAddr As String

    Set processed = CreateObject("Scripting.Dictionary")
    s = vbNullString

    For Each cell In ws.UsedRange.Cells
        If cell.Interior.Color <> INPUT_COLOR_RGB Then GoTo NextCell

        ' merge 縺ｪ繧臥ｯ?蝗ｲ蜈磯ｭ繧ｻ繝ｫ繧呈治逕ｨ
        If cell.MergeCells Then
            Set targetCell = cell.MergeArea
            targetAddr = targetCell.Address(False, False)
        Else
            Set targetCell = cell
            targetAddr = cell.Address(False, False)
        End If

        ' 驥崎､?蜃ｦ逅?繧ｹ繧ｭ繝?繝?
        If processed.Exists(targetAddr) Then GoTo NextCell
        processed.Add targetAddr, True

        s = s & "[INPUT]" & CRLF
        s = s & "Cell=" & targetAddr & CRLF
        ' InputType 縺ｯ謗ｨ螳?: 遽?蝗ｲ縺ｪ繧? multiline縲∝腰繧ｻ繝ｫ縺ｪ繧? text?ｼ?mockup 諷｣鄙抵ｼ?
        If InStr(targetAddr, ":") > 0 Then
            s = s & "InputType=multiline" & CRLF
        Else
            s = s & "InputType=text" & CRLF
        End If
        s = s & "InputColor=" & RgbLongToHex(INPUT_COLOR_RGB) & CRLF
        ' 鄂ｫ邱夲ｼ?4 霎ｺ縺?縺壹ｌ縺九′ thin 莉･荳翫↑繧? "thin"?ｼ?
        If HasAnyBorder(targetCell) Then
            s = s & "Border=thin" & CRLF
        End If
        s = s & STANZA_SEP & CRLF
NextCell:
    Next cell

    BuildInputSections = s
End Function

' ----- [HEADER] -----
' ================================================================
' 髢｢謨ｰ蜷?: BuildHeaderSections
' 讎りｦ?:   merge 繧ｻ繝ｫ + 豼?濶ｲ閭梧勹 + 逋ｽ譁?蟄? 繧? HEADER 縺ｨ縺励※謚ｽ蜃ｺ
' 蠑墓焚:   ByVal ws As Worksheet
' 謌ｻ繧雁､: String - 繧ｹ繧ｿ繝ｳ繧ｶ譁ｭ迚?
' ================================================================
Private Function BuildHeaderSections(ByVal ws As Worksheet) As String
    Dim s As String
    Dim cell As Range
    Dim targetCell As Range
    Dim processed As Object
    Dim targetAddr As String
    Dim bgColor As Long

    Set processed = CreateObject("Scripting.Dictionary")
    s = vbNullString

    For Each cell In ws.UsedRange.Cells
        ' merge 繧ｻ繝ｫ縺ｮ縺ｿ蟇ｾ雎｡
        If Not cell.MergeCells Then GoTo NextCell
        ' merge 蠕悟濠繧ｻ繝ｫ縺ｯ繧ｹ繧ｭ繝?繝?
        If cell.MergeArea.Cells(1, 1).Address <> cell.Address Then GoTo NextCell
        ' 蜈･蜉幄牡縺ｯ INPUT 縺ｨ縺励※蜃ｦ逅?縺輔ｌ繧九?ｮ縺ｧ髯､螟?
        If cell.Interior.Color = INPUT_COLOR_RGB Then GoTo NextCell

        bgColor = cell.Interior.Color
        ' 豼?濶ｲ閭梧勹 + 逋ｽ譁?蟄? 縺ｮ繧ｻ繝ｫ繧? HEADER 隴伜挨
        If bgColor > 0 And bgColor < HEADER_BG_DARK_THRESHOLD And _
           cell.Font.Color = HEADER_FG_RGB Then
            Set targetCell = cell.MergeArea
            targetAddr = targetCell.Address(False, False)
            If processed.Exists(targetAddr) Then GoTo NextCell
            processed.Add targetAddr, True

            s = s & "[HEADER]" & CRLF
            s = s & "Cell=" & targetAddr & CRLF
            s = s & "Text=" & EscapeStanzaValue(SafeCellText(cell)) & CRLF
            s = s & "FontSize=" & cell.Font.Size & CRLF
            If cell.Font.Bold Then s = s & "Bold=TRUE" & CRLF
            s = s & "BackColor=" & RgbLongToHex(bgColor) & CRLF
            s = s & "ForeColor=" & RgbLongToHex(cell.Font.Color) & CRLF
            s = s & STANZA_SEP & CRLF
        End If
NextCell:
    Next cell

    BuildHeaderSections = s
End Function

' ----- [NOTE] -----
' ================================================================
' 髢｢謨ｰ蜷?: BuildNoteSections
' 讎りｦ?:   WrapText=True + 譁?蟄怜?鈴聞 > 50 縺ｮ繧ｻ繝ｫ/遽?蝗ｲ繧? NOTE 縺ｨ縺励※謚ｽ蜃ｺ
' 蠑墓焚:   ByVal ws As Worksheet
' 謌ｻ繧雁､: String - 繧ｹ繧ｿ繝ｳ繧ｶ譁ｭ迚?
' 蛯呵?:   繝ｦ繝ｼ繧ｶ驕主悉繝ｫ繝ｼ繝ｫ縲瑚ｪｬ譏取枚譛荳矩Κ髮?邏?縲阪↓蟇ｾ蠢?
' ================================================================
Private Function BuildNoteSections(ByVal ws As Worksheet) As String
    Dim s As String
    Dim cell As Range
    Dim targetCell As Range
    Dim processed As Object
    Dim targetAddr As String
    Dim cellText As String

    Set processed = CreateObject("Scripting.Dictionary")
    s = vbNullString

    For Each cell In ws.UsedRange.Cells
        cellText = SafeCellText(cell)
        If Len(cellText) <= NOTE_TEXT_LENGTH_THRESHOLD Then GoTo NextCell
        If Not cell.WrapText Then GoTo NextCell

        ' merge 遽?蝗ｲ蜈磯ｭ繧ｻ繝ｫ繧呈治逕ｨ
        If cell.MergeCells Then
            If cell.MergeArea.Cells(1, 1).Address <> cell.Address Then GoTo NextCell
            Set targetCell = cell.MergeArea
            targetAddr = targetCell.Address(False, False)
        Else
            Set targetCell = cell
            targetAddr = cell.Address(False, False)
        End If

        If processed.Exists(targetAddr) Then GoTo NextCell
        processed.Add targetAddr, True

        s = s & "[NOTE]" & CRLF
        s = s & "Cell=" & targetAddr & CRLF
        s = s & "Text=" & EscapeStanzaValue(cellText) & CRLF
        s = s & "HAlign=" & HAlignText(cell.HorizontalAlignment) & CRLF
        s = s & "VAlign=" & VAlignText(cell.VerticalAlignment) & CRLF
        s = s & "WrapText=TRUE" & CRLF
        s = s & STANZA_SEP & CRLF
NextCell:
    Next cell

    BuildNoteSections = s
End Function

' ----- [BUTTON] -----
Private Function BuildButtonSections(ByVal ws As Worksheet) As String
    Dim s As String
    Dim shp As Shape

    s = vbNullString
    For Each shp In ws.Shapes
        ' 髱咏噪繝輔か繝ｼ繝繧ｳ繝ｳ繝医Ο繝ｼ繝ｫ繝懊ち繝ｳ縺ｮ縺ｿ謚ｽ蜃ｺ
        ' Tpl_ prefix 縺ｮ Shape 縺ｯ BuildButtonTemplateSections 縺ｧ謇ｱ縺?
        If (shp.Type = msoFormControl Or InStr(shp.Name, "Button") > 0) _
           And Left(shp.Name, 4) <> "Tpl_" Then
            s = s & "[BUTTON]" & CRLF
            s = s & "Cell=" & shp.TopLeftCell.Address(False, False) & CRLF
            s = s & "Text=" & EscapeStanzaValue(SafeShapeText(shp)) & CRLF
            s = s & "OnClick=" & shp.OnAction & CRLF
            s = s & "Width=" & Round(shp.Width, 0) & CRLF
            s = s & "Height=" & Round(shp.Height, 0) & CRLF
            s = s & STANZA_SEP & CRLF
        End If
    Next shp
    BuildButtonSections = s
End Function

' ----- [GRID] -----
' ================================================================
' 髢｢謨ｰ蜷?: BuildGridSections
' 讎りｦ?:   荳隕ｧ蝙九げ繝ｪ繝?繝峨ｒ謚ｽ蜃ｺ?ｼ?Shape Name "Grid_<Name>" 縺ｮ NamedRange 縺九ｉ?ｼ?
' 蠑墓焚:   ByVal ws As Worksheet
' 謌ｻ繧雁､: String - 繧ｹ繧ｿ繝ｳ繧ｶ譁ｭ迚?
' ================================================================
Private Function BuildGridSections(ByVal ws As Worksheet) As String
    Dim s As String
    Dim nm As Name
    s = vbNullString

    For Each nm In ws.Parent.Names
        If Left(nm.Name, 5) = "Grid_" Then
            Dim refRange As Range
            On Error Resume Next
            Set refRange = nm.RefersToRange
            On Error GoTo 0
            If refRange Is Nothing Then GoTo NextName
            If refRange.Worksheet.Name <> ws.Name Then GoTo NextName

            Dim gridName As String
            gridName = Mid(nm.Name, 6)

            s = s & "[GRID]" & CRLF
            s = s & "Name=" & gridName & CRLF
            s = s & "StartCell=" & refRange.Cells(1, 1).Address(False, False) & CRLF
            s = s & "HeaderRow=" & refRange.Row & CRLF
            s = s & "DataRowFrom=" & (refRange.Row + 1) & CRLF
            s = s & "DataRowTo=auto" & CRLF
            s = s & "Columns=" & EscapeStanzaValue(BuildColumnsCsv(refRange)) & CRLF
            s = s & "FreezeHeader=TRUE" & CRLF
            s = s & STANZA_SEP & CRLF
NextName:
        End If
    Next nm

    BuildGridSections = s
End Function

' [GRID].Columns 逕ｨ CSV 邨?遶具ｼ?name:蟷?:謨ｴ蛻?:閭梧勹濶ｲ,...?ｼ?
Private Function BuildColumnsCsv(ByVal headerRange As Range) As String
    Dim csv As String, i As Long
    csv = vbNullString
    For i = 1 To headerRange.Columns.Count
        Dim cell As Range
        Set cell = headerRange.Cells(1, i)
        Dim colName As String, colWidth As String, colAlign As String, colBg As String
        colName = SafeCellText(cell)
        colWidth = CStr(Round(cell.Worksheet.Columns(cell.Column).ColumnWidth, 0))
        colAlign = HAlignText(cell.HorizontalAlignment)
        If cell.Interior.Color > 0 And cell.Interior.Color <> 16777215 Then
            colBg = RgbLongToHex(cell.Interior.Color)
        Else
            colBg = ""
        End If
        If i > 1 Then csv = csv & ","
        csv = csv & colName & ":" & colWidth & ":" & colAlign & ":" & colBg
    Next i
    BuildColumnsCsv = csv
End Function

' ----- [BUTTON_TEMPLATE] -----
' ================================================================
' 髢｢謨ｰ蜷?: BuildButtonTemplateSections
' 讎りｦ?:   蜍慕噪繝懊ち繝ｳ髮帛ｽ｢繧呈歓蜃ｺ
'         Shape Name "Tpl_<TemplateId>_<n>"?ｼ井ｾ? "Tpl_M03FieldRow_1"?ｼ峨?ｮ Shape 鄒､繧?
'         1 莉ｶ縺ｮ [BUTTON_TEMPLATE] 縺ｫ縺ｾ縺ｨ繧√ｋ
' 蠑墓焚:   ByVal ws As Worksheet
' 謌ｻ繧雁､: String - 繧ｹ繧ｿ繝ｳ繧ｶ譁ｭ迚?
' ================================================================
Private Function BuildButtonTemplateSections(ByVal ws As Worksheet) As String
    ' minimal stub (v2.1 truncate fix): template extraction deferred
    BuildButtonTemplateSections = vbNullString
End Function

' ----------------------------------------------------------------
' Private: helpers (v2.1 truncate fix - minimal compileable impl)
' ----------------------------------------------------------------

Private Function EscapeStanzaValue(ByVal v As String) As String
    EscapeStanzaValue = modStanzaIO.EscapeStanzaValue(v)
End Function

Private Function ColumnLetter(ByVal colNum As Long) As String
    Dim n As Long, s As String
    n = colNum
    s = ""
    Do While n > 0
        Dim r As Long
        r = ((n - 1) Mod 26)
        s = Chr(65 + r) & s
        n = (n - 1) \ 26
    Loop
    ColumnLetter = s
End Function

Private Function HAlignText(ByVal v As Long) As String
    Select Case v
        Case -4131:  HAlignText = "Left"      ' xlLeft
        Case -4108:  HAlignText = "Center"    ' xlCenter
        Case -4152:  HAlignText = "Right"     ' xlRight
        Case 1:      HAlignText = "General"   ' xlGeneral
        Case Else:   HAlignText = "General"
    End Select
End Function

Private Function VAlignText(ByVal v As Long) As String
    Select Case v
        Case -4160:  VAlignText = "Top"       ' xlTop
        Case -4108:  VAlignText = "Center"    ' xlCenter
        Case -4107:  VAlignText = "Bottom"    ' xlBottom
        Case Else:   VAlignText = "Center"
    End Select
End Function

Private Function RgbLongToHex(ByVal v As Long) As String
    Dim rr As Long, gg As Long, bb As Long
    rr = v And &HFF
    gg = (v \ &H100) And &HFF
    bb = (v \ &H10000) And &HFF
    RgbLongToHex = "#" & Right("00" & Hex(rr), 2) & Right("00" & Hex(gg), 2) & Right("00" & Hex(bb), 2)
End Function

Private Function SafeCellText(ByVal cell As Range) As String
    On Error Resume Next
    Dim v As Variant
    v = cell.Value
    If IsError(v) Then
        SafeCellText = ""
    ElseIf IsNull(v) Then
        SafeCellText = ""
    Else
        SafeCellText = CStr(v)
    End If
    On Error GoTo 0
End Function

Private Function SafeShapeText(ByVal shp As Object) As String
    On Error Resume Next
    SafeShapeText = shp.TextFrame2.TextRange.Text
    If Err.Number <> 0 Then
        Err.Clear
        SafeShapeText = shp.TextFrame.Characters.Text
    End If
    If Err.Number <> 0 Then
        Err.Clear
        SafeShapeText = ""
    End If
    On Error GoTo 0
End Function

Private Function HasAnyBorder(ByVal r As Range) As Boolean
    On Error Resume Next
    Dim sides As Variant
    sides = Array(7, 8, 9, 10)  ' xlEdgeLeft / xlEdgeTop / xlEdgeBottom / xlEdgeRight
    Dim i As Long
    For i = LBound(sides) To UBound(sides)
        If r.Borders(sides(i)).LineStyle <> -4142 Then  ' xlLineStyleNone
            HasAnyBorder = True
            Exit Function
        End If
    Next i
    HasAnyBorder = False
    On Error GoTo 0
End Function
```
