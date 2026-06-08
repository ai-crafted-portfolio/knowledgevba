---
title: modButtonWiring.bas
description: modButtonWiring.bas のソースコード（コピペ用）
---

# modButtonWiring.bas

**配置先**: `共通モジュール (3 ブック全て)` 用の VBA モジュール  
**種類**: 標準モジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\`
- ファイル名: `modButtonWiring.bas`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。

---

## ソースコード

```vb
Attribute VB_Name = "modButtonWiring"
' ============================================================
' modButtonWiring (�ｿｽ�ｿｽ�ｿｽ[�ｿｽe�ｿｽB�ｿｽ�ｿｽ�ｿｽe�ｿｽB�ｿｽw, v2.1)
' �ｿｽ�ｿｽ�ｿｽ�ｿｽ:
'   UI �ｿｽ�ｿｽ`�ｿｽX�ｿｽ^�ｿｽ�ｿｽ�ｿｽU�ｿｽ�ｿｽ [BUTTON] / [BUTTON_TEMPLATE] �ｿｽZ�ｿｽN�ｿｽV�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ
'   �ｿｽt�ｿｽH�ｿｽ[�ｿｽ�ｿｽ�ｿｽR�ｿｽ�ｿｽ�ｿｽg�ｿｽ�ｿｽ�ｿｽ[�ｿｽ�ｿｽ�ｿｽ{�ｿｽ^�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽs�ｿｽ�ｿｽ�ｿｽﾉ趣ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ OnAction �ｿｽ�ｿｽ
'   �ｿｽ�ｿｽ�ｿｽ闢厄ｿｽﾄゑｿｽB�ｿｽe�ｿｽX�ｿｽg�ｿｽ�ｿｽp�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ E2E_AutoWiring �ｿｽ�ｿｽ
'   Shapes.AddFormControl �ｿｽz�ｿｽu�ｿｽ�ｿｽ�ｿｽW�ｿｽb�ｿｽN�ｿｽｻ品�ｿｽ�ｿｽ�ｿｽW�ｿｽ�ｿｽ�ｿｽ[�ｿｽ�ｿｽ�ｿｽﾖ擾ｿｽ�ｿｽi�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽB
' �ｿｽﾄ出�ｿｽ�ｿｽ:
'   modUILoader.ApplyButton / modUILoader.ApplyButtonTemplate
' �ｿｽﾋ托ｿｽ�ｿｽ�ｿｽ:
'   ClsStanzaSection (UI �ｿｽ�ｿｽ`�ｿｽZ�ｿｽN�ｿｽV�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽﾌ値�ｿｽI�ｿｽu�ｿｽW�ｿｽF�ｿｽN�ｿｽg)
' �ｿｽﾝ計�ｿｽ�ｿｽ�ｿｽ�ｿｽ:
'   - �ｿｽ{�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽﾍ厄ｿｽ�ｿｽu�ｿｽb�ｿｽN�ｿｽI�ｿｽ[�ｿｽv�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽAclsSetupOrchestrator.RunFullSetup
'     -> ApplyUiToSheet �ｿｽo�ｿｽR�ｿｽﾅ再趣ｿｽ�ｿｽs�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽB�ｿｽ�ｿｽ�ｿｽO�ｿｽ�ｿｽ
'     clsSheetRenderer.ClearScreen �ｿｽ�ｿｽ�ｿｽS Shape �ｿｽ�ｿｽ�ｿｽ尞懶ｿｽ�ｿｽ�ｿｽ驍ｽ�ｿｽﾟ、
'     �ｿｽﾄ撰ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽﾄゑｿｽ�ｿｽd�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽﾈゑｿｽ (idempotent)�ｿｽB
'   - ClearScreen �ｿｽｺゑｿｽﾈゑｿｽ�ｿｽﾄ適�ｿｽp�ｿｽﾅゑｿｽ�ｿｽc�ｿｽ[�ｿｽ�ｿｽ�ｿｽo�ｿｽﾈゑｿｽ�ｿｽ謔､�ｿｽA�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽO�ｿｽ�ｿｽ
'     �ｿｽ�ｿｽ�ｿｽ�ｿｽ Shape �ｿｽ�ｿｽ�ｿｽ尞懶ｿｽ�ｿｽ�ｿｽ�ｿｽB
'   - OnAction �ｿｽﾍ厄ｿｽ�ｿｽC�ｿｽ�ｿｽ�ｿｽ}�ｿｽN�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽﾅ設定す�ｿｽ�ｿｽBExcel �ｿｽﾍ設抵ｿｽ�ｿｽﾌ読み戻ゑｿｽ
'     �ｿｽ�ｿｽ�ｿｽﾉブ�ｿｽb�ｿｽN�ｿｽ�ｿｽ�ｿｽ�ｿｽO�ｿｽu (Book.xlsm!Macro �ｿｽ`�ｿｽ�ｿｽ) �ｿｽ�ｿｽ�ｿｽ驍ｪ�ｿｽA�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ
'     Excel �ｿｽﾌ仕�ｿｽl�ｿｽﾅゑｿｽ�ｿｽ�ｿｽs�ｿｽ�合�ｿｽﾅはなゑｿｽ�ｿｽB
'   - [BUTTON].Width / Height �ｿｽﾍピ�ｿｽN�ｿｽZ�ｿｽ�ｿｽ�ｿｽ�ｿｽ (schema 3.7)�ｿｽB
'     AddFormControl �ｿｽﾍポ�ｿｽC�ｿｽ�ｿｽ�ｿｽg�ｿｽP�ｿｽﾊのゑｿｽ�ｿｽ�ｿｽ px -> pt �ｿｽ�ｿｽ�ｿｽZ�ｿｽ�ｿｽ�ｿｽ�ｿｽB
' ASCII-only inside VBA string literals (CP932 mojibake avoidance).
' ============================================================
Option Explicit

Private Const MOD_NAME As String = "modButtonWiring"
Private Const FC_BUTTON As Long = 0           ' xlButtonControl
Private Const HA_CENTER As Long = -4108       ' xlHAlignCenter
Private Const PX_PER_PT As Double = 0.75      ' 96dpi: 1px = 0.75pt
Private Const BTN_MIN_W As Double = 24#       ' �ｿｽ{�ｿｽ^�ｿｽ�ｿｽ�ｿｽﾅ擾ｿｽ�ｿｽ�ｿｽ (pt)
Private Const BTN_MIN_H As Double = 12#       ' �ｿｽ{�ｿｽ^�ｿｽ�ｿｽ�ｿｽﾅ擾ｿｽ�ｿｽ�ｿｽ (pt)
Private Const DEFAULT_ROW_COUNT As Long = 20  ' DataRowTo=auto �ｿｽ�ｿｽ�ｿｽﾌ暫�ｿｽ�ｿｽs�ｿｽ�ｿｽ
Private Const NAME_PREFIX As String = "btnwire_"

' ============================================================
' �ｿｽﾖ撰ｿｽ�ｿｽ�ｿｽ: PlaceButton
' �ｿｽT�ｿｽv:   [BUTTON] �ｿｽX�ｿｽ^�ｿｽ�ｿｽ�ｿｽU 1 �ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽﾃ的�ｿｽ{�ｿｽ^�ｿｽ�ｿｽ�ｿｽｶ撰ｿｽ�ｿｽ�ｿｽ OnAction
'         �ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ闢厄ｿｽﾄゑｿｽB
' �ｿｽ�ｿｽ�ｿｽ�ｿｽ:   ws           - �ｿｽz�ｿｽu�ｿｽ�ｿｽV�ｿｽ[�ｿｽg
'         cellAddr     - �ｿｽA�ｿｽ�ｿｽ�ｿｽJ�ｿｽ[�ｿｽZ�ｿｽ�ｿｽ (A1 �ｿｽ\�ｿｽL�ｿｽA�ｿｽﾍ囲なら左�ｿｽ�ｿｽ)
'         caption      - �ｿｽ{�ｿｽ^�ｿｽ�ｿｽ�ｿｽL�ｿｽ�ｿｽ�ｿｽv�ｿｽV�ｿｽ�ｿｽ�ｿｽ�ｿｽ
'         onClick      - �ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ}�ｿｽN�ｿｽ�ｿｽ�ｿｽ�ｿｽ (�ｿｽ�ｿｽ�ｿｽC�ｿｽ�ｿｽ)
'         widthPx      - �ｿｽ�ｿｽ (�ｿｽs�ｿｽN�ｿｽZ�ｿｽ�ｿｽ)
'         heightPx     - �ｿｽ�ｿｽ�ｿｽ�ｿｽ (�ｿｽs�ｿｽN�ｿｽZ�ｿｽ�ｿｽ)
'         foreColorHex - �ｿｽL�ｿｽ�ｿｽ�ｿｽv�ｿｽV�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽF (#RRGGBB�ｿｽA�ｿｽ�ｿｽ�ｿｽ)
' ============================================================
Public Sub PlaceButton(ByVal ws As Worksheet, ByVal cellAddr As String, _
                       ByVal caption As String, ByVal onClick As String, _
                       ByVal widthPx As Double, ByVal heightPx As Double, _
                       ByVal foreColorHex As String, _
                       Optional ByVal backColorLong As Long = -1)
    On Error GoTo ErrHandler
    If ws Is Nothing Then Exit Sub
    If Len(cellAddr) = 0 Then Exit Sub

    Dim anchor As Range
    Set anchor = ws.Range(cellAddr).Cells(1, 1)

    Dim btnName As String
    btnName = NAME_PREFIX & "s_" & SafeToken(cellAddr)
    DeleteShapeByName ws, btnName

    Dim wPt As Double, hPt As Double
    wPt = PxToPt(widthPx, BTN_MIN_W, 120#)
    hPt = PxToPt(heightPx, BTN_MIN_H, 24#)

    Dim shp As Shape
    If backColorLong >= 0 Then
        ' Phase R-3-ﾏ�-M03-style (蟾ｮ3): 濶ｲ莉倥″繝懊ち繝ｳ縺ｯ Shape(隗剃ｸｸ遏ｩ蠖｢)縺ｧ謠冗判縲�
        ' Form Control 繝懊ち繝ｳ縺ｯ蝪励ｊ縺､縺ｶ縺嶺ｸ榊庄縺ｮ縺溘ａ縲。ackColor 謖�螳壽凾縺ｮ縺ｿ Shape 蛹悶�
        Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, anchor.Left, anchor.Top, wPt, hPt)
        shp.Name = btnName
        shp.Fill.ForeColor.RGB = backColorLong
        shp.Line.Visible = msoFalse
        If Len(onClick) > 0 Then shp.OnAction = onClick
        On Error Resume Next
        With shp.TextFrame2.TextRange
            .Text = caption
            .Font.Bold = msoTrue
            .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        End With
        shp.TextFrame2.HorizontalAnchor = 2          ' msoAnchorCenter
        shp.TextFrame2.VerticalAnchor = 3            ' msoAnchorMiddle
        ' Phase R-3-psi-M03-style-beta: 6-char captions ("field add" etc.) wrapped+clipped
        ' in the short button. Disable wrap + shrink internal margins so the label fits 1 line.
        shp.TextFrame2.WordWrap = msoFalse
        shp.TextFrame2.MarginLeft = 1
        shp.TextFrame2.MarginRight = 1
        shp.TextFrame2.MarginTop = 0
        shp.TextFrame2.MarginBottom = 0
        On Error GoTo ErrHandler
    Else
        Set shp = ws.Shapes.AddFormControl(FC_BUTTON, anchor.Left, anchor.Top, wPt, hPt)
        shp.Name = btnName
        If Len(onClick) > 0 Then shp.OnAction = onClick
        ApplyCaption ws, shp, caption, foreColorHex
    End If
    Exit Sub

ErrHandler:
    Debug.Print "[" & MOD_NAME & ".PlaceButton] " & cellAddr & ": " & _
                Err.Number & " " & Err.Description
End Sub

' ============================================================
' �ｿｽﾖ撰ｿｽ�ｿｽ�ｿｽ: PlaceTemplateButtons
' �ｿｽT�ｿｽv:   [BUTTON_TEMPLATE] �ｿｽX�ｿｽ^�ｿｽ�ｿｽ�ｿｽU 1 �ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽA�ｿｽﾎ会ｿｽ [GRID] �ｿｽ�ｿｽ
'         �ｿｽf�ｿｽ[�ｿｽ^�ｿｽs�ｿｽﾍ囲ぶゑｿｽ per-row �ｿｽ{�ｿｽ^�ｿｽ�ｿｽ�ｿｽQ�ｿｽｶ撰ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽB
' �ｿｽ�ｿｽ�ｿｽ�ｿｽ:   ws      - �ｿｽz�ｿｽu�ｿｽ�ｿｽV�ｿｽ[�ｿｽg
'         sec     - [BUTTON_TEMPLATE] �ｿｽZ�ｿｽN�ｿｽV�ｿｽ�ｿｽ�ｿｽ�ｿｽ
'         gridSec - GridRef �ｿｽ�ｿｽ�ｿｽw�ｿｽ�ｿｽ [GRID] �ｿｽZ�ｿｽN�ｿｽV�ｿｽ�ｿｽ�ｿｽ�ｿｽ (Nothing �ｿｽ�ｿｽ)
' �ｿｽ�ｿｽ�ｿｽl:   gridSec �ｿｽ�ｿｽ Nothing �ｿｽﾜゑｿｽ�ｿｽ�ｿｽ DataRowTo=auto �ｿｽﾌ場合�ｿｽ�ｿｽ
'         RowRangeFrom �ｿｽ�ｿｽ�ｿｽ�ｿｽ DEFAULT_ROW_COUNT �ｿｽs�ｿｽﾔん生撰ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽB
' ============================================================
Public Sub PlaceTemplateButtons(ByVal ws As Worksheet, _
                                ByVal sec As ClsStanzaSection, _
                                ByVal gridSec As ClsStanzaSection)
    On Error GoTo ErrHandler
    If ws Is Nothing Then Exit Sub
    If sec Is Nothing Then Exit Sub

    Dim colOffset As String
    colOffset = sec.GetValue("ColumnOffset")
    If Len(colOffset) = 0 Then Exit Sub

    Dim rowFrom As Long
    rowFrom = ToLong(sec.GetValue("RowRangeFrom"), 0)
    If rowFrom <= 0 Then Exit Sub

    Dim btnCount As Long
    btnCount = ToLong(sec.GetValue("ButtonCount"), 0)
    If btnCount <= 0 Then Exit Sub

    Dim templateId As String
    templateId = sec.GetValue("TemplateId")
    Dim argConv As String
    argConv = LCase(sec.GetValue("OnClickArgConvention"))
    Dim rowTo As Long
    rowTo = ResolveRowTo(gridSec, rowFrom)

    Dim r As Long, n As Long
    For r = rowFrom To rowTo
        Dim anchor As Range
        Set anchor = ws.Range(colOffset & CStr(r))
        Dim cumLeft As Double
        cumLeft = 0#
        For n = 1 To btnCount
            cumLeft = cumLeft + PlaceOneTemplateButton( _
                ws, sec, templateId, n, r, anchor, cumLeft, argConv)
        Next n
    Next r
    Exit Sub

ErrHandler:
    Debug.Print "[" & MOD_NAME & ".PlaceTemplateButtons] " & _
                Err.Number & " " & Err.Description
End Sub

' ============================================================
' �ｿｽﾖ撰ｿｽ�ｿｽ�ｿｽ: PlaceOneTemplateButton
' �ｿｽT�ｿｽv:   per-row �ｿｽ{�ｿｽ^�ｿｽ�ｿｽ 1 �ｿｽﾂを生撰ｿｽ�ｿｽB�ｿｽﾟゑｿｽl�ｿｽﾍ配�ｿｽu�ｿｽ�ｿｽ (pt) �ｿｽﾅ、
'         �ｿｽﾄ出�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽs�ｿｽﾌ趣ｿｽ�ｿｽ{�ｿｽ^�ｿｽ�ｿｽ�ｿｽﾌ搾ｿｽ�ｿｽI�ｿｽt�ｿｽZ�ｿｽb�ｿｽg�ｿｽﾝ積に使�ｿｽ�ｿｽ�ｿｽB
' �ｿｽ�ｿｽ�ｿｽ�ｿｽ:   ws, sec, templateId, n(�ｿｽ{�ｿｽ^�ｿｽ�ｿｽ�ｿｽﾔ搾ｿｽ), rowNum(�ｿｽs�ｿｽﾔ搾ｿｽ),
'         anchor(�ｿｽs�ｿｽA�ｿｽ�ｿｽ�ｿｽJ�ｿｽ[), leftOffsetPt(�ｿｽ�ｿｽ�ｿｽﾝ撰ｿｽ), argConv
' �ｿｽﾟゑｿｽl: Double - �ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ{�ｿｽ^�ｿｽ�ｿｽ�ｿｽﾌ包ｿｽ (pt)
' ============================================================
Private Function PlaceOneTemplateButton( _
    ByVal ws As Worksheet, ByVal sec As ClsStanzaSection, _
    ByVal templateId As String, ByVal n As Long, ByVal rowNum As Long, _
    ByVal anchor As Range, ByVal leftOffsetPt As Double, _
    ByVal argConv As String) As Double

    Dim wPt As Double
    wPt = PxToPt(ToDouble(sec.GetValue("Button" & CStr(n) & "_Width"), 60#), _
                 BTN_MIN_W, 60#)
    PlaceOneTemplateButton = wPt

    On Error GoTo ErrHandler
    Dim hPt As Double
    hPt = PxToPt(ToDouble(sec.GetValue("Button" & CStr(n) & "_Height"), 20#), _
                 BTN_MIN_H, 20#)

    Dim btnName As String
    ' Excel shape names have a length cap; long names make .Name fail
    ' silently, leaving the form-control default name/caption. Keep short.
    btnName = NAME_PREFIX & "t" & Left(SafeToken(templateId), 6) & _
              "_" & CStr(rowNum) & "_" & CStr(n)
    DeleteShapeByName ws, btnName

    Dim shp As Shape
    Set shp = ws.Shapes.AddFormControl(FC_BUTTON, _
        anchor.Left + leftOffsetPt, anchor.Top, wPt, hPt)
    shp.Name = btnName

    ' �ｿｽL�ｿｽ�ｿｽ�ｿｽv�ｿｽV�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ OnAction �ｿｽﾝ抵ｿｽ�ｿｽ�ｿｽ�ｿｽﾉ行�ｿｽ�ｿｽ�ｿｽBOnAction �ｿｽﾌ托ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽﾂ具ｿｽ�ｿｽ�ｿｽ
    ' �ｿｽ�ｿｽ�ｿｽ�ｿｽﾄ趣ｿｽ�ｿｽs�ｿｽ�ｿｽ�ｿｽG�ｿｽ�ｿｽ�ｿｽ[�ｿｽﾉなゑｿｽ�ｿｽﾄゑｿｽ Button{n}_Text �ｿｽﾌキ�ｿｽ�ｿｽ�ｿｽv�ｿｽV�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽm�ｿｽ�ｿｽ�ｿｽ�ｿｽ
    ' �ｿｽ�ｿｽ�ｿｽf�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ謔､�ｿｽﾉゑｿｽ�ｿｽ�ｿｽ (�ｿｽ�ｿｽ�ｿｽ@�ｿｽ�ｿｽ�ｿｽﾘで費ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ "�ｿｽ{�ｿｽ^" �ｿｽs�ｿｽ�合�ｿｽﾌ搾ｿｽ�ｿｽ�ｿｽ:
    ' �ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽﾍ茨ｿｽ�ｿｽ�ｿｽ�ｿｽt�ｿｽ�ｿｽ OnAction �ｿｽ�ｿｽ�ｿｽ�ｿｽﾌ趣ｿｽ�ｿｽs�ｿｽ�ｿｽ ErrHandler �ｿｽo�ｿｽR�ｿｽ�ｿｽ
    ' ApplyCaption �ｿｽ�ｿｽ�ｿｽX�ｿｽL�ｿｽb�ｿｽv�ｿｽ�ｿｽ�ｿｽA�ｿｽ�ｿｽ�ｿｽ阮ｼ "�ｿｽ{�ｿｽ^�ｿｽ�ｿｽ N" �ｿｽ�ｿｽ�ｿｽc�ｿｽ�ｿｽ�ｿｽﾄゑｿｽ�ｿｽ�ｿｽ)�ｿｽB
    ApplyCaption ws, shp, sec.GetValue("Button" & CStr(n) & "_Text"), _
                 sec.GetValue("Button" & CStr(n) & "_ForeColor")

    ' OnAction �ｿｽﾝ抵ｿｽBrow_index �ｿｽK�ｿｽ�ｿｽﾍ茨ｿｽ�ｿｽ�ｿｽ�ｿｽt�ｿｽ�ｿｽ�ｿｽﾄ出�ｿｽﾌゑｿｽ�ｿｽﾟ単�ｿｽ�ｿｽ�ｿｽ�ｿｽp�ｿｽ�ｿｽ�ｿｽﾅ奇ｿｽ�ｿｽ�ｿｽ
    ' ("'Macro �ｿｽs�ｿｽﾔ搾ｿｽ'")�ｿｽB�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽs�ｿｽﾍ局擾ｿｽ�ｿｽI�ｿｽﾉ茨ｿｽ�ｿｽ�ｿｽﾂぶゑｿｽ�ｿｽA�ｿｽz�ｿｽu�ｿｽ�ｿｽ
    ' �ｿｽL�ｿｽ�ｿｽ�ｿｽv�ｿｽV�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽﾍ撰ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽﾆゑｿｽ�ｿｽ�ｿｽB
    Dim onClick As String
    onClick = sec.GetValue("Button" & CStr(n) & "_OnClick")
    If Len(onClick) > 0 Then
        On Error Resume Next
        If argConv = "row_index" Then
            shp.OnAction = "'" & onClick & " " & CStr(rowNum) & "'"
        Else
            shp.OnAction = onClick
        End If
        On Error GoTo ErrHandler
    End If
    Exit Function

ErrHandler:
    Debug.Print "[" & MOD_NAME & ".PlaceOneTemplateButton] r" & CStr(rowNum) & _
                " n" & CStr(n) & ": " & Err.Number & " " & Err.Description
End Function

' ============================================================
' �ｿｽﾖ撰ｿｽ�ｿｽ�ｿｽ: ApplyCaption
' �ｿｽT�ｿｽv:   �ｿｽ{�ｿｽ^�ｿｽ�ｿｽ�ｿｽﾉキ�ｿｽ�ｿｽ�ｿｽv�ｿｽV�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽﾝ抵ｿｽ (�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ + �ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽF)�ｿｽB
'         FormControl �ｿｽ{�ｿｽ^�ｿｽ�ｿｽ�ｿｽﾌキ�ｿｽ�ｿｽ�ｿｽv�ｿｽV�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ ws.Buttons(name).Caption
'         �ｿｽo�ｿｽR�ｿｽ�ｿｽ�ｿｽm�ｿｽ�ｿｽ�ｿｽBTextFrame.Characters.Text �ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽﾆ奇ｿｽ�ｿｽ阨ｶ�ｿｽ�ｿｽ
'         "�ｿｽ{�ｿｽ^�ｿｽ�ｿｽ N" �ｿｽ�ｿｽ�ｿｽc�ｿｽ�ｿｽA�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽs�ｿｽ{�ｿｽ^�ｿｽ�ｿｽ�ｿｽ�ｿｽ "�ｿｽ{�ｿｽ^" �ｿｽﾆ鯉ｿｽ�ｿｽﾘゑｿｽ�ｿｽﾂ具ｿｽ�ｿｽ�ｿｽ
'         �ｿｽ�ｿｽ�ｿｽ�ｿｽ (�ｿｽ�ｿｽ�ｿｽ@�ｿｽ�ｿｽ�ｿｽﾘで確�ｿｽF�ｿｽA2026-05-22)�ｿｽB�ｿｽﾝ定失�ｿｽs�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽz�ｿｽu�ｿｽﾍ撰ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽB
' ============================================================
Private Sub ApplyCaption(ByVal ws As Worksheet, ByVal shp As Shape, _
                         ByVal caption As String, ByVal foreColorHex As String)
    If Len(caption) = 0 Then Exit Sub
    On Error Resume Next
    ' FormControl �ｿｽ{�ｿｽ^�ｿｽ�ｿｽ�ｿｽﾌ表�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ Buttons �ｿｽR�ｿｽ�ｿｽ�ｿｽN�ｿｽV�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽo�ｿｽR�ｿｽﾅ設定す�ｿｽ�ｿｽB
    shp.TextFrame.Characters.Text = caption
    shp.TextFrame.HorizontalAlignment = HA_CENTER
    Dim btn As Object
    Set btn = ws.Buttons(shp.Name)
    If Not (btn Is Nothing) Then btn.Text = caption
    ' v2.3 (2026-05-26): Form Control �ｿｽﾍ背�ｿｽi�ｿｽD�ｿｽﾅ抵ｿｽﾌゑｿｽ�ｿｽﾟ費ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ (#FFFFFF) �ｿｽﾍ読めなゑｿｽ�ｿｽB
    ' ui_seed �ｿｽ�ｿｽ�ｿｽﾅ搾ｿｽ�ｿｽﾉ難ｿｽ�ｿｽ�ｿｽﾏゑｿｽ�ｿｽ�ｿｽ�ｿｽA�ｿｽﾅ終 fallback �ｿｽﾆゑｿｽ�ｿｽﾄ搾ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽﾂ読撰ｿｽ�ｿｽ�ｿｽﾛ証ゑｿｽ�ｿｽ�ｿｽB
    ' �ｿｽ�ｿｽ�ｿｽF�ｿｽw�ｿｽ�ｿｽ (#FFFFFF) �ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ鼾��ｿｽﾍ搾ｿｽ�ｿｽﾉ擾ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽB
    Dim effectiveHex As String
    effectiveHex = foreColorHex
    If Len(effectiveHex) = 0 Then effectiveHex = "#000000"
    If UCase(effectiveHex) = "#FFFFFF" Then effectiveHex = "#000000"
    shp.TextFrame.Characters.Font.Color = HexToLong(effectiveHex)
    On Error GoTo 0
End Sub

' ============================================================
' �ｿｽﾖ撰ｿｽ�ｿｽ�ｿｽ: DeleteShapeByName
' �ｿｽT�ｿｽv:   �ｿｽw�ｿｽ阮ｼ�ｿｽ�ｿｽ Shape �ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽﾝゑｿｽ�ｿｽ�ｿｽﾎ削除 (idempotent �ｿｽp)�ｿｽB
' ============================================================
Private Sub DeleteShapeByName(ByVal ws As Worksheet, ByVal shapeName As String)
    On Error Resume Next
    Dim shp As Shape
    Set shp = ws.Shapes(shapeName)
    If Not shp Is Nothing Then shp.Delete
    On Error GoTo 0
End Sub

' ============================================================
' �ｿｽﾖ撰ｿｽ�ｿｽ�ｿｽ: ResolveRowTo
' �ｿｽT�ｿｽv:   per-row �ｿｽ{�ｿｽ^�ｿｽ�ｿｽ�ｿｽﾌ最終�ｿｽs�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽﾟゑｿｽBgridSec �ｿｽ�ｿｽ DataRowTo �ｿｽ�ｿｽ
'         �ｿｽ�ｿｽ�ｿｽﾆゑｿｽ�ｿｽAauto / �ｿｽ�ｿｽ�ｿｽw�ｿｽ�ｿｽ / �ｿｽs�ｿｽ�ｿｽ�ｿｽﾈゑｿｽb�ｿｽ�ｿｽs�ｿｽ�ｿｽ�ｿｽﾅ補う�ｿｽB
' �ｿｽﾟゑｿｽl: Long - �ｿｽﾅ終�ｿｽs�ｿｽﾔ搾ｿｽ
' ============================================================
Private Function ResolveRowTo(ByVal gridSec As ClsStanzaSection, _
                              ByVal rowFrom As Long) As Long
    Dim fallback As Long
    fallback = rowFrom + DEFAULT_ROW_COUNT - 1
    If gridSec Is Nothing Then
        ResolveRowTo = fallback
        Exit Function
    End If
    Dim v As String
    v = LCase(Trim(gridSec.GetValue("DataRowTo")))
    If Len(v) = 0 Or v = "auto" Then
        ResolveRowTo = fallback
        Exit Function
    End If
    Dim rt As Long
    rt = ToLong(v, 0)
    If rt < rowFrom Then
        ResolveRowTo = fallback
    Else
        ResolveRowTo = rt
    End If
End Function

' ============================================================
' �ｿｽﾖ撰ｿｽ�ｿｽ�ｿｽ: PxToPt
' �ｿｽT�ｿｽv:   �ｿｽs�ｿｽN�ｿｽZ�ｿｽ�ｿｽ�ｿｽl�ｿｽ�ｿｽ�ｿｽ|�ｿｽC�ｿｽ�ｿｽ�ｿｽg�ｿｽﾉ奇ｿｽ�ｿｽZ�ｿｽB0 �ｿｽﾈ会ｿｽ�ｿｽﾈゑｿｽ�ｿｽ�ｿｽ�ｿｽ px �ｿｽ�ｿｽ�ｿｽg�ｿｽ�ｿｽ�ｿｽA
'         �ｿｽ�ｿｽ�ｿｽZ�ｿｽ�ｿｽﾉ最擾ｿｽ�ｿｽ�ｿｽ�ｿｽ@�ｿｽK�ｿｽ[�ｿｽh (minPt) �ｿｽ�ｿｽK�ｿｽp�ｿｽ�ｿｽ�ｿｽ�ｿｽB
' ============================================================
Private Function PxToPt(ByVal px As Double, ByVal minPt As Double, _
                        ByVal defaultPx As Double) As Double
    Dim srcPx As Double
    srcPx = px
    If srcPx <= 0# Then srcPx = defaultPx
    Dim pt As Double
    pt = srcPx * PX_PER_PT
    If pt < minPt Then pt = minPt
    PxToPt = pt
End Function

' ============================================================
' �ｿｽﾖ撰ｿｽ�ｿｽ�ｿｽ: HexToLong
' �ｿｽT�ｿｽv:   "#RRGGBB" �ｿｽ`�ｿｽ�ｿｽ�ｿｽﾌ色�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ VBA �ｿｽﾌ色 Long �ｿｽﾉ変奇ｿｽ�ｿｽB
'         �ｿｽs�ｿｽ�ｿｽ�ｿｽﾈ値�ｿｽ�ｿｽ 0 (�ｿｽ�ｿｽ) �ｿｽ�ｿｽﾔゑｿｽ�ｿｽB
' ============================================================
Private Function HexToLong(ByVal hexStr As String) As Long
    On Error GoTo ErrHandler
    Dim s As String
    s = Trim(hexStr)
    If Len(s) > 0 Then
        If Left(s, 1) = "#" Then s = Mid(s, 2)
    End If
    If Len(s) <> 6 Then
        HexToLong = 0
        Exit Function
    End If
    HexToLong = RGB(CLng("&H" & Mid(s, 1, 2)), _
                    CLng("&H" & Mid(s, 3, 2)), _
                    CLng("&H" & Mid(s, 5, 2)))
    Exit Function
ErrHandler:
    HexToLong = 0
End Function

' ============================================================
' �ｿｽﾖ撰ｿｽ�ｿｽ�ｿｽ: SafeToken
' �ｿｽT�ｿｽv:   Shape �ｿｽ�ｿｽ�ｿｽﾉ使�ｿｽ�ｿｽ�ｿｽ�ｿｽ謔､�ｿｽL�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽE�ｿｽu�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽB
' ============================================================
Private Function SafeToken(ByVal s As String) As String
    Dim t As String
    t = Replace(s, ":", "_")
    t = Replace(t, " ", "_")
    t = Replace(t, "$", "")
    SafeToken = t
End Function

' ============================================================
' �ｿｽﾖ撰ｿｽ�ｿｽ�ｿｽ: ToLong
' �ｿｽT�ｿｽv:   �ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ Long �ｿｽﾉ変奇ｿｽ�ｿｽB�ｿｽ�ｿｽ / �ｿｽﾏ奇ｿｽ�ｿｽ�ｿｽ�ｿｽs�ｿｽ�ｿｽ�ｿｽﾍ奇ｿｽ�ｿｽ�ｿｽl�ｿｽ�ｿｽﾔゑｿｽ�ｿｽB
' ============================================================
Private Function ToLong(ByVal s As String, ByVal dflt As Long) As Long
    On Error GoTo ErrHandler
    If Len(Trim(s)) = 0 Then
        ToLong = dflt
    Else
        ToLong = CLng(Trim(s))
    End If
    Exit Function
ErrHandler:
    ToLong = dflt
End Function

' ============================================================
' �ｿｽﾖ撰ｿｽ�ｿｽ�ｿｽ: ToDouble
' �ｿｽT�ｿｽv:   �ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ Double �ｿｽﾉ変奇ｿｽ�ｿｽB�ｿｽ�ｿｽ / �ｿｽﾏ奇ｿｽ�ｿｽ�ｿｽ�ｿｽs�ｿｽ�ｿｽ�ｿｽﾍ奇ｿｽ�ｿｽ�ｿｽl�ｿｽ�ｿｽﾔゑｿｽ�ｿｽB
' ============================================================
Private Function ToDouble(ByVal s As String, ByVal dflt As Double) As Double
    On Error GoTo ErrHandler
    If Len(Trim(s)) = 0 Then
        ToDouble = dflt
    Else
        ToDouble = CDbl(Trim(s))
    End If
    Exit Function
ErrHandler:
    ToDouble = dflt
End Function
```
