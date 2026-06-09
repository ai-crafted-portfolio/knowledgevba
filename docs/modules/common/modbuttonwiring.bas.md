---
title: modButtonWiring.bas
description: modButtonWiring.bas のソースコード（コピペ用）
---

# modButtonWiring.bas

**配置先**: 共通モジュール（3 ブック共通）
**種類**: 標準モジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `modButtonWiring.bas`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。
> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。
> 改行コードは CRLF（Windows 標準）のままで OK です。

---

## ソースコード

```vb
Attribute VB_Name = "modButtonWiring"
' ============================================================
' modButtonWiring (???[?e?B???e?B?w, v2.1)
' ????:
'   UI ??`?X?^???U?? [BUTTON] / [BUTTON_TEMPLATE] ?Z?N?V????????
'   ?t?H?[???R???g???[???{?^???????s????????????? OnAction ??
'   ???蓖???B?e?X?g??p?????? E2E_AutoWiring ??
'   Shapes.AddFormControl ?z?u???W?b?N??i???W???[??????i?????B
' ??o??:
'   modUILoader.ApplyButton / modUILoader.ApplyButtonTemplate
' ?????:
'   ClsStanzaSection (UI ??`?Z?N?V??????l?I?u?W?F?N?g)
' ??v????:
'   - ?{????????u?b?N?I?[?v?????AclsSetupOrchestrator.RunFullSetup
'     -> ApplyUiToSheet ?o?R?????s?????B???O??
'     clsSheetRenderer.ClearScreen ???S Shape ??????????A
'     ???????????d??????? (idempotent)?B
'   - ClearScreen ???????K?p????c?[???o??????A?????O??
'     ???? Shape ????????B
'   - OnAction ????C???}?N??????????BExcel ??????????
'     ????u?b?N????O?u (Book.xlsm!Macro ?`??) ???邪?A?????
'     Excel ??d?l?????s????????B
'   - [BUTTON].Width / Height ??s?N?Z???? (schema 3.7)?B
'     AddFormControl ??|?C???g?P?????? px -> pt ???Z????B
' ASCII-only inside VBA string literals (CP932 mojibake avoidance).
' ============================================================
Option Explicit

Private Const MOD_NAME As String = "modButtonWiring"
Private Const FC_BUTTON As Long = 0           ' xlButtonControl
Private Const HA_CENTER As Long = -4108       ' xlHAlignCenter
Private Const PX_PER_PT As Double = 0.75      ' 96dpi: 1px = 0.75pt
Private Const BTN_MIN_W As Double = 24#       ' ?{?^??????? (pt)
Private Const BTN_MIN_H As Double = 12#       ' ?{?^??????? (pt)
Private Const DEFAULT_ROW_COUNT As Long = 20  ' DataRowTo=auto ????b??s??
Private Const NAME_PREFIX As String = "btnwire_"

' ============================================================
' ?????: PlaceButton
' ?T?v:   [BUTTON] ?X?^???U 1 ???????I?{?^??????? OnAction
'         ?????蓖???B
' ????:   ws           - ?z?u??V?[?g
'         cellAddr     - ?A???J?[?Z?? (A1 ?\?L?A???????)
'         caption      - ?{?^???L???v?V????
'         onClick      - ?????}?N???? (???C??)
'         widthPx      - ?? (?s?N?Z??)
'         heightPx     - ???? (?s?N?Z??)
'         foreColorHex - ?L???v?V?????????F (#RRGGBB?A???)
' ============================================================
Public Sub PlaceButton(ByVal ws As Worksheet, ByVal cellAddr As String, _
                       ByVal caption As String, ByVal onClick As String, _
                       ByVal widthPx As Double, ByVal heightPx As Double, _
                       ByVal foreColorHex As String, _
                       Optional ByVal backColorLong As Long = -1)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0903] modButtonWiring.PlaceButton ENTER"  ' [ADR-0100]
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
        ' Phase R-3-ψ-M03-style (差3): 色付きボタンは Shape(角丸矩形)で描画。
        ' Form Control ボタンは塗りつぶし不可のため、BackColor 指定時のみ Shape 化。
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0904] modButtonWiring.PlaceButton EXIT-OK"  ' [ADR-0100]
    Exit Sub

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0905] modButtonWiring.PlaceButton EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Debug.Print "[" & MOD_NAME & ".PlaceButton] " & cellAddr & ": " & _
                Err.Number & " " & Err.Description
End Sub

' ============================================================
' ?????: PlaceTemplateButtons
' ?T?v:   [BUTTON_TEMPLATE] ?X?^???U 1 ??????A??? [GRID] ??
'         ?f?[?^?s????? per-row ?{?^???Q???????B
' ????:   ws      - ?z?u??V?[?g
'         sec     - [BUTTON_TEMPLATE] ?Z?N?V????
'         gridSec - GridRef ???w?? [GRID] ?Z?N?V???? (Nothing ??)
' ???l:   gridSec ?? Nothing ????? DataRowTo=auto ?????
'         RowRangeFrom ???? DEFAULT_ROW_COUNT ?s????????B
' ============================================================
Public Sub PlaceTemplateButtons(ByVal ws As Worksheet, _
                                ByVal sec As ClsStanzaSection, _
                                ByVal gridSec As ClsStanzaSection)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0906] modButtonWiring.PlaceTemplateButtons ENTER"  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0907] modButtonWiring.PlaceTemplateButtons EXIT-OK"  ' [ADR-0100]
    Exit Sub

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0908] modButtonWiring.PlaceTemplateButtons EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Debug.Print "[" & MOD_NAME & ".PlaceTemplateButtons] " & _
                Err.Number & " " & Err.Description
End Sub

' ============================================================
' ?????: PlaceOneTemplateButton
' ?T?v:   per-row ?{?^?? 1 ?????B???l??z?u?? (pt) ??A
'         ??o????????s????{?^??????I?t?Z?b?g????g???B
' ????:   ws, sec, templateId, n(?{?^?????), rowNum(?s???),
'         anchor(?s?A???J?[), leftOffsetPt(?????), argConv
' ???l: Double - ?????????{?^????? (pt)
' ============================================================
Private Function PlaceOneTemplateButton( _
    ByVal ws As Worksheet, ByVal sec As ClsStanzaSection, _
    ByVal templateId As String, ByVal n As Long, ByVal rowNum As Long, _
    ByVal anchor As Range, ByVal leftOffsetPt As Double, _
    ByVal argConv As String) As Double
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0909] modButtonWiring.PlaceOneTemplateButton ENTER"  ' [ADR-0100]

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

    ' ?L???v?V?????? OnAction ???????s???BOnAction ???????????
    ' ???????s???G???[??????? Button{n}_Text ??L???v?V???????m????
    ' ???f??????????? (???@???????????? "?{?^" ?s???????:
    ' ????????????t?? OnAction ???????s?? ErrHandler ?o?R??
    ' ApplyCaption ???X?L?b?v???A???? "?{?^?? N" ???c???????)?B
    ApplyCaption ws, shp, sec.GetValue("Button" & CStr(n) & "_Text"), _
                 sec.GetValue("Button" & CStr(n) & "_ForeColor")

    ' OnAction ???Brow_index ?K???????t????o?????P????p???????
    ' ("'Macro ?s???'")?B??????s?????I?????????A?z?u??
    ' ?L???v?V??????????????????B
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0910] modButtonWiring.PlaceOneTemplateButton EXIT-OK"  ' [ADR-0100]
    Exit Function

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0911] modButtonWiring.PlaceOneTemplateButton EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Debug.Print "[" & MOD_NAME & ".PlaceOneTemplateButton] r" & CStr(rowNum) & _
                " n" & CStr(n) & ": " & Err.Number & " " & Err.Description
End Function

' ============================================================
' ?????: ApplyCaption
' ?T?v:   ?{?^????L???v?V???????? (?????? + ?????F)?B
'         FormControl ?{?^????L???v?V?????? ws.Buttons(name).Caption
'         ?o?R???m???BTextFrame.Characters.Text ?????????????
'         "?{?^?? N" ???c??A?????s?{?^???? "?{?^" ???????????
'         ???? (???@?????m?F?A2026-05-22)?B????s?????z?u???????????B
' ============================================================
Private Sub ApplyCaption(ByVal ws As Worksheet, ByVal shp As Shape, _
                         ByVal caption As String, ByVal foreColorHex As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0912] modButtonWiring.ApplyCaption ENTER"  ' [ADR-0100]
    If Len(caption) = 0 Then Exit Sub
    On Error Resume Next
    ' FormControl ?{?^????\?????? Buttons ?R???N?V?????o?R??????B
    shp.TextFrame.Characters.Text = caption
    shp.TextFrame.HorizontalAlignment = HA_CENTER
    Dim btn As Object
    Set btn = ws.Buttons(shp.Name)
    If Not (btn Is Nothing) Then btn.Text = caption
    ' v2.3 (2026-05-26): Form Control ??w?i?D???????????? (#FFFFFF) ???????B
    ' ui_seed ???????????????A??I fallback ?????????????????????????B
    ' ???F?w?? (#FFFFFF) ??????????????????????B
    Dim effectiveHex As String
    effectiveHex = foreColorHex
    If Len(effectiveHex) = 0 Then effectiveHex = "#000000"
    If UCase(effectiveHex) = "#FFFFFF" Then effectiveHex = "#000000"
    shp.TextFrame.Characters.Font.Color = HexToLong(effectiveHex)
    On Error GoTo 0
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0913] modButtonWiring.ApplyCaption EXIT-OK"  ' [ADR-0100]
End Sub

' ============================================================
' ?????: DeleteShapeByName
' ?T?v:   ?w???? Shape ??????????? (idempotent ?p)?B
' ============================================================
Private Sub DeleteShapeByName(ByVal ws As Worksheet, ByVal shapeName As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0914] modButtonWiring.DeleteShapeByName ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim shp As Shape
    Set shp = ws.Shapes(shapeName)
    If Not shp Is Nothing Then shp.Delete
    On Error GoTo 0
End Sub

' ============================================================
' ?????: ResolveRowTo
' ?T?v:   per-row ?{?^?????I?s???????BgridSec ?? DataRowTo ??
'         ??????Aauto / ???w?? / ?s?????b??s??????B
' ???l: Long - ??I?s???
' ============================================================
Private Function ResolveRowTo(ByVal gridSec As ClsStanzaSection, _
                              ByVal rowFrom As Long) As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0915] modButtonWiring.ResolveRowTo ENTER"  ' [ADR-0100]
    Dim fallback As Long
    fallback = rowFrom + DEFAULT_ROW_COUNT - 1
    If gridSec Is Nothing Then
        ResolveRowTo = fallback
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0916] modButtonWiring.ResolveRowTo EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    Dim v As String
    v = LCase(Trim(gridSec.GetValue("DataRowTo")))
    If Len(v) = 0 Or v = "auto" Then
        ResolveRowTo = fallback
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0917] modButtonWiring.ResolveRowTo EXIT-OK"  ' [ADR-0100]
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
' ?????: PxToPt
' ?T?v:   ?s?N?Z???l???|?C???g????Z?B0 ????????? px ???g???A
'         ???Z????????@?K?[?h (minPt) ??K?p????B
' ============================================================
Private Function PxToPt(ByVal px As Double, ByVal minPt As Double, _
                        ByVal defaultPx As Double) As Double
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0918] modButtonWiring.PxToPt ENTER"  ' [ADR-0100]
    Dim srcPx As Double
    srcPx = px
    If srcPx <= 0# Then srcPx = defaultPx
    Dim pt As Double
    pt = srcPx * PX_PER_PT
    If pt < minPt Then pt = minPt
    PxToPt = pt
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0919] modButtonWiring.PxToPt EXIT-OK"  ' [ADR-0100]
End Function

' ============================================================
' ?????: HexToLong
' ?T?v:   "#RRGGBB" ?`????F??????? VBA ??F Long ?????B
'         ?s????l?? 0 (??) ?????B
' ============================================================
Private Function HexToLong(ByVal hexStr As String) As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0920] modButtonWiring.HexToLong ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Dim s As String
    s = Trim(hexStr)
    If Len(s) > 0 Then
        If Left(s, 1) = "#" Then s = Mid(s, 2)
    End If
    If Len(s) <> 6 Then
        HexToLong = 0
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0921] modButtonWiring.HexToLong EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    HexToLong = RGB(CLng("&H" & Mid(s, 1, 2)), _
                    CLng("&H" & Mid(s, 3, 2)), _
                    CLng("&H" & Mid(s, 5, 2)))
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0922] modButtonWiring.HexToLong EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0923] modButtonWiring.HexToLong EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    HexToLong = 0
End Function

' ============================================================
' ?????: SafeToken
' ?T?v:   Shape ????g??????L?????????E?u??????B
' ============================================================
Private Function SafeToken(ByVal s As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0924] modButtonWiring.SafeToken ENTER"  ' [ADR-0100]
    Dim t As String
    t = Replace(s, ":", "_")
    t = Replace(t, " ", "_")
    t = Replace(t, "$", "")
    SafeToken = t
End Function

' ============================================================
' ?????: ToLong
' ?T?v:   ??????? Long ?????B?? / ??????s???????l?????B
' ============================================================
Private Function ToLong(ByVal s As String, ByVal dflt As Long) As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0925] modButtonWiring.ToLong ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    If Len(Trim(s)) = 0 Then
        ToLong = dflt
    Else
        ToLong = CLng(Trim(s))
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0926] modButtonWiring.ToLong EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0927] modButtonWiring.ToLong EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ToLong = dflt
End Function

' ============================================================
' ?????: ToDouble
' ?T?v:   ??????? Double ?????B?? / ??????s???????l?????B
' ============================================================
Private Function ToDouble(ByVal s As String, ByVal dflt As Double) As Double
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0928] modButtonWiring.ToDouble ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    If Len(Trim(s)) = 0 Then
        ToDouble = dflt
    Else
        ToDouble = CDbl(Trim(s))
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0929] modButtonWiring.ToDouble EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0930] modButtonWiring.ToDouble EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ToDouble = dflt
End Function
```
