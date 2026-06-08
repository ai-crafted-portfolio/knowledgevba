---
title: modSheetButtons.bas
description: modSheetButtons.bas 縺ｮ繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝会ｼ医さ繝斐・逕ｨ・・---

# modSheetButtons.bas

**驟咲ｽｮ蜈・*: `蜈ｨ繝悶ャ繧ｯ蜈ｱ騾啻 逕ｨ縺ｮ VBA 繝｢繧ｸ繝･繝ｼ繝ｫ
**遞ｮ鬘・*: 讓呎ｺ悶Δ繧ｸ繝･繝ｼ繝ｫ

---

## 繝輔ぃ繧､繝ｫ縺ｨ縺励※菫晏ｭ・
繝｡繝｢蟶ｳ・医∪縺溘・莉ｻ諢上・繝・く繧ｹ繝医お繝・ぅ繧ｿ・峨↓荳九・繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝牙・譁・ｒ雋ｼ繧贋ｻ倥￠縲・*`modSheetButtons.bas`** 縺ｨ縺・≧蜷榊燕縺ｧ `installer\vba_modules\common\` 驟堺ｸ九↓菫晏ｭ倥＠縺ｦ縺上□縺輔＞縲よ枚蟄励さ繝ｼ繝峨・ ANSI・・hift-JIS・峨∵隼陦後・ CRLF 縺ｫ縺励※縺上□縺輔＞縲・
---

## 繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝・
```vb
Attribute VB_Name = "modSheetButtons"
' ================================================================
' modSheetButtons (v2.3 Phase R-3, 2026-05-28)
' Place form-control buttons on M-05/M-06/M-08 sheets.
' (M-09 Knowledge Display sheet was retired 2026-05-31; the view UserForm
'  is launched only via M-08 grid DoubleClick -> OpenViewWithId.)
' All JP literals use ChrW(&Hxxxx) to survive CP932 round-trips.
' ================================================================
Option Explicit

' --- JP literal helpers (ChrW so CP932 round-trip preserves chars) ---
Private Function NM_M05_DISPLAY() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1349] modSheetButtons.NM_M05_DISPLAY ENTER"  ' [ADR-0100]
    NM_M05_DISPLAY = ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & ChrW(&H767B) & ChrW(&H9332)
End Function
Private Function NM_M06_DISPLAY() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1350] modSheetButtons.NM_M06_DISPLAY ENTER"  ' [ADR-0100]
    NM_M06_DISPLAY = ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & ChrW(&H4FEE) & ChrW(&H6B63)
End Function
Private Function NM_M08_DISPLAY() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1351] modSheetButtons.NM_M08_DISPLAY ENTER"  ' [ADR-0100]
    NM_M08_DISPLAY = ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & ChrW(&H691C) & ChrW(&H7D22)
End Function
' NM_M09_DISPLAY removed 2026-05-31: M-09 sheet retired (see header comment).

Public Sub PlaceV23SheetButtons()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1352] modSheetButtons.PlaceV23SheetButtons ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    PlaceM05Buttons
    PlaceM06Buttons
    PlaceM08Buttons
    ' PlaceM09Buttons removed 2026-05-31: M-09 sheet retired.
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1353] modSheetButtons.PlaceV23SheetButtons EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1354] modSheetButtons.PlaceV23SheetButtons EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Debug.Print "[ERR] PlaceV23SheetButtons: " & Err.Number & " " & Err.Description
End Sub

Private Sub PlaceM05Buttons()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1355] modSheetButtons.PlaceM05Buttons ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ResolveSheet(NM_M05_DISPLAY(), "M-05")
    If ws Is Nothing Then Exit Sub
    ws.Unprotect
    ClearV23Buttons ws
    ' Phase R-3-omega: top 10->40 so buttons sit below the navy title bar (row1).
    AddV23Button ws, 10, 40, 140, 28, ChrW(&H65B0) & ChrW(&H898F) & ChrW(&H767B) & ChrW(&H9332), "TestPhaseP_ShowRegister"
    AddV23Button ws, 160, 40, 100, 28, ChrW(&H30AF) & ChrW(&H30EA) & ChrW(&H30A2), "TestPhaseP_Clear"
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1356] modSheetButtons.PlaceM05Buttons EXIT-OK"  ' [ADR-0100]
End Sub

Private Sub PlaceM06Buttons()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1357] modSheetButtons.PlaceM06Buttons ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ResolveSheet(NM_M06_DISPLAY(), "M-06")
    If ws Is Nothing Then Exit Sub
    ws.Unprotect
    ClearV23Buttons ws
    ' Phase R-3-omega: vestigial A3 label removed (number entry is in the UserForm);
    ' button top 50->40 so it sits below the navy title bar.
    AddV23Button ws, 10, 40, 200, 28, _
        ChrW(&H4FEE) & ChrW(&H6B63) & ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & ChrW(&H30E0) & ChrW(&H3092) & ChrW(&H958B) & ChrW(&H304F), _
        "Btn_OpenEditForm"
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1358] modSheetButtons.PlaceM06Buttons EXIT-OK"  ' [ADR-0100]
End Sub

' Phase R-3-omega (2026-05-29): M-08 は ui_seed/検索/M-08.txt の [SUBHEADER]/[LABEL]/
' [INPUT]/[BUTTON]/[GRID] stanza で完全描画する (beta-style)。検索/条件クリア ボタンは
' ui_seed [BUTTON] (OnClick=Btn_SearchV23 / Btn_SearchClearV23) で colored shape 配置。
' input セルは ApplyInput が Locked=False にするため保護下でも入力可。よって本 Sub は no-op。
Private Sub PlaceM08Buttons()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1359] modSheetButtons.PlaceM08Buttons ENTER"  ' [ADR-0100]
    ' intentionally empty: M-08 buttons/labels are now ui_seed-driven.
End Sub

' 2026-05-31: PlaceM09Buttons removed. M-09 sheet (Knowledge Display) was retired;
' the view UserForm is launched only via M-08 grid DoubleClick ->
' OpenViewWithId. DropRetiredScreens in clsSetupOrchestrator deletes any
' surviving M-09 sheet on install.

Private Function ResolveSheet(ByVal displayName As String, ByVal idName As String) As Worksheet
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1360] modSheetButtons.ResolveSheet ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(displayName)
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets(idName)
    Set ResolveSheet = ws
    On Error GoTo 0
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1361] modSheetButtons.ResolveSheet EXIT-OK"  ' [ADR-0100]
End Function

' Phase R-3-omega (2026-05-29): v23 buttons are now colored rounded-rectangle
' Shapes (not Form Control buttons), so removal must scan ws.Shapes by name.
Private Sub ClearV23Buttons(ByVal ws As Worksheet)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1362] modSheetButtons.ClearV23Buttons ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim shp As Object
    Dim toRemove As Collection
    Set toRemove = New Collection
    For Each shp In ws.Shapes
        If Left(shp.Name, 8) = "v23_btn_" Then toRemove.Add shp
    Next shp
    Dim i As Long
    For i = 1 To toRemove.Count
        toRemove.Item(i).Delete
    Next i
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1363] modSheetButtons.ClearV23Buttons EXIT-OK"  ' [ADR-0100]
End Sub

' Phase R-3-omega (2026-05-29): colored shape button (#4472C4, white bold, WordWrap off)
' to match the admin beta-style. OnAction macro wiring is unchanged.
Private Sub AddV23Button(ByVal ws As Worksheet, ByVal left As Single, ByVal top As Single, _
                          ByVal w As Single, ByVal h As Single, _
                          ByVal caption As String, ByVal macroName As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1364] modSheetButtons.AddV23Button ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, left, top, w, h)
    shp.Name = "v23_btn_" & Replace(macroName, "Btn_", "")
    shp.Fill.ForeColor.RGB = RGB(68, 114, 196)   ' #4472C4
    shp.Line.Visible = msoFalse
    shp.OnAction = macroName
    With shp.TextFrame2.TextRange
        .Text = caption
        .Font.Bold = msoTrue
        .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With
    shp.TextFrame2.HorizontalAnchor = 2          ' msoAnchorCenter
    shp.TextFrame2.VerticalAnchor = 3            ' msoAnchorMiddle
    shp.TextFrame2.WordWrap = msoFalse
    shp.TextFrame2.MarginLeft = 1
    shp.TextFrame2.MarginRight = 1
    shp.TextFrame2.MarginTop = 0
    shp.TextFrame2.MarginBottom = 0
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1365] modSheetButtons.AddV23Button EXIT-OK"  ' [ADR-0100]
End Sub

Public Sub TestPhaseP_Clear()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1366] modSheetButtons.TestPhaseP_Clear ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim lg As clsLogger
    Set lg = New clsLogger
    lg.Init ThisWorkbook.Worksheets("LOG")
    lg.LogInfo "modSheetButtons", "TestPhaseP_Clear", "M-05 sheet has no fields to clear (UserForm-only)", "", "LOG-M05-CLEAR-NOOP"
    On Error GoTo 0
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1367] modSheetButtons.TestPhaseP_Clear EXIT-OK"  ' [ADR-0100]
End Sub
```