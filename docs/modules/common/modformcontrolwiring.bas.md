---
title: modFormControlWiring.bas
description: modFormControlWiring.bas のソースコード（コピペ用）
---

# modFormControlWiring.bas

**配置先**: 共通モジュール（検索.xlsm / 管理.xlsm 共通）
**種類**: 標準モジュール
**更新日**: 2026-06-30 14:44 JST

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `modFormControlWiring.bas`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。
> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。
> 改行コードは CRLF（Windows 標準）のままで OK です。

---

## ソースコード

```vb
Attribute VB_Name = "modFormControlWiring"
' ============================================================
' modFormControlWiring (utility layer, v2.3 Phase FC)
' Role:
'   Dynamically place Form Control / data validation widgets
'   from UI stanzas: [VALIDATION], [CHECKBOX], [LISTBOX].
'   Hybrid strategy per ADR-0096:
'     [VALIDATION] -> Range.Validation.Add (xlValidateList).
'                     Most reliable across Excel COM versions.
'     [CHECKBOX]   -> Shapes.AddFormControl(xlCheckBox, ...).
'                     On AddFormControl failure (some Excel COM
'                     hosts return 1004 / NoForm), falls back
'                     transparently to a 2-value Validation list
'                     ('TRUE,FALSE') on the same cell.
'     [LISTBOX]    -> Shapes.AddFormControl(xlListBox, ...).
'                     Same Validation fallback path on failure.
' Linked cells:
'   CHECKBOX / LISTBOX LinkedCell is the cell that mirrors the
'   current widget value (TRUE/FALSE for CheckBox, 1..N index
'   for ListBox). M-11 / similar sheets keep LinkedCell on the
'   hidden side (HideColumnsFrom=H), so the visible layout is
'   not disturbed.
' Shape names:
'   FC_VAL_<token>  : marker shape for a Validation list (none
'                      placed; tracking is via the cell itself).
'   FC_CHK_<token>  : Form Control checkbox shape (token is the
'                      stanza Name or sanitized caption).
'   FC_LST_<token>  : Form Control listbox shape.
' Coding rule:
'   ADR-0094 D1. CP932 + CRLF. ASCII-only inside string literals
'   (no JP in source body); ChrW() for any JP runtime label.
' ============================================================
Option Explicit

Private Const MOD_NAME As String = "modFormControlWiring"
Private Const FC_CHECKBOX As Long = 1     ' xlCheckBox
Private Const FC_LISTBOX As Long = 6      ' xlListBox
Private Const FC_OPTION As Long = 7       ' xlOptionButton (reserved)
Private Const PX_PER_PT As Double = 0.75  ' 96dpi: 1px = 0.75pt
Private Const CHK_NAME_PREFIX As String = "FC_CHK_"
Private Const LST_NAME_PREFIX As String = "FC_LST_"
Private Const CHK_DEFAULT_W_PT As Double = 220#
Private Const CHK_DEFAULT_H_PT As Double = 18#
Private Const LST_DEFAULT_W_PT As Double = 140#
Private Const LST_DEFAULT_H_PT As Double = 80#

' ============================================================
' Public Sub: PlaceValidationList
' Role: attach an Excel data-validation list to a cell.
' Args:
'   ws            - target worksheet (must not be Nothing)
'   cellAddr      - A1 address (e.g. "B5"). First cell is used
'                   when a range is given.
'   listValues    - CSV value list (e.g. "OFF,ERROR,WARN,INFO")
'   defaultValue  - optional initial cell value. When omitted
'                   the first CSV element is used.
' Returns: nothing. Errors are swallowed (debug-printed) so a
'   bad stanza never breaks RunFullSetup.
' ============================================================
Public Sub PlaceValidationList(ByVal ws As Worksheet, _
                                ByVal cellAddr As String, _
                                ByVal listValues As String, _
                                Optional ByVal defaultValue As String = "")
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1222] modFormControlWiring.PlaceValidationList ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    If ws Is Nothing Then Exit Sub
    If Len(cellAddr) = 0 Then Exit Sub
    If Len(listValues) = 0 Then Exit Sub

    Dim r As Range
    Set r = ws.Range(cellAddr).Cells(1, 1)
    r.Locked = False

    On Error Resume Next
    r.Validation.Delete
    On Error GoTo ErrHandler

    r.Validation.Add Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
        Formula1:=listValues
    r.Validation.InCellDropdown = True
    r.Validation.IgnoreBlank = False

    Dim dv As String
    dv = defaultValue
    If Len(dv) = 0 Then
        Dim parts() As String
        parts = Split(listValues, ",")
        If UBound(parts) >= LBound(parts) Then dv = Trim(parts(LBound(parts)))
    End If
    If Len(dv) > 0 Then r.Value = dv
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1223] modFormControlWiring.PlaceValidationList EXIT-OK"  ' [ADR-0100]
    Exit Sub

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1224] modFormControlWiring.PlaceValidationList EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Debug.Print "[" & MOD_NAME & ".PlaceValidationList] " & cellAddr & ": " & _
                Err.Number & " " & Err.Description
End Sub

' ============================================================
' Public Sub: PlaceCheckBox
' Role: place a Form Control CheckBox at a cell anchor.
'       On Shapes.AddFormControl failure, falls back to a 2-value
'       Validation list ("TRUE,FALSE") on the same cell so the
'       user always has *some* way to pick a value.
' Args:
'   ws             - target worksheet
'   cellAddr       - anchor cell (A1 form)
'   captionText    - visible caption (may contain JP characters
'                    supplied by caller; the stanza loader hands
'                    them in pre-decoded)
'   onActionMacro  - OnAction macro ("" to skip)
'   linkedCellAddr - LinkedCell address (TRUE/FALSE mirror cell).
'                    Empty string skips LinkedCell wiring.
'   nameToken      - stanza Name (Optional); used for shape Name.
'                    Falls back to a SafeToken(cellAddr) when omitted.
'   defaultBool    - initial Value (False/True).
' ============================================================
Public Sub PlaceCheckBox(ByVal ws As Worksheet, _
                          ByVal cellAddr As String, _
                          ByVal captionText As String, _
                          ByVal onActionMacro As String, _
                          ByVal linkedCellAddr As String, _
                          Optional ByVal nameToken As String = "", _
                          Optional ByVal defaultBool As Boolean = False)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1225] modFormControlWiring.PlaceCheckBox ENTER"  ' [ADR-0100]
    On Error GoTo FallbackA
    If ws Is Nothing Then Exit Sub
    If Len(cellAddr) = 0 Then Exit Sub

    Dim anchor As Range
    Set anchor = ws.Range(cellAddr).Cells(1, 1)

    Dim shpName As String
    shpName = CHK_NAME_PREFIX
    If Len(nameToken) > 0 Then
        shpName = shpName & SafeToken(nameToken)
    Else
        shpName = shpName & SafeToken(cellAddr)
    End If
    DeleteShapeByName ws, shpName

    Dim shp As Shape
    Set shp = ws.Shapes.AddFormControl(FC_CHECKBOX, _
        anchor.Left, anchor.Top, CHK_DEFAULT_W_PT, CHK_DEFAULT_H_PT)
    shp.Name = shpName

    ' Caption via the Buttons-style accessor (CheckBoxes collection).
    On Error Resume Next
    ws.CheckBoxes(shpName).Caption = captionText
    If Err.Number <> 0 Then
        Err.Clear
        shp.TextFrame.Characters.Text = captionText
    End If

    If Len(onActionMacro) > 0 Then
        shp.OnAction = onActionMacro
    End If

    If Len(linkedCellAddr) > 0 Then
        ws.CheckBoxes(shpName).LinkedCell = ws.Range(linkedCellAddr).Address(External:=False)
    End If

    ' xlOn = 1, xlOff = -4146. Use Form Control accessor for stability.
    If defaultBool Then
        ws.CheckBoxes(shpName).Value = 1            ' xlOn
        If Len(linkedCellAddr) > 0 Then ws.Range(linkedCellAddr).Value = True
    Else
        ws.CheckBoxes(shpName).Value = -4146         ' xlOff
        If Len(linkedCellAddr) > 0 Then ws.Range(linkedCellAddr).Value = False
    End If
    On Error GoTo FallbackA

    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1226] modFormControlWiring.PlaceCheckBox EXIT-OK"  ' [ADR-0100]
    Exit Sub

FallbackA:
    ' Form Control CheckBox failed (rare but observed on some Excel COM
    ' hosts). Fall through to a 2-value Validation list so the user can
    ' still pick TRUE/FALSE on the same cell.
    On Error Resume Next
    Debug.Print "[" & MOD_NAME & ".PlaceCheckBox] FALLBACK validation " & _
                cellAddr & ": " & Err.Number & " " & Err.Description
    Err.Clear
    Dim fb As String
    If defaultBool Then fb = "TRUE" Else fb = "FALSE"
    PlaceValidationList ws, cellAddr, "TRUE,FALSE", fb
    On Error GoTo 0
End Sub

' ============================================================
' Public Function: GetCheckBoxValue
' Role: read the current TRUE/FALSE value of a CheckBox by
'       caption (matches stanza Caption text). Falls through to
'       LinkedCell value when the shape itself is unreadable
'       (e.g. fallback Validation-only case).
' Return: True / False (False on any failure).
' ============================================================
Public Function GetCheckBoxValue(ByVal ws As Worksheet, _
                                  ByVal captionText As String) As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1227] modFormControlWiring.GetCheckBoxValue ENTER"  ' [ADR-0100]
    On Error GoTo Fallback
    If ws Is Nothing Then GoTo Fallback
    If Len(captionText) = 0 Then GoTo Fallback

    Dim cb As Object
    For Each cb In ws.CheckBoxes
        If cb.Caption = captionText Then
            GetCheckBoxValue = (cb.Value = 1)        ' xlOn
            If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1228] modFormControlWiring.GetCheckBoxValue EXIT-OK"  ' [ADR-0100]
            Exit Function
        End If
    Next cb
Fallback:
    GetCheckBoxValue = False
End Function

' ============================================================
' Public Sub: SetCheckBoxValue
' Role: update a CheckBox value by caption (Shape.Value +
'       LinkedCell mirror). Silent no-op when the caption is not
'       found.
' ============================================================
Public Sub SetCheckBoxValue(ByVal ws As Worksheet, _
                             ByVal captionText As String, _
                             ByVal valueBool As Boolean)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1229] modFormControlWiring.SetCheckBoxValue ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    If ws Is Nothing Then Exit Sub
    If Len(captionText) = 0 Then Exit Sub

    Dim cb As Object
    For Each cb In ws.CheckBoxes
        If cb.Caption = captionText Then
            If valueBool Then
                cb.Value = 1                          ' xlOn
            Else
                cb.Value = -4146                      ' xlOff
            End If
            Dim lc As String
            lc = CStr(cb.LinkedCell)
            If Len(lc) > 0 Then
                On Error Resume Next
                ws.Range(lc).Value = valueBool
                On Error GoTo ErrHandler
            End If
            If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1230] modFormControlWiring.SetCheckBoxValue EXIT-OK"  ' [ADR-0100]
            Exit Sub
        End If
    Next cb
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1231] modFormControlWiring.SetCheckBoxValue EXIT-OK"  ' [ADR-0100]
    Exit Sub

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1232] modFormControlWiring.SetCheckBoxValue EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Debug.Print "[" & MOD_NAME & ".SetCheckBoxValue] " & captionText & _
                ": " & Err.Number & " " & Err.Description
End Sub

' ============================================================
' Public Sub: PlaceListBox
' Role: place a Form Control ListBox at a cell anchor.
'       Fallback to a Validation list on the same cell if
'       AddFormControl fails.
' Args:
'   ws             - target worksheet
'   cellAddr       - anchor cell
'   listValues     - CSV value list
'   linkedCellAddr - LinkedCell (1..N index mirror)
'   nameToken      - optional shape Name token
' ============================================================
Public Sub PlaceListBox(ByVal ws As Worksheet, _
                         ByVal cellAddr As String, _
                         ByVal listValues As String, _
                         ByVal linkedCellAddr As String, _
                         Optional ByVal nameToken As String = "")
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1233] modFormControlWiring.PlaceListBox ENTER"  ' [ADR-0100]
    On Error GoTo FallbackA
    If ws Is Nothing Then Exit Sub
    If Len(cellAddr) = 0 Then Exit Sub
    If Len(listValues) = 0 Then Exit Sub

    Dim anchor As Range
    Set anchor = ws.Range(cellAddr).Cells(1, 1)

    Dim shpName As String
    shpName = LST_NAME_PREFIX
    If Len(nameToken) > 0 Then
        shpName = shpName & SafeToken(nameToken)
    Else
        shpName = shpName & SafeToken(cellAddr)
    End If
    DeleteShapeByName ws, shpName

    Dim shp As Shape
    Set shp = ws.Shapes.AddFormControl(FC_LISTBOX, _
        anchor.Left, anchor.Top, LST_DEFAULT_W_PT, LST_DEFAULT_H_PT)
    shp.Name = shpName

    ' Populate items. ListFillRange would point at a worksheet range;
    ' here we use AddItem directly so the values are self-contained.
    Dim parts() As String
    parts = Split(listValues, ",")
    Dim i As Long
    On Error Resume Next
    For i = LBound(parts) To UBound(parts)
        ws.ListBoxes(shpName).AddItem Trim(parts(i))
    Next i
    If Len(linkedCellAddr) > 0 Then
        ws.ListBoxes(shpName).LinkedCell = ws.Range(linkedCellAddr).Address(External:=False)
    End If
    On Error GoTo FallbackA
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1234] modFormControlWiring.PlaceListBox EXIT-OK"  ' [ADR-0100]
    Exit Sub

FallbackA:
    On Error Resume Next
    Debug.Print "[" & MOD_NAME & ".PlaceListBox] FALLBACK validation " & _
                cellAddr & ": " & Err.Number & " " & Err.Description
    Err.Clear
    PlaceValidationList ws, cellAddr, listValues, ""
    On Error GoTo 0
End Sub

' ============================================================
' Private Sub: DeleteShapeByName
' Role: idempotent shape removal. Used before every Add to keep
'       repeated RunFullSetup runs from stacking duplicate shapes.
' ============================================================
Private Sub DeleteShapeByName(ByVal ws As Worksheet, ByVal shapeName As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1235] modFormControlWiring.DeleteShapeByName ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim shp As Shape
    Set shp = ws.Shapes(shapeName)
    If Not shp Is Nothing Then shp.Delete
    On Error GoTo 0
End Sub

' ============================================================
' Private Function: SafeToken
' Role: sanitize a string for use inside a Shape name. Strips
'       characters Shape.Name rejects.
' ============================================================
Private Function SafeToken(ByVal s As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1236] modFormControlWiring.SafeToken ENTER"  ' [ADR-0100]
    Dim t As String
    t = s
    t = Replace(t, ":", "_")
    t = Replace(t, " ", "_")
    t = Replace(t, "$", "")
    t = Replace(t, "!", "_")
    t = Replace(t, "#", "_")
    t = Replace(t, "/", "_")
    t = Replace(t, "\", "_")
    SafeToken = t
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1237] modFormControlWiring.SafeToken EXIT-OK"  ' [ADR-0100]
End Function
```
