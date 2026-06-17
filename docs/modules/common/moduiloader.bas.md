---
title: modUILoader.bas
description: modUILoader.bas のソースコード（コピペ用）
---

# modUILoader.bas

**配置先**: 共通モジュール（3 ブック共通）
**種類**: 標準モジュール
**更新日**: 2026-06-12 23:05

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `modUILoader.bas`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。
> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。
> 改行コードは CRLF（Windows 標準）のままで OK です。

---

## ソースコード

```vb
Attribute VB_Name = "modUILoader"
' ================================================================
' ???W???[??: modUILoader
' (?????W???[???w?b?_?R?????g????: ?????G???R?[?h?j?? (UTF-8 ??) ????
'  ???{??R?????g?? U+FFFD ??u?????????s?\?B???? impl_phase3 ???O?Q??)
' ================================================================
' v2.1-btnwire (2026-05-22):
'   ApplyButton / ApplyButtonTemplate ?? modButtonWiring ???????A
'   ???t?H?[???R???g???[???{?^???????????????? (???????????z?u
'   ?O???p?~)?B?????? modButtonWiring ?????B
' ================================================================
' v2.2-phase3 (2026-05-23) migration_plan 0.6.2 task 3.11 / 3.12:
'   task 3.11: ApplyGrid ?????? [GRID] ?X?L?[?}???]???C?B?w?b?_?a???E
'     ???E????E?w?i?F??\????p?L?[ ColumnHeaders ????`????
'     (v2_ui_stanza_schema.md 3.11.1?AOP-1 ?m??)?BColumns (?_????) ??
'     ?w?b?_?`???g???? (?f?[?^???o??R?[?h clsGridIO/modEntrySearch ??p)?B
'   task 3.12: ?V?Z?N?V?????^ [FORM_FROM_FORMAT] / [GRID_FROM_FORMAT] ??
'     ?n???h?? ApplyFormFromFormat / ApplyGridFromFormat ??V?K????
'     (v2_ui_stanza_schema.md 3.12 / 3.13?AOP-2 ?m??)?B?t?H?[?}?b?g??`??
'     [FIELD] ?Q????????E?^???????t?H?[??/?O???b?h???I?`????B
'   ????? (?R?[?h??????): modStanzaIO, modConfigHolder, modButtonWiring,
'     modFormatLoader, modKnowledgeFileIO, ClsStanzaSection?B
'   ??: ?{?t?@?C????Z?b?V?????J?n?O???? UTF-8 ?j????????{??R?????g
'     484 ?????????? U+FFFD ????????? (?N???[????????????????)?B?j????
'     ?R?????g????R?[?h??????B?{???C?? CP932+CRLF ????????A?????s?\??
'     ?R?????g??v???[?X?z???_??u?????? (???????????????????????)?B
' ================================================================
Option Explicit

' ----------------------------------------------------------------
' (???R?????g????: ?????G???R?[?h?j???????????s?\)
' ----------------------------------------------------------------
' ??: [FORM_FROM_FORMAT] / [GRID_FROM_FORMAT] ???I?t?H?[?????? 4 ??
' (v2_ui_stanza_schema.md 3.12 / 3.13?A?\?????i #/?K?{/?t?B?[???h??+?^/?????)
Private Const FORM_COL_NO As Long = 0          ' ?A?? (#) ???I?t?Z?b?g
Private Const FORM_COL_REQUIRED As Long = 1    ' ?K?{?}?[?N???I?t?Z?b?g
Private Const FORM_COL_FIELD As Long = 2       ' ?t?B?[???h??+?^ ???I?t?Z?b?g
Private Const FORM_COL_INPUT As Long = 3       ' ????????I?t?Z?b?g
' ??????????w?i?F (v2_ui_stanza_schema.md 3.6 ?? [INPUT] ???????l)
Private Const INPUT_DEFAULT_COLOR As String = "#FFFFCC"
' ????K??????E?l (?V?K?R?[?h??}?W?b?N?i???o?[???)
Private Const VALIDATE_INT_MIN As String = "-2147483647"
Private Const VALIDATE_INT_MAX As String = "2147483647"
Private Const VALIDATE_DATE_MIN As String = "1900/1/1"
Private Const VALIDATE_DATE_MAX As String = "9999/12/31"

' ----------------------------------------------------------------
' Public I/F
' ----------------------------------------------------------------

' ================================================================
' (???R?????g????: ?????G???R?[?h?j???????????s?\)
' ================================================================
Public Function LoadUiDefinition(ByVal xlsmName As String, ByVal screenId As String) As Collection
    Debug.Print "[D-1453] modUILoader.LoadUiDefinition ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler

    Dim filePath As String
    filePath = modConfigHolder.GetUiDir() & xlsmName & "\" & screenId & ".txt"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        Set LoadUiDefinition = New Collection
        Debug.Print "[D-1454] modUILoader.LoadUiDefinition EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If

    Set LoadUiDefinition = modStanzaIO.ParseStanzaFile(filePath)
    Debug.Print "[D-1455] modUILoader.LoadUiDefinition EXIT-OK"  ' [ADR-0100]
    Exit Function

ErrHandler:
    Debug.Print "[D-1456] modUILoader.LoadUiDefinition EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Set LoadUiDefinition = New Collection
End Function

' ================================================================
' Public Function: ResolveCellByInputDataKey (iter19, ADR-0090 SSOT helper)
' Role:    Look up the cell address bound to a given inputDataKey within
'          the ui_seed definition for (xlsmName, screenId). Generic SSOT
'          cell resolver that replaces hard-coded addresses throughout
'          both production and test code.
' Args:    xlsmName     - role-binding xlsm name (admin / register / search JP literal)
'          screenId     - screen identifier (e.g. "M-02", "M-03", "M-12")
'          inputDataKey - the [INPUT] inputDataKey value to match
' Return:  Cell address (e.g. "C3") if a match exists, "" otherwise.
'          When a range is bound (e.g. "C3:D3") the first cell is returned.
' ADR:     ADR-0090 (hard-code prohibition).
' ================================================================
Public Function ResolveCellByInputDataKey( _
    ByVal xlsmName As String, _
    ByVal screenId As String, _
    ByVal inputDataKey As String _
) As String
    Debug.Print "[D-1457] modUILoader.ResolveCellByInputDataKey ENTER"  ' [ADR-0100]
    On Error GoTo Fallback
    Dim ui As Collection
    Set ui = LoadUiDefinition(xlsmName, screenId)
    If ui Is Nothing Then GoTo Fallback
    If ui.Count = 0 Then GoTo Fallback
    Dim i As Long
    For i = 1 To ui.Count
        Dim sec As ClsStanzaSection
        Set sec = ui.Item(i)
        If sec.SectionName = "INPUT" Then
            Dim key As String
            key = sec.GetValue("inputDataKey")
            If key = inputDataKey Then
                Dim cellExpr As String
                cellExpr = sec.GetValue("Cell")
                If Len(cellExpr) > 0 Then
                    Dim colonPos As Long
                    colonPos = InStr(cellExpr, ":")
                    If colonPos > 0 Then
                        ResolveCellByInputDataKey = Left(cellExpr, colonPos - 1)
                    Else
                        ResolveCellByInputDataKey = cellExpr
                    End If
                    Debug.Print "[D-1458] modUILoader.ResolveCellByInputDataKey EXIT-OK"  ' [ADR-0100]
                    Exit Function
                End If
            End If
        End If
    Next i
Fallback:
    ResolveCellByInputDataKey = ""
End Function

' ================================================================
' (???R?????g????: ?????G???R?[?h?j???????????s?\)
' ================================================================
Public Function LoadUiList(ByVal xlsmName As String) As Collection
    Debug.Print "[D-1459] modUILoader.LoadUiList ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler

    Dim result As Collection
    Set result = New Collection

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folderPath As String
    folderPath = modConfigHolder.GetUiDir() & xlsmName & "\"

    If Not fso.FolderExists(folderPath) Then
        Set LoadUiList = result
        Debug.Print "[D-1460] modUILoader.LoadUiList EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If

    Dim folder As Object
    Set folder = fso.GetFolder(folderPath)
    Dim file As Object
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "txt" Then
            result.Add fso.GetBaseName(file.Name)
        End If
    Next file

    Set LoadUiList = result
    Debug.Print "[D-1461] modUILoader.LoadUiList EXIT-OK"  ' [ADR-0100]
    Exit Function

ErrHandler:
    Debug.Print "[D-1462] modUILoader.LoadUiList EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Set LoadUiList = New Collection
End Function

' ================================================================
' (???R?????g????: ?????G???R?[?h?j???????????s?\)
' ================================================================
Public Function ApplyUiToSheet( _
    ByVal xlsmName As String, _
    ByVal screenId As String, _
    ByVal ws As Worksheet _
) As Boolean
    Debug.Print "[D-1463] modUILoader.ApplyUiToSheet ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler

    ' uiSchemaFailMode (v2_ui_stanza_schema.md 5): safeDefault vs abort.
    ' canonical config uses warn/strict tokens; abort==strict, safeDefault==warn.
    Dim failMode As String
    failMode = GetUiSchemaFailMode()

    Dim filePath As String
    filePath = modConfigHolder.GetUiDir() & xlsmName & "\" & screenId & ".txt"

    ' (???R?????g????: ?????G???R?[?h?j???????????s?\)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        Debug.Print "[UILoader] file not found: " & filePath
        If failMode = "abort" Then
            Err.Raise vbObjectError + 9301, "modUILoader.ApplyUiToSheet", _
                "UI definition file not found (abort mode): " & filePath
        End If
        ApplyUiToSheet = False
        Debug.Print "[D-1464] modUILoader.ApplyUiToSheet EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If

    ' Phase R-2-Fix (2026-05-28): unprotect sheet before writing. A previous
    ' install may have left sheets protected, which silently blocks Cells.Clear
    ' and Range.Value writes (caught by ApplyUiStanzas outer error handler so
    ' cells stay empty).
    On Error Resume Next
    ws.Unprotect Password:=""
    On Error GoTo ErrHandler

    ' parse
    Dim sections As Collection
    Set sections = modStanzaIO.ParseStanzaFile(filePath)
    If sections.Count = 0 Then
        ' grammar error / empty parse result (schema 5 error condition)
        Debug.Print "[UILoader] parse failed: " & filePath
        If failMode = "abort" Then
            Err.Raise vbObjectError + 9302, "modUILoader.ApplyUiToSheet", _
                "UI definition grammar error (abort mode): " & filePath
        End If
        ApplyUiToSheet = False
        Debug.Print "[D-1465] modUILoader.ApplyUiToSheet EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If

    ' (???R?????g????: ?????G???R?[?h?j???????????s?\)
    ApplySectionsByName ws, sections, "SHEET"
    ApplySectionsByName ws, sections, "COLUMN"
    ApplySectionsByName ws, sections, "ROW"
    ApplySectionsByName ws, sections, "FONT"
    ApplySectionsByName ws, sections, "STYLE"
    ' Phase R-3-ψ-M03-style: SCREEN(タイトル帯) は STYLE(SheetBackColor) の後に適用し、
    ' A1 帯の塗りが白で上書きされないようにする(残差1 修正)。SUBHEADER も STYLE 後で帯維持。
    ApplySectionsByName ws, sections, "HEADER"
    ApplySectionsByName ws, sections, "SUBHEADER"
    ApplySectionsByName ws, sections, "SCREEN"
    ApplySectionsByName ws, sections, "LABEL"
    ApplySectionsByName ws, sections, "INPUT"
    ApplySectionsByName ws, sections, "CHECK"
    ApplySectionsByName ws, sections, "VALIDATION"
    ApplySectionsByName ws, sections, "CHECKBOX"
    ApplySectionsByName ws, sections, "LISTBOX"
    ApplySectionsByName ws, sections, "NOTE"
    ApplySectionsByName ws, sections, "GRID"
    ApplySectionsByName ws, sections, "FORM_FROM_FORMAT"
    ApplySectionsByName ws, sections, "GRID_FROM_FORMAT"
    ApplySectionsByName ws, sections, "BUTTON"
    ApplySectionsByName ws, sections, "BUTTON_TEMPLATE"
    ApplySectionsByName ws, sections, "FREEZE"

    ApplyUiToSheet = True
    Debug.Print "[D-1466] modUILoader.ApplyUiToSheet EXIT-OK"  ' [ADR-0100]
    Exit Function

ErrHandler:
    Debug.Print "[D-1467] modUILoader.ApplyUiToSheet EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Debug.Print "[UILoader] " & screenId & " - " & Err.Description
    ' abort mode: re-raise so the caller can halt all-screen startup.
    ' safeDefault mode: swallow, return False (this screen only fails).
    If GetUiSchemaFailMode() = "abort" Then
        Err.Raise Err.Number, "modUILoader.ApplyUiToSheet", Err.Description
    End If
    ApplyUiToSheet = False
End Function

' ================================================================
' (???R?????g????: ?????G???R?[?h?j???????????s?\)
' ================================================================
Private Function GetUiSchemaFailMode() As String
    Debug.Print "[D-1468] modUILoader.GetUiSchemaFailMode ENTER"  ' [ADR-0100]
    Dim s As String
    s = LCase(modConfigHolder.GetValueOrDefault("uiSchemaFailMode", "safeDefault"))
    If s = "abort" Or s = "strict" Then
        GetUiSchemaFailMode = "abort"
    Else
        GetUiSchemaFailMode = "safeDefault"
    End If
    Debug.Print "[D-1469] modUILoader.GetUiSchemaFailMode EXIT-OK"  ' [ADR-0100]
End Function

' ----------------------------------------------------------------
' (???R?????g????: ?????G???R?[?h?j???????????s?\)
' ----------------------------------------------------------------

Private Sub ApplySectionsByName( _
    ByVal ws As Worksheet, _
    ByVal sections As Collection, _
    ByVal targetName As String _
)
    Debug.Print "[D-1470] modUILoader.ApplySectionsByName ENTER"  ' [ADR-0100]
    Dim sec As ClsStanzaSection
    For Each sec In sections
        If sec.SectionName = targetName Then
            Select Case targetName
                Case "SHEET":  ApplySheet ws, sec
                Case "COLUMN": ApplyColumn ws, sec
                Case "ROW":    ApplyRow ws, sec
                Case "FONT":   ApplyFont ws, sec
                Case "HEADER": ApplyHeader ws, sec
                Case "LABEL":  ApplyLabel ws, sec
                Case "INPUT":  ApplyInput ws, sec
                Case "NOTE":   ApplyNote ws, sec
                Case "GRID":   ApplyGrid ws, sec
                Case "FORM_FROM_FORMAT": ApplyFormFromFormat ws, sec
                Case "GRID_FROM_FORMAT": ApplyGridFromFormat ws, sec
                Case "BUTTON": ApplyButton ws, sec
                Case "BUTTON_TEMPLATE": ApplyButtonTemplate ws, sec, sections
                ' v2.3 Phase A2 STYLE dispatch (re-enabled 2026-05-27 after bisect).
                Case "STYLE"
                    ApplyStyle ws, sec
                ' Phase R-2-Fix (2026-05-28): SCREEN/SUBHEADER/FREEZE handlers added.
                ' Previously the dispatcher had no Case for these section names so the
                ' auto-appended Phase A2 stanza was silent-skipped, leaving 管?Exlsm
                ' sheets fully blank. See ADR-0074 and R2_root_cause_report.md.
                Case "SCREEN": ApplyScreen ws, sec
                Case "SUBHEADER": ApplySubheader ws, sec
                Case "FREEZE": ApplyFreeze ws, sec
                ' Phase R-3-Ω-CHECK (2026-05-29): CHECK section for Form Control checkbox.
                ' M-10 ui_seed has [CHECK] stanza but dispatcher had no Case, so silent skip.
                Case "CHECK": ApplyCheck ws, sec
                ' Phase FC (2026-06-03 ADR-0096): hybrid FormControl dispatch.
                ' Phase FT (2026-06-03 ADR-0097): per-Case Resume Next so a single bad stanza
                ' never blocks the rest of ApplySectionsByName. ApplyXxx itself
                ' already has GoTo FCErr; this is a belt-and-braces layer.
                Case "VALIDATION"
                    Debug.Print "[FC-DISPATCH] VALIDATION sheet=" & ws.Name
                    On Error Resume Next
                    ApplyValidation ws, sec
                    If Err.Number <> 0 Then
                        Debug.Print "[FC-DISPATCH-ERR] VALIDATION " & Err.Number
                        Err.Clear
                    End If
                    On Error GoTo 0
                Case "CHECKBOX"
                    Debug.Print "[FC-DISPATCH] CHECKBOX sheet=" & ws.Name
                    On Error Resume Next
                    ApplyCheckBox ws, sec
                    If Err.Number <> 0 Then
                        Debug.Print "[FC-DISPATCH-ERR] CHECKBOX " & Err.Number
                        Err.Clear
                    End If
                    On Error GoTo 0
                Case "LISTBOX"
                    Debug.Print "[FC-DISPATCH] LISTBOX sheet=" & ws.Name
                    On Error Resume Next
                    ApplyListBox ws, sec
                    If Err.Number <> 0 Then
                        Debug.Print "[FC-DISPATCH-ERR] LISTBOX " & Err.Number
                        Err.Clear
                    End If
                    On Error GoTo 0
            End Select
        End If
    Next sec
    Debug.Print "[D-1471] modUILoader.ApplySectionsByName EXIT-OK"  ' [ADR-0100]
End Sub

' ----------------------------------------------------------------
' (???R?????g????: ?????G???R?[?h?j???????????s?\)
' ----------------------------------------------------------------

Private Sub ApplySheet(ByVal ws As Worksheet, ByVal sec As ClsStanzaSection)
    Debug.Print "[D-1472] modUILoader.ApplySheet ENTER"  ' [ADR-0100]
    ' v2.3 fix (2026-05-27): skip TabColor when "no value" token
    ' (empty / "-" / U+2015 horizontal bar) so ApplyTabColors result is preserved.
    Dim tabColor As String
    tabColor = sec.GetValue("TabColor")
    If Len(Trim(tabColor)) > 0 And Trim(tabColor) <> "-" Then
        ws.Tab.Color = HexToRgbLong(tabColor)
    End If
    ' v2.3 (2026-05-27): rename sheet to display name when SheetName differs.
    ' Per ui_seed [SHEET] SheetName=<display name>. Wrapped in On Error since
    ' name conflicts (existing sheet with same name) raise 1004 - we just keep
    ' the M-NN id in that case rather than hang.
    Dim displayName As String
    displayName = sec.GetValue("SheetName")
    If Len(Trim(displayName)) > 0 And Trim(displayName) <> "-" Then
        If ws.Name <> displayName Then
            On Error Resume Next
            ws.Name = displayName
            On Error GoTo 0
        End If
    End If
    ' Phase R-3-ψ-M03-style (右側余白除去): HideColumnsFrom=<列> 指定時、その列～末尾(XFD)を非表示。
    ' mock は A～J で完結。grid(A:G)+補助文オーバーフロー(～J)の右の空白列を全て??す。
    Dim hideFrom As String
    hideFrom = Trim(sec.GetValue("HideColumnsFrom"))
    If Len(hideFrom) > 0 Then
        On Error Resume Next
        ws.Range(hideFrom & ":XFD").EntireColumn.Hidden = True
        On Error GoTo 0
    End If
    ' (???R?????g????: ?????G???R?[?h?j???????????s?\)
    Debug.Print "[D-1473] modUILoader.ApplySheet EXIT-OK"  ' [ADR-0100]
End Sub

' Phase R-2-Fix: SCREEN section handler. Writes Title to A1.
Private Sub ApplyScreen(ByVal ws As Worksheet, ByVal sec As ClsStanzaSection)
    Debug.Print "[D-1474] modUILoader.ApplyScreen ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim t As String
    t = Trim(sec.GetValue("Title"))
    If Len(t) > 0 Then
        ' Phase R-3-ψ-M03-style: タイトルも全幅色帯対応(mock #1F3864/白/h30)。
        ' BackColor 指定時は A1..EndCol を帯化、未指定時は従来の素テキスト。
        Dim lc As Long
        Dim ec As String
        ec = Trim(sec.GetValue("EndCol"))
        If Len(ec) > 0 Then lc = ws.Range(ec & "1").Column Else lc = 1
        Dim bar As Range
        Set bar = ws.Range(ws.Cells(1, 1), ws.Cells(1, lc))
        Dim bg As String
        bg = Trim(sec.GetValue("BackColor"))
        If Len(bg) > 0 Then bar.Interior.Color = HexToRgbLong(bg)
        ws.Cells(1, 1).Value = t
        bar.Font.Bold = True
        bar.Font.Size = 14
        Dim fc As String
        fc = Trim(sec.GetValue("FontColor"))
        If Len(fc) > 0 Then bar.Font.Color = HexToRgbLong(fc) Else bar.Font.Color = RGB(45, 75, 130)
        Dim h As String
        h = Trim(sec.GetValue("BarHeight"))
        If IsNumeric(h) Then ws.Rows(1).RowHeight = CDbl(h)
    End If
    On Error GoTo 0
    Debug.Print "[D-1475] modUILoader.ApplyScreen EXIT-OK"  ' [ADR-0100]
End Sub

' Phase R-2-Fix: SUBHEADER section handler. Writes Title to next available row in column A.
' We track placement via a sheet-scoped counter (R2_subRow custom property).
Private Sub ApplySubheader(ByVal ws As Worksheet, ByVal sec As ClsStanzaSection)
    Debug.Print "[D-1476] modUILoader.ApplySubheader ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim t As String
    t = Trim(sec.GetValue("Title"))
    If Len(t) = 0 Then Exit Sub
    Dim tgt As Range
    ' Phase R-3-ψ-M03: explicit placement (Cell) で配置 ? サブヘッダー間に部品の
    ' 行を確保するため。Cell 未指定時は従来の auto-flow(行3から+2)。
    Dim cellK As String
    cellK = Trim(sec.GetValue("Cell"))
    If Len(cellK) > 0 Then
        Set tgt = ws.Range(cellK).Cells(1, 1)
    Else
        Dim r As Long
        r = 3
        Do While Len(Trim(CStr(ws.Cells(r, 1).Value))) > 0 And r <= 50
            r = r + 2   ' leave 1-row gap between subheaders
        Loop
        If r > 50 Then Exit Sub
        Set tgt = ws.Cells(r, 1)
    End If
    ' Phase R-3-ψ-M03-style (差1): 全幅色付き帯。BackColor + EndCol で帯化、FontColor + Height。
    ' mock ??測: BackColor=#2F5597 / 白文字 / 高さ24pt / 全幅(A..EndCol)。
    Dim lastC As Long
    Dim endCol As String
    endCol = Trim(sec.GetValue("EndCol"))
    If Len(endCol) > 0 Then
        lastC = ws.Range(endCol & "1").Column
    Else
        lastC = tgt.Column
    End If
    Dim bar As Range
    Set bar = ws.Range(ws.Cells(tgt.Row, tgt.Column), ws.Cells(tgt.Row, lastC))
    Dim bg As String
    bg = Trim(sec.GetValue("BackColor"))
    If Len(bg) > 0 Then bar.Interior.Color = HexToRgbLong(bg)
    tgt.Value = t
    bar.Font.Bold = True
    bar.Font.Size = 12
    Dim fc As String
    fc = Trim(sec.GetValue("FontColor"))
    If Len(fc) > 0 Then
        bar.Font.Color = HexToRgbLong(fc)
    Else
        bar.Font.Color = RGB(45, 75, 130)
    End If
    Dim shh As String
    shh = Trim(sec.GetValue("Height"))
    If IsNumeric(shh) Then ws.Rows(tgt.Row).RowHeight = CDbl(shh)
    On Error GoTo 0
    Debug.Print "[D-1477] modUILoader.ApplySubheader EXIT-OK"  ' [ADR-0100]
End Sub

' Phase R-2-Fix: FREEZE section handler. Applies FreezePanes at given Rows/Cols.
Private Sub ApplyFreeze(ByVal ws As Worksheet, ByVal sec As ClsStanzaSection)
    Debug.Print "[D-1478] modUILoader.ApplyFreeze ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim enabled As String
    enabled = LCase(Trim(sec.GetValue("Enabled")))
    If enabled <> "on" And enabled <> "true" And enabled <> "1" Then Exit Sub
    Dim rowsS As String, colsS As String
    rowsS = Trim(sec.GetValue("Rows"))
    colsS = Trim(sec.GetValue("Cols"))
    Dim rowsN As Long, colsN As Long
    If IsNumeric(rowsS) Then rowsN = CLng(rowsS) Else rowsN = 0
    If IsNumeric(colsS) Then colsN = CLng(colsS) Else colsN = 0
    ws.Activate
    ' 2026-06-06 v7 fix: ヘャドレス COM (Visible=False) で Activate しても Window.Height が小さい バグ
    ' 碕乞: Cells.Activate + FreezePanes = SplitRow となる聞扱づき
    ' 対策: Window サイズを強制大きくセットしてから FreezePanes
    Dim oldWidth As Double, oldHeight As Double
    oldWidth = ws.Application.ActiveWindow.Width
    oldHeight = ws.Application.ActiveWindow.Height
    ws.Application.ActiveWindow.WindowState = xlNormal
    ws.Application.ActiveWindow.Width = 1024
    ws.Application.ActiveWindow.Height = 768
    ws.Application.ActiveWindow.FreezePanes = False
    ws.Application.ActiveWindow.ScrollRow = 1
    ws.Application.ActiveWindow.ScrollColumn = 1
    Dim freezeRow As Long, freezeCol As Long
    freezeRow = rowsN + 1
    freezeCol = colsN + 1
    If freezeRow < 1 Then freezeRow = 1
    If freezeCol < 1 Then freezeCol = 1
    ws.Cells(freezeRow, freezeCol).Activate
    ws.Application.ActiveWindow.FreezePanes = True
    ws.Application.ActiveWindow.ScrollRow = 1
    ws.Application.ActiveWindow.ScrollColumn = 1
    On Error Resume Next
    ' [DIAG3 removed 2026-06-06] previously wrote debug info to A30, removed per user complaint
    On Error GoTo 0
    Debug.Print "[D-1479] modUILoader.ApplyFreeze EXIT-OK"  ' [ADR-0100]
End Sub

Private Sub ApplyColumn(ByVal ws As Worksheet, ByVal sec As ClsStanzaSection)
    Debug.Print "[D-1480] modUILoader.ApplyColumn ENTER"  ' [ADR-0100]
    Dim key As Variant
    Dim colLetter As String
    Dim widthVal As Double
    For Each key In sec.KeyValues.Keys
        If Left(CStr(key), 12) = "ColumnWidth_" Then
            colLetter = Mid(CStr(key), 13)
            widthVal = CDbl(sec.GetValue(CStr(key)))
            ws.Columns(colLetter).ColumnWidth = widthVal
        End If
    Next key
    Debug.Print "[D-1481] modUILoader.ApplyColumn EXIT-OK"  ' [ADR-0100]
End Sub

Private Sub ApplyRow(ByVal ws As Worksheet, ByVal sec As ClsStanzaSection)
    Debug.Print "[D-1482] modUILoader.ApplyRow ENTER"  ' [ADR-0100]
    Dim key As Variant
    Dim rowNum As Long
    Dim heightVal As Double
    For Each key In sec.KeyValues.Keys
        If Left(CStr(key), 10) = "RowHeight_" Then
            rowNum = CLng(Mid(CStr(key), 11))
            heightVal = CDbl(sec.GetValue(CStr(key)))
            ws.Rows(rowNum).RowHeight = heightVal
        End If
    Next key
    Debug.Print "[D-1483] modUILoader.ApplyRow EXIT-OK"  ' [ADR-0100]
End Sub

Private Sub ApplyFont(ByVal ws As Worksheet, ByVal sec As ClsStanzaSection)
    Debug.Print "[D-1484] modUILoader.ApplyFont ENTER"  ' [ADR-0100]
    Dim defaultFont As String
    Dim defaultSize As String
    defaultFont = sec.GetValue("DefaultFont")
    defaultSize = sec.GetValue("DefaultFontSize")
    If Len(defaultFont) > 0 Then
        ws.Cells.Font.Name = defaultFont
    End If
    If Len(defaultSize) > 0 Then
        ws.Cells.Font.Size = CDbl(defaultSize)
    End If
    Debug.Print "[D-1485] modUILoader.ApplyFont EXIT-OK"  ' [ADR-0100]
End Sub

Private Sub ApplyHeader(ByVal ws As Worksheet, ByVal sec As ClsStanzaSection)
    Debug.Print "[D-1486] modUILoader.ApplyHeader ENTER"  ' [ADR-0100]
    Dim cellAddr As String
    cellAddr = sec.GetValue("Cell")
    If Len(cellAddr) = 0 Then Exit Sub

    Dim targetRange As Range
    Set targetRange = ws.Range(cellAddr)

    ' (???R?????g????: ?????G???R?[?h?j???????????s?\)
    If InStr(cellAddr, ":") > 0 Then
        On Error Resume Next
        targetRange.Merge
        On Error GoTo 0
    End If

    ' (???R?????g????: ?????G???R?[?h?j???????????s?\)
    targetRange.Cells(1, 1).Value = UnescapeStanzaValue(sec.GetValue("Text"))
    If sec.HasKey("BackColor") Then
        targetRange.Interior.Color = HexToRgbLong(sec.GetValue("BackColor"))
    End If
    If sec.HasKey("ForeColor") Then
        targetRange.Font.Color = HexToRgbLong(sec.GetValue("ForeColor"))
    End If
    If sec.HasKey("FontSize") Then
        targetRange.Font.Size = CDbl(sec.GetValue("FontSize"))
    End If
    If sec.GetValue("Bold") = "TRUE" Then
        targetRange.Font.Bold = True
    End If
    Debug.Print "[D-1487] modUILoader.ApplyHeader EXIT-OK"  ' [ADR-0100]
End Sub

Private Sub ApplyLabel(ByVal ws As Worksheet, ByVal sec As ClsStanzaSection)
    Debug.Print "[D-1488] modUILoader.ApplyLabel ENTER"  ' [ADR-0100]
    Dim cellAddr As String
    cellAddr = sec.GetValue("Cell")
    If Len(cellAddr) = 0 Then Exit Sub
    Dim r As Range
    Set r = ws.Range(cellAddr)
    r.Value = modStanzaIO.UnescapeStanzaValue(sec.GetValue("Text"))
    ' Phase R-3-ψ-M03-style: 補助文ラベル等の色/サイズ対応(mock の灰色説明文)。
    On Error Resume Next
    If sec.HasKey("ForeColor") Then r.Font.Color = HexToRgbLong(sec.GetValue("ForeColor"))
    If sec.HasKey("FontSize") Then r.Font.Size = CDbl(sec.GetValue("FontSize"))
    On Error GoTo 0
    Debug.Print "[D-1489] modUILoader.ApplyLabel EXIT-OK"  ' [ADR-0100]
End Sub

' [CHECK] glyph-based pseudo-checkbox (M-10 row selector). Kept for the
' historical M-10 row picker; new screens should prefer the dynamic
' [CHECKBOX] / [LISTBOX] / [VALIDATION] stanzas dispatched via
' modFormControlWiring (Phase FC, ADR-0096) which use real Form Control
' shapes with a Validation-list fallback when AddFormControl fails.
Private Sub ApplyCheck(ByVal ws As Worksheet, ByVal sec As ClsStanzaSection)
    Debug.Print "[D-1490] modUILoader.ApplyCheck ENTER"  ' [ADR-0100]
    Dim cellAddr As String
    cellAddr = Trim(sec.GetValue("Cell"))
    If Len(cellAddr) = 0 Then Exit Sub

    Dim defaultChecked As Boolean
    defaultChecked = (UCase(Trim(sec.GetValue("DefaultChecked"))) = "TRUE")

    Dim r As Range
    Set r = ws.Range(cellAddr)
    If defaultChecked Then
        r.Value = ChrW(&H2611)  ' BALLOT BOX WITH CHECK
    Else
        r.Value = ChrW(&H2610)  ' BALLOT BOX
    End If
    r.HorizontalAlignment = xlCenter
    r.VerticalAlignment = xlCenter
    r.Font.Size = 14
    Debug.Print "[D-1491] modUILoader.ApplyCheck EXIT-OK"  ' [ADR-0100]
End Sub

' Q5=B full body: input cell color / border / data validation per schema 3.6.
Private Sub ApplyInput(ByVal ws As Worksheet, ByVal sec As ClsStanzaSection)
    Debug.Print "[D-1492] modUILoader.ApplyInput ENTER"  ' [ADR-0100]
    Dim cellAddr As String
    cellAddr = sec.GetValue("Cell")
    If Len(cellAddr) = 0 Then Exit Sub
    Dim r As Range
    Set r = ws.Range(cellAddr)
    ' Phase R-3-ψ-M03-style: range 指定(C4:E4 等)は merge して 1 矩形に
    ' (mock 準拠。未 merge だと各セル個別 border で 3 つの矩形に見える不具合の修正)。
    If InStr(cellAddr, ":") > 0 Then
        On Error Resume Next
        r.Merge
        On Error GoTo 0
    End If

    ' input background color: explicit InputColor, else schema 3.6 default #FFFFCC
    If sec.HasKey("InputColor") Then
        r.Interior.Color = HexToRgbLong(sec.GetValue("InputColor"))
    Else
        r.Interior.Color = HexToRgbLong("#FFFFCC")
    End If

    ' border (schema 3.6: none / thin / medium / thick, default thin)
    Dim borderVal As String
    If sec.HasKey("Border") Then
        borderVal = LCase(sec.GetValue("Border"))
    Else
        borderVal = "thin"
    End If
    ApplyBorder r, borderVal

    ' data validation by InputType (schema 3.6 enum)
    Dim inputType As String
    inputType = LCase(sec.GetValue("InputType"))
    ' [B28 2026-06-12] see ApplyDropdownValidation
    On Error Resume Next
    r.Worksheet.Unprotect
    On Error GoTo 0
    On Error Resume Next
    r.Validation.Delete
    On Error GoTo 0
    Select Case inputType
        Case "dropdown"
            ApplyDropdownValidation r, sec.GetValue("DropdownSource")
        Case "number"
            r.Validation.Add Type:=xlValidateWholeNumber, _
                AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
                Formula1:="-2147483647", Formula2:="2147483647"
        Case "date", "date_range"
            r.Validation.Add Type:=xlValidateDate, _
                AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
                Formula1:="1900/1/1", Formula2:="9999/12/31"
    End Select
    ' MaxLength (text-length data validation)
    If sec.HasKey("MaxLength") Then
        Dim maxLen As String
        maxLen = sec.GetValue("MaxLength")
        If Len(maxLen) > 0 Then
            On Error Resume Next
            r.Validation.Delete
            r.Validation.Add Type:=xlValidateTextLength, _
                AlertStyle:=xlValidAlertStop, Operator:=xlLessEqual, _
                Formula1:=maxLen
            On Error GoTo 0
        End If
    End If

    ' optional initial value (LinkedField pre-fill not in scope here)
    If sec.HasKey("Value") Then
        r.Cells(1, 1).Value = modStanzaIO.UnescapeStanzaValue(sec.GetValue("Value"))
    End If

    ' Phase R-3-omega: input cells must be editable while the sheet is protected
    ' (ProtectSheet "light"). Unlock so the user can type (e.g. M-08 keyword/format).
    On Error Resume Next
    r.Locked = False
    On Error GoTo 0
    Debug.Print "[D-1493] modUILoader.ApplyInput EXIT-OK"  ' [ADR-0100]
End Sub

' Apply a border style to a range. styleName per schema: none/thin/medium/thick.
Private Sub ApplyBorder(ByVal r As Range, ByVal styleName As String)
    Debug.Print "[D-1494] modUILoader.ApplyBorder ENTER"  ' [ADR-0100]
    Select Case styleName
        Case "none"
            r.Borders.LineStyle = xlLineStyleNone
        Case "thin"
            r.Borders.LineStyle = xlContinuous
            r.Borders.Weight = xlThin
        Case "medium"
            r.Borders.LineStyle = xlContinuous
            r.Borders.Weight = xlMedium
        Case "thick"
            r.Borders.LineStyle = xlContinuous
            r.Borders.Weight = xlThick
        Case Else
            r.Borders.LineStyle = xlContinuous
            r.Borders.Weight = xlThin
    End Select
    Debug.Print "[D-1495] modUILoader.ApplyBorder EXIT-OK"  ' [ADR-0100]
End Sub

' Apply list data validation from a DropdownSource string.
' Supported (schema 3.6): static:v1,v2,...  /  format:<id>:fields:<no> (deferred).
Private Sub ApplyDropdownValidation(ByVal r As Range, ByVal source As String)
    Debug.Print "[D-1496] modUILoader.ApplyDropdownValidation ENTER"  ' [ADR-0100]
    ' [B28 2026-06-12] Validation writes fail silently on a protected sheet even
    ' with UserInterfaceOnly (refresh path lost dropdowns). Unprotect here;
    ' the pipeline re-protects at its ApplyProtection step.
    On Error Resume Next
    r.Worksheet.Unprotect
    On Error GoTo 0
    If Len(source) = 0 Then Exit Sub
    Dim listCsv As String
    If Left(source, 7) = "static:" Then
        listCsv = Mid(source, 8)
    ElseIf Left(source, 7) = "format:" Then
        ' format:list -> dynamic load all format names from format_dir
        ' [USER-REQ 2026-06-09] Display FormatName instead of FormatID.
        If LCase(Mid(source, 8)) = "list" Then
            Dim formats As Collection
            On Error Resume Next
            Set formats = modFormatLoader.LoadFormatList()
            On Error GoTo 0
            If formats Is Nothing Then
                Debug.Print "[D-1497] modUILoader.ApplyDropdownValidation EXIT-OK"  ' [ADR-0100]
                Exit Sub
            End If
            Dim fi As Long
            For fi = 1 To formats.Count
                Dim ent As Object
                Set ent = formats.Item(fi)
                Dim nmStr As String
                nmStr = ""
                On Error Resume Next
                If ent.Exists("Description") Then nmStr = CStr(ent("Description"))
                If Len(nmStr) = 0 And ent.Exists("FormatID") Then nmStr = CStr(ent("FormatID"))
                On Error GoTo 0
                If Len(nmStr) > 0 Then
                    If Len(listCsv) > 0 Then listCsv = listCsv & ","
                    listCsv = listCsv & nmStr
                End If
            Next fi
        Else
            Debug.Print "[D-1497] modUILoader.ApplyDropdownValidation EXIT-OK"  ' [ADR-0100]
            Exit Sub
        End If
    Else
        ' bare CSV fallback
        listCsv = source
    End If
    If Len(listCsv) = 0 Then Exit Sub
    On Error Resume Next
    r.Validation.Delete
    r.Validation.Add Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
        Formula1:=listCsv
    On Error GoTo 0
End Sub

' S3-RND-02 minimal body: Q29 NOTE subkey (Text only); Position/FontSize/Italic deferred.
Private Sub ApplyNote(ByVal ws As Worksheet, ByVal sec As ClsStanzaSection)
    Debug.Print "[D-1498] modUILoader.ApplyNote ENTER"  ' [ADR-0100]
    Dim cellAddr As String
    cellAddr = sec.GetValue("Cell")
    If Len(cellAddr) = 0 Then Exit Sub
    Dim r As Range
    Set r = ws.Range(cellAddr)
    If InStr(cellAddr, ":") > 0 Then
        On Error Resume Next
        r.Merge
        On Error GoTo 0
    End If
    r.Cells(1, 1).Value = modStanzaIO.UnescapeStanzaValue(sec.GetValue("Text"))
    r.WrapText = True
    ' Phase R-3-ψ-M03-style (差5): 背景色 + 枠。mock ??測 BackColor=#DEEBF7 / 枠#BFBFBF。
    On Error Resume Next
    If sec.HasKey("BackColor") Then r.Interior.Color = HexToRgbLong(sec.GetValue("BackColor"))
    Dim nb As String
    nb = LCase(Trim(sec.GetValue("Border")))
    If Len(nb) > 0 Then ApplyBorder r, nb
    r.VerticalAlignment = xlTop
    On Error GoTo 0
    Debug.Print "[D-1499] modUILoader.ApplyNote EXIT-OK"  ' [ADR-0100]
End Sub

' ================================================================
' ?????: ApplyGrid (v2.2 task 3.11 ???C)
' ?T?v:   [GRID] ?Z?N?V??????w?b?_?s??`????B?w?b?_?a???E???E????E
'         ?w?i?F??\????p?L?[ ColumnHeaders (OP-1 ?m??A3.11.1) ??????B
'         Columns (?_????) ??w?b?_?`???g?????B?_??????f?[?^???o??
'         ?R?[?h (clsGridIO / modEntrySearch) ??p??{??????G?????B
'         ?f?[?^?s (StartCell..EndCell) ??l????? cls ???[?U????B
' ????:   ByVal ws As Worksheet         - ?K?p??V?[?g
'         ByVal sec As ClsStanzaSection - [GRID] ?Z?N?V????
' ???l:   StartCell ?? v2.2 ????f?[?^?Z????m?? (3.11)?B?w?b?_?s??????L?[
'         HeaderRow ??`?????? StartCell ????s????B
' ================================================================
Private Sub ApplyGrid(ByVal ws As Worksheet, ByVal sec As ClsStanzaSection)
    Debug.Print "[D-1500] modUILoader.ApplyGrid ENTER"  ' [ADR-0100]
    Dim startCell As String
    startCell = sec.GetValue("StartCell")
    If Len(startCell) = 0 Then Exit Sub

    ' StartCell = ???f?[?^?Z?? (3.11 v2.2 ?m??)
    Dim anchor As Range
    Set anchor = ws.Range(startCell).Cells(1, 1)

    ' ?w?b?_?s???: ???? HeaderRow?A??????? StartCell ?s
    Dim headerRow As Long
    If sec.HasKey("HeaderRow") And Len(sec.GetValue("HeaderRow")) > 0 Then
        headerRow = CLng(sec.GetValue("HeaderRow"))
    Else
        headerRow = anchor.Row
    End If

    ' Columns (?_???? CSV) ??????w?b?_???`??A??????? Name ?????u??
    Dim colsCsv As String
    colsCsv = sec.GetValue("Columns")
    If Len(colsCsv) > 0 Then
        ApplyGridColumns ws, sec, anchor, headerRow, colsCsv
    Else
        Dim gridName As String
        gridName = sec.GetValue("Name")
        If Len(gridName) > 0 Then
            anchor.Value = modStanzaIO.UnescapeStanzaValue(gridName)
        End If
    End If

    ' Phase R-3-ψ-M03-style (差4): 見出し行 + 空データ行(InitialRows)に罫線(#BFBFBF)。
    ' 親確定=案B(空白罫線のみ、??データ表示時に上書き)。見出しは太字 + 任意 HeaderBackColor。
    On Error Resume Next
    Dim nCols As Long
    If Len(colsCsv) > 0 Then
        nCols = UBound(Split(colsCsv, ",")) - LBound(Split(colsCsv, ",")) + 1
    Else
        nCols = 1
    End If
    Dim initRows As Long
    initRows = 0
    If sec.HasKey("InitialRows") And IsNumeric(sec.GetValue("InitialRows")) Then initRows = CLng(sec.GetValue("InitialRows"))
    If nCols > 0 Then
        Dim gl As Long, gr As Long
        gl = anchor.Column
        gr = anchor.Column + nCols - 1
        Dim grng As Range
        Set grng = ws.Range(ws.Cells(headerRow, gl), ws.Cells(headerRow + initRows, gr))
        grng.Borders.LineStyle = xlContinuous
        grng.Borders.Weight = xlThin
        grng.Borders.Color = HexToRgbLong("#BFBFBF")
        Dim hrng As Range
        Set hrng = ws.Range(ws.Cells(headerRow, gl), ws.Cells(headerRow, gr))
        hrng.Font.Bold = True
        If sec.HasKey("HeaderBackColor") Then hrng.Interior.Color = HexToRgbLong(sec.GetValue("HeaderBackColor"))
        ' [USERINPUT 2026-06-06] Unlock data cells (header+1 to header+initRows)
        ' so user can type into GRID even with sheet protection enabled.
        Dim drng As Range
        Set drng = ws.Range(ws.Cells(headerRow + 1, gl), ws.Cells(headerRow + initRows, gr))
        drng.Locked = False
        ' [USERINPUT 2026-06-06] Pre-fill No column + add data validation dropdowns
        ' Iterate column names from Columns CSV.
        Dim colArr() As String
        colArr = Split(colsCsv, ",")
        Dim ci As Long
        For ci = LBound(colArr) To UBound(colArr)
            Dim colName As String
            colName = Trim$(colArr(ci))
            Dim colAbs As Long
            colAbs = gl + ci
            Dim colRng As Range
            Set colRng = ws.Range(ws.Cells(headerRow + 1, colAbs), ws.Cells(headerRow + initRows, colAbs))
            Select Case LCase$(colName)
                Case "seq"
                    Dim ni As Long
                    For ni = 1 To initRows
                        ws.Cells(headerRow + ni, colAbs).Value = ni
                    Next ni
                Case "fieldtype"
                    colRng.Validation.Delete
                    colRng.Validation.Add Type:=xlValidateList, Formula1:= _
                        ChrW(&H0031) & ChrW(&H884C) & ChrW(&H30C6) & ChrW(&H30AD) & ChrW(&H30B9) & ChrW(&H30C8) & "," & _
                        ChrW(&H8907) & ChrW(&H6570) & ChrW(&H884C) & ChrW(&H30C6) & ChrW(&H30AD) & ChrW(&H30B9) & ChrW(&H30C8) & "," & _
                        ChrW(&H6570) & ChrW(&H5024) & "," & _
                        ChrW(&H65E5) & ChrW(&H4ED8) & "," & _
                        ChrW(&H9078) & ChrW(&H629E) & "," & _
                        ChrW(&H30C1) & ChrW(&H30A7) & ChrW(&H30C3) & ChrW(&H30AF)
                    colRng.Validation.IgnoreBlank = True
                    colRng.Validation.InCellDropdown = True
                Case "required"
                    colRng.Validation.Delete
                    colRng.Validation.Add Type:=xlValidateList, Formula1:="TRUE,FALSE"
                    colRng.Validation.IgnoreBlank = True
                    colRng.Validation.InCellDropdown = True
                Case "scroll"
                    colRng.Validation.Delete
                    colRng.Validation.Add Type:=xlValidateList, Formula1:="TRUE,FALSE"
                    colRng.Validation.IgnoreBlank = True
                    colRng.Validation.InCellDropdown = True
                ' [B36 2026-06-12] searchTarget was missing from this Select
                ' Case, so seed/refresh never applied the TRUE/FALSE dropdown
                ' to column I while fieldType/required/scroll got theirs
                ' (UAT re-report 2026-06-12: M-03 search-target listbox).
                Case "searchtarget"
                    colRng.Validation.Delete
                    colRng.Validation.Add Type:=xlValidateList, Formula1:="TRUE,FALSE"
                    colRng.Validation.IgnoreBlank = True
                    colRng.Validation.InCellDropdown = True
            End Select
        Next ci
        ' [USERINPUT 2026-06-06] M-03 preset rows: only when GRID data is empty (initial render)
        ' Preset demonstrates field type / required / rows / options / scroll variations.
        If Len(CStr(ws.Cells(headerRow + 1, gl + 1).Value)) = 0 Then
            ' Determine column positions by column name
            Dim colIdx As Object
            Set colIdx = CreateObject("Scripting.Dictionary")
            colIdx.CompareMode = 1
            For ci = LBound(colArr) To UBound(colArr)
                colIdx(Trim$(colArr(ci))) = gl + ci
            Next ci
            ' Preset row 1: title (single-line, required)
            If colIdx.Exists("fieldName") Then ws.Cells(headerRow + 1, colIdx("fieldName")).Value = ChrW(&H4EF6) & ChrW(&H540D)
            If colIdx.Exists("fieldType") Then ws.Cells(headerRow + 1, colIdx("fieldType")).Value = ChrW(&H0031) & ChrW(&H884C) & ChrW(&H30C6) & ChrW(&H30AD) & ChrW(&H30B9) & ChrW(&H30C8)
            If colIdx.Exists("required") Then ws.Cells(headerRow + 1, colIdx("required")).Value = "TRUE"
            ' Preset row 2: content (multi-line, required, scroll, rows=5)
            If colIdx.Exists("fieldName") Then ws.Cells(headerRow + 2, colIdx("fieldName")).Value = ChrW(&H5185) & ChrW(&H5BB9)
            If colIdx.Exists("fieldType") Then ws.Cells(headerRow + 2, colIdx("fieldType")).Value = ChrW(&H8907) & ChrW(&H6570) & ChrW(&H884C) & ChrW(&H30C6) & ChrW(&H30AD) & ChrW(&H30B9) & ChrW(&H30C8)
            If colIdx.Exists("required") Then ws.Cells(headerRow + 2, colIdx("required")).Value = "TRUE"
            If colIdx.Exists("rows") Then ws.Cells(headerRow + 2, colIdx("rows")).Value = 5
            If colIdx.Exists("scroll") Then ws.Cells(headerRow + 2, colIdx("scroll")).Value = "TRUE"
            ' Preset row 3: notes (multi-line, optional, no-scroll, rows=3)
            If colIdx.Exists("fieldName") Then ws.Cells(headerRow + 3, colIdx("fieldName")).Value = ChrW(&H5099) & ChrW(&H8003)
            If colIdx.Exists("fieldType") Then ws.Cells(headerRow + 3, colIdx("fieldType")).Value = ChrW(&H8907) & ChrW(&H6570) & ChrW(&H884C) & ChrW(&H30C6) & ChrW(&H30AD) & ChrW(&H30B9) & ChrW(&H30C8)
            If colIdx.Exists("required") Then ws.Cells(headerRow + 3, colIdx("required")).Value = "FALSE"
            If colIdx.Exists("rows") Then ws.Cells(headerRow + 3, colIdx("rows")).Value = 3
            If colIdx.Exists("scroll") Then ws.Cells(headerRow + 3, colIdx("scroll")).Value = "FALSE"
            ' Preset row 4: created date (date, optional)
            If colIdx.Exists("fieldName") Then ws.Cells(headerRow + 4, colIdx("fieldName")).Value = ChrW(&H4F5C) & ChrW(&H6210) & ChrW(&H65E5)
            If colIdx.Exists("fieldType") Then ws.Cells(headerRow + 4, colIdx("fieldType")).Value = ChrW(&H65E5) & ChrW(&H4ED8)
            If colIdx.Exists("required") Then ws.Cells(headerRow + 4, colIdx("required")).Value = "FALSE"
            ' Preset row 5: category (select, options)
            If colIdx.Exists("fieldName") Then ws.Cells(headerRow + 5, colIdx("fieldName")).Value = ChrW(&H30AB) & ChrW(&H30C6) & ChrW(&H30B4) & ChrW(&H30EA)
            If colIdx.Exists("fieldType") Then ws.Cells(headerRow + 5, colIdx("fieldType")).Value = ChrW(&H9078) & ChrW(&H629E)
            If colIdx.Exists("required") Then ws.Cells(headerRow + 5, colIdx("required")).Value = "FALSE"
            If colIdx.Exists("options") Then ws.Cells(headerRow + 5, colIdx("options")).Value = ChrW(&H91CD) & ChrW(&H8981) & ChrW(&H5EA6) & "1," & ChrW(&H91CD) & ChrW(&H8981) & ChrW(&H5EA6) & "2," & ChrW(&H91CD) & ChrW(&H8981) & ChrW(&H5EA6) & "3"
        End If
    End If
    On Error GoTo 0

    ' FreezeHeader (schema 3.11?A???? TRUE)?Bheadless ??? ActiveWindow ?s???
    Dim freezeHdr As String
    If sec.HasKey("FreezeHeader") Then
        freezeHdr = UCase(sec.GetValue("FreezeHeader"))
    Else
        freezeHdr = "TRUE"
    End If
    If freezeHdr = "TRUE" Then
        On Error Resume Next
        ws.Activate
        ws.Cells(headerRow + 1, anchor.Column).Select
        ws.Application.ActiveWindow.FreezePanes = True
        On Error GoTo 0
    End If
    Debug.Print "[D-1501] modUILoader.ApplyGrid EXIT-OK"  ' [ADR-0100]
End Sub

' ================================================================
' ?????: ApplyGridColumns (v2.2 task 3.11)
' ?T?v:   [GRID] ??w?b?_??? 1 ????`????BColumns ??_???? CSV?A
'         ?\???????? ColumnHeaders (?a??:??:????:?w?i?F ?? CSV) ??????B
'         ?? i ?? Columns[i] ?? ColumnHeaders[i] ?? 1:1 ??? (3.11.1)?B
'         ColumnHeaders ?? Columns ??????/?s?????A?_??????
'         ?w?b?_???????t?H?[???o?b?N????B
' ????:   ByVal ws As Worksheet         - ?K?p??V?[?g
'         ByVal sec As ClsStanzaSection - [GRID] ?Z?N?V????
'         ByVal anchor As Range         - ???f?[?^?Z?? (StartCell)
'         ByVal headerRow As Long       - ?w?b?_?s???
'         ByVal colsCsv As String       - Columns (?_???? CSV)
' ================================================================
Private Sub ApplyGridColumns( _
    ByVal ws As Worksheet, _
    ByVal sec As ClsStanzaSection, _
    ByVal anchor As Range, _
    ByVal headerRow As Long, _
    ByVal colsCsv As String _
)
    Debug.Print "[D-1502] modUILoader.ApplyGridColumns ENTER"  ' [ADR-0100]
    Dim logicalCols() As String
    logicalCols = Split(colsCsv, ",")

    ' ?\???????L?[ ColumnHeaders ????? (?C??L?[?AOP-1 ?m??)
    Dim hdrCsv As String
    hdrCsv = sec.GetValue("ColumnHeaders")
    Dim hdrDefs() As String
    Dim hdrCount As Long
    hdrCount = 0
    If Len(hdrCsv) > 0 Then
        hdrDefs = Split(hdrCsv, ",")
        hdrCount = UBound(hdrDefs) - LBound(hdrDefs) + 1
    Else
        ' ColumnHeaders ?s??: ?S???_???????w?b?_??t?H?[???o?b?N + WARN
        Debug.Print "[UILoader] WARN: [GRID] '" & sec.GetValue("Name") & _
            "' ?? ColumnHeaders ?s??????_???????w?b?_??g?p"
    End If

    ' ??????f?[?^?????K?p??????f?[?^???????s???????
    Dim dataEndRow As Long
    dataEndRow = GetGridDataEndRow(ws, sec)

    Dim i As Long
    For i = LBound(logicalCols) To UBound(logicalCols)
        Dim colIdx As Long
        colIdx = i - LBound(logicalCols)
        Dim headerCell As Range
        Set headerCell = ws.Cells(headerRow, anchor.Column + colIdx)
        If colIdx < hdrCount Then
            ApplyGridHeaderColumn ws, headerCell, anchor.Column + colIdx, _
                anchor.Row, dataEndRow, Trim(hdrDefs(LBound(hdrDefs) + colIdx))
        Else
            ' ColumnHeaders ?s??????_???????w?b?_???????t?H?[???o?b?N
            headerCell.Value = Trim(logicalCols(i))
        End If
    Next i
    Debug.Print "[D-1503] modUILoader.ApplyGridColumns EXIT-OK"  ' [ADR-0100]
End Sub

' ================================================================
' ?????: ApplyGridHeaderColumn (v2.2 task 3.11)
' ?T?v:   ColumnHeaders ?? 1 ?g?[?N?? (?a??:??:????:?w?i?F) ?? 1 ???K?p?B
'         ?a??=?w?b?_?Z???l?A??=???A????=?w?b?_+?f?[?^??????????A
'         ?w?i?F=?w?b?_?Z???w?i?B??t?B?[???h??K?p????? (3.11.1 ??? 3)?B
' ????:   ByVal ws As Worksheet       - ?K?p??V?[?g
'         ByVal headerCell As Range   - ?w?b?_?Z??
'         ByVal colNumber As Long     - ?????? (?????)
'         ByVal dataStartRow As Long  - ?f?[?^?????s
'         ByVal dataEndRow As Long    - ?f?[?^???????s (0 ???j????f?[?^
'                                       ???K?p????? = EndCell ?s????)
'         ByVal headerToken As String - ColumnHeaders ?? 1 ?g?[?N??
' ================================================================
Private Sub ApplyGridHeaderColumn( _
    ByVal ws As Worksheet, _
    ByVal headerCell As Range, _
    ByVal colNumber As Long, _
    ByVal dataStartRow As Long, _
    ByVal dataEndRow As Long, _
    ByVal headerToken As String _
)
    Debug.Print "[D-1504] modUILoader.ApplyGridHeaderColumn ENTER"  ' [ADR-0100]
    Dim parts() As String
    parts = Split(headerToken, ":")

    ' ?a?? (part 0): ?w?b?_?Z????l
    If UBound(parts) >= 0 Then
        headerCell.Value = modStanzaIO.UnescapeStanzaValue(Trim(parts(0)))
    End If

    ' ?? (part 1): ???Y????
    If UBound(parts) >= 1 Then
        If Len(Trim(parts(1))) > 0 Then
            ws.Columns(colNumber).ColumnWidth = CDbl(Trim(parts(1)))
        End If
    End If

    ' ???? (part 2): ?w?b?_?Z?? + ?f?[?^??????????
    If UBound(parts) >= 2 Then
        If Len(Trim(parts(2))) > 0 Then
            Dim hAlign As Long
            hAlign = HAlignToEnum(Trim(parts(2)))
            headerCell.HorizontalAlignment = hAlign
            If dataEndRow >= dataStartRow Then
                ws.Range(ws.Cells(dataStartRow, colNumber), _
                         ws.Cells(dataEndRow, colNumber)).HorizontalAlignment = hAlign
            End If
        End If
    End If

    ' ?w?i?F (part 3): ?w?b?_?Z???w?i
    If UBound(parts) >= 3 Then
        If Len(Trim(parts(3))) > 0 Then
            headerCell.Interior.Color = HexToRgbLong(Trim(parts(3)))
        End If
    End If
    Debug.Print "[D-1505] modUILoader.ApplyGridHeaderColumn EXIT-OK"  ' [ADR-0100]
End Sub

' ================================================================
' ?????: GetGridDataEndRow (v2.2 task 3.11)
' ?T?v:   [GRID].EndCell (v2.2 ??K?{???????f?[?^??E???Z??) ????
'         ?f?[?^???????s????????BEndCell ?s??E?s?????? 0 ?????B
' ????:   ByVal ws As Worksheet         - ???V?[?g
'         ByVal sec As ClsStanzaSection - [GRID] ?Z?N?V????
' ???l: Long - ?f?[?^???????s??? (0 = ?s??)
' ================================================================
Private Function GetGridDataEndRow( _
    ByVal ws As Worksheet, _
    ByVal sec As ClsStanzaSection _
) As Long
    Debug.Print "[D-1506] modUILoader.GetGridDataEndRow ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Dim endCell As String
    endCell = sec.GetValue("EndCell")
    If Len(endCell) = 0 Then
        GetGridDataEndRow = 0
        Debug.Print "[D-1507] modUILoader.GetGridDataEndRow EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    GetGridDataEndRow = ws.Range(endCell).Cells(1, 1).Row
    Debug.Print "[D-1508] modUILoader.GetGridDataEndRow EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    Debug.Print "[D-1509] modUILoader.GetGridDataEndRow EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    GetGridDataEndRow = 0
End Function

' ----------------------------------------------------------------
' Public/Private: ???I?`??n???h?? (v2.2 task 3.12?AOP-2 ?m??)
' [FORM_FROM_FORMAT] / [GRID_FROM_FORMAT] ??t?H?[?}?b?g??` .txt ??
' [FIELD] ?Q????A??????E?^???????t?H?[??/?O???b?h???I?`????B
' ApplyFormFromFormat / ApplyGridFromFormat ????J I/F?B?????
' ApplyUiToSheet ???A?t?H?[?}?b?g?I???X?E?i???b?W????????? cls ??
' ???o?????`???? (v2_ui_stanza_schema.md 3.12 ??? 7)?B
' ----------------------------------------------------------------

' ================================================================
' ?????: ApplyFormFromFormat (v2.2 task 3.12?A???J I/F)
' ?T?v:   [FORM_FROM_FORMAT] ?Z?N?V??????`????BSourceFormatRef ?Z????
'         FormatID ????t?H?[?}?b?g??`????A[FIELD] ?Q?? 1 ?s 1 ?????
'         ????t?H?[??????I?`????BPopulateFromKnowledge ?w????????l
'         ?????????[?U???AReadOnly / ReadOnlyFormat ?????????b?N????
'         (v2_ui_stanza_schema.md 3.12 ??? 1-6?B?f?[?^?w???????????
'         ?t?H?[???AM-05 ?o?^ / M-06 ?C?? / M-09 ?\?????g?p)?B
' ????:   ByVal ws As Worksheet         - ?K?p??V?[?g
'         ByVal sec As ClsStanzaSection - [FORM_FROM_FORMAT] ?Z?N?V????
' ================================================================
Public Sub ApplyFormFromFormat(ByVal ws As Worksheet, ByVal sec As ClsStanzaSection)
    Debug.Print "[D-1510] modUILoader.ApplyFormFromFormat ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler

    ' FieldName -> ??????Z????n ?????\ (?[?U?E???b?N??g?p)
    Dim fieldCells As Object
    Set fieldCells = CreateObject("Scripting.Dictionary")

    ' ??? 1-4: ?w?b?_ + [FIELD] ?f?[?^?s??`??
    Dim fieldCount As Long
    fieldCount = RenderFormatDriven(ws, sec, fieldCells)
    If fieldCount < 0 Then Exit Sub  ' ?t?H?[?}?b?g???I?? (?w?b?_???`???)

    ' ??? 5: PopulateFromKnowledge ?w?????i???b?W????l?????????[?U
    PopulateFormCells ws, sec, fieldCells

    ' ??? 6: ReadOnly=TRUE ?????????AReadOnlyFormat=TRUE ??
    '         SourceFormatRef ?Z???????b?N????
    ApplyFormLocks ws, sec, fieldCells

    Debug.Print "[D-1511] modUILoader.ApplyFormFromFormat EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    Debug.Print "[D-1512] modUILoader.ApplyFormFromFormat EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Debug.Print "[UILoader] ApplyFormFromFormat error: " & Err.Description
End Sub

' ================================================================
' ?????: ApplyGridFromFormat (v2.2 task 3.12?A???J I/F)
' ?T?v:   [GRID_FROM_FORMAT] ?Z?N?V??????`????BSourceFormatRef ??
'         FormatID ????t?H?[?}?b?g??`????A[FIELD] ?Q?? 1 ?s 1 ?????
'         ?v???r???[?O???b?h????I?`????B?f?[?^?w???????????????????
'         ?????p?????A?i???b?W?[?U?E???b?N??s????
'         (v2_ui_stanza_schema.md 3.13?A??? 1-4 ???BM-04 ???g?p)?B
' ????:   ByVal ws As Worksheet         - ?K?p??V?[?g
'         ByVal sec As ClsStanzaSection - [GRID_FROM_FORMAT] ?Z?N?V????
' ================================================================
Public Sub ApplyGridFromFormat(ByVal ws As Worksheet, ByVal sec As ClsStanzaSection)
    Debug.Print "[D-1513] modUILoader.ApplyGridFromFormat ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler

    Dim fieldCells As Object
    Set fieldCells = CreateObject("Scripting.Dictionary")

    ' ??? 1-4 ??? (?[?U?E???b?N???)
    RenderFormatDriven ws, sec, fieldCells

    Debug.Print "[D-1514] modUILoader.ApplyGridFromFormat EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    Debug.Print "[D-1515] modUILoader.ApplyGridFromFormat EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Debug.Print "[UILoader] ApplyGridFromFormat error: " & Err.Description
End Sub

' ================================================================
' ?????: RenderFormatDriven (v2.2 task 3.12)
' ?T?v:   [FORM_FROM_FORMAT] / [GRID_FROM_FORMAT] ?????`???? (3.12
'         ??? 1-4)?B4 ??w?b?_??`????ASourceFormatRef ????t?H?[?}?b?g??
'         ???????? [FIELD] ?Q?? 1 ?s????`????B
' ????:   ByVal ws As Worksheet           - ?K?p??V?[?g
'         ByVal sec As ClsStanzaSection   - ???Z?N?V????
'         ByRef outFieldCells As Object   - ?o??: FieldName -> ???????n
' ???l: Long - ?`????? [FIELD] ?????BStartCell ?s?????? SourceFormatRef
'         ?? (?t?H?[?}?b?g???I??) ???? -1
' ================================================================
Private Function RenderFormatDriven( _
    ByVal ws As Worksheet, _
    ByVal sec As ClsStanzaSection, _
    ByRef outFieldCells As Object _
) As Long
    Debug.Print "[D-1516] modUILoader.RenderFormatDriven ENTER"  ' [ADR-0100]
    Dim startCell As String
    startCell = sec.GetValue("StartCell")
    If Len(startCell) = 0 Then
        RenderFormatDriven = -1
        Debug.Print "[D-1517] modUILoader.RenderFormatDriven EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If

    ' StartCell = ???I?t?H?[???????Z?? (?w?b?_?s)?B?f?[?^?s????? 1 ?s??????
    Dim anchor As Range
    Set anchor = ws.Range(startCell).Cells(1, 1)

    ' ??? 3 ?O??: 4 ??w?b?_?s??`?? (?t?H?[?}?b?g???I??????w?b?_??`??)
    DrawFormatHeaderRow ws, sec, anchor

    ' ??? 1: SourceFormatRef ?Z?? -> FormatID
    Dim formatId As String
    formatId = GetSourceFormatId(ws, sec)
    If Len(formatId) = 0 Then
        Debug.Print "[UILoader] WARN: " & sec.SectionName & _
            " ?? SourceFormatRef ????????w?b?_???`?? (?t?H?[?}?b?g???I??)"
        RenderFormatDriven = -1
        Debug.Print "[D-1518] modUILoader.RenderFormatDriven EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If

    ' ??? 2: ?t?H?[?}?b?g??`?? [FIELD] ?Q???L?q?? (= ?\????) ????
    Dim fields As Collection
    Set fields = CollectFieldSections(formatId)

    ' ??? 4: [FIELD] ????? 1 ?s????f?[?^?s??`?? (?w?b?_?s?? i ?s??)
    Dim i As Long
    For i = 1 To fields.Count
        Dim fieldSec As ClsStanzaSection
        Set fieldSec = fields.Item(i)
        DrawFormatFieldRow ws, anchor, anchor.Row + i, fieldSec, outFieldCells
    Next i

    RenderFormatDriven = fields.Count
End Function

' ================================================================
' ?????: GetFormatDrivenFieldCells (v2.2 ?c???, BuildDictFromCells ?A?g)
' ?T?v:   [FORM_FROM_FORMAT] / [GRID_FROM_FORMAT] ???`???????I?t?H?[????
'         FieldName -> ???????n ?????\???A??`?????Z?o???????B
'         RenderFormatDriven ????????C?A?E?g?K?? (StartCell ?s +1+i?A
'         ????? = StartCell ?? + FORM_COL_INPUT) ?????L??????A?Z?????
'         ?????????s??????????p?N?G???????B
'         modEntryKnowledge.BuildDictFromCells ?????I?t?H?[???????l??
'         ???W???????g?p???? (v2_ui_stanza_schema.md ??3.12 ??? 4)?B
' ????:   ByVal ws As Worksheet         - ???V?[?g
'         ByVal sec As ClsStanzaSection - [FORM_FROM_FORMAT] ?Z?N?V????
' ???l: Object - FieldName -> ???????n ?? Dictionary?B?t?H?[?}?b?g
'                  ???I???E[FIELD] 0 ???EStartCell ?s????????
'                  Dictionary (Nothing ????????o????????s?v?????)?B
' ================================================================
Public Function GetFormatDrivenFieldCells( _
    ByVal ws As Worksheet, _
    ByVal sec As ClsStanzaSection _
) As Object
    Debug.Print "[D-1519] modUILoader.GetFormatDrivenFieldCells ENTER"  ' [ADR-0100]
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    Set GetFormatDrivenFieldCells = result
    On Error GoTo ErrHandler

    ' StartCell = ???I?t?H?[???????Z?? (?w?b?_?s)
    Dim startCell As String
    startCell = sec.GetValue("StartCell")
    If Len(startCell) = 0 Then Exit Function
    Dim anchor As Range
    Set anchor = ws.Range(startCell).Cells(1, 1)

    ' SourceFormatRef ?Z?? -> FormatID?B???I?? (??) ????????????
    Dim formatId As String
    formatId = GetSourceFormatId(ws, sec)
    If Len(formatId) = 0 Then Exit Function

    ' ?t?H?[?}?b?g??`?? [FIELD] ?Q???L?q?? (= ?\????) ????
    Dim fields As Collection
    Set fields = CollectFieldSections(formatId)

    ' [FIELD] i ?????? = ?w?b?_?s +i ?s??EStartCell ?? +FORM_COL_INPUT
    Dim i As Long
    For i = 1 To fields.Count
        Dim fieldSec As ClsStanzaSection
        Set fieldSec = fields.Item(i)
        Dim fieldName As String
        fieldName = fieldSec.GetValue("FieldName")
        If Len(fieldName) > 0 Then
            result(fieldName) = _
                ws.Cells(anchor.Row + i, anchor.Column + FORM_COL_INPUT).Address
        End If
    Next i
    Debug.Print "[D-1520] modUILoader.GetFormatDrivenFieldCells EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    Debug.Print "[D-1521] modUILoader.GetFormatDrivenFieldCells EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Debug.Print "[UILoader] GetFormatDrivenFieldCells error: " & Err.Description
End Function

' ================================================================
' ?????: GetSourceFormatId (v2.2 task 3.12)
' ?T?v:   SourceFormatRef ?L?[ (A1 ?Q??A?V?[?g?C????) ???w???Z???????????A
'         ???????????? FormatID ??????????B??E?????s?\?????????B
' ????:   ByVal ws As Worksheet         - ??T?V?[?g
'         ByVal sec As ClsStanzaSection - ???Z?N?V????
' ???l: String - FormatID (???I???E???s?????????)
' ================================================================
Private Function GetSourceFormatId( _
    ByVal ws As Worksheet, _
    ByVal sec As ClsStanzaSection _
) As String
    Debug.Print "[D-1522] modUILoader.GetSourceFormatId ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Dim ref As String
    ref = sec.GetValue("SourceFormatRef")
    If Len(ref) = 0 Then
        GetSourceFormatId = vbNullString
        Debug.Print "[D-1523] modUILoader.GetSourceFormatId EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    Dim refCell As Range
    Set refCell = ResolveCellRef(ws, ref)
    If refCell Is Nothing Then
        GetSourceFormatId = vbNullString
        Debug.Print "[D-1524] modUILoader.GetSourceFormatId EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    GetSourceFormatId = Trim(CStr(refCell.Cells(1, 1).Value))
    Debug.Print "[D-1525] modUILoader.GetSourceFormatId EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    Debug.Print "[D-1526] modUILoader.GetSourceFormatId EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    GetSourceFormatId = vbNullString
End Function

' ================================================================
' ?????: ResolveCellRef (v2.2 task 3.12)
' ?T?v:   A1 ?Q???????? Range ?????????B"!" ????????V?[?g?C???Q??
'         (?? M-03!C3) ????????u?b?N???V?[?g???Q?????B
' ????:   ByVal ws As Worksheet     - ??T?V?[?g (??C???Q????????)
'         ByVal refStr As String    - A1 ?Q???????
' ???l: Range - ???????????B?????s?\???? Nothing
' ================================================================
Private Function ResolveCellRef( _
    ByVal ws As Worksheet, _
    ByVal refStr As String _
) As Range
    Debug.Print "[D-1527] modUILoader.ResolveCellRef ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Dim sepPos As Long
    sepPos = InStr(refStr, "!")
    If sepPos = 0 Then
        Set ResolveCellRef = ws.Range(refStr)
        Debug.Print "[D-1528] modUILoader.ResolveCellRef EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    ' ?V?[?g?C???Q??: ?V?[?g?? + ?Z????n?????
    Dim sheetName As String
    Dim cellAddr As String
    sheetName = Left(refStr, sepPos - 1)
    cellAddr = Mid(refStr, sepPos + 1)
    If Len(sheetName) >= 2 Then
        If Left(sheetName, 1) = "'" And Right(sheetName, 1) = "'" Then
            sheetName = Mid(sheetName, 2, Len(sheetName) - 2)
        End If
    End If
    Set ResolveCellRef = ws.Parent.Worksheets(sheetName).Range(cellAddr)
    Debug.Print "[D-1529] modUILoader.ResolveCellRef EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    Debug.Print "[D-1530] modUILoader.ResolveCellRef EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Set ResolveCellRef = Nothing
End Function

' ================================================================
' ?????: CollectFieldSections (v2.2 task 3.12)
' ?T?v:   formatId ??t?H?[?}?b?g??` .txt ?? modFormatLoader ?o?R????A
'         [FIELD] ?Z?N?V???????L?q?? (= ???\????) ?? Collection ?????B
' ????:   ByVal formatId As String - ???t?H?[?}?b?g ID
' ???l: Collection - ?v?f?? [FIELD] ?? ClsStanzaSection (0 ????????)
' ================================================================
Private Function CollectFieldSections(ByVal formatId As String) As Collection
    Debug.Print "[D-1531] modUILoader.CollectFieldSections ENTER"  ' [ADR-0100]
    Dim result As Collection
    Set result = New Collection
    On Error GoTo ErrHandler

    Dim formatSections As Collection
    Set formatSections = modFormatLoader.LoadFormat(formatId)
    If formatSections Is Nothing Then
        Set CollectFieldSections = result
        Debug.Print "[D-1532] modUILoader.CollectFieldSections EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If

    Dim sec As ClsStanzaSection
    For Each sec In formatSections
        If sec.SectionName = "FIELD" Then
            result.Add sec
        End If
    Next sec

    If result.Count = 0 Then
        Debug.Print "[UILoader] WARN: ?t?H?[?}?b?g '" & formatId & _
            "' ?? [FIELD] ?????? (??`?t?@?C???s???????)"
    End If
    Set CollectFieldSections = result
    Debug.Print "[D-1533] modUILoader.CollectFieldSections EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    Debug.Print "[D-1534] modUILoader.CollectFieldSections EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Set CollectFieldSections = result
End Function

' ================================================================
' ?????: DrawFormatHeaderRow (v2.2 task 3.12)
' ?T?v:   ???I?t?H?[??/?O???b?h?? 4 ??w?b?_?s??`????BColumns ?L?[??
'         ???o??:?? ?? CSV (??? 4 ???\?????i #/?K?{/?t?B?[???h??+?^/?????)?B
'         ???o?????w?b?_?Z????????A?????Y??????K?p????B
' ????:   ByVal ws As Worksheet         - ?K?p??V?[?g
'         ByVal sec As ClsStanzaSection - ???Z?N?V????
'         ByVal anchor As Range         - StartCell (?w?b?_?s?????Z??)
' ================================================================
Private Sub DrawFormatHeaderRow( _
    ByVal ws As Worksheet, _
    ByVal sec As ClsStanzaSection, _
    ByVal anchor As Range _
)
    Debug.Print "[D-1535] modUILoader.DrawFormatHeaderRow ENTER"  ' [ADR-0100]
    Dim colsCsv As String
    colsCsv = sec.GetValue("Columns")
    If Len(colsCsv) = 0 Then Exit Sub

    Dim colDefs() As String
    colDefs = Split(colsCsv, ",")
    Dim i As Long
    For i = LBound(colDefs) To UBound(colDefs)
        Dim colIdx As Long
        colIdx = i - LBound(colDefs)
        Dim parts() As String
        parts = Split(colDefs(i), ":")
        Dim headerCell As Range
        Set headerCell = ws.Cells(anchor.Row, anchor.Column + colIdx)
        ' ???o?? (part 0)
        If UBound(parts) >= 0 Then
            headerCell.Value = modStanzaIO.UnescapeStanzaValue(Trim(parts(0)))
        End If
        ' ?? (part 1)
        If UBound(parts) >= 1 Then
            If Len(Trim(parts(1))) > 0 Then
                ws.Columns(anchor.Column + colIdx).ColumnWidth = CDbl(Trim(parts(1)))
            End If
        End If
    Next i
    Debug.Print "[D-1536] modUILoader.DrawFormatHeaderRow EXIT-OK"  ' [ADR-0100]
End Sub

' ================================================================
' ?????: DrawFormatFieldRow (v2.2 task 3.12)
' ?T?v:   1 ???? [FIELD] ???I?t?H?[???? 1 ?f?[?^?s??`????B# ??=
'         FieldNo?A?K?{??=Required ??? ???A?t?B?[???h??+?^ ??=FieldName
'         (FieldType)?A???????=FieldType ????G?o????????Z???B?????
'         ?Z?????n?? outFieldCells ?? FieldName ???L?[??L?^????B
' ????:   ByVal ws As Worksheet              - ?K?p??V?[?g
'         ByVal anchor As Range              - StartCell (?w?b?_?s?????)
'         ByVal rowNum As Long               - ?`????f?[?^?s???
'         ByVal fieldSec As ClsStanzaSection - ??? [FIELD] ?Z?N?V????
'         ByRef outFieldCells As Object      - ?o??: FieldName -> ???????n
' ================================================================
Private Sub DrawFormatFieldRow( _
    ByVal ws As Worksheet, _
    ByVal anchor As Range, _
    ByVal rowNum As Long, _
    ByVal fieldSec As ClsStanzaSection, _
    ByRef outFieldCells As Object _
)
    Debug.Print "[D-1537] modUILoader.DrawFormatFieldRow ENTER"  ' [ADR-0100]
    Dim baseCol As Long
    baseCol = anchor.Column

    ' # ??: FieldNo
    ws.Cells(rowNum, baseCol + FORM_COL_NO).Value = fieldSec.GetValue("FieldNo")

    ' ?K?{??: Required=TRUE ??? ???A?????O???
    If UCase(fieldSec.GetValue("Required")) = "TRUE" Then
        ws.Cells(rowNum, baseCol + FORM_COL_REQUIRED).Value = FORM_REQUIRED_MARK
    Else
        ws.Cells(rowNum, baseCol + FORM_COL_REQUIRED).Value = vbNullString
    End If

    ' ?t?B?[???h??+?^ ??: FieldName (FieldType)
    Dim fieldName As String
    fieldName = fieldSec.GetValue("FieldName")
    ws.Cells(rowNum, baseCol + FORM_COL_FIELD).Value = _
        fieldName & "?i" & fieldSec.GetValue("FieldType") & "?j"

    ' ???????: FieldType ???????Z?????o?E????
    Dim inputCell As Range
    Set inputCell = ws.Cells(rowNum, baseCol + FORM_COL_INPUT)
    ApplyFormatInputCell inputCell, fieldSec

    ' ?[?U?E???b?N????? FieldName -> ???????n ???L?^
    If Len(fieldName) > 0 Then
        outFieldCells(fieldName) = inputCell.Address
    End If
    Debug.Print "[D-1538] modUILoader.DrawFormatFieldRow EXIT-OK"  ' [ADR-0100]
End Sub

' ================================================================
' ?????: ApplyFormatInputCell (v2.2 task 3.12)
' ?T?v:   ???I?t?H?[?????????Z??????????B?w?i?F?E?r????K?p???A
'         FieldType ???????????K?? ([INPUT].InputType ?? 1:1 ???) ??
'         DefaultValue ???f????B
' ????:   ByVal r As Range                   - ??????Z??
'         ByVal fieldSec As ClsStanzaSection - ??? [FIELD] ?Z?N?V????
' ================================================================
Private Sub ApplyFormatInputCell( _
    ByVal r As Range, _
    ByVal fieldSec As ClsStanzaSection _
)
    Debug.Print "[D-1539] modUILoader.ApplyFormatInputCell ENTER"  ' [ADR-0100]
    ' ???????w?i?F?E?r?? (schema 3.6 ?? [INPUT] ????????)
    r.Interior.Color = HexToRgbLong(INPUT_DEFAULT_COLOR)
    ApplyBorder r, "thin"

    ' FieldType ???????????K??
    Dim fieldType As String
    fieldType = LCase(fieldSec.GetValue("FieldType"))
    On Error Resume Next
    r.Validation.Delete
    On Error GoTo 0
    ApplyFieldValidation r, fieldType, fieldSec

    ' DefaultValue ???????l???????f
    Dim defVal As String
    defVal = fieldSec.GetValue("DefaultValue")
    If Len(defVal) > 0 Then
        r.Cells(1, 1).Value = modStanzaIO.UnescapeStanzaValue(defVal)
    End If
    Debug.Print "[D-1540] modUILoader.ApplyFormatInputCell EXIT-OK"  ' [ADR-0100]
End Sub

' ================================================================
' ?????: ApplyFieldValidation (v2.2 task 3.12)
' ?T?v:   FieldType ???????????K?????????Z????K?p????Bdropdown ??
'         DropdownOptions ??????Anumber / date ????K???Atext /
'         multiline ?? MaxLength ????????K????K?p????B
' ????:   ByVal r As Range                   - ??????Z??
'         ByVal fieldType As String          - FieldType (??????)
'         ByVal fieldSec As ClsStanzaSection - ??? [FIELD] ?Z?N?V????
' ???l:   checkbox ?? ApplyInput ????l?A??p?????K????????????B
' ================================================================
Private Sub ApplyFieldValidation( _
    ByVal r As Range, _
    ByVal fieldType As String, _
    ByVal fieldSec As ClsStanzaSection _
)
    Debug.Print "[D-1541] modUILoader.ApplyFieldValidation ENTER"  ' [ADR-0100]
    Select Case fieldType
        Case "dropdown"
            ApplyDropdownValidation r, fieldSec.GetValue("DropdownOptions")
        Case "number"
            On Error Resume Next
            r.Validation.Add Type:=xlValidateWholeNumber, _
                AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
                Formula1:=VALIDATE_INT_MIN, Formula2:=VALIDATE_INT_MAX
            On Error GoTo 0
        Case "date"
            On Error Resume Next
            r.Validation.Add Type:=xlValidateDate, _
                AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
                Formula1:=VALIDATE_DATE_MIN, Formula2:=VALIDATE_DATE_MAX
            On Error GoTo 0
        Case "text", "multiline"
            ApplyMaxLengthValidation r, fieldSec.GetValue("MaxLength")
    End Select
    Debug.Print "[D-1542] modUILoader.ApplyFieldValidation EXIT-OK"  ' [ADR-0100]
End Sub

' ================================================================
' ?????: ApplyMaxLengthValidation (v2.2 task 3.12)
' ?T?v:   ??????Z?????????????????K????K?p????BmaxLen ??????
'         ??????????B
' ????:   ByVal r As Range          - ??????Z??
'         ByVal maxLen As String    - ??????? ([FIELD].MaxLength)
' ================================================================
Private Sub ApplyMaxLengthValidation(ByVal r As Range, ByVal maxLen As String)
    Debug.Print "[D-1543] modUILoader.ApplyMaxLengthValidation ENTER"  ' [ADR-0100]
    If Len(Trim(maxLen)) = 0 Then Exit Sub
    On Error Resume Next
    r.Validation.Add Type:=xlValidateTextLength, _
        AlertStyle:=xlValidAlertStop, Operator:=xlLessEqual, _
        Formula1:=Trim(maxLen)
    On Error GoTo 0
    Debug.Print "[D-1544] modUILoader.ApplyMaxLengthValidation EXIT-OK"  ' [ADR-0100]
End Sub

' ================================================================
' ?????: PopulateFormCells (v2.2 task 3.12)
' ?T?v:   [FORM_FROM_FORMAT].PopulateFromKnowledge ???w???Z??????
'         knowledgeNo ????AmodKnowledgeFileIO ?o?R??i???b?W???????A
'         FieldName ??v??e??????????l??[?U???? (3.12 ??? 5)?B
'         PopulateFromKnowledge ???w????????????? (M-05 ?V?K?o?^???)?B
' ????:   ByVal ws As Worksheet         - ?K?p??V?[?g
'         ByVal sec As ClsStanzaSection - [FORM_FROM_FORMAT] ?Z?N?V????
'         ByVal fieldCells As Object    - FieldName -> ???????n ?????\
' ================================================================
Private Sub PopulateFormCells( _
    ByVal ws As Worksheet, _
    ByVal sec As ClsStanzaSection, _
    ByVal fieldCells As Object _
)
    Debug.Print "[D-1545] modUILoader.PopulateFormCells ENTER"  ' [ADR-0100]
    Dim pfkRef As String
    pfkRef = sec.GetValue("PopulateFromKnowledge")
    If Len(pfkRef) = 0 Then Exit Sub

    Dim knCell As Range
    Set knCell = ResolveCellRef(ws, pfkRef)
    If knCell Is Nothing Then Exit Sub
    Dim knowledgeNo As String
    knowledgeNo = Trim(CStr(knCell.Cells(1, 1).Value))
    If Len(knowledgeNo) = 0 Then Exit Sub

    Dim knowledge As Object
    Set knowledge = modKnowledgeFileIO.LoadKnowledge(knowledgeNo)
    If knowledge Is Nothing Then Exit Sub
    If knowledge.Count = 0 Then Exit Sub

    ' FieldName ??v??e??????????l??[?U
    Dim fieldName As Variant
    For Each fieldName In fieldCells.Keys
        If knowledge.Exists(CStr(fieldName)) Then
            ws.Range(CStr(fieldCells(fieldName))).Cells(1, 1).Value = _
                knowledge(CStr(fieldName))
        End If
    Next fieldName
    Debug.Print "[D-1546] modUILoader.PopulateFormCells EXIT-OK"  ' [ADR-0100]
End Sub

' ================================================================
' ?????: ApplyFormLocks (v2.2 task 3.12)
' ?T?v:   [FORM_FROM_FORMAT] ????b?N???? (3.12 ??? 6)?BReadOnly=TRUE ??
'         ?S??????????b?N (?\?????[?h?AOP-3)?AFALSE ???W??????B
'         ReadOnlyFormat=TRUE ?? SourceFormatRef ?Z???????b?N????B
' ????:   ByVal ws As Worksheet         - ?K?p??V?[?g
'         ByVal sec As ClsStanzaSection - [FORM_FROM_FORMAT] ?Z?N?V????
'         ByVal fieldCells As Object    - FieldName -> ???????n ?????\
' ???l:   ?Z?????b?N??V?[?g???????L???B?????????????S???B
' ================================================================
Private Sub ApplyFormLocks( _
    ByVal ws As Worksheet, _
    ByVal sec As ClsStanzaSection, _
    ByVal fieldCells As Object _
)
    Debug.Print "[D-1547] modUILoader.ApplyFormLocks ENTER"  ' [ADR-0100]
    ' ReadOnly: TRUE ???????????b?N?AFALSE ???W?? (???? FALSE)
    Dim isReadOnly As Boolean
    isReadOnly = (UCase(sec.GetValue("ReadOnly")) = "TRUE")
    Dim fieldName As Variant
    For Each fieldName In fieldCells.Keys
        ws.Range(CStr(fieldCells(fieldName))).Cells(1, 1).Locked = isReadOnly
    Next fieldName

    ' ReadOnlyFormat=TRUE ?? SourceFormatRef ?Z???????b?N (??I??s??)
    If UCase(sec.GetValue("ReadOnlyFormat")) = "TRUE" Then
        Dim refCell As Range
        Set refCell = ResolveCellRef(ws, sec.GetValue("SourceFormatRef"))
        If Not refCell Is Nothing Then
            refCell.Cells(1, 1).Locked = True
        End If
    End If
    Debug.Print "[D-1548] modUILoader.ApplyFormLocks EXIT-OK"  ' [ADR-0100]
End Sub

' v2.1-btnwire (2026-05-22): [BUTTON] handler?BmodButtonWiring ??????
' ???t?H?[???R???g???[???{?^????????????????B????????u?A???J?[?Z?????
' ?L???v?V?????????? + ?????R???g???[????? bind ???v??p?~ (???p???
' ????z?u??v????????A?e???f 2026-05-22)?B?{????????u?b?N?I?[?v????
' clsSetupOrchestrator.RunFullSetup -> ApplyUiToSheet ?o?R?????s????A
' ???O?? clsSheetRenderer.ClearScreen ???S Shape ?????????? idempotent?B
Private Sub ApplyButton(ByVal ws As Worksheet, ByVal sec As ClsStanzaSection)
    Debug.Print "[D-1549] modUILoader.ApplyButton ENTER"  ' [ADR-0100]
    Dim cellAddr As String
    cellAddr = sec.GetValue("Cell")
    If Len(cellAddr) = 0 Then Exit Sub

    Dim caption As String
    caption = modStanzaIO.UnescapeStanzaValue(sec.GetValue("Text"))
    Dim onClick As String
    onClick = sec.GetValue("OnClick")
    Dim widthPx As Double, heightPx As Double
    widthPx = ParseDoubleOr(sec.GetValue("Width"), 120#)
    heightPx = ParseDoubleOr(sec.GetValue("Height"), 24#)

    ' Phase R-3-ψ-M03-style (差3): BackColor 指定時は色付き Shape ボタン化。
    Dim bcl As Long
    bcl = -1
    If sec.HasKey("BackColor") Then bcl = HexToRgbLong(sec.GetValue("BackColor"))
    modButtonWiring.PlaceButton ws, cellAddr, caption, onClick, _
        widthPx, heightPx, sec.GetValue("ForeColor"), bcl
    Debug.Print "[D-1550] modUILoader.ApplyButton EXIT-OK"  ' [ADR-0100]
End Sub

' v2.1-btnwire: [BUTTON_TEMPLATE] handler?B??? [GRID] ??f?[?^?s?????
' per-row ?{?^???? modButtonWiring ?????????????B
Private Sub ApplyButtonTemplate(ByVal ws As Worksheet, ByVal sec As ClsStanzaSection, _
                                ByVal sections As Collection)
    Debug.Print "[D-1551] modUILoader.ApplyButtonTemplate ENTER"  ' [ADR-0100]
    Dim gridRef As String
    gridRef = sec.GetValue("GridRef")
    Dim gridSec As ClsStanzaSection
    Set gridSec = FindGridByName(sections, gridRef)
    modButtonWiring.PlaceTemplateButtons ws, sec, gridSec
    Debug.Print "[UILoader] BUTTON_TEMPLATE applied: " & sec.GetValue("TemplateId") & _
        " grid=" & gridRef
    Debug.Print "[D-1552] modUILoader.ApplyButtonTemplate EXIT-OK"  ' [ADR-0100]
End Sub

' sections ?R???N?V???????? Name ??v?? [GRID] ?Z?N?V?????????B
' ???????????? Nothing?B
Private Function FindGridByName(ByVal sections As Collection, _
                                ByVal gridName As String) As ClsStanzaSection
    Debug.Print "[D-1553] modUILoader.FindGridByName ENTER"  ' [ADR-0100]
    Dim sec As ClsStanzaSection
    For Each sec In sections
        If sec.SectionName = "GRID" Then
            If sec.GetValue("Name") = gridName Then
                Set FindGridByName = sec
                Debug.Print "[D-1554] modUILoader.FindGridByName EXIT-OK"  ' [ADR-0100]
                Exit Function
            End If
        End If
    Next sec
    Set FindGridByName = Nothing
End Function

' ??????? Double ?????B?? / ??????s???????l?????B
Private Function ParseDoubleOr(ByVal s As String, ByVal dflt As Double) As Double
    Debug.Print "[D-1555] modUILoader.ParseDoubleOr ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    If Len(Trim(s)) = 0 Then
        ParseDoubleOr = dflt
    Else
        ParseDoubleOr = CDbl(Trim(s))
    End If
    Debug.Print "[D-1556] modUILoader.ParseDoubleOr EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    Debug.Print "[D-1557] modUILoader.ParseDoubleOr EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ParseDoubleOr = dflt
End Function

' ----------------------------------------------------------------
' Private: helpers (v2.1 truncate fix)
' ----------------------------------------------------------------

' Convert a schema HAlign token (left/center/right) to an Excel enum.
Private Function HAlignToEnum(ByVal token As String) As Long
    Debug.Print "[D-1558] modUILoader.HAlignToEnum ENTER"  ' [ADR-0100]
    Select Case LCase(token)
        Case "left":   HAlignToEnum = xlLeft
        Case "center": HAlignToEnum = xlCenter
        Case "right":  HAlignToEnum = xlRight
        Case Else:     HAlignToEnum = xlGeneral
    End Select
    Debug.Print "[D-1559] modUILoader.HAlignToEnum EXIT-OK"  ' [ADR-0100]
End Function

Private Function HexToRgbLong(ByVal hexStr As String) As Long
    Debug.Print "[D-1560] modUILoader.HexToRgbLong ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Dim s As String
    s = hexStr
    If Left(s, 1) = "#" Then s = Mid(s, 2)
    If Len(s) <> 6 Then
        HexToRgbLong = 0
        Debug.Print "[D-1561] modUILoader.HexToRgbLong EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    Dim rr As Long, gg As Long, bb As Long
    rr = CLng("&H" & Mid(s, 1, 2))
    gg = CLng("&H" & Mid(s, 3, 2))
    bb = CLng("&H" & Mid(s, 5, 2))
    HexToRgbLong = RGB(rr, gg, bb)
    Debug.Print "[D-1562] modUILoader.HexToRgbLong EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    Debug.Print "[D-1563] modUILoader.HexToRgbLong EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    HexToRgbLong = 0
End Function

' ============================================================
' v2.3 Phase A2 restore (2026-05-27 rewrite): [STYLE] reader/applier
' TEMP DISABLED for compile-error triage. Whole block guarded by
' #If PHASE_A2_STYLE = 1 Then ... #End If (PHASE_A2_STYLE = 0).
' ============================================================
#Const PHASE_A2_STYLE = 1
#If PHASE_A2_STYLE = 1 Then
Private Sub ApplyStyle(ByVal ws As Worksheet, ByVal sec As ClsStanzaSection)
    Debug.Print "[D-1564] modUILoader.ApplyStyle ENTER"  ' [ADR-0100]
    On Error Resume Next
    If ws Is Nothing Then Exit Sub
    If sec Is Nothing Then Exit Sub
    Dim v As String

    v = sec.GetValue("BaseFont")
    If StyleHasValue(v) Then ws.Cells.Font.Name = v

    v = sec.GetValue("BaseFontSize")
    If StyleHasValue(v) Then
        If IsNumeric(v) Then ws.Cells.Font.Size = CDbl(v)
    End If

    v = sec.GetValue("BaseFontColor")
    If StyleHasValue(v) Then ws.Cells.Font.Color = StyleColor(v)

    v = sec.GetValue("SheetBackColor")
    If StyleHasValue(v) Then ws.Cells.Interior.Color = StyleColor(v)

    v = sec.GetValue("BaseAlignH")
    If StyleHasValue(v) Then ws.Cells.HorizontalAlignment = HAlignToEnum(v)

    v = sec.GetValue("BaseAlignV")
    If StyleHasValue(v) Then ws.Cells.VerticalAlignment = StyleVAlign(v)

    v = sec.GetValue("BaseWrap")
    If StyleHasValue(v) Then ws.Cells.WrapText = (LCase(v) = "on")

    v = sec.GetValue("BaseNumberFormat")
    If StyleHasValue(v) Then
        If LCase(v) <> "standard" Then ws.Cells.NumberFormat = v
    End If

    On Error GoTo 0
    Debug.Print "[D-1565] modUILoader.ApplyStyle EXIT-OK"  ' [ADR-0100]
End Sub

' Returns False for empty / "-" / starts-with horizontal-bar (design "no value").
Private Function StyleHasValue(ByVal v As String) As Boolean
    Debug.Print "[D-1566] modUILoader.StyleHasValue ENTER"  ' [ADR-0100]
    StyleHasValue = False
    If Len(v) = 0 Then Exit Function
    Dim t As String
    t = Trim(v)
    If Len(t) = 0 Then Exit Function
    If t = "-" Then Exit Function
    ' U+2015 HORIZONTAL BAR = full-width dash used in design spec
    If Left(t, 1) = ChrW(8213) Then Exit Function
    StyleHasValue = True
    Debug.Print "[D-1567] modUILoader.StyleHasValue EXIT-OK"  ' [ADR-0100]
End Function

' Map color tokens to RGB Long. Accepts Japanese name, English name, or #hex.
Private Function StyleColor(ByVal token As String) As Long
    Debug.Print "[D-1568] modUILoader.StyleColor ENTER"  ' [ADR-0100]
    Dim t As String
    t = LCase(Trim(token))
    Select Case t
        Case "white"
            StyleColor = RGB(255, 255, 255)
        Case "black"
            StyleColor = RGB(0, 0, 0)
        Case "darkgray"
            StyleColor = RGB(64, 64, 64)
        Case "lightgray"
            StyleColor = RGB(200, 200, 200)
        Case "red"
            StyleColor = RGB(255, 0, 0)
        Case "blue"
            StyleColor = RGB(0, 0, 255)
        Case "green"
            StyleColor = RGB(0, 128, 0)
        Case Else
            If Left(t, 1) = "#" Then
                StyleColor = HexToRgbLong(t)
            Else
                StyleColor = StyleColorJP(token)
            End If
    End Select
    Debug.Print "[D-1569] modUILoader.StyleColor EXIT-OK"  ' [ADR-0100]
End Function

' Japanese color names. 2026-06-06: replaced no-op vbBlack with explicit JP mapping (M-02 dark-on-dark bug fix).
Private Function StyleColorJP(ByVal token As String) As Long
    Debug.Print "[D-1570] modUILoader.StyleColorJP ENTER"  ' [ADR-0100]
    Dim t As String
    t = Trim(token)
    Select Case t
        Case ChrW(&H767D): StyleColorJP = RGB(255, 255, 255)
        Case ChrW(&H9ED2): StyleColorJP = RGB(0, 0, 0)
        Case ChrW(&H6FC3) & ChrW(&H3044) & ChrW(&H30B0) & ChrW(&H30EC) & ChrW(&H30FC): StyleColorJP = RGB(64, 64, 64)
        Case ChrW(&H30B0) & ChrW(&H30EC) & ChrW(&H30FC): StyleColorJP = RGB(128, 128, 128)
        Case ChrW(&H8584) & ChrW(&H3044) & ChrW(&H30B0) & ChrW(&H30EC) & ChrW(&H30FC): StyleColorJP = RGB(200, 200, 200)
        Case ChrW(&H8D64): StyleColorJP = RGB(255, 0, 0)
        Case ChrW(&H9752): StyleColorJP = RGB(0, 0, 255)
        Case ChrW(&H7DD1): StyleColorJP = RGB(0, 128, 0)
        Case ChrW(&H9EC4): StyleColorJP = RGB(255, 255, 0)
        Case ChrW(&H8336): StyleColorJP = RGB(139, 69, 19)
        Case ChrW(&H7D2B): StyleColorJP = RGB(128, 0, 128)
        Case ChrW(&H30AA) & ChrW(&H30EC) & ChrW(&H30F3) & ChrW(&H30B8): StyleColorJP = RGB(255, 165, 0)
        Case Else: StyleColorJP = vbBlack
    End Select
    Debug.Print "[D-1571] modUILoader.StyleColorJP EXIT-OK"  ' [ADR-0100]
End Function

' Map "top|center|bottom" to xlVAlign enum.
Private Function StyleVAlign(ByVal v As String) As Long
    Debug.Print "[D-1572] modUILoader.StyleVAlign ENTER"  ' [ADR-0100]
    Dim t As String
    t = LCase(Trim(v))
    Select Case t
        Case "top"
            StyleVAlign = xlTop
        Case "center"
            StyleVAlign = xlCenter
        Case "bottom"
            StyleVAlign = xlBottom
        Case Else
            StyleVAlign = xlCenter
    End Select
    Debug.Print "[D-1573] modUILoader.StyleVAlign EXIT-OK"  ' [ADR-0100]
End Function

#End If
' (end of PHASE_A2_STYLE = 1 conditional block)

' ============================================================
' Phase FC (2026-06-03, ADR-0096): FormControl stanza handlers.
' Forwarded straight to modFormControlWiring; this module only
' decodes stanza keys.
' ============================================================

Private Sub ApplyValidation(ByVal ws As Worksheet, ByVal sec As ClsStanzaSection)
    Dim cellAddr As String
    cellAddr = sec.GetValue("Cell")
    If Len(cellAddr) = 0 Then Exit Sub
    Dim listValues As String
    listValues = sec.GetValue("List")
    If Len(listValues) = 0 Then listValues = sec.GetValue("Choices")
    If Len(listValues) = 0 Then Exit Sub
    modFormControlWiring.PlaceValidationList ws, cellAddr, listValues, sec.GetValue("DefaultValue")
End Sub

Private Sub ApplyCheckBox(ByVal ws As Worksheet, ByVal sec As ClsStanzaSection)
    Dim cellAddr As String
    cellAddr = sec.GetValue("Cell")
    If Len(cellAddr) = 0 Then Exit Sub
    modFormControlWiring.PlaceCheckBox ws, cellAddr, sec.GetValue("Caption"), sec.GetValue("OnClick"), sec.GetValue("LinkedCell"), sec.GetValue("Name"), False
End Sub

Private Sub ApplyListBox(ByVal ws As Worksheet, ByVal sec As ClsStanzaSection)
    Dim cellAddr As String
    cellAddr = sec.GetValue("Cell")
    If Len(cellAddr) = 0 Then Exit Sub
    Dim listValues As String
    listValues = sec.GetValue("List")
    If Len(listValues) = 0 Then listValues = sec.GetValue("Choices")
    If Len(listValues) = 0 Then Exit Sub
    modFormControlWiring.PlaceListBox ws, cellAddr, listValues, sec.GetValue("LinkedCell"), sec.GetValue("Name")
End Sub

' ADR-0006/0090/0094 JP literal removal:
Private Property Get FORM_REQUIRED_MARK() As String
    FORM_REQUIRED_MARK = ChrW(&H203B)
End Property
```
