---
title: modEntrySearch.bas
description: modEntrySearch.bas のソースコード（コピペ用）
---

# modEntrySearch.bas

**配置先**: `検索.xlsm` 用の VBA モジュール
**種類**: 標準モジュール
**更新日**: 2026-06-12 01:22

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\search\`
- ファイル名: `modEntrySearch.bas`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。
> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。
> 改行コードは CRLF（Windows 標準）のままで OK です。

---

## ソースコード

```vb
Attribute VB_Name = "modEntrySearch"
Option Explicit

' ================================================================
' v2.3 search entry points (Phase Q-1, 2026-05-27)
' Read M-08 search form input, walk data_dir .txt files via
' modKnowledgeFileIO, and write matching records into the grid.
'
' M-08 cells (v2.3):
'   C8  = keyword
'   C10 = format id filter
'   A14+ = result grid start row
' ================================================================

Private Const M08_SHEET_DISPLAY_NAME As String = "M-08"
Private Const M08_CELL_KEYWORD As String = "C8"
Private Const M08_CELL_FORMAT As String = "C10"
Private Const M08_RESULT_START_ROW As Long = 14
Private Const M08_RESULT_LAST_ROW As Long = 200

Public Sub Btn_SearchV23()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1792] modEntrySearch.Btn_SearchV23 ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    ' [BTN-GUARD-PRELUDE-BEGIN] auto-injected by inject_btn_template.py
    Const BTN As String = "Btn_SearchV23"
    Dim XLSM As String
    XLSM = ChrW(&H691C) & ChrW(&H7D22)
    modBtnGuard.LogEnter BTN, XLSM
    If Not modBtnGuard.CheckPrereq(BTN, "config;data_dir", XLSM) Then
        modBtnGuard.LogExit BTN, XLSM, False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1793] modEntrySearch.Btn_SearchV23 EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If
    ' [BTN-GUARD-PRELUDE-END]
    Dim ws As Worksheet
    Set ws = ResolveSearchSheet()
    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    ws.Unprotect
    On Error GoTo ErrHandler

    Dim keyword As String, fmtFilter As String
    keyword = Trim(CStr(ws.Range(M08_CELL_KEYWORD).Value))
    fmtFilter = Trim(CStr(ws.Range(M08_CELL_FORMAT).Value))

    ' [BUG-B8 2026-06-11] C10 dropdown holds FormatName (display name) but
    ' matching compares FormatID (same root cause as B5). Resolve display
    ' name to ID; unknown values pass through so direct ID input still works.
    fmtFilter = ResolveFormatIdFromDisplay(fmtFilter)

    ' [SPEC T-24 / BUG-B10 2026-06-11] zero search conditions must show a
    ' message and NOT fall back to listing every knowledge file.
    If Len(keyword) = 0 And Len(fmtFilter) = 0 Then
        If Not IsTestMode() Then
            MsgBox MsgNoCondition(), vbExclamation, XLSM
        End If
        modBtnGuard.LogExit BTN, XLSM, True
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1792b] modEntrySearch.Btn_SearchV23 EXIT-OK (no condition)"  ' [ADR-0100]
        Exit Sub
    End If

    Dim clearRng As Range
    Set clearRng = ws.Range(ws.Cells(M08_RESULT_START_ROW, 1), ws.Cells(M08_RESULT_LAST_ROW, 8))
    On Error Resume Next
    clearRng.ClearContents
    ' [N4 2026-06-12] strip previous run's borders/zebra so a smaller result
    ' set does not leave stale formatting below the new rows.
    clearRng.Borders.LineStyle = xlNone
    clearRng.Interior.ColorIndex = xlNone
    ' 2026-05-31 UX: reset column A font (blue + underline hyperlink-look)
    ' that was applied to the previous search result.
    Dim aColRng As Range
    Set aColRng = ws.Range(ws.Cells(M08_RESULT_START_ROW, 1), ws.Cells(M08_RESULT_LAST_ROW, 1))
    aColRng.Font.ColorIndex = xlAutomatic
    aColRng.Font.Underline = xlUnderlineStyleNone
    On Error GoTo ErrHandler

    Dim dataDir As String
    dataDir = modConfigHolder.GetDataDir()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(dataDir) Then
        ws.Range("A" & M08_RESULT_START_ROW).Value = "(data_dir not found: " & dataDir & ")"
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1794] modEntrySearch.Btn_SearchV23 EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If

    Dim outRow As Long
    outRow = M08_RESULT_START_ROW
    Dim hitCount As Long
    hitCount = 0

    ' [USER-REQ 2026-06-09] Collect hits first so we can sort by UpdatedAt desc.
    Dim hits As Collection
    Set hits = New Collection
    Dim f As Object
    For Each f In fso.GetFolder(dataDir).Files
        If LCase(fso.GetExtensionName(f.Name)) = "txt" Then
            Dim knwNo As String
            knwNo = fso.GetBaseName(f.Name)
            Dim data As Object
            Set data = modKnowledgeFileIO.LoadKnowledge(knwNo)
            If Not data Is Nothing Then
                If MatchesSearchV23(data, keyword, fmtFilter) Then
                    Dim hit As Object
                    Set hit = CreateObject("Scripting.Dictionary")
                    hit("knwNo") = knwNo
                    Set hit("data") = data
                    hit("updatedAt") = SafeStr(data, "UpdatedAt")
                    hits.Add hit
                End If
            End If
        End If
    Next f

    ' [USER-REQ 2026-06-09] Sort hits by UpdatedAt descending.
    Dim ai As Long, bi As Long
    For ai = 1 To hits.Count - 1
        For bi = 1 To hits.Count - ai
            Dim ha As Object, hb As Object
            Set ha = hits.Item(bi)
            Set hb = hits.Item(bi + 1)
            If CStr(ha("updatedAt")) < CStr(hb("updatedAt")) Then
                hits.Remove bi + 1
                If bi = 1 Then
                    hits.Add hb, , 1
                Else
                    hits.Add hb, , bi
                End If
            End If
        Next bi
    Next ai

    Dim hi As Long
    For hi = 1 To hits.Count
        If outRow > M08_RESULT_LAST_ROW Then Exit For
        Dim hRec As Object
        Set hRec = hits.Item(hi)
        Dim hData As Object
        Set hData = hRec("data")
        Dim hKnwNo As String
        hKnwNo = CStr(hRec("knwNo"))

        ws.Cells(outRow, 1).Value = hitCount + 1
        ' [USER-REQ 2026-06-09] Reorder columns: A=No, B=FormatName, C-E=search fields, F=UpdatedAt, G=knwNo (hidden)
        Dim fmtNm As String
        fmtNm = ResolveFormatName(SafeStr(hData, "FormatID"))
        ws.Cells(outRow, 2).Value = fmtNm
        Dim sFields(1 To 3) As String
        GetSearchTargetFieldValues hData, SafeStr(hData, "FormatID"), sFields
        ws.Cells(outRow, 3).Value = TrimExcerpt(sFields(1))
        ws.Cells(outRow, 4).Value = TrimExcerpt(sFields(2))
        ws.Cells(outRow, 5).Value = TrimExcerpt(sFields(3))
        ws.Cells(outRow, 6).Value = SafeStr(hData, "UpdatedAt")
        ws.Cells(outRow, 7).Value = hKnwNo
        With ws.Cells(outRow, 1)
            .Font.Color = RGB(0, 0, 255)
            .Font.Underline = xlUnderlineStyleSingle
        End With
        outRow = outRow + 1
        hitCount = hitCount + 1
    Next hi

    ' 2026-06-08 UX: extend grid borders/styling to all written rows so
    ' result rows past the initial 10 don't appear as bare cells.
    If hitCount > 0 Then
        Dim styRng As Range
        Set styRng = ws.Range(ws.Cells(M08_RESULT_START_ROW, 1), ws.Cells(outRow - 1, 7))
        styRng.Borders.LineStyle = xlContinuous
        styRng.Borders.Weight = xlThin
        styRng.Borders.Color = RGB(191, 191, 191)
        ' [USER-REQ 2026-06-09] Cells must not overflow their column width.
        ' WrapText off + IndentLevel 0; columns to right will hide spill via opaque fill.
        styRng.WrapText = False
        Dim rr As Long
        For rr = M08_RESULT_START_ROW To outRow - 1
            If (rr - M08_RESULT_START_ROW) Mod 2 = 1 Then
                ws.Range(ws.Cells(rr, 1), ws.Cells(rr, 7)).Interior.Color = RGB(247, 247, 247)
            Else
                ws.Range(ws.Cells(rr, 1), ws.Cells(rr, 7)).Interior.ColorIndex = xlNone
            End If
        Next rr
    End If

    ' [USER-REQ 2026-06-09] Force column G (knowledgeNo) to be hidden after each search.
    ' Setup_search applies HideColumnsFrom=G but Excel sometimes loses it on workbook reopen.
    On Error Resume Next
    ws.Columns("G:G").Hidden = True
    On Error GoTo 0

    ' [N1 2026-06-12] Unprotect was one-way - restore light protection
    modCommon.ReProtectLight ws

    On Error Resume Next
    Dim lg As clsLogger
    Set lg = New clsLogger
    lg.Init ThisWorkbook.Worksheets("LOG")
    lg.LogInfo "modEntrySearch", "Btn_SearchV23", _
        "search done keyword=" & keyword & " fmt=" & fmtFilter & " hits=" & hitCount, _
        "", "LOG-M08-SEARCH-OK"
    On Error GoTo 0
    ' [BTN-GUARD-EXIT-OK] auto-injected
    modBtnGuard.LogExit BTN, XLSM, True
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1795] modEntrySearch.Btn_SearchV23 EXIT-OK"  ' [ADR-0100]
    Exit Sub

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1796] modEntrySearch.Btn_SearchV23 EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ' [BTN-GUARD-ERR-LOG] auto-injected
    modBtnGuard.LogExit BTN, XLSM, False
    Debug.Print "[ERR] Btn_SearchV23: " & Err.Number & " " & Err.Description
End Sub

Public Sub Btn_SearchClearV23()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1797] modEntrySearch.Btn_SearchClearV23 ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    ' [BTN-GUARD-PRELUDE-BEGIN] auto-injected by inject_btn_template.py
    Const BTN As String = "Btn_SearchClearV23"
    Dim XLSM As String
    XLSM = ChrW(&H691C) & ChrW(&H7D22)
    modBtnGuard.LogEnter BTN, XLSM
    If Not modBtnGuard.CheckPrereq(BTN, "config", XLSM) Then
        modBtnGuard.LogExit BTN, XLSM, False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1798] modEntrySearch.Btn_SearchClearV23 EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If
    ' [BTN-GUARD-PRELUDE-END]
    Dim ws As Worksheet
    Set ws = ResolveSearchSheet()
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ws.Unprotect
    On Error GoTo ErrHandler
    ws.Range(M08_CELL_KEYWORD).Value = ""
    ws.Range(M08_CELL_FORMAT).Value = ""
    Dim clearRng As Range
    Set clearRng = ws.Range(ws.Cells(M08_RESULT_START_ROW, 1), ws.Cells(M08_RESULT_LAST_ROW, 8))
    clearRng.ClearContents
    ' 2026-05-31 UX: also reset column A hyperlink-look font on clear.
    Dim aColRng As Range
    Set aColRng = ws.Range(ws.Cells(M08_RESULT_START_ROW, 1), ws.Cells(M08_RESULT_LAST_ROW, 1))
    aColRng.Font.ColorIndex = xlAutomatic
    aColRng.Font.Underline = xlUnderlineStyleNone
    ' 2026-06-08 UX: also clear extended grid borders + alternating row tint
    clearRng.Borders.LineStyle = xlNone
    clearRng.Interior.ColorIndex = xlNone
    ' [N3 2026-06-12] restore the seed look: thin border on the initial 10 rows
    Dim seedRng As Range
    Set seedRng = ws.Range(ws.Cells(M08_RESULT_START_ROW, 1), ws.Cells(M08_RESULT_START_ROW + 9, 7))
    seedRng.Borders.LineStyle = xlContinuous
    seedRng.Borders.Weight = xlThin
    seedRng.Borders.Color = RGB(191, 191, 191)
    ' [N2 2026-06-12] re-protect (Unprotect was one-way)
    modCommon.ReProtectLight ws
    ' [BTN-GUARD-EXIT-OK] auto-injected
    modBtnGuard.LogExit BTN, XLSM, True
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1799] modEntrySearch.Btn_SearchClearV23 EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1800] modEntrySearch.Btn_SearchClearV23 EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ' [BTN-GUARD-ERR-LOG] auto-injected
    modBtnGuard.LogExit BTN, XLSM, False
    Debug.Print "[ERR] Btn_SearchClearV23: " & Err.Number & " " & Err.Description
End Sub

' --- helpers ---

Private Function ResolveSearchSheet() As Worksheet
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1801] modEntrySearch.ResolveSearchSheet ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim jp As String
    jp = ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & ChrW(&H691C) & ChrW(&H7D22)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(jp)
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets("M-08")
    Set ResolveSearchSheet = ws
    On Error GoTo 0
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1802] modEntrySearch.ResolveSearchSheet EXIT-OK"  ' [ADR-0100]
End Function

Private Function MatchesSearchV23(ByVal data As Object, ByVal keyword As String, ByVal fmtFilter As String) As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1803] modEntrySearch.MatchesSearchV23 ENTER"  ' [ADR-0100]
    If Len(fmtFilter) > 0 Then
        Dim actual As String
        actual = SafeStr(data, "FormatID")
        If actual <> fmtFilter Then
            MatchesSearchV23 = False
            If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1804] modEntrySearch.MatchesSearchV23 EXIT-OK"  ' [ADR-0100]
            Exit Function
        End If
    End If
    If Len(keyword) = 0 Then
        MatchesSearchV23 = True
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1805] modEntrySearch.MatchesSearchV23 EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    ' [USER-REQ 2026-06-09] Keyword matching restricted to searchTarget=TRUE fields.
    ' [BUG-B9 2026-06-11] Space-separated keywords (half/full width) are AND
    ' conditions over the combined searchTarget text, not one long literal.
    Dim searchFields As Collection
    Set searchFields = GetSearchTargetFieldNames(SafeStr(data, "FormatID"))
    Dim combined As String
    combined = ""
    Dim sfi As Long
    For sfi = 1 To searchFields.Count
        Dim fname As String
        fname = CStr(searchFields(sfi))
        If data.Exists(fname) Then combined = combined & " " & CStr(data(fname))
    Next sfi
    MatchesSearchV23 = ContainsAllTerms(combined, keyword)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1806] modEntrySearch.MatchesSearchV23 EXIT-OK"  ' [ADR-0100]
End Function

' [USER-REQ 2026-06-09] Returns all FIELD names where searchTarget=TRUE
' for the given format. Empty Collection if no format or no marked fields.
Private Function GetSearchTargetFieldNames(ByVal fmtId As String) As Collection
    Dim result As Collection
    Set result = New Collection
    On Error Resume Next
    If Len(fmtId) = 0 Then
        Set GetSearchTargetFieldNames = result
        Exit Function
    End If
    Dim secs As Collection
    Set secs = modFormatLoader.LoadFormat(fmtId)
    If secs Is Nothing Then
        Set GetSearchTargetFieldNames = result
        Exit Function
    End If
    Dim sec As ClsStanzaSection
    Dim i As Long
    For i = 1 To secs.Count
        Set sec = secs.Item(i)
        If sec.SectionName = "FIELD" Then
            Dim st As String
            st = ""
            If sec.HasKey("searchTarget") Then st = LCase(Trim(sec.GetValue("searchTarget")))
            If st = "true" Or st = "1" Or st = "yes" Then
                Dim fn As String
                fn = ""
                If sec.HasKey("FieldName") Then fn = sec.GetValue("FieldName")
                If Len(fn) > 0 Then result.Add fn
            End If
        End If
    Next i
    Set GetSearchTargetFieldNames = result
End Function

Private Function SafeStr(ByVal data As Object, ByVal key As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1807] modEntrySearch.SafeStr ENTER"  ' [ADR-0100]
    If data Is Nothing Then Exit Function
    If data.Exists(key) Then SafeStr = CStr(data(key))
End Function

' Phase R-3-Omega legacy helper (kept for now).
Private Function NthFieldValue(ByVal data As Object, ByVal n As Long) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1808] modEntrySearch.NthFieldValue ENTER"  ' [ADR-0100]
    If data Is Nothing Then Exit Function
    Dim meta As String
    meta = "|KnowledgeNo|FormatID|FormatId|FormatVersion|CreatedAt|UpdatedAt|"
    Dim cnt As Long
    cnt = 0
    Dim k As Variant
    For Each k In data.Keys
        If InStr(meta, "|" & CStr(k) & "|") = 0 Then
            cnt = cnt + 1
            If cnt = n Then
                Dim v As String
                v = CStr(data(CStr(k)))
                If Len(v) > 50 Then v = Left(v, 50) & "..."
                NthFieldValue = v
                If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1809] modEntrySearch.NthFieldValue EXIT-OK"  ' [ADR-0100]
                Exit Function
            End If
        End If
    Next k
End Function

Private Function FirstTextFieldValue(ByVal data As Object) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1810] modEntrySearch.FirstTextFieldValue ENTER"  ' [ADR-0100]
    If data Is Nothing Then Exit Function
    Dim k As Variant
    For Each k In data.Keys
        Dim ks As String
        ks = CStr(k)
        If ks <> "KnowledgeNo" And ks <> "FormatID" Then
            Dim v As String
            v = CStr(data(ks))
            If Len(v) > 0 Then
                If Len(v) > 50 Then v = Left(v, 50) & "..."
                FirstTextFieldValue = v
                If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1811] modEntrySearch.FirstTextFieldValue EXIT-OK"  ' [ADR-0100]
                Exit Function
            End If
        End If
    Next k
End Function

' SPEC_DRIFT-NEW3 (2026-05-31): trim a value for the subject / excerpt result columns.
Private Function TrimExcerpt(ByVal v As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1812] modEntrySearch.TrimExcerpt ENTER"  ' [ADR-0100]
    Dim s As String
    s = v
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Trim(s)
    If Len(s) > 50 Then s = Left(s, 50) & "..."
    TrimExcerpt = s
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1813] modEntrySearch.TrimExcerpt EXIT-OK"  ' [ADR-0100]
End Function

' SPEC_DRIFT-NEW3 (2026-05-31): resolve subject/excerpt keys for a record.
Private Sub SubjectAndExcerptKeys(ByVal data As Object, _
                                  ByVal formatId As String, _
                                  ByRef subjKey As String, _
                                  ByRef excKey As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1814] modEntrySearch.SubjectAndExcerptKeys ENTER"  ' [ADR-0100]
    subjKey = ""
    excKey = ""

    Dim L_KENMEI As String, L_JISHO As String, L_TANITSU As String, L_FUKUSU As String
    Dim L_TEJUN As String
    L_KENMEI = ChrW(&H4EF6) & ChrW(&H540D)
    L_JISHO  = ChrW(&H4E8B) & ChrW(&H8C61)
    L_TANITSU = ChrW(&H5358) & ChrW(&H4E00) & ChrW(&H884C)
    L_FUKUSU  = ChrW(&H8907) & ChrW(&H6570) & ChrW(&H884C)
    L_TEJUN  = ChrW(&H624B) & ChrW(&H9806)

    On Error Resume Next
    Dim sections As Collection
    If Len(formatId) > 0 Then
        Set sections = modFormatLoader.LoadFormat(formatId)
    End If
    On Error GoTo 0

    Dim firstSingleLine As String
    Dim firstMultiLine As String
    Dim firstTejunLike As String
    firstSingleLine = ""
    firstMultiLine = ""
    firstTejunLike = ""

    If Not sections Is Nothing Then
        Dim i As Long
        For i = 1 To sections.Count
            Dim sec As Object
            Set sec = sections.Item(i)
            If sec.SectionName = "FIELD" Then
                Dim fName As String, fType As String
                fName = sec.GetValue("FieldName")
                fType = sec.GetValue("FieldType")

                If Len(subjKey) = 0 Then
                    If InStr(fName, L_KENMEI) > 0 Then subjKey = fName
                End If
                If Len(firstSingleLine) = 0 Then
                    If fType = L_TANITSU Then firstSingleLine = fName
                End If

                If Len(excKey) = 0 Then
                    If InStr(fName, L_JISHO) > 0 Then excKey = fName
                End If
                If Len(firstMultiLine) = 0 Then
                    If fType = L_FUKUSU Then firstMultiLine = fName
                End If
                If Len(firstTejunLike) = 0 Then
                    If InStr(fName, L_TEJUN) > 0 Then firstTejunLike = fName
                End If
            End If
        Next i

        If Len(subjKey) = 0 Then subjKey = firstSingleLine
        If Len(excKey) = 0 Then
            If Len(firstMultiLine) > 0 Then
                excKey = firstMultiLine
            ElseIf Len(firstTejunLike) > 0 Then
                excKey = firstTejunLike
            End If
        End If
    End If

    If Len(subjKey) = 0 Or Len(excKey) = 0 Then
        Dim meta As String
        meta = "|KnowledgeNo|FormatID|FormatId|FormatVersion|CreatedAt|UpdatedAt|"
        Dim cnt As Long
        cnt = 0
        Dim k As Variant
        For Each k In data.Keys
            Dim ks As String
            ks = CStr(k)
            If InStr(meta, "|" & ks & "|") = 0 Then
                cnt = cnt + 1
                If cnt = 1 And Len(subjKey) = 0 Then subjKey = ks
                If cnt = 2 And Len(excKey) = 0 Then excKey = ks
                If cnt >= 2 Then Exit For
            End If
        Next k
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1815] modEntrySearch.SubjectAndExcerptKeys EXIT-OK"  ' [ADR-0100]
End Sub

' ================================================================
' Public Sub: Btn_OpenViewFromSheet
' Invoked from the "View" button on M-09 (display sheet, legacy).
' ADR-0090 (2026-06-01): cell address is SSOT-driven via ui_seed.
' M-09 sheet itself is retired per ADR-0085 (SPEC_DRIFT-NEW1); this
' handler is kept for legacy installs that still have the sheet.
' IsTestMode guard skips modal MsgBox during E2E runs.
' ================================================================
Public Sub Btn_OpenViewFromSheet()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1816] modEntrySearch.Btn_OpenViewFromSheet ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    ' [BTN-GUARD-PRELUDE-BEGIN] auto-injected by inject_btn_template.py
    Const BTN As String = "Btn_OpenViewFromSheet"
    Dim XLSM As String
    XLSM = ChrW(&H691C) & ChrW(&H7D22)
    modBtnGuard.LogEnter BTN, XLSM
    If Not modBtnGuard.CheckPrereq(BTN, "config;data_dir", XLSM) Then
        modBtnGuard.LogExit BTN, XLSM, False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1817] modEntrySearch.Btn_OpenViewFromSheet EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If
    ' [BTN-GUARD-PRELUDE-END]
    Dim jpName As String
    jpName = ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & ChrW(&H8868) & ChrW(&H793A)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(jpName)
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets("M-09")
    On Error GoTo ErrHandler
    If ws Is Nothing Then
        Debug.Print "[INFO] Btn_OpenViewFromSheet: M-09 sheet retired (ADR-0085)"
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1818] modEntrySearch.Btn_OpenViewFromSheet EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If

    Dim kidCell As String
    kidCell = ResolveM09KidCell()
    Dim kid As String
    kid = Trim(CStr(ws.Range(kidCell).Value))

    If Len(kid) = 0 Then
        If Not IsTestMode() Then
            Dim msgEmpty As String
            msgEmpty = kidCell & ChrW(&H306B) & ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & ChrW(&H756A) & ChrW(&H53F7) & ChrW(&H3092) & ChrW(&H5165) & ChrW(&H529B) & ChrW(&H3057) & ChrW(&H3066) & ChrW(&H304F) & ChrW(&H3060) & ChrW(&H3055) & ChrW(&H3044)
            MsgBox msgEmpty, vbExclamation, ChrW(&H8A73) & ChrW(&H7D30) & ChrW(&H8868) & ChrW(&H793A)
        End If
        Debug.Print "[WARN] Btn_OpenViewFromSheet: kid empty cell=" & kidCell
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1819] modEntrySearch.Btn_OpenViewFromSheet EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If

    On Error Resume Next
    Dim engine As clsSearchEngine
    Set engine = NewSearchEngine()
    If Not engine Is Nothing Then engine.DisplayKnowledge kid
    On Error GoTo ErrHandler
    ' [BTN-GUARD-EXIT-OK] auto-injected
    modBtnGuard.LogExit BTN, XLSM, True
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1820] modEntrySearch.Btn_OpenViewFromSheet EXIT-OK"  ' [ADR-0100]
    Exit Sub

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1821] modEntrySearch.Btn_OpenViewFromSheet EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ' [BTN-GUARD-ERR-LOG] auto-injected
    modBtnGuard.LogExit BTN, XLSM, False
    Debug.Print "[ERR] Btn_OpenViewFromSheet: " & Err.Number & " " & Err.Description
End Sub

Private Function ResolveM09KidCell() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1822] modEntrySearch.ResolveM09KidCell ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim uiCol As Collection
    Set uiCol = modUILoader.LoadUiDefinition("search", "M-09")
    If Err.Number = 0 And Not uiCol Is Nothing Then
        Dim sec As ClsStanzaSection
        Dim i As Long
        For i = 1 To uiCol.Count
            Set sec = uiCol.Item(i)
            If sec.SectionName = "INPUT" Then
                Dim nm As String
                nm = sec.GetValue("Name")
                If InStr(nm, ChrW(&H756A) & ChrW(&H53F7)) > 0 Then
                    ResolveM09KidCell = sec.GetValue("Cell")
                    If Len(ResolveM09KidCell) > 0 Then Exit Function
                End If
            End If
        Next i
    End If
    Err.Clear
    On Error GoTo 0
    ResolveM09KidCell = "C5"
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1823] modEntrySearch.ResolveM09KidCell EXIT-OK"  ' [ADR-0100]
End Function

Private Function IsTestMode() As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1824] modEntrySearch.IsTestMode ENTER"  ' [ADR-0100]
    On Error Resume Next
    IsTestMode = False
    Dim v As String
    v = Environ$("KNW_TEST_MODE")
    If LCase(v) = "1" Or LCase(v) = "true" Then IsTestMode = True
    On Error GoTo 0
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1825] modEntrySearch.IsTestMode EXIT-OK"  ' [ADR-0100]
End Function

Private Function NewSearchEngine() As clsSearchEngine
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1826] modEntrySearch.NewSearchEngine ENTER"  ' [ADR-0100]
    On Error Resume Next
    Set NewSearchEngine = New clsSearchEngine
    On Error GoTo 0
End Function


' ================================================================
' Function: ContainsAllTerms ([BUG-B9 2026-06-11])
' Outline:  AND match - every space-separated term in terms must appear
'           in target (case-insensitive). Full-width spaces normalized.
' Args:     target - combined searchTarget field text
'           terms  - raw keyword cell value
' Returns:  Boolean - True when ALL terms hit (empty terms -> True)
' ================================================================
Private Function ContainsAllTerms(ByVal target As String, ByVal terms As String) As Boolean
    Dim norm As String
    norm = Replace(terms, ChrW(&H3000), " ")
    Dim parts() As String
    parts = Split(Trim$(norm), " ")
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 Then
            If InStr(1, target, parts(i), vbTextCompare) = 0 Then
                ContainsAllTerms = False
                Exit Function
            End If
        End If
    Next i
    ContainsAllTerms = True
End Function

' ================================================================
' Function: ResolveFormatIdFromDisplay ([BUG-B8 2026-06-11])
' Outline:  Resolve format display name (FormatName) to FormatID via
'           modFormatLoader.LoadFormatList (same pattern as the B5 fix
'           clsUserFormRenderer.ResolveDisplayToFormatId).
' Args:     disp - dropdown cell value (display name or raw FormatID)
' Returns:  String - FormatID when resolved, otherwise disp pass-through
' ================================================================
Private Function ResolveFormatIdFromDisplay(ByVal disp As String) As String
    On Error GoTo ErrHandler
    ResolveFormatIdFromDisplay = disp
    If Len(disp) = 0 Then Exit Function
    Dim formats As Collection
    Set formats = modFormatLoader.LoadFormatList()
    If formats Is Nothing Then Exit Function
    Dim i As Long
    For i = 1 To formats.Count
        Dim ent As Object
        Set ent = formats.Item(i)
        If ent.Exists("Description") Then
            If CStr(ent("Description")) = disp Then
                If ent.Exists("FormatID") Then ResolveFormatIdFromDisplay = CStr(ent("FormatID"))
                Exit Function
            End If
        End If
    Next i
    Exit Function
ErrHandler:
    ' keep pass-through value on error (no resource to release)
End Function

' ================================================================
' Function: MsgNoCondition ([SPEC T-24 / BUG-B10 2026-06-11])
' Outline:  "Enter a search condition" message. JP literal built with
'           ChrW per ADR-0006/0090/0094 (no raw JP literals in code).
' Returns:  String - message text
' ================================================================
Private Function MsgNoCondition() As String
    MsgNoCondition = ChrW(&H691C) & ChrW(&H7D22) & ChrW(&H6761) & ChrW(&H4EF6) & _
                     ChrW(&H3092) & ChrW(&H5165) & ChrW(&H529B) & ChrW(&H3057) & _
                     ChrW(&H3066) & ChrW(&H304F) & ChrW(&H3060) & ChrW(&H3055) & ChrW(&H3044)
End Function

' [USER-REQ 2026-06-09] Look up format display name from format file
Private Function ResolveFormatName(ByVal fmtId As String) As String
    On Error Resume Next
    ResolveFormatName = fmtId
    If Len(fmtId) = 0 Then Exit Function
    Dim secs As Collection
    Set secs = modFormatLoader.LoadFormat(fmtId)
    If secs Is Nothing Then Exit Function
    Dim sec As ClsStanzaSection
    Dim i As Long
    For i = 1 To secs.Count
        Set sec = secs.Item(i)
        If sec.SectionName = "FORMAT" Then
            If sec.HasKey("FormatName") Then
                Dim nm As String
                nm = sec.GetValue("FormatName")
                If Len(nm) > 0 Then ResolveFormatName = nm
            End If
            Exit For
        End If
    Next i
End Function

' [USER-REQ 2026-06-09] Return the first 3 searchTarget=TRUE field VALUES from data
' (field order from format). Empty string if fewer than 3 such fields.
Private Sub GetSearchTargetFieldValues(ByVal data As Object, ByVal fmtId As String, ByRef out() As String)
    On Error Resume Next
    out(1) = "": out(2) = "": out(3) = ""
    If Len(fmtId) = 0 Then Exit Sub
    Dim secs As Collection
    Set secs = modFormatLoader.LoadFormat(fmtId)
    If secs Is Nothing Then Exit Sub
    Dim sec As ClsStanzaSection
    Dim i As Long, k As Long
    k = 0
    For i = 1 To secs.Count
        Set sec = secs.Item(i)
        If sec.SectionName = "FIELD" Then
            Dim st As String
            st = ""
            If sec.HasKey("searchTarget") Then st = LCase(Trim(sec.GetValue("searchTarget")))
            If st = "true" Or st = "1" Or st = "yes" Then
                k = k + 1
                If k <= 3 Then
                    Dim fName As String
                    fName = ""
                    If sec.HasKey("FieldName") Then fName = sec.GetValue("FieldName")
                    If Len(fName) > 0 Then out(k) = SafeStr(data, fName)
                End If
                If k >= 3 Then Exit For
            End If
        End If
    Next i
End Sub
```
