---
title: modEntrySearch.bas
description: modEntrySearch.bas のソースコード（コピペ用）
---

# modEntrySearch.bas

**配置先**: `検索.xlsm` 用の VBA モジュール  
**種類**: 標準モジュール

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
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ResolveSearchSheet()
    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    ws.Unprotect
    On Error GoTo ErrHandler

    Dim keyword As String, fmtFilter As String
    keyword = Trim$(CStr(ws.Range(M08_CELL_KEYWORD).Value))
    fmtFilter = Trim$(CStr(ws.Range(M08_CELL_FORMAT).Value))

    Dim clearRng As Range
    Set clearRng = ws.Range(ws.Cells(M08_RESULT_START_ROW, 1), ws.Cells(M08_RESULT_LAST_ROW, 8))
    On Error Resume Next
    clearRng.ClearContents
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
        Exit Sub
    End If

    Dim outRow As Long
    outRow = M08_RESULT_START_ROW
    Dim hitCount As Long
    hitCount = 0

    Dim f As Object
    For Each f In fso.GetFolder(dataDir).Files
        If LCase$(fso.GetExtensionName(f.Name)) = "txt" Then
            Dim knwNo As String
            knwNo = fso.GetBaseName(f.Name)
            Dim data As Object
            Set data = modKnowledgeFileIO.LoadKnowledge(knwNo)
            If Not data Is Nothing Then
                If MatchesSearchV23(data, keyword, fmtFilter) Then
                    ' SPEC_DRIFT-NEW3 fix (2026-05-31): format-aware subject/excerpt
                    ' resolution. Old NthFieldValue(data,1/2) placed FldA into 件名 col
                    ' and FldB into 事象抜粋 col for SAGYO - wrong order.
                    Dim subjKey As String, excKey As String
                    SubjectAndExcerptKeys data, SafeStr(data, "FormatID"), subjKey, excKey

                    ws.Cells(outRow, 1).Value = hitCount + 1
                    ws.Cells(outRow, 2).Value = knwNo
                    ws.Cells(outRow, 3).Value = SafeStr(data, "FormatID")
                    ws.Cells(outRow, 4).Value = SafeStr(data, "CreatedAt")
                    ws.Cells(outRow, 5).Value = SafeStr(data, "UpdatedAt")
                    ws.Cells(outRow, 6).Value = TrimExcerpt(SafeStr(data, subjKey))
                    ws.Cells(outRow, 7).Value = TrimExcerpt(SafeStr(data, excKey))
                    ' 2026-05-31 UX: paint column A (No) blue + underline
                    ' (hyperlink-look) so the user notices it is clickable.
                    ' DoubleClick on any of A..G in the row triggers
                    ' OpenViewWithId (ThisWorkbook_search.cls allows A..G).
                    With ws.Cells(outRow, 1)
                        .Font.Color = RGB(0, 0, 255)
                        .Font.Underline = xlUnderlineStyleSingle
                    End With
                    outRow = outRow + 1
                    hitCount = hitCount + 1
                    If outRow > M08_RESULT_LAST_ROW Then Exit For
                End If
            End If
        End If
    Next f

    On Error Resume Next
    Dim lg As clsLogger
    Set lg = New clsLogger
    lg.Init ThisWorkbook.Worksheets("LOG")
    lg.LogInfo "modEntrySearch", "Btn_SearchV23", _
        "search done keyword=" & keyword & " fmt=" & fmtFilter & " hits=" & hitCount, _
        "", "LOG-M08-SEARCH-OK"
    On Error GoTo 0
    Exit Sub

ErrHandler:
    Debug.Print "[ERR] Btn_SearchV23: " & Err.Number & " " & Err.Description
End Sub

Public Sub Btn_SearchClearV23()
    On Error GoTo ErrHandler
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
    Exit Sub
ErrHandler:
    Debug.Print "[ERR] Btn_SearchClearV23: " & Err.Number & " " & Err.Description
End Sub

' --- helpers ---

Private Function ResolveSearchSheet() As Worksheet
    On Error Resume Next
    Dim jp As String
    jp = ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & ChrW(&H691C) & ChrW(&H7D22)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(jp)
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets("M-08")
    Set ResolveSearchSheet = ws
    On Error GoTo 0
End Function

Private Function MatchesSearchV23(ByVal data As Object, ByVal keyword As String, ByVal fmtFilter As String) As Boolean
    If Len(fmtFilter) > 0 Then
        Dim actual As String
        actual = SafeStr(data, "FormatID")
        If actual <> fmtFilter Then
            MatchesSearchV23 = False
            Exit Function
        End If
    End If
    If Len(keyword) = 0 Then
        MatchesSearchV23 = True
        Exit Function
    End If
    Dim k As Variant
    For Each k In data.Keys
        Dim v As String
        v = CStr(data(CStr(k)))
        If InStr(1, v, keyword, vbTextCompare) > 0 Then
            MatchesSearchV23 = True
            Exit Function
        End If
    Next k
    MatchesSearchV23 = False
End Function

Private Function SafeStr(ByVal data As Object, ByVal key As String) As String
    If data Is Nothing Then Exit Function
    If data.Exists(key) Then SafeStr = CStr(data(key))
End Function

' Phase R-3-Omega legacy helper (kept for now).
Private Function NthFieldValue(ByVal data As Object, ByVal n As Long) As String
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
                If Len(v) > 50 Then v = Left$(v, 50) & "..."
                NthFieldValue = v
                Exit Function
            End If
        End If
    Next k
End Function

Private Function FirstTextFieldValue(ByVal data As Object) As String
    If data Is Nothing Then Exit Function
    Dim k As Variant
    For Each k In data.Keys
        Dim ks As String
        ks = CStr(k)
        If ks <> "KnowledgeNo" And ks <> "FormatID" Then
            Dim v As String
            v = CStr(data(ks))
            If Len(v) > 0 Then
                If Len(v) > 50 Then v = Left$(v, 50) & "..."
                FirstTextFieldValue = v
                Exit Function
            End If
        End If
    Next k
End Function

' SPEC_DRIFT-NEW3 (2026-05-31): trim a value for the subject / excerpt result columns.
Private Function TrimExcerpt(ByVal v As String) As String
    Dim s As String
    s = v
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Trim$(s)
    If Len(s) > 50 Then s = Left$(s, 50) & "..."
    TrimExcerpt = s
End Function

' SPEC_DRIFT-NEW3 (2026-05-31): resolve subject/excerpt keys for a record.
Private Sub SubjectAndExcerptKeys(ByVal data As Object, _
                                  ByVal formatId As String, _
                                  ByRef subjKey As String, _
                                  ByRef excKey As String)
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
    On Error GoTo ErrHandler
    Dim jpName As String
    jpName = ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & ChrW(&H8868) & ChrW(&H793A)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(jpName)
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets("M-09")
    On Error GoTo ErrHandler
    If ws Is Nothing Then
        Debug.Print "[INFO] Btn_OpenViewFromSheet: M-09 sheet retired (ADR-0085)"
        Exit Sub
    End If

    Dim kidCell As String
    kidCell = ResolveM09KidCell()
    Dim kid As String
    kid = Trim$(CStr(ws.Range(kidCell).Value))

    If Len(kid) = 0 Then
        If Not IsTestMode() Then
            Dim msgEmpty As String
            msgEmpty = kidCell & ChrW(&H306B) & ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & ChrW(&H756A) & ChrW(&H53F7) & ChrW(&H3092) & ChrW(&H5165) & ChrW(&H529B) & ChrW(&H3057) & ChrW(&H3066) & ChrW(&H304F) & ChrW(&H3060) & ChrW(&H3055) & ChrW(&H3044)
            MsgBox msgEmpty, vbExclamation, "Btn_OpenViewFromSheet"
        End If
        Debug.Print "[WARN] Btn_OpenViewFromSheet: kid empty cell=" & kidCell
        Exit Sub
    End If

    On Error Resume Next
    Dim engine As clsSearchEngine
    Set engine = NewSearchEngine()
    If Not engine Is Nothing Then engine.DisplayKnowledge kid
    On Error GoTo ErrHandler
    Exit Sub

ErrHandler:
    Debug.Print "[ERR] Btn_OpenViewFromSheet: " & Err.Number & " " & Err.Description
End Sub

Private Function ResolveM09KidCell() As String
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
End Function

Private Function IsTestMode() As Boolean
    On Error Resume Next
    IsTestMode = False
    Dim v As String
    v = Environ$("KNW_TEST_MODE")
    If LCase$(v) = "1" Or LCase$(v) = "true" Then IsTestMode = True
    On Error GoTo 0
End Function

Private Function NewSearchEngine() As clsSearchEngine
    On Error Resume Next
    Set NewSearchEngine = New clsSearchEngine
    On Error GoTo 0
End Function
```
