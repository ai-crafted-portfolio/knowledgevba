---
title: modEntrySearch.bas
---

# modEntrySearch.bas

| 項目 | 内容 |
|---|---|
| 層 | エントリポイント層 |
| 種別 | 標準モジュール (.bas) |
| 配置ブック | 検索.xlsm |
| 役割 | ナレッジ検索画面のボタン処理。検索条件セルの読み取りと結果グリッドへの書き出し |
| 行数 | 173 行 |

## 取り込み先

標準モジュール（.bas）です。下記コードをコピーし、`modEntrySearch.bas` というファイル名で保存して、VBE の「ファイル → ファイルのインポート」で取り込みます。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
Attribute VB_Name = "modEntrySearch"
' ============================================================
' modEntrySearch (Phase 3 / slim, Sprint2 SRP split, ASCII-only)
' Bridge + Workflow + Btn_, helpers extracted
' ============================================================
Option Explicit

Public Function BuildSearchDictFromCells(ByVal target As Object, ByVal uiSections As Collection) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    If uiSections Is Nothing Then
        Set BuildSearchDictFromCells = result
        Exit Function
    End If
    Dim mode As String
    mode = LCase(modConfigHolder.GetValueOrDefault("uiSchemaFailMode", "warn"))
    Dim sec As ClsStanzaSection
    Dim i As Long
    For i = 1 To uiSections.Count
        Set sec = uiSections.Item(i)
        If sec.SectionName = "INPUT" Then
            Dim cellAddr As String, fieldName As String
            cellAddr = sec.GetValue("Cell")
            fieldName = sec.GetValue("Name")
            If Len(fieldName) = 0 Or Len(cellAddr) = 0 Then
                If mode = "strict" Then
                    Err.Raise vbObjectError + 9101, "modEntrySearch.BuildSearchDictFromCells", _
                              "INPUT stanza missing Name or Cell key"
                End If
            Else
                Dim v As String
                v = clsCellIO.ReadCellValue(target, cellAddr)
                If Len(v) > 0 Then result(fieldName) = v
            End If
        End If
    Next i
    Set BuildSearchDictFromCells = result
End Function

Public Sub WriteResultsToGrid(ByVal target As Object, ByVal gridSection As ClsStanzaSection, ByVal results As Collection)
    If gridSection Is Nothing Then Exit Sub
    If results Is Nothing Then Set results = New Collection
    Dim startCellAddr As String, endCellAddr As String, columnsCsv As String
    startCellAddr = gridSection.GetValue("StartCell")
    endCellAddr = gridSection.GetValue("EndCell")
    columnsCsv = gridSection.GetValue("Columns")
    If Len(startCellAddr) = 0 Then Exit Sub
    Dim columnNames() As String
    If Len(columnsCsv) > 0 Then
        columnNames = Split(columnsCsv, ",")
    Else
        ReDim columnNames(0)
    End If
    Dim maxCols As Long
    maxCols = UBound(columnNames) - LBound(columnNames) + 1
    If maxCols < 1 Then maxCols = 1
    Dim maxRows As Long
    maxRows = clsCellAddrHelper.CalcGridRowCapacity(startCellAddr, endCellAddr, 100)
    Dim writtenRows As Long
    writtenRows = 0
    Dim rowIdx As Long
    For rowIdx = 1 To results.Count
        If rowIdx > maxRows Then Exit For
        Dim rowDict As Object
        Set rowDict = results.Item(rowIdx)
        Dim c As Long
        For c = 0 To maxCols - 1
            Dim colName As String, cellAddr As String
            colName = Trim(columnNames(c))
            cellAddr = clsCellAddrHelper.OffsetCellAddr(startCellAddr, rowIdx - 1, c)
            If Not (rowDict Is Nothing) And rowDict.Exists(colName) Then
                clsCellIO.WriteCellValue target, cellAddr, CStr(rowDict(colName))
            Else
                clsCellIO.WriteCellValue target, cellAddr, ""
            End If
        Next c
        writtenRows = writtenRows + 1
    Next rowIdx
    Dim r As Long
    For r = writtenRows + 1 To maxRows
        Dim cc As Long
        For cc = 0 To maxCols - 1
            clsCellIO.WriteCellValue target, clsCellAddrHelper.OffsetCellAddr(startCellAddr, r - 1, cc), ""
        Next cc
    Next r
    If results.Count > maxRows Then
        Dim overflowAddr As String
        overflowAddr = clsCellAddrHelper.OffsetCellAddr(startCellAddr, maxRows - 1, 0)
        clsCellIO.WriteCellValue target, overflowAddr, "[+" & (results.Count - maxRows) & "]"
    End If
End Sub

Public Function GetSelectedRow(ByVal target As Object, ByVal gridSection As ClsStanzaSection) As Long
    If gridSection Is Nothing Then
        GetSelectedRow = 0
        Exit Function
    End If
    If TypeName(target) = "Dictionary" Then
        GetSelectedRow = 0
        Exit Function
    End If
    On Error GoTo ErrHandler
    Dim startCellAddr As String
    startCellAddr = gridSection.GetValue("StartCell")
    If Len(startCellAddr) = 0 Then
        GetSelectedRow = 0
        Exit Function
    End If
    Dim startRow As Long
    startRow = target.Range(startCellAddr).Row
    Dim activeRow As Long
    activeRow = Application.ActiveCell.Row
    Dim diff As Long
    diff = activeRow - startRow + 1
    If diff < 1 Then
        GetSelectedRow = 0
    Else
        GetSelectedRow = diff
    End If
    Exit Function
ErrHandler:
    GetSelectedRow = 0
End Function

Public Function Search_Workflow(ByVal searchTarget As Object, ByVal searchUi As Collection, _
                                  ByVal resultTarget As Object, ByVal resultUi As Collection, _
                                  ByVal searchExecutor As Object) As Long
    On Error GoTo ErrHandler
    Dim searchDict As Object
    Set searchDict = BuildSearchDictFromCells(searchTarget, searchUi)
    Dim results As Collection
    If searchExecutor Is Nothing Then
        Set results = New Collection
    Else
        Set results = searchExecutor.ExecuteSearch(searchDict)
    End If
    If results Is Nothing Then Set results = New Collection
    Dim gridSec As ClsStanzaSection
    Dim sec As ClsStanzaSection
    Dim i As Long
    For i = 1 To resultUi.Count
        Set sec = resultUi.Item(i)
        If sec.SectionName = "GRID" Then
            Set gridSec = sec
            Exit For
        End If
    Next i
    If Not gridSec Is Nothing Then
        WriteResultsToGrid resultTarget, gridSec, results
    End If
    Search_Workflow = results.Count
    ' S5-LOG-02: SAVE-EXIT-OK-II-008 (Search_Workflow success exit, grid write 完了)
    On Error Resume Next
    Dim oLogger008 As clsLogger
    Set oLogger008 = New clsLogger
    oLogger008.Init ThisWorkbook.Worksheets("LOG")
    oLogger008.LogInfo "modEntrySearch", "Search_Workflow", "検索完了 count=" & results.Count, "", "SAVE-EXIT-OK-II-008"
    On Error GoTo 0
    Exit Function
ErrHandler:
    Search_Workflow = 0
End Function

Public Sub Btn_Search_v21()
    On Error GoTo ErrHandler
    Dim searchUi As Collection: Set searchUi = modUILoader.LoadUiDefinition("Kensaku", "M-08")
    Dim resultUi As Collection: Set resultUi = modUILoader.LoadUiDefinition("Kensaku", "M-07")
    If searchUi.Count = 0 Or resultUi.Count = 0 Then Exit Sub
    Search_Workflow ActiveSheet, searchUi, ActiveSheet, resultUi, Nothing
    Exit Sub
ErrHandler:
    Debug.Print "[ERR] Btn_Search_v21: " & Err.Number & " " & Err.Description
End Sub
```
