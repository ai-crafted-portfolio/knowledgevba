---
title: clsGridIO.cls
description: clsGridIO.cls のソースコード（コピペ用）
---

# clsGridIO.cls

**配置先**: `共通モジュール (3 ブック全て)` 用の VBA モジュール  
**種類**: クラス モジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\`
- ファイル名: `clsGridIO.cls`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。

---

## ソースコード

```vb
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
End
Attribute VB_Name = "clsGridIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================
' Grid (table) I/O helper (Sprint2 SRP split, ASCII-only)
' VB_PredeclaredId = True : callable as clsGridIO.Xxx
' ============================================================
Option Explicit

Private Const DEFAULT_GRID_ROWS As Long = 50

' Read GRID rows from target into Collection of Dict (1 Dict per row)
' columns from GridSection.Columns (csv) or default Name,Type,Required
Public Function ReadGridFields(ByVal target As Object, ByVal gridSec As Object, ByVal mode As String) As Collection
    Dim fields As New Collection
    Dim startCellAddr As String, endCellAddr As String, columnsCsv As String
    startCellAddr = gridSec.GetValue("StartCell")
    endCellAddr = gridSec.GetValue("EndCell")
    columnsCsv = gridSec.GetValue("Columns")
    If Len(startCellAddr) = 0 Then
        Set ReadGridFields = fields
        Exit Function
    End If
    Dim columnNames() As String
    columnNames = SplitColumns(columnsCsv)
    Dim colCount As Long
    colCount = UBound(columnNames) - LBound(columnNames) + 1
    Dim maxRows As Long
    maxRows = clsCellAddrHelper.CalcGridRowCapacity(startCellAddr, endCellAddr, DEFAULT_GRID_ROWS)
    Dim rowIdx As Long
    For rowIdx = 1 To maxRows
        Dim rowDict As Object
        Set rowDict = CreateObject("Scripting.Dictionary")
        Dim hasAny As Boolean
        hasAny = False
        Dim c As Long
        For c = 0 To colCount - 1
            Dim colName As String, cellAddr As String, v As String
            colName = Trim(columnNames(c))
            cellAddr = clsCellAddrHelper.OffsetCellAddr(startCellAddr, rowIdx - 1, c)
            v = clsCellIO.ReadCellValue(target, cellAddr)
            rowDict(colName) = v
            If Len(v) > 0 Then hasAny = True
        Next c
        If hasAny Then
            Dim nameVal As String
            nameVal = ""
            If rowDict.Exists("Name") Then nameVal = CStr(rowDict("Name"))
            If Len(nameVal) > 0 Then
                If rowDict.Exists("Type") Then
                    Dim t As String
                    t = LCase(CStr(rowDict("Type")))
                    If Not IsValidFieldType(t) Then
                        If mode = "strict" Then
                            Err.Raise vbObjectError + 9203, "clsGridIO.ReadGridFields", _
                                      "Invalid FieldType: " & t
                        End If
                        GoTo NextRow
                    End If
                End If
                If rowDict.Exists("Options") Then
                    Dim optStr As String
                    optStr = CStr(rowDict("Options"))
                    If Len(optStr) > 0 Then
                        Dim opts() As String
                        opts = Split(optStr, vbLf)
                        rowDict("Options") = opts
                    End If
                End If
                fields.Add rowDict
            End If
        End If
NextRow:
    Next rowIdx
    Set ReadGridFields = fields
End Function

' Write Collection of Dict to GRID rows on target
Public Sub WriteGridFields(ByVal target As Object, ByVal gridSec As Object, ByVal fields As Object)
    Dim startCellAddr As String, endCellAddr As String, columnsCsv As String
    startCellAddr = gridSec.GetValue("StartCell")
    endCellAddr = gridSec.GetValue("EndCell")
    columnsCsv = gridSec.GetValue("Columns")
    If Len(startCellAddr) = 0 Then Exit Sub
    Dim columnNames() As String
    columnNames = SplitColumns(columnsCsv)
    Dim colCount As Long
    colCount = UBound(columnNames) - LBound(columnNames) + 1
    Dim maxRows As Long
    maxRows = clsCellAddrHelper.CalcGridRowCapacity(startCellAddr, endCellAddr, DEFAULT_GRID_ROWS)
    Dim writtenRows As Long
    writtenRows = 0
    If Not (fields Is Nothing) Then
        Dim i As Long
        For i = 1 To fields.Count
            If i > maxRows Then Exit For
            Dim rowItem As Object
            Set rowItem = fields.Item(i)
            Dim c As Long
            For c = 0 To colCount - 1
                Dim colName As String, cellAddr As String, valStr As String
                colName = Trim(columnNames(c))
                cellAddr = clsCellAddrHelper.OffsetCellAddr(startCellAddr, i - 1, c)
                valStr = ""
                If rowItem.Exists(colName) Then
                    Dim v As Variant
                    v = rowItem(colName)
                    If IsArray(v) Then
                        valStr = JoinArray(v, vbLf)
                    ElseIf IsNull(v) Then
                        valStr = ""
                    Else
                        valStr = CStr(v)
                    End If
                End If
                clsCellIO.WriteCellValue target, cellAddr, valStr
            Next c
            writtenRows = writtenRows + 1
        Next i
    End If
    Dim r As Long
    For r = writtenRows + 1 To maxRows
        Dim cc As Long
        For cc = 0 To colCount - 1
            clsCellIO.WriteCellValue target, clsCellAddrHelper.OffsetCellAddr(startCellAddr, r - 1, cc), ""
        Next cc
    Next r
End Sub

Public Function IsValidFieldType(ByVal t As String) As Boolean
    Select Case LCase(t)
        Case "string", "text", "long", "integer", "double", "number", _
             "date", "datetime", "boolean", "dropdown", "list", ""
            IsValidFieldType = True
        Case Else
            IsValidFieldType = False
    End Select
End Function

Public Function JoinArray(ByVal arr As Variant, ByVal sep As String) As String
    Dim s As String
    s = ""
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If i > LBound(arr) Then s = s & sep
        s = s & CStr(arr(i))
    Next i
    JoinArray = s
End Function

' Parse "Name,Type,Required" -> array. Default if empty.
Private Function SplitColumns(ByVal columnsCsv As String) As String()
    Dim cols() As String
    If Len(columnsCsv) > 0 Then
        cols = Split(columnsCsv, ",")
    Else
        ReDim cols(2)
        cols(0) = "Name"
        cols(1) = "Type"
        cols(2) = "Required"
    End If
    SplitColumns = cols
End Function
```
