---
title: clsGridIO.cls
description: clsGridIO.cls のソースコード（コピペ用）
---

# clsGridIO.cls

**配置先**: 共通モジュール（検索.xlsm / 管理.xlsm 共通）
**種類**: クラスモジュール
**更新日**: 2026-06-30 14:44 JST

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `clsGridIO.cls`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。
> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。
> 改行コードは CRLF（Windows 標準）のままで OK です。

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
' iter19 ADR-0090: ui_seed Columns CSV uses v2.3 keys
'   (fieldName/fieldType/required/options) but downstream (SaveFormat_Workflow,
'   WriteFormatDictToCells.Row->FieldName etc.) expects canonical keys
'   (Name/Type/Required/Options). NormalizeColumnNameToCanonical bridges both
'   without mutating ui_seed (which is SSOT for visible column headers).
Public Function ReadGridFields(ByVal target As Object, ByVal gridSec As Object, ByVal mode As String) As Collection
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0488] clsGridIO.ReadGridFields ENTER"  ' [ADR-0100]
    Dim fields As New Collection
    Dim startCellAddr As String, endCellAddr As String, columnsCsv As String
    startCellAddr = gridSec.GetValue("StartCell")
    endCellAddr = gridSec.GetValue("EndCell")
    columnsCsv = gridSec.GetValue("Columns")
    If Len(startCellAddr) = 0 Then
        Set ReadGridFields = fields
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0489] clsGridIO.ReadGridFields EXIT-OK"  ' [ADR-0100]
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
            Dim canonicalName As String
            colName = Trim(columnNames(c))
            canonicalName = NormalizeColumnNameToCanonical(colName)
            cellAddr = clsCellAddrHelper.OffsetCellAddr(startCellAddr, rowIdx - 1, c)
            v = clsCellIO.ReadCellValue(target, cellAddr)
            ' Store under canonical key so downstream Workflow accessors
            ' (Name/Type/Required/Options) find the value regardless of CSV form.
            rowDict(canonicalName) = v
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
                        ' 2026-06-11 fix: M-03 cell input is comma-separated (spec),
                        ' the canonical file format is pipe-separated (FAULT/SAGYO),
                        ' and the old code split on vbLf (matched nothing).
                        ' Normalize comma/vbLf to pipe, then split.
                        optStr = Replace(optStr, vbLf, "|")
                        optStr = Replace(optStr, ",", "|")
                        optStr = Replace(optStr, ChrW(&HFF0C), "|")   ' full-width comma
                        Dim opts() As String
                        opts = Split(optStr, "|")
                        Dim oi As Long
                        For oi = LBound(opts) To UBound(opts)
                            opts(oi) = Trim(opts(oi))
                        Next oi
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0490] clsGridIO.WriteGridFields ENTER"  ' [ADR-0100]
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
                Dim canonicalName2 As String
                colName = Trim(columnNames(c))
                ' iter19 ADR-0090: Read canonical key from row dict regardless
                ' of CSV form ("fieldName" -> "Name", etc.).
                canonicalName2 = NormalizeColumnNameToCanonical(colName)
                cellAddr = clsCellAddrHelper.OffsetCellAddr(startCellAddr, i - 1, c)
                valStr = ""
                If rowItem.Exists(canonicalName2) Then
                    Dim v As Variant
                    v = rowItem(canonicalName2)
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
    ' [USER-REQ 2026-06-09] Dynamic grid border sizing on load.
    ' (a) Apply thin borders to all data rows so data does not visually
    '     overflow past the pre-allocated grid (InitialRows boundary).
    ' (b) Clear borders on rows past data so empty rows do not show
    '     leftover lines from a previous larger grid.
    If writtenRows > 0 Then
        On Error Resume Next
        ' [USER-REQ 2026-06-09] Unlock grid data cells so user can edit them
        ' directly even when sheet protection (ProtectSheet "light") is on.
        ' Spec: M-03 GRID = direct cell edit (短い値はセル直接編集).
        Dim unlockFirst As String, unlockLast As String
        unlockFirst = clsCellAddrHelper.OffsetCellAddr(startCellAddr, 0, 0)
        unlockLast = clsCellAddrHelper.OffsetCellAddr(startCellAddr, writtenRows - 1, colCount - 1)
        Dim unlockRng As Object
        Set unlockRng = target.Range(unlockFirst & ":" & unlockLast)
        Dim sheetWasProtected As Boolean
        sheetWasProtected = False
        If target.ProtectContents Then
            sheetWasProtected = True
            target.Unprotect
        End If
        unlockRng.Locked = False
        If sheetWasProtected Then
            target.Protect Password:="", UserInterfaceOnly:=True, _
                            AllowFormattingCells:=True, _
                            AllowFormattingColumns:=True, _
                            AllowFormattingRows:=True
        End If
        ' (a) borders on data rows
        Dim dataFirstAddr As String, dataLastAddr As String
        dataFirstAddr = clsCellAddrHelper.OffsetCellAddr(startCellAddr, 0, 0)
        dataLastAddr = clsCellAddrHelper.OffsetCellAddr(startCellAddr, writtenRows - 1, colCount - 1)
        Dim dataRng As Object
        Set dataRng = target.Range(dataFirstAddr & ":" & dataLastAddr)
        ' xlEdgeLeft=7, xlEdgeTop=8, xlEdgeBottom=9, xlEdgeRight=10
        ' xlInsideVertical=11, xlInsideHorizontal=12
        ' xlContinuous=1, xlThin=2
        Dim bdIdx As Variant
        For Each bdIdx In Array(7, 8, 9, 10, 11, 12)
            dataRng.Borders(CLng(bdIdx)).LineStyle = 1  ' xlContinuous
            dataRng.Borders(CLng(bdIdx)).Weight = 2     ' xlThin
            dataRng.Borders(CLng(bdIdx)).Color = RGB(191, 191, 191)  ' #BFBFBF per spec
        Next bdIdx
        ' (b) clear borders past data, up to maxRows
        If writtenRows < maxRows Then
            Dim firstEmptyRow As Long, lastEmptyRow As Long
            firstEmptyRow = writtenRows + 1
            lastEmptyRow = maxRows
            Dim firstAddr As String, lastAddr As String
            firstAddr = clsCellAddrHelper.OffsetCellAddr(startCellAddr, firstEmptyRow - 1, 0)
            lastAddr = clsCellAddrHelper.OffsetCellAddr(startCellAddr, lastEmptyRow - 1, colCount - 1)
            Dim emptyRng As Object
            Set emptyRng = target.Range(firstAddr & ":" & lastAddr)
            emptyRng.Borders.LineStyle = -4142  ' xlLineStyleNone
        End If
        On Error GoTo 0
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0491] clsGridIO.WriteGridFields EXIT-OK"  ' [ADR-0100]
End Sub

Public Function IsValidFieldType(ByVal t As String) As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0492] clsGridIO.IsValidFieldType ENTER"  ' [ADR-0100]
    Dim lt As String
    lt = LCase(t)
    Select Case lt
        Case "string", "text", "long", "integer", "double", "number", _
             "date", "datetime", "boolean", "dropdown", "list", ""
            IsValidFieldType = True
            If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0493] clsGridIO.IsValidFieldType EXIT-OK"  ' [ADR-0100]
            Exit Function
    End Select
    ' iter19 ADR-0090 / ADR-0088: accept canonical CJK base-type names
    ' (NormalizeFieldType output: 単一行 / 複数行 / 日付 / 選択). Without this,
    ' ReadGridFields drops every CJK-typed row, leaving SaveFormat with 0
    ' fields and Btn_LoadFormat verifies failing with sample="" (case 17/18).
    Dim single_line As String, multi_line As String
    Dim date_type As String, select_type As String
    single_line = ChrW(&H5358) & ChrW(&H4E00) & ChrW(&H884C)
    multi_line = ChrW(&H8907) & ChrW(&H6570) & ChrW(&H884C)
    date_type = ChrW(&H65E5) & ChrW(&H4ED8)
    select_type = ChrW(&H9078) & ChrW(&H629E)
    ' 2026-06-11 fix (C-03-001): the M-03 grid carries UI labels (6 values)
    ' and they reach this validator BEFORE NormalizeFieldType runs. The four
    ' labels below were rejected, so rows typed with them were silently
    ' dropped in warn mode (fields lost on save).
    Dim lbl_1line As String, lbl_multiline As String
    Dim lbl_number As String, lbl_check As String
    lbl_1line = "1" & ChrW(&H884C) & ChrW(&H30C6) & ChrW(&H30AD) & ChrW(&H30B9) & ChrW(&H30C8)
    lbl_multiline = ChrW(&H8907) & ChrW(&H6570) & ChrW(&H884C) & ChrW(&H30C6) & ChrW(&H30AD) & ChrW(&H30B9) & ChrW(&H30C8)
    lbl_number = ChrW(&H6570) & ChrW(&H5024)
    lbl_check = ChrW(&H30C1) & ChrW(&H30A7) & ChrW(&H30C3) & ChrW(&H30AF)
    If t = single_line Or t = multi_line Or t = date_type Or t = select_type Then
        IsValidFieldType = True
    ElseIf t = lbl_1line Or t = lbl_multiline Or t = lbl_number Or t = lbl_check Then
        IsValidFieldType = True
    Else
        IsValidFieldType = False
    End If
End Function

Public Function JoinArray(ByVal arr As Variant, ByVal sep As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0494] clsGridIO.JoinArray ENTER"  ' [ADR-0100]
    Dim s As String
    s = ""
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If i > LBound(arr) Then s = s & sep
        s = s & CStr(arr(i))
    Next i
    JoinArray = s
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0495] clsGridIO.JoinArray EXIT-OK"  ' [ADR-0100]
End Function

' iter19 ADR-0090: Map ui_seed column-CSV keys to canonical row-dict keys used
' by SaveFormat_Workflow / LoadFormat_Workflow / WriteFormatDictToCells.
' ui_seed (v2.3 M-03.txt) declares Columns=No,seq,fieldName,fieldType,required,
' rows,options. Existing workflows store rows under Name/Type/Required/Options.
' Aliases unify both forms without changing ui_seed (which is the visible header
' SSOT) or Workflow keys (which are the in-memory canonical form).
Public Function NormalizeColumnNameToCanonical(ByVal raw As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0496] clsGridIO.NormalizeColumnNameToCanonical ENTER"  ' [ADR-0100]
    Dim t As String
    t = LCase(Trim(raw))
    Select Case t
        Case "fieldname", "name"
            NormalizeColumnNameToCanonical = "Name"
        Case "fieldtype", "type"
            NormalizeColumnNameToCanonical = "Type"
        Case "required"
            NormalizeColumnNameToCanonical = "Required"
        Case "options"
            NormalizeColumnNameToCanonical = "Options"
        Case Else
            ' Pass-through for No / seq / rows / future columns.
            NormalizeColumnNameToCanonical = Trim(raw)
    End Select
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0497] clsGridIO.NormalizeColumnNameToCanonical EXIT-OK"  ' [ADR-0100]
End Function

' Parse "Name,Type,Required" -> array. Default if empty.
Private Function SplitColumns(ByVal columnsCsv As String) As String()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0498] clsGridIO.SplitColumns ENTER"  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0499] clsGridIO.SplitColumns EXIT-OK"  ' [ADR-0100]
End Function
```
