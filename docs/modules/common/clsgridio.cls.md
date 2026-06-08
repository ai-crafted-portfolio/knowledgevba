---
title: clsGridIO.cls
description: clsGridIO.cls 縺ｮ繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝会ｼ医さ繝斐・逕ｨ・・---

# clsGridIO.cls

**驟咲ｽｮ蜈・*: `蜈ｨ繝悶ャ繧ｯ蜈ｱ騾啻 逕ｨ縺ｮ VBA 繝｢繧ｸ繝･繝ｼ繝ｫ
**遞ｮ鬘・*: 繧ｯ繝ｩ繧ｹ繝｢繧ｸ繝･繝ｼ繝ｫ

---

## 繝輔ぃ繧､繝ｫ縺ｨ縺励※菫晏ｭ・
繝｡繝｢蟶ｳ・医∪縺溘・莉ｻ諢上・繝・く繧ｹ繝医お繝・ぅ繧ｿ・峨↓荳九・繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝牙・譁・ｒ雋ｼ繧贋ｻ倥￠縲・*`clsGridIO.cls`** 縺ｨ縺・≧蜷榊燕縺ｧ `installer\vba_modules\common\` 驟堺ｸ九↓菫晏ｭ倥＠縺ｦ縺上□縺輔＞縲よ枚蟄励さ繝ｼ繝峨・ ANSI・・hift-JIS・峨∵隼陦後・ CRLF 縺ｫ縺励※縺上□縺輔＞縲・
---

## 繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝・
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
    Dim ui_1line As String, ui_multiline As String, ui_number As String
    Dim ui_date As String, ui_select As String, ui_check As String
    ui_1line = "1" & ChrW(&H884C) & ChrW(&H30C6) & ChrW(&H30AD) & ChrW(&H30B9) & ChrW(&H30C8)
    ui_multiline = ChrW(&H8907) & ChrW(&H6570) & ChrW(&H884C) & ChrW(&H30C6) & ChrW(&H30AD) & ChrW(&H30B9) & ChrW(&H30C8)
    ui_number = ChrW(&H6570) & ChrW(&H5024)
    ui_date = ChrW(&H65E5) & ChrW(&H4ED8)
    ui_select = ChrW(&H9078) & ChrW(&H629E)
    ui_check = ChrW(&H30C1) & ChrW(&H30A7) & ChrW(&H30C3) & ChrW(&H30AF)
    If t = single_line Or t = multi_line Or t = date_type Or t = select_type Then
        IsValidFieldType = True
    ElseIf t = ui_1line Or t = ui_multiline Or t = ui_number Or t = ui_date Or t = ui_select Or t = ui_check Then
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