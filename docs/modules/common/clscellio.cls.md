---
title: clsCellIO.cls
description: clsCellIO.cls 縺ｮ繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝会ｼ医さ繝斐・逕ｨ・・---

# clsCellIO.cls

**驟咲ｽｮ蜈・*: `蜈ｨ繝悶ャ繧ｯ蜈ｱ騾啻 逕ｨ縺ｮ VBA 繝｢繧ｸ繝･繝ｼ繝ｫ
**遞ｮ鬘・*: 繧ｯ繝ｩ繧ｹ繝｢繧ｸ繝･繝ｼ繝ｫ

---

## 繝輔ぃ繧､繝ｫ縺ｨ縺励※菫晏ｭ・
繝｡繝｢蟶ｳ・医∪縺溘・莉ｻ諢上・繝・く繧ｹ繝医お繝・ぅ繧ｿ・峨↓荳九・繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝牙・譁・ｒ雋ｼ繧贋ｻ倥￠縲・*`clsCellIO.cls`** 縺ｨ縺・≧蜷榊燕縺ｧ `installer\vba_modules\common\` 驟堺ｸ九↓菫晏ｭ倥＠縺ｦ縺上□縺輔＞縲よ枚蟄励さ繝ｼ繝峨・ ANSI・・hift-JIS・峨∵隼陦後・ CRLF 縺ｫ縺励※縺上□縺輔＞縲・
---

## 繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝・
```vb
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
End
Attribute VB_Name = "clsCellIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================
' Cell value read/write helper (Sprint2 SRP split, ASCII-only)
' target = Worksheet (production) OR Scripting.Dictionary (mock)
' VB_PredeclaredId = True : callable as clsCellIO.Xxx
' ============================================================
Option Explicit

' Read cell value as String (handles Dict mock, Worksheet, errors, Date)
Public Function ReadCellValue(ByVal target As Object, ByVal cellAddr As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0412] clsCellIO.ReadCellValue ENTER"  ' [ADR-0100]
    If TypeName(target) = "Dictionary" Then
        If target.Exists(cellAddr) Then
            Dim dv As Variant
            dv = target(cellAddr)
            If IsNull(dv) Then
                ReadCellValue = ""
            Else
                ReadCellValue = CStr(dv)
            End If
        Else
            ReadCellValue = ""
        End If
    Else
        Dim r As Range
        Set r = target.Range(cellAddr)
        Dim v As Variant
        v = r.Cells(1, 1).Value
        If IsError(v) Then
            ReadCellValue = CStr(CVErr(v))
        ElseIf IsNull(v) Then
            ReadCellValue = ""
        ElseIf IsDate(v) Then
            ReadCellValue = Format(v, "yyyy-mm-dd")
        Else
            ReadCellValue = CStr(v)
        End If
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0413] clsCellIO.ReadCellValue EXIT-OK"  ' [ADR-0100]
End Function

' Write String value to target cell
Public Sub WriteCellValue(ByVal target As Object, ByVal cellAddr As String, ByVal value As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0414] clsCellIO.WriteCellValue ENTER"  ' [ADR-0100]
    If TypeName(target) = "Dictionary" Then
        target(cellAddr) = value
    Else
        Dim r As Range
        Set r = target.Range(cellAddr)
        r.Cells(1, 1).Value = value
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0415] clsCellIO.WriteCellValue EXIT-OK"  ' [ADR-0100]
End Sub
```