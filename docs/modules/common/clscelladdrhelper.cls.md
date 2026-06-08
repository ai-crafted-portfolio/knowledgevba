---
title: clsCellAddrHelper.cls
description: clsCellAddrHelper.cls 縺ｮ繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝会ｼ医さ繝斐・逕ｨ・・---

# clsCellAddrHelper.cls

**驟咲ｽｮ蜈・*: `蜈ｨ繝悶ャ繧ｯ蜈ｱ騾啻 逕ｨ縺ｮ VBA 繝｢繧ｸ繝･繝ｼ繝ｫ
**遞ｮ鬘・*: 繧ｯ繝ｩ繧ｹ繝｢繧ｸ繝･繝ｼ繝ｫ

---

## 繝輔ぃ繧､繝ｫ縺ｨ縺励※菫晏ｭ・
繝｡繝｢蟶ｳ・医∪縺溘・莉ｻ諢上・繝・く繧ｹ繝医お繝・ぅ繧ｿ・峨↓荳九・繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝牙・譁・ｒ雋ｼ繧贋ｻ倥￠縲・*`clsCellAddrHelper.cls`** 縺ｨ縺・≧蜷榊燕縺ｧ `installer\vba_modules\common\` 驟堺ｸ九↓菫晏ｭ倥＠縺ｦ縺上□縺輔＞縲よ枚蟄励さ繝ｼ繝峨・ ANSI・・hift-JIS・峨∵隼陦後・ CRLF 縺ｫ縺励※縺上□縺輔＞縲・
---

## 繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝・
```vb
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
End
Attribute VB_Name = "clsCellAddrHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================
' Cell address helper (Sprint2 SRP split, ASCII-only)
' VB_PredeclaredId = True : callable as clsCellAddrHelper.Xxx
' ============================================================
Option Explicit

' Convert column index (1=A, 26=Z, 27=AA...) to letter
Public Function ColLetter(ByVal colNum As Long) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0394] clsCellAddrHelper.ColLetter ENTER"  ' [ADR-0100]
    Dim result As String
    result = ""
    Do While colNum > 0
        Dim r As Long
        r = ((colNum - 1) Mod 26)
        result = Chr$(Asc("A") + r) & result
        colNum = (colNum - 1) \ 26
    Loop
    ColLetter = result
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0395] clsCellAddrHelper.ColLetter EXIT-OK"  ' [ADR-0100]
End Function

' Parse "C3" -> col=3, row=3 (via ByRef)
Public Sub ParseCellAddr(ByVal addr As String, ByRef colNum As Long, ByRef rowNum As Long)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0396] clsCellAddrHelper.ParseCellAddr ENTER"  ' [ADR-0100]
    Dim i As Long
    Dim letters As String, digits As String
    letters = ""
    digits = ""
    For i = 1 To Len(addr)
        Dim ch As String
        ch = Mid(addr, i, 1)
        If ch >= "0" And ch <= "9" Then
            digits = digits & ch
        ElseIf (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Then
            letters = letters & UCase(ch)
        End If
    Next i
    rowNum = CLng(Val(digits))
    colNum = 0
    For i = 1 To Len(letters)
        colNum = colNum * 26 + (Asc(Mid(letters, i, 1)) - Asc("A") + 1)
    Next i
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0397] clsCellAddrHelper.ParseCellAddr EXIT-OK"  ' [ADR-0100]
End Sub

' Offset a cell address by (rowOffset, colOffset)
Public Function OffsetCellAddr(ByVal baseCellAddr As String, ByVal rowOffset As Long, ByVal colOffset As Long) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0398] clsCellAddrHelper.OffsetCellAddr ENTER"  ' [ADR-0100]
    Dim baseCol As Long, baseRow As Long
    ParseCellAddr baseCellAddr, baseCol, baseRow
    Dim newCol As Long, newRow As Long
    newCol = baseCol + colOffset
    newRow = baseRow + rowOffset
    OffsetCellAddr = ColLetter(newCol) & CStr(newRow)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0399] clsCellAddrHelper.OffsetCellAddr EXIT-OK"  ' [ADR-0100]
End Function

' GRID row capacity: end-start+1 (rows) if endCellAddr provided, else default
Public Function CalcGridRowCapacity(ByVal startCellAddr As String, ByVal endCellAddr As String, ByVal defaultRows As Long) As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0400] clsCellAddrHelper.CalcGridRowCapacity ENTER"  ' [ADR-0100]
    If Len(endCellAddr) > 0 Then
        Dim sCol As Long, sRow As Long, eCol As Long, eRow As Long
        ParseCellAddr startCellAddr, sCol, sRow
        ParseCellAddr endCellAddr, eCol, eRow
        CalcGridRowCapacity = eRow - sRow + 1
    Else
        CalcGridRowCapacity = defaultRows
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0401] clsCellAddrHelper.CalcGridRowCapacity EXIT-OK"  ' [ADR-0100]
End Function
```