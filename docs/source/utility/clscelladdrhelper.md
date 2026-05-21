---
title: clsCellAddrHelper.cls
---

# clsCellAddrHelper.cls

| 項目 | 内容 |
|---|---|
| 層 | ユーティリティ層 |
| 種別 | クラスモジュール (.cls) |
| 配置ブック | 3 ブック共通 |
| 役割 | セル番地の計算（列文字変換・オフセット等）を行うヘルパー |
| 行数 | 71 行 |

## 取り込み先

クラスモジュール（.cls）です。下記コードをコピーし、`clsCellAddrHelper.cls` というファイル名で保存して、VBE の「ファイル → ファイルのインポート」で取り込みます。先頭の `VERSION 1.0 CLASS` から始まる行はクラスモジュールのファイル形式の一部なので、削らずにそのまま保存してください。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
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
    Dim result As String
    result = ""
    Do While colNum > 0
        Dim r As Long
        r = ((colNum - 1) Mod 26)
        result = Chr$(Asc("A") + r) & result
        colNum = (colNum - 1) \ 26
    Loop
    ColLetter = result
End Function

' Parse "C3" -> col=3, row=3 (via ByRef)
Public Sub ParseCellAddr(ByVal addr As String, ByRef colNum As Long, ByRef rowNum As Long)
    Dim i As Long
    Dim letters As String, digits As String
    letters = ""
    digits = ""
    For i = 1 To Len(addr)
        Dim ch As String
        ch = Mid$(addr, i, 1)
        If ch >= "0" And ch <= "9" Then
            digits = digits & ch
        ElseIf (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Then
            letters = letters & UCase(ch)
        End If
    Next i
    rowNum = CLng(Val(digits))
    colNum = 0
    For i = 1 To Len(letters)
        colNum = colNum * 26 + (Asc(Mid$(letters, i, 1)) - Asc("A") + 1)
    Next i
End Sub

' Offset a cell address by (rowOffset, colOffset)
Public Function OffsetCellAddr(ByVal baseCellAddr As String, ByVal rowOffset As Long, ByVal colOffset As Long) As String
    Dim baseCol As Long, baseRow As Long
    ParseCellAddr baseCellAddr, baseCol, baseRow
    Dim newCol As Long, newRow As Long
    newCol = baseCol + colOffset
    newRow = baseRow + rowOffset
    OffsetCellAddr = ColLetter(newCol) & CStr(newRow)
End Function

' GRID row capacity: end-start+1 (rows) if endCellAddr provided, else default
Public Function CalcGridRowCapacity(ByVal startCellAddr As String, ByVal endCellAddr As String, ByVal defaultRows As Long) As Long
    If Len(endCellAddr) > 0 Then
        Dim sCol As Long, sRow As Long, eCol As Long, eRow As Long
        ParseCellAddr startCellAddr, sCol, sRow
        ParseCellAddr endCellAddr, eCol, eRow
        CalcGridRowCapacity = eRow - sRow + 1
    Else
        CalcGridRowCapacity = defaultRows
    End If
End Function
```
