---
title: clsCellAddrHelper.cls
description: clsCellAddrHelper.cls のソースコード（コピペ用）
---

# clsCellAddrHelper.cls

**配置先**: 共通モジュール（3 ブック共通）  
**種類**: クラスモジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `clsCellAddrHelper.cls`
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
