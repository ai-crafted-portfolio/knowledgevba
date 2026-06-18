---
title: clsCellAddrHelper.cls
description: clsCellAddrHelper.cls のソースコード（コピペ用）
---

# clsCellAddrHelper.cls

**配置先**: 共通モジュール（検索.xlsm / 管理.xlsm 共通）
**種類**: クラスモジュール
**更新日**: 2026-06-04 12:30 JST

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
