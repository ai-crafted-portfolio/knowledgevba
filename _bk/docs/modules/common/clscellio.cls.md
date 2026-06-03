---
title: clsCellIO.cls
description: clsCellIO.cls のソースコード（コピペ用）
---

# clsCellIO.cls

**配置先**: `共通モジュール (3 ブック全て)` 用の VBA モジュール  
**種類**: クラス モジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\`
- ファイル名: `clsCellIO.cls`
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
End Function

' Write String value to target cell
Public Sub WriteCellValue(ByVal target As Object, ByVal cellAddr As String, ByVal value As String)
    If TypeName(target) = "Dictionary" Then
        target(cellAddr) = value
    Else
        Dim r As Range
        Set r = target.Range(cellAddr)
        r.Cells(1, 1).Value = value
    End If
End Sub
```
